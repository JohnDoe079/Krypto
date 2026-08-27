"""Parser raportów użytkownika HTX w formacie .xlsx."""

import os
import pandas as pd
from pathlib import Path
from typing import Dict

from models.schemas import (
    ExtractedIdentifiers,
    clean_val,
    extract_email,
    extract_phone,
    is_wallet_address,
)
from config import HTX_SHEETS


class HTXReportParser:
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.xl = pd.ExcelFile(str(file_path), engine="openpyxl")
        self.identifiers = ExtractedIdentifiers(
            source_file=self.file_path.name,
            exchange="htx"
        )

    def parse_all(self) -> ExtractedIdentifiers:
        all_sheets = self.xl.sheet_names
        print(f"  Znaleziono {len(all_sheets)} arkuszy: {', '.join(all_sheets)}")

        for sheet_name in all_sheets:
            try:
                self._parse_sheet(sheet_name)
            except Exception as e:
                print(f"  [!] Błąd w arkuszu '{sheet_name}': {e}")

        if self.identifiers.unknown_sheets:
            print(f"  [!] Nieznane arkusze: {', '.join(self.identifiers.unknown_sheets)}")

        return self.identifiers

    def _parse_sheet(self, sheet_name: str):
        internal_name = HTX_SHEETS.get(sheet_name)

        if sheet_name == "register_1":
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=0, engine="openpyxl")
            print(f"  [ARKUSZ] '{sheet_name}' — {len(df.columns)} kolumn, {len(df)} wierszy")
            self._parse_register_1(df, sheet_name)
            self.identifiers.parsed_sheets.append(sheet_name)
        else:
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=0, engine="openpyxl")
            print(f"  [ARKUSZ] '{sheet_name}' — nieznany, parsowanie generyczne")
            self._parse_generic(df, sheet_name)
            self.identifiers.unknown_sheets.append(sheet_name)
            self.identifiers.parsed_sheets.append(sheet_name)

    def _match_certified_photos(self):
        """Dopasowuje foldery ze zdjęciami do UID. Folder może być prefiksem UID (krótszy o 1+ znaków)."""
        certified_dir = self.file_path.parent / "certified_photos"
        if not certified_dir.exists() or not certified_dir.is_dir():
            print(f"  [!] Nie znaleziono katalogu certified_photos w: {certified_dir}")
            return

        # Zbierz wszystkie foldery numeryczne
        folders = [d for d in certified_dir.iterdir() if d.is_dir() and d.name.isdigit()]
        if not folders:
            print(f"  [!] Katalog certified_photos istnieje, ale nie zawiera folderów numerycznych")
            return

        for uid in sorted(self.identifiers.user_ids):
            matched_folder = None
            # Szukaj folderu który jest prefiksem UID (najdłuższy pasujący)
            best_match = None
            best_len = 0
            for fdir in folders:
                fname = fdir.name
                if uid.startswith(fname) or fname.startswith(uid):
                    if len(fname) > best_len:
                        best_match = fdir
                        best_len = len(fname)

            if best_match:
                photos = []
                for ext in (".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"):
                    photos.extend(sorted(best_match.glob(f"*{ext}")))
                    photos.extend(sorted(best_match.glob(f"*{ext.upper()}")))
                # Usuń duplikaty zachowując kolejność
                seen = set()
                unique_photos = []
                for p in photos:
                    if p.name not in seen:
                        seen.add(p.name)
                        unique_photos.append(str(p))

                if unique_photos:
                    self.identifiers.certified_photos[uid] = unique_photos
                    print(f"  [ZDJĘCIA] UID {uid} <- folder {best_match.name}: {len(unique_photos)} zdjęć")
                else:
                    print(f"  [ZDJĘCIA] UID {uid} <- folder {best_match.name}: brak plików graficznych")
            else:
                print(f"  [ZDJĘCIA] UID {uid}: nie znaleziono pasującego folderu w certified_photos")

    def _parse_register_1(self, df: pd.DataFrame, sheet_name: str):
        """Parsuje arkusz register_1 — dane rejestracyjne / KYC HTX."""
        cols = [str(c).strip().lower().replace("-", "_") for c in df.columns]
        col_map: Dict[str, int] = {}
        for idx, c in enumerate(cols):
            if c == "uid":
                col_map["uid"] = idx
            elif c == "name":
                col_map["name"] = idx
            elif c == "phone":
                col_map["phone"] = idx
            elif c == "email":
                col_map["email"] = idx
            elif c == "idcard":
                col_map["idcard"] = idx
            elif c == "user_address":
                col_map["user_address"] = idx
            elif "created" in c:
                col_map["gmt_created"] = idx
            elif c == "country":
                col_map["country"] = idx
            elif c == "balance":
                col_map["balance"] = idx
            elif c == "bankcard":
                col_map["bankcard"] = idx
            elif c == "alipay":
                col_map["alipay"] = idx
            elif c == "wechat":
                col_map["wechat"] = idx

        print(f"  Kolumny register_1: {list(col_map.keys())}")

        # Przygotuj sekcję w customer_info_sections
        if "HTX Register" not in self.identifiers.customer_info_sections:
            self.identifiers.customer_info_sections["HTX Register"] = {}

        for _, row in df.iterrows():
            uid = clean_val(row.iloc[col_map["uid"]]) if "uid" in col_map else None
            uid_str = str(uid).strip() if uid else None

            # --- UID (właściciel konta) ---
            if uid_str:
                self.identifiers.user_ids.add(uid_str)
                if uid_str not in self.identifiers.wallet_addresses_by_user:
                    self.identifiers.wallet_addresses_by_user[uid_str] = []
                if uid_str not in self.identifiers.certified_photos:
                    self.identifiers.certified_photos[uid_str] = []

            # --- Dane osobowe ---
            name = clean_val(row.iloc[col_map["name"]]) if "name" in col_map else None
            if name:
                self.identifiers.names.add(str(name))

            email = clean_val(row.iloc[col_map["email"]]) if "email" in col_map else None
            if email:
                email_clean = extract_email(str(email))
                if email_clean:
                    self.identifiers.emails.add(email_clean)

            phone = clean_val(row.iloc[col_map["phone"]]) if "phone" in col_map else None
            if phone:
                phone_clean = extract_phone(str(phone))
                if phone_clean:
                    self.identifiers.phones.add(phone_clean)

            idcard = clean_val(row.iloc[col_map["idcard"]]) if "idcard" in col_map else None
            if idcard:
                self.identifiers.id_numbers.add(str(idcard))

            # --- Adres portfela (user_address) ---
            user_address = clean_val(row.iloc[col_map["user_address"]]) if "user_address" in col_map else None
            if user_address:
                addr = str(user_address).strip()
                if is_wallet_address(addr):
                    self.identifiers.wallet_addresses.add(addr)
                    if uid_str:
                        if addr not in self.identifiers.wallet_addresses_by_user[uid_str]:
                            self.identifiers.wallet_addresses_by_user[uid_str].append(addr)

            # --- Kraj / Geolokalizacja ---
            country = clean_val(row.iloc[col_map["country"]]) if "country" in col_map else None
            if country:
                self.identifiers.geolocations.add(str(country))

            # --- Data rejestracji ---
            gmt_created = clean_val(row.iloc[col_map["gmt_created"]]) if "gmt_created" in col_map else None
            if gmt_created:
                self.identifiers.time_ranges[sheet_name] = {
                    "from": str(gmt_created),
                    "to": str(gmt_created)
                }

            # --- Pozostałe dane płatnicze (bankcard, alipay, wechat) ---
            bankcard = clean_val(row.iloc[col_map["bankcard"]]) if "bankcard" in col_map else None
            if bankcard:
                self.identifiers.account_numbers.add(f"bankcard:{bankcard}")

            alipay = clean_val(row.iloc[col_map["alipay"]]) if "alipay" in col_map else None
            if alipay:
                self.identifiers.account_numbers.add(f"alipay:{alipay}")

            wechat = clean_val(row.iloc[col_map["wechat"]]) if "wechat" in col_map else None
            if wechat:
                self.identifiers.account_numbers.add(f"wechat:{wechat}")

            # --- Zapis pełnego wiersza do customer_info_sections ---
            row_data = {}
            for col_name, col_idx in col_map.items():
                val = clean_val(row.iloc[col_idx])
                if val is not None:
                    row_data[col_name] = str(val)

            if uid_str:
                self.identifiers.customer_info_sections["HTX Register"][uid_str] = row_data
            else:
                # Jeśli brak UID, zapisz pod indeksem wiersza
                row_key = f"row_{_}"
                self.identifiers.customer_info_sections["HTX Register"][row_key] = row_data

        # Podsumowanie
        print(f"  Sparsowano {len(df)} wierszy register_1")
        if self.identifiers.wallet_addresses_by_user:
            for uid_val, addrs in self.identifiers.wallet_addresses_by_user.items():
                print(f"  Adresy portfeli dla UID {uid_val}: {', '.join(addrs)}")

        # --- DOPASOWANIE ZDJĘĆ CERTYFIKOWANYCH (po zebraniu wszystkich UID) ---
        self._match_certified_photos()

    def _parse_generic(self, df: pd.DataFrame, sheet_name: str):
        """Generyczne parsowanie nieznanych arkuszy HTX — szuka portfeli, emaili, IP, TXID."""
        found = {"ips": 0, "wallets": 0, "txids": 0, "emails": 0, "uids": 0}
        for col in df.columns:
            for v in df[col].dropna():
                v = clean_val(v)
                if not v:
                    continue
                if is_wallet_address(v):
                    self.identifiers.wallet_addresses.add(v)
                    found["wallets"] += 1
                else:
                    email = extract_email(v)
                    if email:
                        self.identifiers.emails.add(email)
                        found["emails"] += 1
        total = sum(found.values())
        if total > 0:
            print(f"  Generycznie znaleziono: {found}")
