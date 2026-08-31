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
    is_valid_ip,
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
        self._htx_balances = {}  # uid -> [{currency, balance}, ...]

    def _clean_htx_val(self, val):
        """Czyści wartość z HTX Excela — usuwa .0 z floatów i notację naukową."""
        v = clean_val(val)
        if v is None:
            return None
        s = str(v).strip()
        if s.endswith(".0"):
            s = s[:-2]
        try:
            f = float(s)
            if 'e' in s.lower() or 'E' in s:
                s = f"{f:.12f}".rstrip('0').rstrip('.')
                if s.startswith('.'):
                    s = '0' + s
                if s.startswith('-.'):
                    s = '-0' + s[1:]
        except ValueError:
            pass
        return s if s else None

    def parse_all(self) -> ExtractedIdentifiers:
        all_sheets = self.xl.sheet_names
        print(f"  Znaleziono {len(all_sheets)} arkuszy:")
        for i, sn in enumerate(all_sheets, 1):
            print(f"    {i}. [{repr(sn)}] (len={len(sn)})")
        print(f"  Arkusze: {', '.join(all_sheets)}")

        for sheet_name in all_sheets:
            try:
                self._parse_sheet(sheet_name)
            except Exception as e:
                print(f"  [!] Błąd w arkuszu '{sheet_name}': {e}")

        if self.identifiers.unknown_sheets:
            print(f"  [!] Nieznane arkusze (pominięte): {', '.join(self.identifiers.unknown_sheets)}")

        # Dołącz salda z balance_1 do profilu użytkownika
        self._merge_balances_into_register()

        return self.identifiers

    def _parse_sheet(self, sheet_name: str):
        if sheet_name == "register_1":
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=0, engine="openpyxl")
            print(f"  [ARKUSZ] '{sheet_name}' — {len(df.columns)} kolumn, {len(df)} wierszy")
            self._parse_register_1(df, sheet_name)
            self.identifiers.parsed_sheets.append(sheet_name)
        elif sheet_name == "balance_1":
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=0, engine="openpyxl")
            print(f"  [ARKUSZ] '{sheet_name}' — {len(df.columns)} kolumn, {len(df)} wierszy")
            self._parse_balance_1(df, sheet_name)
            self.identifiers.parsed_sheets.append(sheet_name)
        elif sheet_name == "login_1":
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=0, engine="openpyxl")
            print(f"  [ARKUSZ] '{sheet_name}' — {len(df.columns)} kolumn, {len(df)} wierszy")
            self._parse_login_1(df, sheet_name)
            self.identifiers.parsed_sheets.append(sheet_name)
        else:
            # Pozostałe arkusze (trade_*, DeviceFP_1, deposit&withdraw...)
            # Tylko generyczne skanowanie pierwszych 100 wierszy w poszukiwaniu portfeli/ID
            # NIE ładujemy wszystkich wierszy do pamięci
            print(f"  [ARKUSZ] '{sheet_name}' — pominięty (tylko register_1, balance_1, login_1 są parsowane)")
            self.identifiers.unknown_sheets.append(sheet_name)

    def _detect_columns(self, df: pd.DataFrame, mapping: dict) -> dict:
        """Wykrywa indeksy kolumn na podstawie słownika {nazwa_wewnętrzna: [możliwe_nazwy]}."""
        result = {}
        cols_lower = [str(c).strip().lower().replace("-", "_").replace(" ", "_") for c in df.columns]
        for internal_name, possible_names in mapping.items():
            for idx, col_name in enumerate(cols_lower):
                for possible in possible_names:
                    if possible in col_name:
                        result[internal_name] = idx
                        break
                if internal_name in result:
                    break
        return result

    def _parse_balance_1(self, df: pd.DataFrame, sheet_name: str):
        """Parsuje arkusz balance_1 — salda per waluta per UID."""
        print(f"  [DEBUG balance_1] Kolumny w arkuszu: {list(df.columns)}")
        col_map = self._detect_columns(df, {
            "uid": ["uid", "user_id", "userid"],
            "currency": ["currency", "coin", "asset", "symbol"],
            "balance": ["balance", "total", "all", "amount"],
        })
        print(f"  [DEBUG balance_1] Wykryte kolumny: {list(col_map.keys())}")
        if not col_map:
            print(f"  [DEBUG balance_1] NIE wykryto kolumn — pominięto!")
            return

        for _, row in df.iterrows():
            uid = self._clean_htx_val(row.iloc[col_map["uid"]]) if "uid" in col_map else None
            curr = self._clean_htx_val(row.iloc[col_map["currency"]]) if "currency" in col_map else None
            bal = self._clean_htx_val(row.iloc[col_map["balance"]]) if "balance" in col_map else None
            if uid and curr and bal:
                if uid not in self._htx_balances:
                    self._htx_balances[uid] = []
                self._htx_balances[uid].append({"currency": str(curr), "balance": str(bal)})
        print(f"  Sparsowano balance_1: {len(df)} wierszy")

    def _parse_login_1(self, df: pd.DataFrame, sheet_name: str):
        """Parsuje arkusz login_1 — logowania per UID.
        Kolumny: uid, login_time, login_terminal, ip
        """
        col_map = self._detect_columns(df, {
            "uid": ["uid", "user_id", "userid"],
            "login_time": ["login_time", "time", "date", "timestamp"],
            "login_terminal": ["login_terminal", "terminal", "device", "platform"],
            "ip": ["ip", "ip_address", "address"],
        })
        print(f"  [DEBUG login_1] Kolumny w arkuszu: {list(df.columns)}")
        print(f"  [DEBUG login_1] Wykryte kolumny: {list(col_map.keys())}")

        if not col_map:
            print(f"  [DEBUG login_1] NIE wykryto kolumn — pominięto!")
            return

        for _, row in df.iterrows():
            uid = self._clean_htx_val(row.iloc[col_map["uid"]]) if "uid" in col_map else None
            login_time = self._clean_htx_val(row.iloc[col_map["login_time"]]) if "login_time" in col_map else None
            terminal = self._clean_htx_val(row.iloc[col_map["login_terminal"]]) if "login_terminal" in col_map else None
            ip = self._clean_htx_val(row.iloc[col_map["ip"]]) if "ip" in col_map else None

            if uid:
                if uid not in self.identifiers.login_records:
                    self.identifiers.login_records[uid] = []
                record = {}
                if login_time:
                    record["time"] = str(login_time)
                    # Dodaj do time_ranges per arkusz
                    if sheet_name not in self.identifiers.time_ranges:
                        self.identifiers.time_ranges[sheet_name] = {"from": str(login_time), "to": str(login_time)}
                    else:
                        # Aktualizuj zakres
                        current_from = self.identifiers.time_ranges[sheet_name]["from"]
                        current_to = self.identifiers.time_ranges[sheet_name]["to"]
                        if str(login_time) < current_from:
                            self.identifiers.time_ranges[sheet_name]["from"] = str(login_time)
                        if str(login_time) > current_to:
                            self.identifiers.time_ranges[sheet_name]["to"] = str(login_time)
                if terminal:
                    record["terminal"] = str(terminal)
                if ip:
                    record["ip"] = str(ip)
                    if is_valid_ip(str(ip)):
                        self.identifiers.ips.add(str(ip))

                if record:
                    self.identifiers.login_records[uid].append(record)

        total_records = sum(len(v) for v in self.identifiers.login_records.values())
        print(f"  Sparsowano login_1: {len(df)} wierszy, {total_records} rekordów logowania")

    def _merge_balances_into_register(self):
        """Dołącza salda z balance_1 do profilu użytkownika w customer_info_sections.
        Salda są przechowywane jako lista {currency, balance} pod kluczem 'balances_list',
        a reporter sam je rozbije na osobne wiersze.
        """
        if not self._htx_balances:
            return
        if "HTX Register" not in self.identifiers.customer_info_sections:
            self.identifiers.customer_info_sections["HTX Register"] = {}

        for uid, balances in self._htx_balances.items():
            # Zbuduj czytelny string z sald (dla podsumowania w sekcji 4)
            bal_items = []
            for b in balances:
                bal_items.append(f"{b['currency']}: {b['balance']}")
            bal_str = " | ".join(bal_items)

            if uid in self.identifiers.customer_info_sections["HTX Register"]:
                # Nie dodawaj balances_detail (surowy string) — tylko lista
                self.identifiers.customer_info_sections["HTX Register"][uid]["balances_list"] = bal_str
            else:
                # Jeśli UID z balance_1 nie ma w register_1, utwórz minimalny wpis
                self.identifiers.customer_info_sections["HTX Register"][uid] = {
                    "uid": uid,
                    "balances_list": bal_str,
                }
            print(f"  [SALDA] UID {uid}: {len(balances)} walut")

    def _match_certified_photos(self):
        """Dopasowuje foldery ze zdjęciami do UID. Folder może być prefiksem UID (krótszy o 1+ znaków)."""
        certified_dir = self.file_path.parent / "certified_photos"
        if not certified_dir.exists() or not certified_dir.is_dir():
            print(f"  [!] Nie znaleziono katalogu certified_photos w: {certified_dir}")
            return

        folders = [d for d in certified_dir.iterdir() if d.is_dir() and d.name.isdigit()]
        if not folders:
            print(f"  [!] Katalog certified_photos istnieje, ale nie zawiera folderów numerycznych")
            return

        for uid in sorted(self.identifiers.user_ids):
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
            if c == "uid": col_map["uid"] = idx
            elif c == "name": col_map["name"] = idx
            elif c == "phone": col_map["phone"] = idx
            elif c == "email": col_map["email"] = idx
            elif c == "idcard": col_map["idcard"] = idx
            elif c == "user_address": col_map["user_address"] = idx
            elif "created" in c: col_map["gmt_created"] = idx
            elif c == "country": col_map["country"] = idx
            elif c == "balance": col_map["balance"] = idx
            elif c == "bankcard": col_map["bankcard"] = idx
            elif c == "alipay": col_map["alipay"] = idx
            elif c == "wechat": col_map["wechat"] = idx

        print(f"  Kolumny register_1: {list(col_map.keys())}")

        if "HTX Register" not in self.identifiers.customer_info_sections:
            self.identifiers.customer_info_sections["HTX Register"] = {}

        for _, row in df.iterrows():
            uid = self._clean_htx_val(row.iloc[col_map["uid"]]) if "uid" in col_map else None
            uid_str = str(uid).strip() if uid else None

            if uid_str:
                self.identifiers.user_ids.add(uid_str)
                if uid_str not in self.identifiers.wallet_addresses_by_user:
                    self.identifiers.wallet_addresses_by_user[uid_str] = []
                if uid_str not in self.identifiers.certified_photos:
                    self.identifiers.certified_photos[uid_str] = []

            name = self._clean_htx_val(row.iloc[col_map["name"]]) if "name" in col_map else None
            if name: self.identifiers.names.add(str(name))

            email = self._clean_htx_val(row.iloc[col_map["email"]]) if "email" in col_map else None
            if email:
                email_clean = extract_email(str(email))
                if email_clean: self.identifiers.emails.add(email_clean)

            phone = self._clean_htx_val(row.iloc[col_map["phone"]]) if "phone" in col_map else None
            if phone:
                phone_clean = extract_phone(str(phone))
                if phone_clean: self.identifiers.phones.add(phone_clean)

            idcard = self._clean_htx_val(row.iloc[col_map["idcard"]]) if "idcard" in col_map else None
            if idcard: self.identifiers.id_numbers.add(str(idcard))

            user_address = self._clean_htx_val(row.iloc[col_map["user_address"]]) if "user_address" in col_map else None
            if user_address:
                addr = str(user_address).strip()
                if is_wallet_address(addr):
                    self.identifiers.wallet_addresses.add(addr)
                    if uid_str:
                        if addr not in self.identifiers.wallet_addresses_by_user[uid_str]:
                            self.identifiers.wallet_addresses_by_user[uid_str].append(addr)

            country = self._clean_htx_val(row.iloc[col_map["country"]]) if "country" in col_map else None
            if country: self.identifiers.geolocations.add(str(country))

            gmt_created = self._clean_htx_val(row.iloc[col_map["gmt_created"]]) if "gmt_created" in col_map else None
            if gmt_created:
                self.identifiers.time_ranges[sheet_name] = {
                    "from": str(gmt_created), "to": str(gmt_created)
                }

            bankcard = self._clean_htx_val(row.iloc[col_map["bankcard"]]) if "bankcard" in col_map else None
            if bankcard: self.identifiers.account_numbers.add(f"bankcard:{bankcard}")

            alipay = self._clean_htx_val(row.iloc[col_map["alipay"]]) if "alipay" in col_map else None
            if alipay: self.identifiers.account_numbers.add(f"alipay:{alipay}")

            wechat = self._clean_htx_val(row.iloc[col_map["wechat"]]) if "wechat" in col_map else None
            if wechat: self.identifiers.account_numbers.add(f"wechat:{wechat}")

            row_data = {}
            for col_name, col_idx in col_map.items():
                val = self._clean_htx_val(row.iloc[col_idx])
                if val is not None:
                    row_data[col_name] = str(val)

            if uid_str:
                self.identifiers.customer_info_sections["HTX Register"][uid_str] = row_data
            else:
                row_key = f"row_{_}"
                self.identifiers.customer_info_sections["HTX Register"][row_key] = row_data

        print(f"  Sparsowano {len(df)} wierszy register_1")
        if self.identifiers.wallet_addresses_by_user:
            for uid_val, addrs in self.identifiers.wallet_addresses_by_user.items():
                print(f"  Adresy portfeli dla UID {uid_val}: {', '.join(addrs)}")

        self._match_certified_photos()
