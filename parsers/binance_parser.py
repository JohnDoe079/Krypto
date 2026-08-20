"""Parser raportow uzytkownika Binance w formacie .xlsx."""

import pandas as pd
import re
from pathlib import Path
from typing import Dict, List, Set, Optional
from openpyxl import load_workbook

from models.schemas import (
    ExtractedIdentifiers,
    clean_val,
    is_valid_ip,
    is_wallet_address,
    is_txid,
    extract_email,
    extract_phone,
    extract_time_range,
)
from config import BINANCE_SHEETS


class BinanceReportParser:
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.xl = pd.ExcelFile(str(file_path), engine="openpyxl")
        self.identifiers = ExtractedIdentifiers(
            source_file=self.file_path.name,
            exchange="binance"
        )

    def parse_all(self) -> ExtractedIdentifiers:
        all_sheets = self.xl.sheet_names
        print(f"  Znaleziono {len(all_sheets)} arkuszy: {', '.join(all_sheets)}")

        for sheet_name in all_sheets:
            try:
                self._parse_sheet(sheet_name)
            except Exception as e:
                print(f"  [!] Blad w arkuszu '{sheet_name}': {e}")

        if self.identifiers.unknown_sheets:
            print(f"  [!] Nieznane arkusze (sparsowane generycznie): {', '.join(self.identifiers.unknown_sheets)}")

        # Dedup telefonow: Mobile ma kierunkowy (+380...), SMS bez kierunkowego
        self._dedup_phones()

        return self.identifiers

    def _dedup_phones(self):
        """Jezeli SMS jest podzbiorem Mobile (bez kierunkowego), usun duplikat."""
        phones = sorted(self.identifiers.phones)
        to_remove = set()
        for p1 in phones:
            for p2 in phones:
                if p1 == p2:
                    continue
                p1_clean = p1.lstrip("+").replace(" ", "").replace("-", "")
                p2_clean = p2.lstrip("+").replace(" ", "").replace("-", "")
                if len(p1_clean) > len(p2_clean) and p1_clean.endswith(p2_clean):
                    to_remove.add(p2)
                elif len(p2_clean) > len(p1_clean) and p2_clean.endswith(p1_clean):
                    to_remove.add(p1)
        self.identifiers.phones -= to_remove

    def _add_time_range(self, df: pd.DataFrame, sheet_name: str):
        """Zapisuje zakres czasowy arkusza jezeli znajdzie kolumny czasowe."""
        tr = extract_time_range(df)
        if tr:
            self.identifiers.time_ranges[sheet_name] = tr
            print(f"    Zakres czasowy '{sheet_name}': {tr['from']} -> {tr['to']}")

    def _parse_sheet(self, sheet_name: str):
        if sheet_name == "Customer Information":
            df = pd.read_excel(
                self.file_path, sheet_name=sheet_name, header=None, engine="openpyxl"
            )
            self._parse_customer_info_raw(df)
            self._add_time_range(df, sheet_name)
            self.identifiers.parsed_sheets.append(sheet_name)
        elif sheet_name == "KYC Documents":
            df = pd.read_excel(
                self.file_path, sheet_name=sheet_name, header=None, engine="openpyxl"
            )
            self._parse_kyc_documents(df)
            self._add_time_range(df, sheet_name)
            self.identifiers.parsed_sheets.append(sheet_name)
        else:
            df = pd.read_excel(
                self.file_path, sheet_name=sheet_name, header=0, engine="openpyxl"
            )
            internal_name = BINANCE_SHEETS.get(sheet_name)
            if internal_name:
                method_name = f"_parse_{internal_name}"
                method = getattr(self, method_name, self._parse_generic)
                method(df, sheet_name)
                self._add_time_range(df, sheet_name)
                self.identifiers.parsed_sheets.append(sheet_name)
            else:
                print(f"  [!] Nieznany arkusz: '{sheet_name}' — parsowanie generyczne")
                self._parse_generic(df, sheet_name)
                self._add_time_range(df, sheet_name)
                self.identifiers.unknown_sheets.append(sheet_name)
                self.identifiers.parsed_sheets.append(sheet_name)

    def _parse_customer_info_raw(self, df: pd.DataFrame):
        sections = {}
        current_section = None
        headers = None

        i = 0
        while i < df.shape[0]:
            row = df.iloc[i]
            non_null = row.dropna()

            if len(non_null) == 1 and i > 2:
                val = str(non_null.iloc[0]).strip()
                if val and val not in ("nan", "None"):
                    current_section = val
                    sections[current_section] = {}
                    headers = None
                i += 1
                continue

            if current_section and len(non_null) > 1 and headers is None:
                headers = [clean_val(v) for v in row.values]
                i += 1
                continue

            if current_section and headers is not None and len(non_null) >= 1:
                values = [clean_val(v) for v in row.values]
                for h, v in zip(headers, values):
                    if h is not None:
                        sections[current_section][h] = v if v is not None else ""
                headers = None
                i += 1
                continue

            i += 1

        self.identifiers.customer_info_sections = sections

        for section_name, data in sections.items():
            for col, val in data.items():
                if val is None or val == "":
                    continue
                v = str(val).strip()

                email = extract_email(v)
                if email:
                    self.identifiers.emails.add(email)
                    continue

                phone = extract_phone(v)
                if phone:
                    self.identifiers.phones.add(phone)
                    continue

                if is_valid_ip(v):
                    self.identifiers.ips.add(v)
                    continue

                # User ID z Customer Information = wlasciciel konta
                if re.match(r"^\d{9,10}$", v):
                    self.identifiers.user_ids.add(v)
                    continue

                if re.match(r"^[A-Z]{1,2}\d{6,10}$", v):
                    self.identifiers.id_numbers.add(v)
                    continue

                if col and "nationality" in str(col).lower() and len(v) == 2 and v.isalpha():
                    self.identifiers.nationalities.add(v)
                    continue

    def _parse_kyc_documents(self, df: pd.DataFrame):
        texts = []
        for col in df.columns:
            for val in df[col].dropna():
                v = clean_val(val)
                if v:
                    texts.append(v)
        try:
            wb = load_workbook(self.file_path)
            ws = wb["KYC Documents"]
            images = ws._images
            if images:
                img_dir = Path("output/images")
                img_dir.mkdir(parents=True, exist_ok=True)
                for idx, img in enumerate(images):
                    ext = img.format.lower() if img.format else "png"
                    img_path = img_dir / f"{self.file_path.stem}_kyc_{idx+1}.{ext}"
                    with open(img_path, "wb") as fimg:
                        fimg.write(img._data())
                    self.identifiers.kyc_images.append(str(img_path))
                    print(f"    Zapisano obrazek KYC: {img_path.name}")
            else:
                print(f"    Brak osadzonych obrazkow w KYC Documents (teksty: {texts})")
        except Exception as e:
            print(f"  [!] Blad przy wyciaganiu obrazkow KYC: {e}")

    def _parse_access_logs(self, df: pd.DataFrame, sheet_name: str):
        for col, target in [
            ("Real IP", self.identifiers.ips),
            ("Geolocation", self.identifiers.geolocations),
            ("Browser", self.identifiers.browsers),
        ]:
            if col in df.columns:
                for v in df[col].dropna():
                    v = clean_val(v)
                    if v:
                        if col == "Real IP" and is_valid_ip(v):
                            target.add(v)
                        elif col != "Real IP":
                            target.add(v)
        # User ID z Access Logs = powiazany (nie wlasciciel)
        if "User ID" in df.columns:
            for v in df["User ID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.related_user_ids.add(str(v))

    def _parse_approved_devices(self, df: pd.DataFrame, sheet_name: str):
        if "IP Address" in df.columns:
            for v in df["IP Address"].dropna():
                v = clean_val(v)
                if v and is_valid_ip(v):
                    self.identifiers.ips.add(v)
        if "Geolocation" in df.columns:
            for v in df["Geolocation"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.geolocations.add(v)
        if "Key" in df.columns and "Value" in df.columns:
            for _, row in df.iterrows():
                key = clean_val(row.get("Key"))
                val = clean_val(row.get("Value"))
                if not key or not val:
                    continue
                key_lower = key.lower()
                if "fvideo_id" in key_lower:
                    self.identifiers.fvideo_ids.add(val)
                elif "bnc-uuid" in key_lower or "bnc_uuid" in key_lower:
                    self.identifiers.bnc_uuids.add(val)
                elif "device_id" in key_lower:
                    self.identifiers.device_ids.add(val)
                elif "login_ip" in key_lower and is_valid_ip(val):
                    self.identifiers.ips.add(val)

    def _parse_deposit_history(self, df: pd.DataFrame, sheet_name: str):
        for col in ["Deposit Address", "Source Address"]:
            if col in df.columns:
                for v in df[col].dropna():
                    v = clean_val(v)
                    if v and is_wallet_address(v):
                        self.identifiers.wallet_addresses.add(v)
        if "TXID" in df.columns:
            for v in df["TXID"].dropna():
                v = clean_val(v)
                if v and is_txid(v):
                    self.identifiers.txids.add(v)
        # User ID z Deposit History = powiazany
        if "User ID" in df.columns:
            for v in df["User ID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.related_user_ids.add(str(v))

    def _parse_withdrawal_history(self, df: pd.DataFrame, sheet_name: str):
        if "Destination Address" in df.columns:
            for v in df["Destination Address"].dropna():
                v = clean_val(v)
                if v and is_wallet_address(v):
                    self.identifiers.wallet_addresses.add(v)
        if "txId" in df.columns:
            for v in df["txId"].dropna():
                v = clean_val(v)
                if v and is_txid(v):
                    self.identifiers.txids.add(v)
        # User ID = powiazany
        if "User ID" in df.columns:
            for v in df["User ID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.related_user_ids.add(str(v))

    def _parse_attempted_withdrawal(self, df: pd.DataFrame, sheet_name: str):
        if "Address" in df.columns:
            for v in df["Address"].dropna():
                v = clean_val(v)
                if v and is_wallet_address(v):
                    self.identifiers.wallet_addresses.add(v)
        # User ID = powiazany
        if "User ID" in df.columns:
            for v in df["User ID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.related_user_ids.add(str(v))

    def _parse_fiat_deposit(self, df: pd.DataFrame, sheet_name: str):
        for col, target in [
            ("Card Bin", self.identifiers.card_bins),
            ("Card Last 4 Digital", self.identifiers.card_last4),
            ("Iban", self.identifiers.ibans),
            ("Account Number", self.identifiers.account_numbers),
        ]:
            if col in df.columns:
                for v in df[col].dropna():
                    v = clean_val(v)
                    if v:
                        target.add(v)
        if "User Email" in df.columns:
            for v in df["User Email"].dropna():
                email = extract_email(v)
                if email:
                    self.identifiers.emails.add(email)
        # User Id = powiazany
        if "User Id" in df.columns:
            for v in df["User Id"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.related_user_ids.add(str(v))

    def _parse_fiat_trades(self, df: pd.DataFrame, sheet_name: str):
        if "User Email" in df.columns:
            for v in df["User Email"].dropna():
                email = extract_email(v)
                if email:
                    self.identifiers.emails.add(email)
        if "Order Id" in df.columns:
            for v in df["Order Id"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.order_ids.add(v)
        # User Id = powiazany
        if "User Id" in df.columns:
            for v in df["User Id"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.related_user_ids.add(str(v))

    def _parse_binance_pay(self, df: pd.DataFrame, sheet_name: str):
        if "Counterparty Wallet ID" in df.columns:
            for v in df["Counterparty Wallet ID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.counterparty_ids.add(v)
        if "Counterparty Binance ID" in df.columns:
            for v in df["Counterparty Binance ID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.counterparty_ids.add(v)
        if "Transaction ID" in df.columns:
            for v in df["Transaction ID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.transaction_ids.add(v)
        # User ID = powiazany
        if "User ID" in df.columns:
            for v in df["User ID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.related_user_ids.add(str(v))

    def _parse_p2p(self, df: pd.DataFrame, sheet_name: str):
        if "Order ID" in df.columns:
            for v in df["Order ID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.order_ids.add(v)
        # Target UID = powiazany uzytkownik
        if "Target UID" in df.columns:
            for v in df["Target UID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.related_user_ids.add(str(v))

    def _parse_otc_trading(self, df: pd.DataFrame, sheet_name: str):
        if "OrderId" in df.columns:
            for v in df["OrderId"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.order_ids.add(v)
        # UserId = powiazany
        if "UserId" in df.columns:
            for v in df["UserId"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.related_user_ids.add(str(v))

    def _parse_spot_asset_log(self, df: pd.DataFrame, sheet_name: str):
        self._parse_asset_log(df)

    def _parse_funding_asset_log(self, df: pd.DataFrame, sheet_name: str):
        self._parse_asset_log(df)

    def _parse_asset_log(self, df: pd.DataFrame):
        if "Transaction ID" in df.columns:
            for v in df["Transaction ID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.transaction_ids.add(v)
        # User ID = powiazany
        if "User ID" in df.columns:
            for v in df["User ID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.related_user_ids.add(str(v))

    def _parse_order_history(self, df: pd.DataFrame, sheet_name: str):
        if "Order ID" in df.columns:
            for v in df["Order ID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.order_ids.add(v)
        # User ID = powiazany
        if "User ID" in df.columns:
            for v in df["User ID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.related_user_ids.add(str(v))

    def _parse_generic(self, df: pd.DataFrame, sheet_name: str):
        found = {"ips": 0, "wallets": 0, "txids": 0, "emails": 0}
        for col in df.columns:
            for v in df[col].dropna():
                v = clean_val(v)
                if not v:
                    continue
                if is_valid_ip(v):
                    self.identifiers.ips.add(v)
                    found["ips"] += 1
                elif is_wallet_address(v):
                    self.identifiers.wallet_addresses.add(v)
                    found["wallets"] += 1
                elif is_txid(v):
                    self.identifiers.txids.add(v)
                    found["txids"] += 1
                else:
                    email = extract_email(v)
                    if email:
                        self.identifiers.emails.add(email)
                        found["emails"] += 1
        total = sum(found.values())
        if total > 0:
            print(f"    Generycznie znaleziono: {found}")
