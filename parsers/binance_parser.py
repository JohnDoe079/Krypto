"""Parser raportów użytkownika Binance w formacie .xlsx."""

import pandas as pd
import re
from pathlib import Path
from typing import Dict, List, Set, Optional
from openpyxl import load_workbook

from models.schemas import (
    ExtractedIdentifiers,
    AssetBalance,
    AssetTransaction,
    clean_val,
    is_valid_ip,
    is_wallet_address,
    is_txid,
    extract_email,
    extract_phone,
    extract_time_range,
)
from config import BINANCE_SHEETS


def _fmt_num(val) -> str:
    if val is None or val == "":
        return ""
    try:
        if hasattr(val, "to_eng_string"):
            f = float(val)
        else:
            f = float(val)
        if abs(f) < 1e-12:
            return "0"
        if f == int(f):
            return str(int(f))
        s = f"{f:.12f}".rstrip("0").rstrip(".")
        return s
    except (ValueError, TypeError, OverflowError):
        return str(val).strip()


def _is_zero(val) -> bool:
    if val is None or val == "":
        return True
    try:
        if hasattr(val, "to_eng_string"):
            return val == 0
        return abs(float(val)) < 1e-12
    except (ValueError, TypeError, OverflowError):
        return False


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
                print(f"  [!] Błąd w arkuszu '{sheet_name}': {e}")

        if self.identifiers.unknown_sheets:
            print(f"  [!] Nieznane arkusze (sparsowane generycznie): {', '.join(self.identifiers.unknown_sheets)}")

        self._dedup_phones()
        return self.identifiers

    def _dedup_phones(self):
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
        tr = extract_time_range(df)
        if tr:
            self.identifiers.time_ranges[sheet_name] = tr
            print(f"  Zakres czasowy '{sheet_name}': {tr['from']} -> {tr['to']}")

    def _parse_sheet(self, sheet_name: str):
        if sheet_name == "Customer Information":
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=None, engine="openpyxl")
            self._parse_customer_info_raw(df)
            self._add_time_range(df, sheet_name)
            self.identifiers.parsed_sheets.append(sheet_name)
        elif sheet_name == "KYC Documents":
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=None, engine="openpyxl")
            self._parse_kyc_documents(df)
            self._add_time_range(df, sheet_name)
            self.identifiers.parsed_sheets.append(sheet_name)
        elif sheet_name == "Assets Overview":
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=None, engine="openpyxl")
            self._parse_assets_overview(df)
            self.identifiers.parsed_sheets.append(sheet_name)
        else:
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=0, engine="openpyxl")
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

    def _parse_assets_overview(self, df: pd.DataFrame):
        for i in range(min(15, len(df))):
            for j in range(len(df.columns)):
                val = clean_val(df.iloc[i, j])
                if val and "estimate total balance" in str(val).lower():
                    for di, dj in [(0, 1), (1, 0), (1, 1)]:
                        ni, nj = i + di, j + dj
                        if ni < len(df) and nj < len(df.columns):
                            btc_val = clean_val(df.iloc[ni, nj])
                            if btc_val and not _is_zero(btc_val):
                                btc_str = _fmt_num(btc_val).replace("\n", " ").replace("\r", " ").strip()
                                self.identifiers.estimate_total_btc = btc_str
                                print(f"  Estimate Total Balance(BTC): {self.identifiers.estimate_total_btc}")
                                break

        current_wallet = "Spot"
        i = 0
        while i < len(df):
            row = df.iloc[i]
            row_vals = [str(v).strip().lower() if pd.notna(v) else "" for v in row.values]
            row_str = " ".join(row_vals)

            if "spot" in row_str:
                current_wallet = "Spot"
            elif "futures" in row_str:
                current_wallet = "Futures"
            elif "earn" in row_str:
                current_wallet = "Earn"
            elif "margin" in row_str:
                current_wallet = "Margin"
            elif "pool" in row_str:
                current_wallet = "Pool"
            elif "funding" in row_str:
                current_wallet = "Funding"

            if "currency name" in row_str or "currency code" in row_str or "all positions" in row_str:
                headers = [clean_val(v) for v in row.values]
                col_map = {}
                for idx_h, h in enumerate(headers):
                    if h:
                        h_lower = str(h).lower().strip()
                        if "currency name" in h_lower:
                            col_map["currency_name"] = idx_h
                        elif "currency code" in h_lower or (h_lower == "code" and "currency" not in row_str):
                            col_map["currency_code"] = idx_h
                        elif "all positions" in h_lower or h_lower == "all":
                            col_map["all_positions"] = idx_h
                        elif "available" in h_lower:
                            col_map["available_positions"] = idx_h
                        elif "withdrawal" in h_lower or "in withdrawal" in h_lower:
                            col_map["in_withdrawal"] = idx_h
                        elif "pending" in h_lower:
                            col_map["pending_order"] = idx_h
                        elif "btc equivalent" in h_lower or ("btc" in h_lower and "equivalent" in h_lower):
                            col_map["btc_equivalent"] = idx_h
                        elif "usdt equivalent" in h_lower or ("usdt" in h_lower and "equivalent" in h_lower):
                            col_map["usdt_equivalent"] = idx_h

                if "currency_code" in col_map or "currency_name" in col_map:
                    i += 1
                    while i < len(df):
                        row_data = df.iloc[i]
                        code_idx = col_map.get("currency_code", col_map.get("currency_name", 0))
                        code_val = clean_val(row_data.iloc[code_idx])
                        if not code_val:
                            break

                        all_pos = clean_val(row_data.iloc[col_map.get("all_positions", 0)]) if "all_positions" in col_map else ""
                        if _is_zero(all_pos):
                            i += 1
                            continue

                        bal = AssetBalance(
                            currency_name=str(clean_val(row_data.iloc[col_map.get("currency_name", 0)])) if "currency_name" in col_map else "",
                            currency_code=str(code_val),
                            all_positions=_fmt_num(all_pos),
                            available_positions=_fmt_num(clean_val(row_data.iloc[col_map.get("available_positions", 0)])) if "available_positions" in col_map else "",
                            in_withdrawal=_fmt_num(clean_val(row_data.iloc[col_map.get("in_withdrawal", 0)])) if "in_withdrawal" in col_map else "",
                            pending_order=_fmt_num(clean_val(row_data.iloc[col_map.get("pending_order", 0)])) if "pending_order" in col_map else "",
                            btc_equivalent=_fmt_num(clean_val(row_data.iloc[col_map.get("btc_equivalent", 0)])) if "btc_equivalent" in col_map else "",
                            usdt_equivalent=_fmt_num(clean_val(row_data.iloc[col_map.get("usdt_equivalent", 0)])) if "usdt_equivalent" in col_map else "",
                            wallet_type=current_wallet,
                        )
                        self.identifiers.asset_balances.append(bal)
                        i += 1
                        continue
            i += 1

        wallet_order = {"Spot": 0, "Funding": 1, "Futures": 2, "Earn": 3, "Margin": 4, "Pool": 5}
        self.identifiers.asset_balances.sort(key=lambda b: (wallet_order.get(b.wallet_type, 99), b.currency_code))
        print(f"  Sparsowano {len(self.identifiers.asset_balances)} sald walut (pominięto zera)")

    def _parse_spot_asset_log(self, df: pd.DataFrame, sheet_name: str):
        self._parse_asset_log(df, "Spot")

    def _parse_funding_asset_log(self, df: pd.DataFrame, sheet_name: str):
        self._parse_asset_log(df, "Funding")

    def _parse_asset_log(self, df: pd.DataFrame, wallet_type: str):
        cols = [str(c).strip().lower() for c in df.columns]
        col_map = {}
        for idx, c in enumerate(cols):
            if "time" in c or "date" in c:
                col_map["time"] = idx
            elif c == "currency" or c == "asset" or "currency" in c:
                col_map["currency"] = idx
            elif "change" in c:
                col_map["change"] = idx
            elif "amount" in c:
                col_map["amount"] = idx
            elif "locked" in c:
                col_map["locked"] = idx
            elif "freeze" in c:
                col_map["freeze"] = idx
            elif "processing" in c:
                col_map["processing"] = idx
            elif "reason" in c or "type" in c or "operation" in c:
                col_map["reason"] = idx
            elif "transaction id" in c or "txid" in c or "tx id" in c:
                col_map["transaction_id"] = idx

        print(f"  Kolumny Asset Log ({wallet_type}): {list(col_map.keys())}")

        for _, row in df.iterrows():
            chg_val = clean_val(row.iloc[col_map.get("change", 0)]) if "change" in col_map else None
            if chg_val is None:
                amt_val = clean_val(row.iloc[col_map.get("amount", 0)]) if "amount" in col_map else None
                if amt_val:
                    reason = str(clean_val(row.iloc[col_map.get("reason", 0)])).lower() if "reason" in col_map else ""
                    if any(kw in reason for kw in ["withdrawal", "send", "out", "fee", "sell"]):
                        chg_val = f"-{amt_val}"
                    elif any(kw in reason for kw in ["deposit", "receive", "in", "buy", "reward", "staking", "airdrop"]):
                        chg_val = f"+{amt_val}"
                    else:
                        chg_val = amt_val

            txn = AssetTransaction(
                time=str(clean_val(row.iloc[col_map.get("time", 0)])) if "time" in col_map else "",
                currency=str(clean_val(row.iloc[col_map.get("currency", 0)])) if "currency" in col_map else "",
                amount=_fmt_num(clean_val(row.iloc[col_map.get("amount", 0)])) if "amount" in col_map else "",
                locked=_fmt_num(clean_val(row.iloc[col_map.get("locked", 0)])) if "locked" in col_map else "",
                freeze=_fmt_num(clean_val(row.iloc[col_map.get("freeze", 0)])) if "freeze" in col_map else "",
                processing=_fmt_num(clean_val(row.iloc[col_map.get("processing", 0)])) if "processing" in col_map else "",
                change=_fmt_num(chg_val) if chg_val else "",
                reason=str(clean_val(row.iloc[col_map.get("reason", 0)])) if "reason" in col_map else "",
                transaction_id=str(clean_val(row.iloc[col_map.get("transaction_id", 0)])) if "transaction_id" in col_map else "",
                wallet_type=wallet_type,
            )
            if txn.time or txn.currency:
                if wallet_type == "Spot":
                    self.identifiers.spot_transactions.append(txn)
                else:
                    self.identifiers.funding_transactions.append(txn)

        count = len(self.identifiers.spot_transactions) if wallet_type == "Spot" else len(self.identifiers.funding_transactions)
        print(f"  Sparsowano {count} transakcji {wallet_type}")

        if "Transaction ID" in df.columns:
            for v in df["Transaction ID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.transaction_ids.add(v)
        if "User ID" in df.columns:
            for v in df["User ID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.related_user_ids.add(str(v))

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
                    print(f"  Zapisano obrazek KYC: {img_path.name}")
            else:
                print(f"  Brak osadzonych obrazków w KYC Documents (teksty: {texts})")
        except Exception as e:
            print(f"  [!] Błąd przy wyciąganiu obrazków KYC: {e}")

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
        cols = [str(c).strip().lower() for c in df.columns]
        col_map = {}
        for idx, c in enumerate(cols):
            if "time" in c or "date" in c or "create" in c:
                col_map["time"] = idx
            elif "currency" in c:
                col_map["currency"] = idx
            elif "amount" in c:
                col_map["amount"] = idx
            elif "usdt" in c and "amount" not in c:
                col_map["usdt_value"] = idx
            elif "status" in c:
                col_map["status"] = idx
            elif "txid" in c or "transaction id" in c:
                col_map["txid"] = idx
            elif "deposit address" in c:
                col_map["deposit_address"] = idx
            elif "source address" in c:
                col_map["source_address"] = idx
            elif "counterparty" in c or "counter party" in c:
                col_map["counterparty"] = idx
            elif "user id" in c:
                col_map["user_id"] = idx

        print(f"  Kolumny Deposit History: {list(col_map.keys())}")

        for _, row in df.iterrows():
            curr = str(clean_val(row.iloc[col_map.get("currency", 0)])) if "currency" in col_map else ""
            amt = clean_val(row.iloc[col_map.get("amount", 0)]) if "amount" in col_map else None
            status = str(clean_val(row.iloc[col_map.get("status", 0)])).lower() if "status" in col_map else ""
            time_str = str(clean_val(row.iloc[col_map.get("time", 0)])) if "time" in col_map else ""
            txid = str(clean_val(row.iloc[col_map.get("txid", 0)])) if "txid" in col_map else ""

            if curr and amt:
                txn = AssetTransaction(
                    time=time_str,
                    currency=curr,
                    amount=_fmt_num(amt),
                    change=_fmt_num(amt),
                    reason=f"Deposit ({status})" if status else "Deposit",
                    transaction_id=txid,
                    wallet_type="Deposit",
                )
                self.identifiers.deposit_transactions.append(txn)

            if "deposit_address" in col_map:
                v = clean_val(row.iloc[col_map["deposit_address"]])
                if v and is_wallet_address(v):
                    self.identifiers.wallet_addresses.add(v)
            if "source_address" in col_map:
                v = clean_val(row.iloc[col_map["source_address"]])
                if v and is_wallet_address(v):
                    self.identifiers.wallet_addresses.add(v)
            if "txid" in col_map:
                v = clean_val(row.iloc[col_map["txid"]])
                if v and is_txid(v):
                    self.identifiers.txids.add(v)
            if "user_id" in col_map:
                v = clean_val(row.iloc[col_map["user_id"]])
                if v:
                    self.identifiers.related_user_ids.add(str(v))
            if "counterparty" in col_map:
                v = clean_val(row.iloc[col_map["counterparty"]])
                if v:
                    self.identifiers.counterparty_ids.add(str(v))

        print(f"  Sparsowano {len(self.identifiers.deposit_transactions)} depozytów")

    def _parse_withdrawal_history(self, df: pd.DataFrame, sheet_name: str):
        cols = [str(c).strip().lower() for c in df.columns]
        col_map = {}
        for idx, c in enumerate(cols):
            if "time" in c or "date" in c or "apply" in c:
                col_map["time"] = idx
            elif "currency" in c:
                col_map["currency"] = idx
            elif "amount" in c:
                col_map["amount"] = idx
            elif "usdt" in c and "amount" not in c:
                col_map["usdt_value"] = idx
            elif "status" in c:
                col_map["status"] = idx
            elif "txid" in c or "transaction id" in c:
                col_map["txid"] = idx
            elif "destination address" in c:
                col_map["dest_address"] = idx
            elif "counterparty" in c or "counter party" in c:
                col_map["counterparty"] = idx
            elif "user id" in c:
                col_map["user_id"] = idx

        print(f"  Kolumny Withdrawal History: {list(col_map.keys())}")

        for _, row in df.iterrows():
            curr = str(clean_val(row.iloc[col_map.get("currency", 0)])) if "currency" in col_map else ""
            amt = clean_val(row.iloc[col_map.get("amount", 0)]) if "amount" in col_map else None
            status = str(clean_val(row.iloc[col_map.get("status", 0)])).lower() if "status" in col_map else ""
            time_str = str(clean_val(row.iloc[col_map.get("time", 0)])) if "time" in col_map else ""
            txid = str(clean_val(row.iloc[col_map.get("txid", 0)])) if "txid" in col_map else ""

            if curr and amt:
                txn = AssetTransaction(
                    time=time_str,
                    currency=curr,
                    amount=_fmt_num(amt),
                    change=_fmt_num(f"-{amt}"),
                    reason=f"Withdrawal ({status})" if status else "Withdrawal",
                    transaction_id=txid,
                    wallet_type="Withdrawal",
                )
                self.identifiers.withdrawal_transactions.append(txn)

            if "dest_address" in col_map:
                v = clean_val(row.iloc[col_map["dest_address"]])
                if v and is_wallet_address(v):
                    self.identifiers.wallet_addresses.add(v)
            if "txid" in col_map:
                v = clean_val(row.iloc[col_map["txid"]])
                if v and is_txid(v):
                    self.identifiers.txids.add(v)
            if "user_id" in col_map:
                v = clean_val(row.iloc[col_map["user_id"]])
                if v:
                    self.identifiers.related_user_ids.add(str(v))
            if "counterparty" in col_map:
                v = clean_val(row.iloc[col_map["counterparty"]])
                if v:
                    self.identifiers.counterparty_ids.add(str(v))

        print(f"  Sparsowano {len(self.identifiers.withdrawal_transactions)} wypłat")

    def _parse_attempted_withdrawal(self, df: pd.DataFrame, sheet_name: str):
        if "Address" in df.columns:
            for v in df["Address"].dropna():
                v = clean_val(v)
                if v and is_wallet_address(v):
                    self.identifiers.wallet_addresses.add(v)
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
        if "UserId" in df.columns:
            for v in df["UserId"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.related_user_ids.add(str(v))

    def _parse_order_history(self, df: pd.DataFrame, sheet_name: str):
        if "Order ID" in df.columns:
            for v in df["Order ID"].dropna():
                v = clean_val(v)
                if v:
                    self.identifiers.order_ids.add(v)
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
            print(f"  Generycznie znaleziono: {found}")
