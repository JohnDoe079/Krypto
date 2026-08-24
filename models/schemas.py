"""Modele danych i funkcje pomocnicze do ekstrakcji identyfikatorów."""

import re
import warnings
import pandas as pd
from dataclasses import dataclass, field
from typing import Dict, List, Set, Optional


@dataclass
class AssetBalance:
    """Pojedyncze saldo waluty z Assets Overview."""
    currency_name: str = ""
    currency_code: str = ""
    all_positions: str = ""
    available_positions: str = ""
    in_withdrawal: str = ""
    pending_order: str = ""
    btc_equivalent: str = ""
    usdt_equivalent: str = ""
    wallet_type: str = ""

    def to_dict(self) -> dict:
        return {
            "currency_name": self.currency_name,
            "currency_code": self.currency_code,
            "all_positions": self.all_positions,
            "available_positions": self.available_positions,
            "in_withdrawal": self.in_withdrawal,
            "pending_order": self.pending_order,
            "btc_equivalent": self.btc_equivalent,
            "usdt_equivalent": self.usdt_equivalent,
            "wallet_type": self.wallet_type,
        }


@dataclass
class AssetTransaction:
    """Pojedyncza transakcja z dowolnego arkusza (Spot/Funding/Deposit/Withdrawal/OTC/P2P/Fiat/itp)."""
    time: str = ""
    currency: str = ""
    amount: str = ""
    locked: str = ""
    freeze: str = ""
    processing: str = ""
    change: str = ""
    reason: str = ""
    transaction_id: str = ""
    wallet_type: str = ""          # Spot / Funding / Deposit / Withdrawal / OTC / P2P / Pay / Fiat
    source_sheet: str = ""         # Nazwa oryginalnego arkusza, np. "Spot Asset Log"
    user_id: str = ""              # User ID z wiersza (jeśli inny niż właściciel — do analizy)

    def to_dict(self) -> dict:
        return {
            "time": self.time,
            "currency": self.currency,
            "amount": self.amount,
            "locked": self.locked,
            "freeze": self.freeze,
            "processing": self.processing,
            "change": self.change,
            "reason": self.reason,
            "transaction_id": self.transaction_id,
            "wallet_type": self.wallet_type,
            "source_sheet": self.source_sheet,
            "user_id": self.user_id,
        }


@dataclass
class ExtractedIdentifiers:
    source_file: str = ""
    exchange: str = ""

    user_ids: Set[str] = field(default_factory=set)
    related_user_ids: Set[str] = field(default_factory=set)

    emails: Set[str] = field(default_factory=set)
    phones: Set[str] = field(default_factory=set)
    ips: Set[str] = field(default_factory=set)
    wallet_addresses: Set[str] = field(default_factory=set)
    txids: Set[str] = field(default_factory=set)
    card_bins: Set[str] = field(default_factory=set)
    card_last4: Set[str] = field(default_factory=set)
    formatted_cards: Set[str] = field(default_factory=set)  # np. "5355 57** **** 3305"
    ibans: Set[str] = field(default_factory=set)
    account_numbers: Set[str] = field(default_factory=set)
    device_ids: Set[str] = field(default_factory=set)
    fvideo_ids: Set[str] = field(default_factory=set)
    bnc_uuids: Set[str] = field(default_factory=set)
    order_ids: Set[str] = field(default_factory=set)
    counterparty_ids: Set[str] = field(default_factory=set)
    transaction_ids: Set[str] = field(default_factory=set)
    names: Set[str] = field(default_factory=set)
    nationalities: Set[str] = field(default_factory=set)
    id_numbers: Set[str] = field(default_factory=set)
    geolocations: Set[str] = field(default_factory=set)
    browsers: Set[str] = field(default_factory=set)

    estimate_total_btc: str = ""
    estimate_total_usdt: str = ""
    asset_balances: List[AssetBalance] = field(default_factory=list)
    spot_transactions: List[AssetTransaction] = field(default_factory=list)
    funding_transactions: List[AssetTransaction] = field(default_factory=list)
    deposit_transactions: List[AssetTransaction] = field(default_factory=list)
    withdrawal_transactions: List[AssetTransaction] = field(default_factory=list)
    fiat_deposit_transactions: List[AssetTransaction] = field(default_factory=list)
    fiat_trade_transactions: List[AssetTransaction] = field(default_factory=list)

    time_ranges: Dict[str, Dict[str, str]] = field(default_factory=dict)

    customer_info_sections: Dict[str, Dict[str, str]] = field(default_factory=dict)
    kyc_images: List[str] = field(default_factory=list)
    parsed_sheets: List[str] = field(default_factory=list)
    unknown_sheets: List[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "source_file": self.source_file,
            "exchange": self.exchange,
            "user_ids": sorted(self.user_ids),
            "related_user_ids": sorted(self.related_user_ids),
            "emails": sorted(self.emails),
            "phones": sorted(self.phones),
            "ips": sorted(self.ips),
            "wallet_addresses": sorted(self.wallet_addresses),
            "txids": sorted(self.txids),
            "card_bins": sorted(self.card_bins),
            "card_last4": sorted(self.card_last4),
            "formatted_cards": sorted(self.formatted_cards),
            "ibans": sorted(self.ibans),
            "account_numbers": sorted(self.account_numbers),
            "device_ids": sorted(self.device_ids),
            "fvideo_ids": sorted(self.fvideo_ids),
            "bnc_uuids": sorted(self.bnc_uuids),
            "order_ids": sorted(self.order_ids),
            "counterparty_ids": sorted(self.counterparty_ids),
            "transaction_ids": sorted(self.transaction_ids),
            "names": sorted(self.names),
            "nationalities": sorted(self.nationalities),
            "id_numbers": sorted(self.id_numbers),
            "geolocations": sorted(self.geolocations),
            "browsers": sorted(self.browsers),
            "estimate_total_btc": self.estimate_total_btc,
            "estimate_total_usdt": self.estimate_total_usdt,
            "asset_balances": [b.to_dict() for b in self.asset_balances],
            "spot_transactions": [t.to_dict() for t in self.spot_transactions],
            "funding_transactions": [t.to_dict() for t in self.funding_transactions],
            "deposit_transactions": [t.to_dict() for t in self.deposit_transactions],
            "withdrawal_transactions": [t.to_dict() for t in self.withdrawal_transactions],
            "fiat_deposit_transactions": [t.to_dict() for t in self.fiat_deposit_transactions],
            "fiat_trade_transactions": [t.to_dict() for t in self.fiat_trade_transactions],
            "time_ranges": self.time_ranges,
            "parsed_sheets": self.parsed_sheets,
            "unknown_sheets": self.unknown_sheets,
        }

    def summary(self) -> str:
        lines = [
            f"Plik: {self.source_file}",
            f"Giełda: {self.exchange}",
            f"  Arkuszy sparsowanych: {len(self.parsed_sheets)}",
            f"  ID użytkownika (właściciel): {len(self.user_ids)}",
            f"  ID powiązanych użytkowników: {len(self.related_user_ids)}",
            f"  E-maile: {len(self.emails)}",
            f"  Numery telefonów: {len(self.phones)}",
            f"  IP: {len(self.ips)}",
            f"  Adresy portfeli: {len(self.wallet_addresses)}",
            f"  TXID: {len(self.txids)}",
            f"  BIN kart: {len(self.card_bins)}",
            f"  Ostatnie 4 cyfry kart: {len(self.card_last4)}",
            f"  Sformatowane karty: {len(self.formatted_cards)}",
            f"  IBAN: {len(self.ibans)}",
            f"  Numery kont: {len(self.account_numbers)}",
            f"  ID urządzeń: {len(self.device_ids)}",
            f"  ID Fvideo: {len(self.fvideo_ids)}",
            f"  UUID BNC: {len(self.bnc_uuids)}",
            f"  ID zamówień: {len(self.order_ids)}",
            f"  ID kontrahentów: {len(self.counterparty_ids)}",
            f"  ID transakcji: {len(self.transaction_ids)}",
            f"  Imiona/nazwiska: {len(self.names)}",
            f"  Narodowości: {len(self.nationalities)}",
            f"  Numery dokumentów: {len(self.id_numbers)}",
            f"  Lokalizacje: {len(self.geolocations)}",
            f"  Przeglądarki: {len(self.browsers)}",
            f"  Estimate Total BTC: {self.estimate_total_btc}",
            f"  Salda walut: {len(self.asset_balances)}",
            f"  Transakcje Spot: {len(self.spot_transactions)}",
            f"  Transakcje Funding: {len(self.funding_transactions)}",
            f"  Transakcje Deposit: {len(self.deposit_transactions)}",
            f"  Transakcje Withdrawal: {len(self.withdrawal_transactions)}",
            f"  Transakcje Fiat Deposit: {len(self.fiat_deposit_transactions)}",
            f"  Transakcje Fiat Trades: {len(self.fiat_trade_transactions)}",
            f"  Zakresy czasowe: {len(self.time_ranges)}",
        ]
        return "\n".join(lines)


def clean_val(val) -> Optional[str]:
    if pd.isna(val):
        return None
    v = str(val).strip()
    if v in ("", "N/A", "nan", "None", "NaN", "-"):
        return None
    return v


def _normalize_decimal(val) -> str:
    """Normalizuje string liczbowy do formatu z kropką dziesiętną.

    Obsługuje formaty:
    - US:           1,000.50   → 1000.50
    - Europejski:   1.000,50   → 1000.50
    - Bez tysięcy:  0,05936011 → 0.05936011
    - Ze spacją:    6 020,55   → 6020.55
    """
    if val is None or val == "":
        return ""
    s = str(val).strip()

    # Usuń spacje (zawsze separatory tysięcy)
    s = s.replace(" ", "")

    if "," in s and "." in s:
        last_comma = s.rfind(",")
        last_dot = s.rfind(".")
        if last_comma > last_dot:
            # Europejski: 1.000,50
            s = s.replace(".", "").replace(",", ".")
        else:
            # US: 1,000.50
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) == 2:
            after = parts[1]
            before = parts[0]
            # W krypto po przecinku zazwyczaj jest >2 cyfr (np. 0,05936011)
            if len(after) > 2:
                s = s.replace(",", ".")
            elif len(after) <= 2 and len(before) > 3:
                # Prawdopodobnie tysiące, np. 1,000
                s = s.replace(",", "")
            else:
                s = s.replace(",", ".")
        else:
            s = s.replace(",", ".")

    return s


def is_valid_ip(val: str) -> bool:
    if not val:
        return False
    v = val.strip()
    if not re.match(r"^(\d{1,3}\.){3}\d{1,3}$", v):
        return False
    try:
        return all(0 <= int(x) <= 255 for x in v.split("."))
    except ValueError:
        return False


def is_wallet_address(val: str) -> bool:
    if not val or len(val) < 20:
        return False
    v = val.strip()
    if v.startswith("bc1") or v.startswith("1") or v.startswith("3"):
        return True
    if v.startswith("0x") and len(v) >= 40:
        return True
    if v.startswith("T") and len(v) >= 30:
        return True
    if v.startswith("r") and len(v) >= 25:
        return True
    if len(v) >= 95 and v[0] in "48":
        return True
    return False


def is_txid(val: str) -> bool:
    if not val:
        return False
    v = val.strip()
    return len(v) == 64 and all(c in "0123456789abcdefABCDEF" for c in v)


def extract_email(val: str) -> Optional[str]:
    v = clean_val(val)
    if not v:
        return None
    v = v.lstrip("'")
    if re.match(r"^[^\s@]+@[^\s@]+\.[^\s@]+$", v):
        return v.lower()
    return None


def extract_phone(val: str) -> Optional[str]:
    v = clean_val(val)
    if not v:
        return None
    v = v.lstrip("'")
    if re.match(r"^\d{4}-\d{2}-\d{2}$", v):
        return None
    if re.match(r"^[\+\d][\d\s\-\(\)]{7,25}$", v):
        return v
    return None


def extract_time_range(df: pd.DataFrame) -> Optional[Dict[str, str]]:
    time_cols = [c for c in df.columns if any(
        kw in str(c).lower() for kw in ["time", "date", "create", "update", "login", "timestamp"]
    )]
    if not time_cols:
        return None

    all_dates = []
    for col in time_cols:
        parsed = None
        # Próbuj najpierw najczęstsze formaty z raportów Binance
        for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f", "%d/%m/%Y %H:%M:%S"]:
            try:
                parsed = pd.to_datetime(df[col], format=fmt, errors="coerce")
                if parsed.notna().any():
                    break
            except (ValueError, TypeError):
                continue

        # Fallback z wyciszeniem warningu
        if parsed is None or parsed.isna().all():
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", UserWarning)
                parsed = pd.to_datetime(df[col], errors="coerce")

        valid = parsed.dropna()
        if len(valid) > 0:
            all_dates.extend(valid.tolist())

    if not all_dates:
        return None

    return {
        "from": min(all_dates).strftime("%Y-%m-%d %H:%M:%S"),
        "to": max(all_dates).strftime("%Y-%m-%d %H:%M:%S"),
    }
