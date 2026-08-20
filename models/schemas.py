# models/schemas.py
"""Modele danych i funkcje pomocnicze do ekstrakcji identyfikatorów."""

import re
import pandas as pd
from dataclasses import dataclass, field
from typing import Dict, List, Set, Optional


@dataclass
class ExtractedIdentifiers:
    source_file: str = ""
    exchange: str = ""

    user_ids: Set[str] = field(default_factory=set)
    emails: Set[str] = field(default_factory=set)
    phones: Set[str] = field(default_factory=set)
    ips: Set[str] = field(default_factory=set)
    wallet_addresses: Set[str] = field(default_factory=set)
    txids: Set[str] = field(default_factory=set)
    card_bins: Set[str] = field(default_factory=set)
    card_last4: Set[str] = field(default_factory=set)
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

    customer_info_sections: Dict[str, Dict[str, str]] = field(default_factory=dict)
    kyc_images: List[str] = field(default_factory=list)
    parsed_sheets: List[str] = field(default_factory=list)
    unknown_sheets: List[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "source_file": self.source_file,
            "exchange": self.exchange,
            "user_ids": sorted(self.user_ids),
            "emails": sorted(self.emails),
            "phones": sorted(self.phones),
            "ips": sorted(self.ips),
            "wallet_addresses": sorted(self.wallet_addresses),
            "txids": sorted(self.txids),
            "card_bins": sorted(self.card_bins),
            "card_last4": sorted(self.card_last4),
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
            "parsed_sheets": self.parsed_sheets,
            "unknown_sheets": self.unknown_sheets,
        }

    def summary(self) -> str:
        lines = [
            f"Plik: {self.source_file}",
            f"Giełda: {self.exchange}",
            f"  Arkuszy sparsowanych: {len(self.parsed_sheets)}",
            f"  ID użytkowników: {len(self.user_ids)}",
            f"  E-maile: {len(self.emails)}",
            f"  Numery telefonów: {len(self.phones)}",
            f"  IP: {len(self.ips)}",
            f"  Adresy portfeli: {len(self.wallet_addresses)}",
            f"  TXID: {len(self.txids)}",
            f"  BIN kart: {len(self.card_bins)}",
            f"  Ostatnie 4 cyfry kart: {len(self.card_last4)}",
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
        ]
        return "\n".join(lines)


def clean_val(val) -> Optional[str]:
    if pd.isna(val):
        return None
    v = str(val).strip()
    if v in ("", "N/A", "nan", "None", "NaN"):
        return None
    return v


def is_valid_ip(val: str) -> bool:
    if not val:
        return False
    return bool(re.match(r"^(\d{1,3}\.){3}\d{1,3}$", val.strip()))


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
    if re.match(r"^[\+\d][\d\s\-\(\)]{7,20}$", v):
        return v
    return None
