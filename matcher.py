# matcher.py
"""Porównuje identyfikatory między raportami z różnych giełd."""

from typing import Dict, List
from models.schemas import ExtractedIdentifiers


class ReportComparator:
    def __init__(self, reports: List[ExtractedIdentifiers]):
        self.reports = reports

    def compare(self) -> Dict[str, any]:
        if len(self.reports) < 2:
            return {"error": "Potrzebne są co najmniej 2 raporty do porównania."}

        result = {
            "compared_files": [r.source_file for r in self.reports],
            "common": {},
            "unique_per_file": {},
        }

        fields = [
            "user_ids", "emails", "phones", "ips", "wallet_addresses",
            "txids", "card_bins", "card_last4", "ibans", "account_numbers",
            "device_ids", "order_ids", "counterparty_ids", "transaction_ids",
            "names", "nationalities", "id_numbers", "geolocations",
        ]

        for field_name in fields:
            sets = [set(getattr(r, field_name)) for r in self.reports]
            common = set.intersection(*sets) if sets else set()
            if common:
                result["common"][field_name] = sorted(common)

            for i, r in enumerate(self.reports):
                others = [s for j, s in enumerate(sets) if j != i]
                if others:
                    unique = sets[i] - set.union(*others)
                else:
                    unique = sets[i]
                if unique:
                    key = f"{r.source_file}__{field_name}"
                    result["unique_per_file"][key] = sorted(unique)

        return result

    def print_comparison(self):
        result = self.compare()
        print("\n" + "=" * 70)
        print("PORÓWNANIE RAPORTÓW")
        print("=" * 70)
        print(f"Pliki: {', '.join(result.get('compared_files', []))}")

        common = result.get("common", {})
        if common:
            print("\n🔴 WSPÓLNE IDENTYFIKATORY (POTENCJALNE POWIĄZANIA):")
            print("-" * 70)
            for field, values in common.items():
                print(f"\n  [{field}] — {len(values)} wspólnych:")
                for v in values:
                    print(f"    → {v}")
        else:
            print("\n✅ Nie znaleziono wspólnych identyfikatorów.")

        unique = result.get("unique_per_file", {})
        if unique:
            print("\n📋 UNIKALNE IDENTYFIKATORY (tylko w jednym pliku):")
            print("-" * 70)
            for key, values in unique.items():
                file_name, field = key.split("__", 1)
                print(f"\n  {file_name} :: [{field}] — {len(values)} unikalnych")
