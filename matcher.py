"""Porównuje identyfikatory między raportami z różnych giełd."""

from typing import Dict, List, Any
from models.schemas import ExtractedIdentifiers


class ReportComparator:
    def __init__(self, reports: List[ExtractedIdentifiers]):
        self.reports = reports

    def _get_all_ids_for_field(self, report: ExtractedIdentifiers, field_name: str) -> set:
        """Zwraca wszystkie ID dla danego pola (łączy user_ids + related_user_ids jeżeli to user)."""
        if field_name == "all_user_ids":
            return report.user_ids | report.related_user_ids
        return set(getattr(report, field_name, set()))

    def compare(self) -> Dict[str, Any]:
        if len(self.reports) < 2:
            return {"error": "Potrzebne są co najmniej 2 raporty do porównania."}

        result = {
            "compared_files": [r.source_file for r in self.reports],
            "common": {},
            "unique_per_file": {},
        }

        fields = [
            "user_ids", "related_user_ids", "all_user_ids",
            "emails", "phones", "ips", "wallet_addresses",
            "txids", "card_bins", "card_last4", "ibans", "account_numbers",
            "device_ids", "order_ids", "counterparty_ids", "transaction_ids",
            "names", "nationalities", "id_numbers", "geolocations",
        ]

        for field_name in fields:
            sets = [self._get_all_ids_for_field(r, field_name) for r in self.reports]
            common = set.intersection(*sets) if sets else set()

            if common:
                result["common"][field_name] = []
                for val in sorted(common):
                    # Znajdź pliki, w których występuje ta wartość
                    files_with_val = []
                    time_overview = []
                    for r in self.reports:
                        ids = self._get_all_ids_for_field(r, field_name)
                        if val in ids:
                            files_with_val.append(r.source_file)
                            # Zbierz zakresy czasowe z tego pliku
                            if r.time_ranges:
                                ranges = []
                                for sheet, tr in r.time_ranges.items():
                                    ranges.append(f"{sheet}: {tr['from']} -> {tr['to']}")
                                if ranges:
                                    time_overview.append({
                                        "file": r.source_file,
                                        "ranges": ranges
                                    })

                    entry = {
                        "value": val,
                        "files": files_with_val,
                    }
                    if time_overview:
                        entry["time_context"] = time_overview
                    result["common"][field_name].append(entry)

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
            for field, entries in common.items():
                print(f"\n  [{field}] — {len(entries)} wspólnych:")
                for e in entries:
                    files_str = ", ".join(e["files"])
                    print(f"    → {e['value']}  (występuje w: {files_str})")
                    if "time_context" in e:
                        for tc in e["time_context"]:
                            print(f"       Czas ({tc['file']}):")
                            for r in tc["ranges"][:3]:  # max 3 zakresy
                                print(f"         • {r}")
        else:
            print("\n✅ Nie znaleziono wspólnych identyfikatorów.")

        unique = result.get("unique_per_file", {})
        if unique:
            print("\n📋 UNIKALNE IDENTYFIKATORY (tylko w jednym pliku):")
            print("-" * 70)
            for key, values in unique.items():
                file_name, field = key.split("__", 1)
                print(f"\n  {file_name} :: [{field}] — {len(values)} unikalnych")
