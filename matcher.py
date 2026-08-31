"""Porównuje identyfikatory między raportami z różnych giełd."""

from typing import Dict, List, Any
from models.schemas import ExtractedIdentifiers


class ReportComparator:
    def __init__(self, reports: List[ExtractedIdentifiers]):
        self.reports = reports

    def _get_all_ids_for_field(self, report: ExtractedIdentifiers, field_name: str) -> set:
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
                    files_with_val = []
                    time_overview = []
                    for r in self.reports:
                        ids = self._get_all_ids_for_field(r, field_name)
                        if val in ids:
                            files_with_val.append(r.source_file)
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

        # Dodatkowe porównanie: wspólne IP z login_records HTX (per UID)
        self._compare_htx_login_ips(result)

        return result

    def _compare_htx_login_ips(self, result: Dict[str, Any]):
        """Wykrywa IP współdzielone między użytkownikami HTX w ramach porównywanych raportów."""
        ip_to_users: Dict[str, List[Dict]] = {}
        for r in self.reports:
            if r.exchange != "htx":
                continue
            for uid, records in r.login_records.items():
                for rec in records:
                    ip = rec.get("ip", "").strip()
                    if not ip:
                        continue
                    if ip not in ip_to_users:
                        ip_to_users[ip] = []
                    # Unikalny wpis per (plik, uid)
                    existing = [e for e in ip_to_users[ip] if e["file"] == r.source_file and e["uid"] == uid]
                    if not existing:
                        ip_to_users[ip].append({
                            "file": r.source_file,
                            "uid": uid,
                        })

        shared = {ip: entries for ip, entries in ip_to_users.items() if len(entries) > 1}
        if shared:
            result["htx_shared_login_ips"] = {}
            for ip, entries in sorted(shared.items()):
                users_per_file = {}
                for e in entries:
                    if e["file"] not in users_per_file:
                        users_per_file[e["file"]] = []
                    users_per_file[e["file"]].append(e["uid"])
                result["htx_shared_login_ips"][ip] = {
                    "users_per_file": users_per_file,
                    "total_users": len(entries),
                }

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
                print(f"\n [{field}] — {len(entries)} wspólnych:")
                for e in entries:
                    files_str = ", ".join(e["files"])
                    print(f" → {e['value']} (występuje w: {files_str})")
                    if "time_context" in e:
                        for tc in e["time_context"]:
                            print(f"   Czas ({tc['file']}):")
                            for r in tc["ranges"][:3]:
                                print(f"   • {r}")
        else:
            print("\n✅ Nie znaleziono wspólnych identyfikatorów.")

        # Wspólne IP z login_records HTX
        shared_ips = result.get("htx_shared_login_ips", {})
        if shared_ips:
            print("\n🔴 WSPÓLNE ADRESY IP W LOGOWANIACH HTX:")
            print("-" * 70)
            for ip, data in shared_ips.items():
                print(f"\n IP: {ip} — używane przez {data['total_users']} użytkowników:")
                for fname, uids in data["users_per_file"].items():
                    print(f"   • {fname}: UID {', '.join(sorted(set(uids)))}")

        unique = result.get("unique_per_file", {})
        if unique:
            print("\n📋 UNIKALNE IDENTYFIKATORY (tylko w jednym pliku):")
            print("-" * 70)
            for key, values in unique.items():
                file_name, field = key.split("__", 1)
                print(f"\n {file_name} :: [{field}] — {len(values)} unikalnych")
