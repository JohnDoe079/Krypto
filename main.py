import warnings
warnings.filterwarnings('ignore', message='Workbook contains no default style')
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
main.py
=======
Auto-scan katalogów data/binance/ i data/htx/.
Parsuje WSZYSTKIE pliki .xlsx, porównuje między sobą,
generuje raport DOCX.

Użycie:
 python main.py              # auto-scan + raport
 python main.py --no-docx    # tylko JSON, bez DOCX
 python main.py -v           # szczegółowe logi
"""

import argparse
import json
from pathlib import Path
from typing import List

from parsers import BinanceReportParser, HTXReportParser
from matcher import ReportComparator
from models.schemas import ExtractedIdentifiers
from reporter import ReportGenerator
from config import DATA_DIRS


def scan_directory(dir_path: str) -> List[Path]:
    path = Path(dir_path)
    if not path.exists():
        return []
    return sorted(path.glob("*.xlsx"))


def main():
    parser = argparse.ArgumentParser(
        description="Parser i komparator raportów giełdowych"
    )
    parser.add_argument(
        "-o", "--output", default="parsed_report.json",
        help="Plik wyjściowy JSON"
    )
    parser.add_argument(
        "-r", "--report", default="Raport_Analiza.docx",
        help="Plik raportu DOCX"
    )
    parser.add_argument(
        "--no-docx", action="store_true",
        help="Nie generuj raportu DOCX"
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true",
        help="Szczegółowe logi"
    )
    args = parser.parse_args()

    print("=" * 70)
    print(" CRYPTO EXCHANGE REPORT PARSER & COMPARATOR")
    print("=" * 70)

    reports: List[ExtractedIdentifiers] = []
    file_map = {}

    for exchange, dir_path in DATA_DIRS.items():
        files = scan_directory(dir_path)
        if files:
            print(f"\n📁 {exchange.upper()}: znaleziono {len(files)} plik(ów)")
            for file_path in files:
                print(f" 📄 {file_path.name}")
                if exchange == "binance":
                    p = BinanceReportParser(str(file_path))
                elif exchange == "htx":
                    p = HTXReportParser(str(file_path))
                else:
                    continue

                ids = p.parse_all()
                reports.append(ids)
                file_map[file_path.name] = exchange

                if args.verbose:
                    print("\n" + ids.summary())
        else:
            print(f"\n📁 {exchange.upper()}: brak plików w {dir_path}/")

    if not reports:
        print("\n[!] Nie znaleziono żadnych plików .xlsx.")
        return

    output_data = [r.to_dict() for r in reports]
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(output_data, f, indent=2, ensure_ascii=False)
    print(f"\n💾 Zapisano wyniki JSON do: {args.output}")

    if len(reports) >= 2:
        print(f"\n🔍 Porównywanie {len(reports)} raportów...")
        comp = ReportComparator(reports)
        comp.print_comparison()
        comp_file = args.output.replace(".json", "_comparison.json")
        with open(comp_file, "w", encoding="utf-8") as f:
            json.dump(comp.compare(), f, indent=2, ensure_ascii=False)
        print(f"\n💾 Zapisano porównanie do: {comp_file}")
    else:
        print(f"\nℹ️ Tylko 1 raport — brak porównania.")

    if not args.no_docx:
        print(f"\n📝 Generowanie raportu DOCX: {args.report}...")
        gen = ReportGenerator(args.report)
        gen.generate(reports, file_map)
        print(f"\n✅ Raport gotowy: {args.report}")

    print("\n" + "=" * 70)
    print(" KONIEC")
    print("=" * 70)


if __name__ == "__main__":
    main()
