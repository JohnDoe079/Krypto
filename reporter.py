"""Generuje raport DOCX z wyników analizy raportów giełdowych."""

import os
from datetime import datetime
from typing import Dict, List
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

from models.schemas import ExtractedIdentifiers
from matcher import ReportComparator

# Tłumaczenia sekcji Customer Information
SECTION_TRANSLATIONS = {
    "Basic Information": "Basic Information (Podstawowe informacje)",
    "API Information": "API Information (Informacje API)",
    "KYC Approved Info": "KYC Approved Info (Zatwierdzone dane KYC)",
    "KYC Info": "KYC Info (Informacje KYC)",
    "Address Validation": "Address Validation (Weryfikacja adresu)",
    "EDD Info": "EDD Info (Informacje EDD – Rozszerzona due diligence)",
    "Payment Merchant Info": "Payment Merchant Info (Informacje o sprzedawcy)",
    "Sub-accounts": "Sub-accounts (Subkonta)",
}

# Kolejność wierszy w Basic Information
BASIC_INFO_ORDER = [
    "Registration time", "User ID", "User authentication type",
    "User authentication time", "User nationality", "User ID number",
    "First Name", "Last Name", "Mobile", "Email", "Status",
    "2FA", "2FA Opening time", "2FA Reset Information",
    "SMS", "SMS Opening time", "Tax ID", "TnC", "TnC Sign Date",
]

PAGE_WIDTH_INCHES = 6.1


class ReportGenerator:
    def __init__(self, output_path: str = "Raport_Analiza.docx"):
        self.output_path = output_path
        self.doc = Document()
        self._setup_styles()

    def _setup_styles(self):
        style = self.doc.styles['Normal']
        font = style.font
        font.name = 'Calibri'
        font.size = Pt(11)

    def _add_heading(self, text: str, level: int = 1):
        heading = self.doc.add_heading(text, level=level)
        heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if level == 1:
            for run in heading.runs:
                run.font.color.rgb = RGBColor(0x00, 0x00, 0x80)
                run.font.size = Pt(20)
        elif level == 2:
            for run in heading.runs:
                run.font.color.rgb = RGBColor(0x00, 0x40, 0x80)
                run.font.size = Pt(14)
        elif level == 3:
            for run in heading.runs:
                run.font.color.rgb = RGBColor(0x00, 0x60, 0x80)
                run.font.size = Pt(12)
        elif level == 4:
            for run in heading.runs:
                run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
                run.font.size = Pt(11)
                run.bold = True
        return heading

    def _add_paragraph(self, text: str, bold: bool = False, color: RGBColor = None):
        p = self.doc.add_paragraph()
        run = p.add_run(text)
        run.bold = bold
        if color:
            run.font.color.rgb = color
        return p

    def _add_table(self, headers: List[str], rows: List[List[str]], max_rows: int = 100):
        if not rows:
            self._add_paragraph("(brak danych)")
            return None

        n_cols = len(headers)
        table = self.doc.add_table(rows=1, cols=n_cols)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        table.autofit = False
        table.allow_autofit = False

        col_width = Inches(PAGE_WIDTH_INCHES / n_cols)
        for i in range(n_cols):
            table.columns[i].width = col_width

        # Nagłówki
        hdr_cells = table.rows[0].cells
        for i, h in enumerate(headers):
            hdr_cells[i].text = h
            for paragraph in hdr_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = Pt(10)
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            # Tło nagłówka — ciemny granat
            shading_elm = parse_xml(r'<w:shd {} w:fill="1F4E78"/>'.format(nsdecls('w')))
            hdr_cells[i]._tc.get_or_add_tcPr().append(shading_elm)

        # Wiersze
        for row_data in rows[:max_rows]:
            row_cells = table.add_row().cells
            for i, cell_text in enumerate(row_data):
                txt = str(cell_text) if cell_text is not None else ""
                row_cells[i].text = txt[:250]
                for paragraph in row_cells[i].paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(9)

        if len(rows) > max_rows:
            self._add_paragraph(f" ... i {len(rows) - max_rows} więcej wierszy")
        return table

    def _sort_basic_info(self, data: Dict[str, str]) -> List[List[str]]:
        order_map = {name: idx for idx, name in enumerate(BASIC_INFO_ORDER)}
        items = list(data.items())
        priority = []
        rest = []
        for col, val in items:
            if col in order_map:
                priority.append((order_map[col], col, val))
            else:
                rest.append((col, val))
        priority.sort(key=lambda x: x[0])
        rows = [[col, val if val is not None and str(val).strip() != "" else "(puste)"]
                for _, col, val in priority]
        rows += [[col, val if val is not None and str(val).strip() != "" else "(puste)"]
                 for col, val in rest]
        return rows

    def generate(self, reports: List[ExtractedIdentifiers], file_map: Dict[str, str]):
        # ===== STRONA TYTUŁOWA =====
        self.doc.add_paragraph()
        self.doc.add_paragraph()
        title = self.doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run("RAPORT ANALIZY")
        run.bold = True
        run.font.size = Pt(32)
        run.font.color.rgb = RGBColor(0x00, 0x00, 0x80)

        subtitle = self.doc.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run2 = subtitle.add_run("Raportów giełdowych kryptowalutowych")
        run2.font.size = Pt(18)

        self.doc.add_paragraph()
        date_p = self.doc.add_paragraph()
        date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_p.add_run(f"Data wygenerowania: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

        self.doc.add_page_break()

        # ===== SPIS TREŚCI =====
        self._add_heading("Spis treści", level=1)
        toc_items = [
            "1. Podsumowanie",
            "2. Szczegółowa analiza poszczególnych plików",
        ]
        for i, r in enumerate(reports, 1):
            toc_items.append(f"  2.{i}. {r.source_file}")
        if len(reports) >= 2:
            toc_items.append("3. Porównanie między raportami")
        toc_items.append("4. Pełna lista identyfikatorów")

        for item in toc_items:
            self.doc.add_paragraph(item, style='List Bullet')

        self.doc.add_page_break()

        # ===== 1. PODSUMOWANIE =====
        self._add_heading("1. Podsumowanie", level=1)
        self._add_paragraph(f"Liczba przeanalizowanych plików: {len(reports)}")
        self._add_paragraph(f"Liczba giełd: {len(set(file_map.values()))}")

        summary_rows = []
        for r in reports:
            # Główny zakres czasowy (najszerszy z wszystkich arkuszy)
            all_from = []
            all_to = []
            for tr in r.time_ranges.values():
                all_from.append(tr.get("from", ""))
                all_to.append(tr.get("to", ""))
            time_summary = ""
            if all_from and all_to:
                time_summary = f"{min(all_from)} -> {max(all_to)}"

            summary_rows.append([
                r.source_file,
                r.exchange.upper(),
                str(len(r.parsed_sheets)),
                ", ".join(sorted(r.user_ids)) if r.user_ids else "(brak)",
                str(len(r.related_user_ids)),
                str(len(r.emails)),
                str(len(r.ips)),
                str(len(r.wallet_addresses)),
                str(len(r.txids)),
                time_summary,
            ])

        self._add_table(
            ["Plik", "Giełda", "Arkusze", "ID właściciela", "ID powiązanych",
             "E-maile", "IP", "Portfele", "TXID", "Zakres czasowy"],
            summary_rows
        )

        self.doc.add_page_break()

        # ===== 2. SZCZEGÓŁY KAŻDEGO PLIKU =====
        self._add_heading("2. Szczegółowa analiza poszczególnych plików", level=1)

        for idx, r in enumerate(reports, 1):
            self._add_heading(f"2.{idx}. {r.source_file}", level=2)
            self._add_paragraph(f"Giełda: {r.exchange.upper()}")
            self._add_paragraph(
                f"Przeanalizowane arkusze ({len(r.parsed_sheets)}): {', '.join(r.parsed_sheets)}")
            if r.unknown_sheets:
                self._add_paragraph(
                    f"Nieznane arkusze: {', '.join(r.unknown_sheets)}",
                    color=RGBColor(0xC0, 0x00, 0x00))

            # --- ZAKRESY CZASOWE ---
            if r.time_ranges:
                self._add_heading("Zakresy czasowe per arkusz", level=3)
                time_rows = []
                for sheet_name, tr in sorted(r.time_ranges.items()):
                    time_rows.append([sheet_name, tr.get("from", ""), tr.get("to", "")])
                self._add_table(["Arkusz", "Od", "Do"], time_rows, max_rows=50)

            # --- CUSTOMER INFORMATION ---
            if r.customer_info_sections:
                self._add_heading("Customer Information (Informacje o użytkowniku)", level=3)
                for sec_name, data in r.customer_info_sections.items():
                    if not data:
                        continue
                    display_name = SECTION_TRANSLATIONS.get(sec_name, sec_name)
                    self._add_heading(display_name, level=4)

                    if sec_name == "Basic Information":
                        rows = self._sort_basic_info(data)
                    else:
                        rows = []
                        for col, val in data.items():
                            if val is not None and str(val).strip() != "":
                                rows.append([str(col), str(val)])
                            else:
                                rows.append([str(col), "(puste)"])
                    self._add_table(["Pole", "Wartość"], rows, max_rows=100)

            # --- KYC DOCUMENTS ---
            if "KYC Documents" in r.parsed_sheets:
                self._add_heading("KYC Documents (Dokumenty KYC)", level=3)
                if r.kyc_images:
                    n_cols = 2
                    n_rows = (len(r.kyc_images) + n_cols - 1) // n_cols
                    table = self.doc.add_table(rows=n_rows, cols=n_cols)
                    table.style = 'Table Grid'
                    table.alignment = WD_TABLE_ALIGNMENT.LEFT
                    table.autofit = False
                    table.allow_autofit = False
                    col_w = Inches(PAGE_WIDTH_INCHES / n_cols)
                    for i in range(n_cols):
                        table.columns[i].width = col_w

                    for img_idx, img_path in enumerate(r.kyc_images):
                        row_idx = img_idx // n_cols
                        col_idx = img_idx % n_cols
                        cell = table.cell(row_idx, col_idx)
                        cell.text = ""
                        if os.path.exists(img_path):
                            try:
                                p = cell.paragraphs[0]
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                run = p.add_run()
                                run.add_picture(img_path, width=Inches(2.8))
                            except Exception as e:
                                cell.text = f"[!] Błąd: {e}"
                        else:
                            cell.text = "Brak pliku"
                else:
                    self._add_paragraph("Brak osadzonych obrazków w arkuszu KYC Documents.")

            self.doc.add_page_break()

        # ===== 3. PORÓWNANIE =====
        if len(reports) >= 2:
            self._add_heading("3. Porównanie między raportami", level=1)
            comp = ReportComparator(reports)
            result = comp.compare()

            common = result.get("common", {})
            if common:
                self._add_heading("Wspólne identyfikatory (potencjalne powiązania):", level=2)
                self._add_paragraph(
                    "Poniższe identyfikatory występują w co najmniej dwóch raportach. "
                    "Mogą wskazywać na powiązanie między kontami.",
                    color=RGBColor(0xC0, 0x00, 0x00)
                )

                for field_name, entries in common.items():
                    self._add_heading(f"[{field_name}] — {len(entries)} wspólnych", level=3)
                    rows = []
                    for i, e in enumerate(entries):
                        files_str = ", ".join(e["files"])
                        time_str = ""
                        if "time_context" in e:
                            parts = []
                            for tc in e["time_context"]:
                                file_short = tc["file"]
                                ranges = tc["ranges"]
                                if ranges:
                                    parts.append(f"{file_short}: {ranges[0]}")
                            time_str = "; ".join(parts)
                        rows.append([str(i+1), str(e["value"]), files_str, time_str])
                    self._add_table(["Lp.", "Wartość", "Pliki", "Zakres czasowy"], rows, max_rows=30)
            else:
                self._add_paragraph("Nie znaleziono wspólnych identyfikatorów między raportami.")

            self.doc.add_page_break()

        # ===== 4. PEŁNA LISTA IDENTYFIKATORÓW =====
        self._add_heading("4. Pełna lista identyfikatorów", level=1)
        self._add_paragraph(
            "Szczegółowe dane w formacie JSON zostały zapisane w pliku parsed_report.json.")

        for r in reports:
            self._add_heading(f"{r.source_file}", level=2)

            # ID właściciela (z Customer Information)
            if r.user_ids:
                self._add_paragraph(f"ID właściciela konta ({len(r.user_ids)}):", bold=True)
                rows = [[str(i+1), str(item)] for i, item in enumerate(sorted(r.user_ids))]
                self._add_table(["Lp.", "Wartość"], rows, max_rows=20)

            # ID powiązanych użytkowników
            if r.related_user_ids:
                self._add_paragraph(
                    f"ID powiązanych użytkowników (z P2P, Pay, itp.) ({len(r.related_user_ids)}):",
                    bold=True)
                rows = [[str(i+1), str(item)] for i, item in enumerate(sorted(r.related_user_ids))]
                self._add_table(["Lp.", "Wartość"], rows, max_rows=20)

            id_sections = [
                ("E-maile", r.emails),
                ("Numery telefonów", r.phones),
                ("Adresy IP", r.ips),
                ("Adresy portfeli (krypto)", r.wallet_addresses),
                ("TXID (hash transakcji)", r.txids),
                ("BIN karty", r.card_bins),
                ("Ostatnie 4 cyfry karty", r.card_last4),
                ("IBAN", r.ibans),
                ("Numery kont", r.account_numbers),
                ("ID urządzeń", r.device_ids),
                ("ID Fvideo", r.fvideo_ids),
                ("UUID BNC", r.bnc_uuids),
                ("ID zamówień", r.order_ids),
                ("ID kontrahentów", r.counterparty_ids),
                ("ID transakcji", r.transaction_ids),
                ("Imiona i nazwiska", r.names),
                ("Narodowości", r.nationalities),
                ("Numery dokumentów tożsamości", r.id_numbers),
                ("Lokalizacje GEO", r.geolocations),
                ("Przeglądarki / User Agent", r.browsers),
            ]

            for title, items in id_sections:
                if items:
                    self._add_paragraph(f"{title} ({len(items)}):", bold=True)
                    rows = [[str(i+1), str(item)] for i, item in enumerate(sorted(items))]
                    self._add_table(["Lp.", "Wartość"], rows, max_rows=20)

        self.doc.save(self.output_path)
