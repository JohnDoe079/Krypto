"""Generuje raport DOCX z wynikow analizy raportow gieldowych."""

import os
import re
from datetime import datetime
from typing import Dict, List
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

from models.schemas import ExtractedIdentifiers, AssetTransaction
from matcher import ReportComparator

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

BASIC_INFO_ORDER = [
    "Registration time", "User ID", "User authentication type",
    "User authentication time", "User nationality", "User ID number",
    "First Name", "Last Name", "Mobile", "Email", "Status",
    "2FA", "2FA Opening time", "2FA Reset Information",
    "SMS", "SMS Opening time", "Tax ID", "TnC", "TnC Sign Date",
]

PAGE_WIDTH_INCHES = 7.0


def _to_float(val: str) -> float:
    try:
        return float(str(val).replace(",", "").replace(" ", ""))
    except (ValueError, TypeError):
        return 0.0


def _is_pending_transaction(t: AssetTransaction) -> bool:
    """Sprawdza czy transakcja jest pending/niepotwierdzona (pomijamy w podsumowaniu)."""
    reason = str(t.reason).lower().strip() if t.reason else ""
    if not reason:
        return False
    # Pending keywords
    pending_kw = ["pending", "processing", "initiated", "created", "request", "order"]
    # Operacje ktore musza miec 'success' aby byc liczone
    success_required = ["withdrawal", "deposit", "transfer"]

    # Jezeli zawiera pending keyword i nie ma success
    has_pending = any(kw in reason for kw in pending_kw)
    has_success = "success" in reason
    if has_pending and not has_success:
        return True

    # Jezeli to operacja wymagajaca success, a go nie ma
    for op in success_required:
        if op in reason and not has_success:
            # Wyjatki: "fee", "commission", "reward" nie wymagaja success
            if any(x in reason for x in ["fee", "commission", "reward", "interest", "staking", "airdrop"]):
                return False
            return True

    return False


def _sum_asset_flow(transactions: List[AssetTransaction]) -> Dict[str, Dict[str, float]]:
    """Podsumowanie przeplywow per waluta: przychody, rozchody, netto.
    Zwraca tylko waluty z niezerowym przeplywem. Pomija pending/niepotwierdzone."""
    result = {}
    skipped = 0
    for t in transactions:
        # Pomin pending
        if _is_pending_transaction(t):
            skipped += 1
            continue
        curr = t.currency.upper() if t.currency else "UNKNOWN"
        if curr not in result:
            result[curr] = {"in": 0.0, "out": 0.0, "count_in": 0, "count_out": 0}
        chg = _to_float(t.change)
        if chg > 1e-12:
            result[curr]["in"] += chg
            result[curr]["count_in"] += 1
        elif chg < -1e-12:
            result[curr]["out"] += abs(chg)
            result[curr]["count_out"] += 1

    if skipped > 0:
        print(f"    Pominięto {skipped} transakcji pending/niepotwierdzonych")

    # Usun waluty z zerowym przeplywem
    return {k: v for k, v in result.items() if v["in"] > 1e-12 or v["out"] > 1e-12}


class ReportGenerator:
    def __init__(self, output_path: str = "Raport_Analiza.docx"):
        self.output_path = output_path
        self.doc = Document()
        self._setup_styles()
        self._setup_margins()

    def _setup_margins(self):
        section = self.doc.sections[0]
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.0)
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)

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

    def _add_table(self, headers: List[str], rows: List[List[str]], max_rows: int = None):
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

        hdr_cells = table.rows[0].cells
        for i, h in enumerate(headers):
            hdr_cells[i].text = h
            for paragraph in hdr_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = Pt(10)
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            shading_elm = parse_xml(r'<w:shd {} w:fill="1F4E78"/>'.format(nsdecls('w')))
            hdr_cells[i]._tc.get_or_add_tcPr().append(shading_elm)

        # Jezeli max_rows jest None — pokazujemy wszystko
        limit = max_rows if max_rows is not None else len(rows)
        for row_data in rows[:limit]:
            row_cells = table.add_row().cells
            for i, cell_text in enumerate(row_data):
                txt = str(cell_text) if cell_text is not None else ""
                row_cells[i].text = txt[:250]
                for paragraph in row_cells[i].paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(9)

        if max_rows is not None and len(rows) > max_rows:
            self._add_paragraph(f" ... i {len(rows) - max_rows} wiecej wierszy")
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

    def _add_time_range_for_sheet(self, r: ExtractedIdentifiers, sheet_name: str):
        if sheet_name in r.time_ranges:
            tr = r.time_ranges[sheet_name]
            self._add_paragraph(f"    Zakres czasowy: {tr['from']}  →  {tr['to']}", color=RGBColor(0x40, 0x40, 0x40))

    def _render_customer_info(self, r: ExtractedIdentifiers):
        if not r.customer_info_sections:
            return
        self._add_heading("Customer Information (Informacje o uzytkowniku)", level=3)
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
            self._add_table(["Pole", "Wartosc"], rows, max_rows=100)

    def _render_kyc(self, r: ExtractedIdentifiers):
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
                        cell.text = f"[!] Blad: {e}"
                else:
                    cell.text = "Brak pliku"
        else:
            self._add_paragraph("Brak osadzonych obrazkow w arkuszu KYC Documents.")

    def _render_assets_overview(self, r: ExtractedIdentifiers):
        self._add_heading("Assets Overview (Przeglad aktywow)", level=3)
        if r.estimate_total_btc:
            self._add_paragraph(f"Estimate Total Balance (BTC): {r.estimate_total_btc}", bold=True)
        if r.asset_balances:
            bal_rows = []
            for b in r.asset_balances:
                waluta = b.currency_code
                if b.currency_name and b.currency_name != "—":
                    waluta = f"{b.currency_code} ({b.currency_name})"
                bal_rows.append([
                    waluta,
                    b.all_positions,
                    b.btc_equivalent,
                    b.wallet_type,
                ])
            self._add_table(
                ["Waluta", "Saldo", "Wartosc BTC", "Portfel"],
                bal_rows, max_rows=100)
        else:
            self._add_paragraph("Brak danych o saldach walut (wszystkie salda to 0 lub arkusz nie zawiera tabeli).")

    def _render_asset_log(self, r: ExtractedIdentifiers, transactions: List[AssetTransaction], title: str, sheet_name: str):
        if not transactions:
            return
        self._add_heading(title, level=3)
        self._add_time_range_for_sheet(r, sheet_name)

        # Notka informacyjna
        self._add_paragraph(
            "Uwaga: Ponizszy log pokazuje ruchy srodkow (wpłaty, wypłaty, transfery, rozliczenia). "
            "Szczegoly kupna/sprzedazy (cena, kontrahent) znajduja sie w arkuszu 'Order History'. "
            "Szczegoly wpłat/wypłat (adresy, TXID) znajduja sie w 'Deposit/Withdrawal History'.",
            color=RGBColor(0x60, 0x60, 0x60))

        # Tylko potwierdzone transakcje
        confirmed = [t for t in transactions if not _is_pending_transaction(t)]
        pending_count = len(transactions) - len(confirmed)
        if pending_count > 0:
            self._add_paragraph(f"Pominięto {pending_count} transakcji pending/niepotwierdzonych.", color=RGBColor(0x80, 0x60, 0x00))

        # Podsumowanie per waluta
        flow = _sum_asset_flow(confirmed)

        def fmt(v):
            if abs(v) < 1e-12:
                return "0"
            s = f"{v:.10f}".rstrip("0").rstrip(".")
            if s.startswith("."):
                s = "0" + s
            return s
        def fmt_signed(v):
            if abs(v) < 1e-12:
                return "0"
            s = fmt(v)
            return f"+{s}" if v > 0 else f"-{s}"

        # Grupujemy transakcje per waluta
        txns_by_currency = {}
        for t in sorted(confirmed, key=lambda x: x.time or ""):
            curr = t.currency.upper() if t.currency else "UNKNOWN"
            if curr not in txns_by_currency:
                txns_by_currency[curr] = []
            txns_by_currency[curr].append(t)

        if not flow:
            self._add_paragraph(
                "Brak wykrytych przeplywow (kolumna Change/Amount moze byc pusta lub zawierac same zera).",
                color=RGBColor(0x80, 0x60, 0x00))
            return

        # Dla kazdej waluty: podsumowanie (1 wiersz) + szczegoly transakcji
        for curr in sorted(flow.keys()):
            data = flow[curr]
            netto = data["in"] - data["out"]

            # Podsumowanie waluty — 1 wiersz
            self._add_paragraph(f"Waluta {curr}:", bold=True)
            summary_row = [[
                fmt_signed(data["in"]),
                fmt_signed(-data["out"]),
                fmt_signed(netto),
                str(data["count_in"]),
                str(data["count_out"]),
            ]]
            self._add_table(
                ["Przychody", "Rozchody", "Netto", "L. przych.", "L. rozch."],
                summary_row)

            # Szczegoly transakcji dla tej waluty
            curr_txns = txns_by_currency.get(curr, [])
            if curr_txns:
                txn_rows = []
                for t in curr_txns:
                    chg = _to_float(t.change)
                    chg_str = fmt_signed(chg)
                    txn_rows.append([
                        t.time[:19] if t.time else "",
                        chg_str,
                        t.reason if t.reason else "—",
                        t.transaction_id[:20] if t.transaction_id else "—",
                    ])
                self._add_table(
                    ["Czas", "Zmiana", "Powod", "TxID"],
                    txn_rows)
            self._add_paragraph("")  # odstep miedzy walutami

    def generate(self, reports: List[ExtractedIdentifiers], file_map: Dict[str, str]):
        # ===== STRONA TYTULOWA =====
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
        run2 = subtitle.add_run("Raportow gieldowych kryptowalutowych")
        run2.font.size = Pt(18)

        self.doc.add_paragraph()
        date_p = self.doc.add_paragraph()
        date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_p.add_run(f"Data wygenerowania: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

        self.doc.add_page_break()

        # ===== SPIS TRESCI =====
        self._add_heading("Spis tresci", level=1)
        toc_items = ["1. Podsumowanie", "2. Szczegolowa analiza poszczegolnych plikow"]
        for i, r in enumerate(reports, 1):
            toc_items.append(f"  2.{i}. {r.source_file}")
        if len(reports) >= 2:
            toc_items.append("3. Porownanie miedzy raportami")
        toc_items.append("4. Pelna lista identyfikatorow")
        for item in toc_items:
            self.doc.add_paragraph(item, style='List Bullet')
        self.doc.add_page_break()

        # ===== 1. PODSUMOWANIE =====
        self._add_heading("1. Podsumowanie", level=1)
        self._add_paragraph(f"Liczba przeanalizowanych plikow: {len(reports)}")
        self._add_paragraph(f"Liczba gield: {len(set(file_map.values()))}")

        summary_rows = []
        for r in reports:
            all_from = []
            all_to = []
            for tr in r.time_ranges.values():
                all_from.append(tr.get("from", ""))
                all_to.append(tr.get("to", ""))
            time_summary = ""
            if all_from and all_to:
                time_summary = f"{min(all_from)} -> {max(all_to)}"
            summary_rows.append([
                r.source_file, r.exchange.upper(), str(len(r.parsed_sheets)),
                ", ".join(sorted(r.user_ids)) if r.user_ids else "(brak)",
                str(len(r.related_user_ids)), str(len(r.emails)),
                str(len(r.ips)), str(len(r.wallet_addresses)), str(len(r.txids)),
                r.estimate_total_btc if r.estimate_total_btc else "(brak)",
                str(len(r.asset_balances)), time_summary,
            ])
        self._add_table(
            ["Plik", "Gielda", "Arkusze", "ID wlasciciela", "ID powiazanych",
             "E-maile", "IP", "Portfele", "TXID", "Est. Total BTC", "Salda", "Zakres czasowy"],
            summary_rows)
        self.doc.add_page_break()

        # ===== 2. SZCZEGOLY KAZDEGO PLIKU =====
        self._add_heading("2. Szczegolowa analiza poszczegolnych plikow", level=1)

        for idx, r in enumerate(reports, 1):
            self._add_heading(f"2.{idx}. {r.source_file}", level=2)
            self._add_paragraph(f"Gielda: {r.exchange.upper()}")
            # ID uzytkownika — wyraznie na poczatku
            if r.user_ids:
                uid_str = ", ".join(sorted(r.user_ids))
                self._add_paragraph(f"ID uzytkownika (wlasciciel konta): {uid_str}", bold=True, color=RGBColor(0x00, 0x00, 0x80))
            self._add_paragraph(
                f"Przeanalizowane arkusze ({len(r.parsed_sheets)}): {', '.join(r.parsed_sheets)}")
            if r.unknown_sheets:
                self._add_paragraph(
                    f"Nieznane arkusze: {', '.join(r.unknown_sheets)}",
                    color=RGBColor(0xC0, 0x00, 0x00))

            # GLOBALNY ZAKRES CZASOWY KONTA
            all_from = []
            all_to = []
            for tr in r.time_ranges.values():
                all_from.append(tr.get("from", ""))
                all_to.append(tr.get("to", ""))
            if all_from and all_to:
                self._add_paragraph(
                    f"Globalny zakres czasowy konta: {min(all_from)}  →  {max(all_to)}",
                    bold=True, color=RGBColor(0x00, 0x40, 0x80))
                self._add_paragraph("")

            # Renderujemy w kolejnosci parsed_sheets
            for sheet_name in r.parsed_sheets:
                if sheet_name == "Customer Information":
                    self._render_customer_info(r)
                elif sheet_name == "KYC Documents":
                    self._render_kyc(r)
                elif sheet_name == "Assets Overview":
                    self._render_assets_overview(r)
                elif sheet_name == "Spot Asset Log":
                    self._render_asset_log(r, r.spot_transactions,
                        "Spot Asset Log (Historia ruchow Spot)", "Spot Asset Log")
                elif sheet_name == "Funding Asset Log":
                    self._render_asset_log(r, r.funding_transactions,
                        "Funding Asset Log (Historia ruchow Funding)", "Funding Asset Log")
                # Inne arkusze beda dodawane pozniej

            self.doc.add_page_break()

        # ===== 3. POROWNANIE =====
        if len(reports) >= 2:
            self._add_heading("3. Porownanie miedzy raportami", level=1)
            comp = ReportComparator(reports)
            result = comp.compare()
            common = result.get("common", {})
            if common:
                self._add_heading("Wspolne identyfikatory (potencjalne powiazania):", level=2)
                self._add_paragraph(
                    "Ponizsze identyfikatory wystepuja w co najmniej dwoch raportach. "
                    "Moga wskazywac na powiazanie miedzy kontami.",
                    color=RGBColor(0xC0, 0x00, 0x00))
                for field_name, entries in common.items():
                    self._add_heading(f"[{field_name}] — {len(entries)} wspolnych", level=3)
                    rows = []
                    for i, e in enumerate(entries):
                        files_str = ", ".join(e["files"])
                        time_str = ""
                        if "time_context" in e:
                            parts = []
                            for tc in e["time_context"]:
                                if tc["ranges"]:
                                    parts.append(f"{tc['file']}: {tc['ranges'][0]}")
                            time_str = "; ".join(parts)
                        rows.append([str(i+1), str(e["value"]), files_str, time_str])
                    self._add_table(["Lp.", "Wartosc", "Pliki", "Zakres czasowy"], rows, max_rows=30)
            else:
                self._add_paragraph("Nie znaleziono wspolnych identyfikatorow miedzy raportami.")
            self.doc.add_page_break()

        # ===== 4. PELNA LISTA IDENTYFIKATOROW =====
        self._add_heading("4. Pelna lista identyfikatorow", level=1)
        self._add_paragraph("Szczegolowe dane w formacie JSON zostaly zapisane w pliku parsed_report.json.")

        for r in reports:
            self._add_heading(f"{r.source_file}", level=2)
            if r.user_ids:
                self._add_paragraph(f"ID wlasciciela konta ({len(r.user_ids)}):", bold=True)
                rows = [[str(i+1), str(item)] for i, item in enumerate(sorted(r.user_ids))]
                self._add_table(["Lp.", "Wartosc"], rows, max_rows=20)
            if r.related_user_ids:
                self._add_paragraph(f"ID powiazanych uzytkownikow ({len(r.related_user_ids)}):", bold=True)
                rows = [[str(i+1), str(item)] for i, item in enumerate(sorted(r.related_user_ids))]
                self._add_table(["Lp.", "Wartosc"], rows, max_rows=20)

            id_sections = [
                ("E-maile", r.emails), ("Numery telefonow", r.phones),
                ("Adresy IP", r.ips), ("Adresy portfeli (krypto)", r.wallet_addresses),
                ("TXID (hash transakcji)", r.txids), ("BIN karty", r.card_bins),
                ("Ostatnie 4 cyfry karty", r.card_last4), ("IBAN", r.ibans),
                ("Numery kont", r.account_numbers), ("ID urzadzen", r.device_ids),
                ("ID Fvideo", r.fvideo_ids), ("UUID BNC", r.bnc_uuids),
                ("ID zamowien", r.order_ids), ("ID kontrahentow", r.counterparty_ids),
                ("ID transakcji", r.transaction_ids), ("Imiona i nazwiska", r.names),
                ("Narodowosci", r.nationalities), ("Numery dokumentow tozsamosci", r.id_numbers),
                ("Lokalizacje GEO", r.geolocations), ("Przegladarki / User Agent", r.browsers),
            ]
            for title, items in id_sections:
                if items:
                    self._add_paragraph(f"{title} ({len(items)}):", bold=True)
                    rows = [[str(i+1), str(item)] for i, item in enumerate(sorted(items))]
                    self._add_table(["Lp.", "Wartosc"], rows, max_rows=20)

        self.doc.save(self.output_path)
