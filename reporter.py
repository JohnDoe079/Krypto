"""Generuje raport DOCX z wyników analizy raportów giełdowych."""

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

from models.schemas import ExtractedIdentifiers, AssetTransaction, _normalize_decimal
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
    """Konwertuje string na float, obsługuje przecinek dziesiętny."""
    try:
        normalized = _normalize_decimal(val)
        return float(normalized)
    except (ValueError, TypeError):
        return 0.0


def _is_pending_transaction(t: AssetTransaction) -> bool:
    """Sprawdza czy transakcja jest pending/niepotwierdzona (pomijamy w podsumowaniu)."""
    reason = str(t.reason).lower().strip() if t.reason else ""
    if not reason:
        return False

    ok_statuses = ["completed", "success", "filled", "confirmed"]
    if any(s in reason for s in ok_statuses):
        return False

    pending_kw = ["pending", "processing", "initiated", "created", "request", "order"]
    has_pending = any(kw in reason for kw in pending_kw)
    has_success = "success" in reason or "completed" in reason
    if has_pending and not has_success:
        return True

    success_required = ["withdrawal", "deposit", "transfer"]
    for op in success_required:
        if op in reason and not has_success:
            if any(x in reason for x in ["fee", "commission", "reward", "interest", "staking", "airdrop"]):
                return False
            return True

    return False


def _sum_asset_flow(transactions: List[AssetTransaction]) -> Dict[str, Dict[str, float]]:
    """Podsumowanie przepływów per waluta: przychody, rozchody, netto."""
    result = {}
    skipped = 0
    for t in transactions:
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
        print(f"  Pominięto {skipped} transakcji pending/niepotwierdzonych")

    return {k: v for k, v in result.items() if v["in"] > 1e-12 or v["out"] > 1e-12}


def _fmt(v: float) -> str:
    if abs(v) < 1e-12:
        return "0"
    s = f"{abs(v):.10f}".rstrip("0").rstrip(".")
    if s.startswith("."):
        s = "0" + s
    return s


def _fmt_signed(v: float) -> str:
    if abs(v) < 1e-12:
        return "0"
    s = _fmt(v)
    return f"+{s}" if v > 0 else f"-{s}"


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
            shading_elm = parse_xml(r'<w:shd {} w:fill="1F4E79"/>'.format(nsdecls('w')))
            hdr_cells[i]._tc.get_or_add_tcPr().append(shading_elm)

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

    def _add_time_range_for_sheet(self, r: ExtractedIdentifiers, sheet_name: str):
        if sheet_name in r.time_ranges:
            tr = r.time_ranges[sheet_name]
            self._add_paragraph(f"  Zakres czasowy: {tr['from']} → {tr['to']}", color=RGBColor(0x40, 0x40, 0x40))

    def _render_customer_info(self, r: ExtractedIdentifiers):
        if not r.customer_info_sections:
            return
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
                        cell.text = f"[!] Błąd: {e}"
                else:
                    cell.text = "Brak pliku"
        else:
            self._add_paragraph("Brak osadzonych obrazków w arkuszu KYC Documents.")

    def _render_assets_overview(self, r: ExtractedIdentifiers):
        self._add_heading("Assets Overview (Przegląd aktywów)", level=3)
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
                ["Waluta", "Saldo", "Wartość BTC", "Portfel"],
                bal_rows, max_rows=100)
        else:
            self._add_paragraph("Brak danych o saldach walut (wszystkie salda to 0 lub arkusz nie zawiera tabeli).")

    def _render_transaction_section(self, r: ExtractedIdentifiers, transactions: List[AssetTransaction],
                                     title: str, sheet_name: str, note: str = None):
        if not transactions:
            return
        self._add_heading(title, level=3)
        self._add_time_range_for_sheet(r, sheet_name)

        if note:
            self._add_paragraph(note, color=RGBColor(0x60, 0x60, 0x60))

        confirmed = [t for t in transactions if not _is_pending_transaction(t)]
        pending_count = len(transactions) - len(confirmed)
        if pending_count > 0:
            self._add_paragraph(f"Pominięto {pending_count} transakcji pending/niepotwierdzonych.",
                                color=RGBColor(0x80, 0x60, 0x00))

        flow = _sum_asset_flow(confirmed)

        txns_by_currency = {}
        for t in sorted(confirmed, key=lambda x: x.time or ""):
            curr = t.currency.upper() if t.currency else "UNKNOWN"
            if curr not in txns_by_currency:
                txns_by_currency[curr] = []
            txns_by_currency[curr].append(t)

        if not flow:
            self._add_paragraph(
                "Brak wykrytych przepływów (kolumna Change/Amount może być pusta lub zawierać same zera).",
                color=RGBColor(0x80, 0x60, 0x00))
            return

        self._add_paragraph(
            "Uwaga: Saldo poniżej to BILANS ZMIAN w okresie raportu (przychody minus rozchody). "
            "Jeżeli saldo jest ujemne, oznacza to że w okresie raportu rozchody przewyższyły przychody.",
            color=RGBColor(0x60, 0x60, 0x60))

        for curr in sorted(flow.keys()):
            data = flow[curr]
            netto = data["in"] - data["out"]
            curr_txns = txns_by_currency.get(curr, [])

            times = [t.time for t in curr_txns if t.time]
            time_range_str = ""
            if times:
                time_range_str = f" (zakres: {min(times)[:19]} → {max(times)[:19]})"

            self._add_paragraph(f"Waluta {curr}:{time_range_str}", bold=True)
            summary_row = [[
                f"+{_fmt(data['in'])}",
                f"-{_fmt(data['out'])}",
                _fmt_signed(netto),
                str(data["count_in"]),
                str(data["count_out"]),
            ]]
            self._add_table(
                ["Przychody", "Rozchody", "Saldo", "L. przych.", "L. rozch."],
                summary_row)

            asset_balance = None
            for b in r.asset_balances:
                if b.currency_code.upper() == curr:
                    asset_balance = b
                    break

            if asset_balance:
                bal_all = _to_float(asset_balance.all_positions)
                if abs(netto - bal_all) > 1e-8:
                    diff = bal_all - netto
                    self._add_paragraph(
                        f" ℹ️ Różnica między logiem ({_fmt_signed(netto)}) a Assets Overview ({_fmt_signed(bal_all)}): {_fmt_signed(diff)} {curr}. "
                        f"Brakujące transakcje (depozyty, wypłaty, transfery) znajdują się w innych arkuszach.",
                        color=RGBColor(0x00, 0x60, 0x80))

            if netto < -1e-12:
                self._add_paragraph(
                    f" ⚠️ UWAGA: Ujemne saldo w logu ({_fmt_signed(netto)}). "
                    f"Oznacza to że w okresie objętym tym logiem rozchody przewyższyły przychody. "
                    f"Nie oznacza to debetu — brakuje tu depozytów/wypłat z arkuszy Deposit/Withdrawal History.",
                    color=RGBColor(0xC0, 0x00, 0x00))

            nonzero_txns = [t for t in curr_txns if abs(_to_float(t.change)) > 1e-12]
            if nonzero_txns:
                txn_rows = []
                for t in nonzero_txns:
                    chg = _to_float(t.change)
                    chg_str = _fmt_signed(chg)
                    reason_display = t.reason if t.reason else "—"
                    if len(reason_display) > 60:
                        reason_display = reason_display[:57] + "..."
                    txn_rows.append([
                        t.time[:19] if t.time else "",
                        chg_str,
                        reason_display,
                        t.transaction_id[:20] if t.transaction_id else "—",
                    ])
                self._add_table(
                    ["Czas", "Zmiana", "Powód", "TxID"],
                    txn_rows)
            self._add_paragraph("")

    def _render_combined_flow(self, r: ExtractedIdentifiers):
        """Łączy wszystkie transakcje (Spot + Funding + Deposit + Withdrawal) i pokazuje pełny bilans per waluta."""
        all_txns = (r.spot_transactions + r.funding_transactions +
                    r.deposit_transactions + r.withdrawal_transactions)
        if not all_txns:
            return

        self._add_heading("Pełny bilans przepływów (wszystkie źródła)", level=3)
        self._add_paragraph(
            "Poniższa tabela łączy dane ze wszystkich arkuszy: Spot Asset Log, Funding Asset Log, "
            "Deposit History oraz Withdrawal History. Dla każdej waluty pokazuje całkowite przychody, rozchody i netto.",
            color=RGBColor(0x60, 0x60, 0x60))

        confirmed = [t for t in all_txns if not _is_pending_transaction(t)]
        flow = _sum_asset_flow(confirmed)

        if not flow:
            self._add_paragraph("Brak potwierdzonych przepływów.")
            return

        rows = []
        for curr in sorted(flow.keys()):
            data = flow[curr]
            netto = data["in"] - data["out"]

            asset_balance = None
            for b in r.asset_balances:
                if b.currency_code.upper() == curr:
                    asset_balance = b
                    break

            bal_str = asset_balance.all_positions if asset_balance else "(brak)"
            diff_str = ""
            if asset_balance:
                bal_all = _to_float(asset_balance.all_positions)
                diff = bal_all - netto
                diff_str = _fmt_signed(diff)

            rows.append([
                curr,
                f"+{_fmt(data['in'])}",
                f"-{_fmt(data['out'])}",
                _fmt_signed(netto),
                bal_str,
                diff_str,
            ])

        self._add_table(
            ["Waluta", "Przychody (wszystkie)", "Rozchody (wszystkie)", "Netto", "Saldo z Assets Overview", "Różnica"],
            rows)

        self._add_paragraph(
            "ℹ️ Kolumna 'Różnica' pokazuje różnicę między pełnym bilansem przepływów a saldem z Assets Overview. "
            "Jeżeli różnica jest niezerowa, oznacza to że część środków została przeniesiona między portfelami (np. Spot → Funding) "
            "lub znajduje się w innych produktach (Futures, Earn, Margin, Pool).",
            color=RGBColor(0x00, 0x60, 0x80))

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
        toc_items = ["1. Podsumowanie", "2. Szczegółowa analiza poszczególnych plików"]
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
            ["Plik", "Giełda", "Arkusze", "ID właściciela", "ID powiązanych",
             "E-maile", "IP", "Portfele", "TXID", "Est. Total BTC", "Salda", "Zakres czasowy"],
            summary_rows)
        self.doc.add_page_break()

        # ===== 2. SZCZEGÓŁY KAŻDEGO PLIKU =====
        self._add_heading("2. Szczegółowa analiza poszczególnych plików", level=1)

        for idx, r in enumerate(reports, 1):
            self._add_heading(f"2.{idx}. {r.source_file}", level=2)
            self._add_paragraph(f"Giełda: {r.exchange.upper()}")
            if r.user_ids:
                uid_str = ", ".join(sorted(r.user_ids))
                self._add_paragraph(f"ID użytkownika (właściciel konta): {uid_str}", bold=True, color=RGBColor(0x00, 0x00, 0x80))
            self._add_paragraph(
                f"Przeanalizowane arkusze ({len(r.parsed_sheets)}): {', '.join(r.parsed_sheets)}")
            if r.unknown_sheets:
                self._add_paragraph(
                    f"Nieznane arkusze: {', '.join(r.unknown_sheets)}",
                    color=RGBColor(0xC0, 0x00, 0x00))

            all_from = []
            all_to = []
            for tr in r.time_ranges.values():
                all_from.append(tr.get("from", ""))
                all_to.append(tr.get("to", ""))
            if all_from and all_to:
                self._add_paragraph(
                    f"Globalny zakres czasowy konta: {min(all_from)} → {max(all_to)}",
                    bold=True, color=RGBColor(0x00, 0x40, 0x80))
            self._add_paragraph("")

            for sheet_name in r.parsed_sheets:
                if sheet_name == "Customer Information":
                    self._render_customer_info(r)
                elif sheet_name == "KYC Documents":
                    self._render_kyc(r)
                elif sheet_name == "Assets Overview":
                    self._render_assets_overview(r)
                elif sheet_name == "Spot Asset Log":
                    self._render_transaction_section(r, r.spot_transactions,
                        "Spot Asset Log (Historia ruchów Spot)", "Spot Asset Log",
                        "Uwaga: Ponizszy log pokazuje ruchy środków (wpłaty, wypłaty, transfery, rozliczenia). "
                        "Szczegóły kupna/sprzedaży (cena, kontrahent) znajdują się w arkuszu 'Order History'. "
                        "Szczegóły wpłat/wypłat (adresy, TXID) znajdują się w 'Deposit/Withdrawal History'.")
                elif sheet_name == "Funding Asset Log":
                    self._render_transaction_section(r, r.funding_transactions,
                        "Funding Asset Log (Historia ruchów Funding)", "Funding Asset Log",
                        "Uwaga: Funding Wallet obsługuje P2P, Binance Pay, Binance Card i inne usługi.")
                elif sheet_name == "Deposit History":
                    self._render_transaction_section(r, r.deposit_transactions,
                        "Deposit History (Historia wpłat)", "Deposit History",
                        "Uwaga: Wpłaty z zewnętrznych adresów blockchain lub z innych kont Binance.")
                elif sheet_name == "Withdrawal History":
                    self._render_transaction_section(r, r.withdrawal_transactions,
                        "Withdrawal History (Historia wypłat)", "Withdrawal History",
                        "Uwaga: Wypłaty na zewnętrzne adresy blockchain lub do innych kont Binance.")

            self._render_combined_flow(r)
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
                    color=RGBColor(0xC0, 0x00, 0x00))
                for field_name, entries in common.items():
                    self._add_heading(f"[{field_name}] — {len(entries)} wspólnych", level=3)
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
                    self._add_table(["Lp.", "Wartość", "Pliki", "Zakres czasowy"], rows, max_rows=30)
            else:
                self._add_paragraph("Nie znaleziono wspólnych identyfikatorów między raportami.")
            self.doc.add_page_break()

        # ===== 4. PEŁNA LISTA IDENTYFIKATORÓW =====
        self._add_heading("4. Pełna lista identyfikatorów", level=1)
        self._add_paragraph("Szczegółowe dane w formacie JSON zostały zapisane w pliku parsed_report.json.")

        for r in reports:
            self._add_heading(f"{r.source_file}", level=2)
            if r.user_ids:
                self._add_paragraph(f"ID właściciela konta ({len(r.user_ids)}):", bold=True)
                rows = [[str(i+1), str(item)] for i, item in enumerate(sorted(r.user_ids))]
                self._add_table(["Lp.", "Wartość"], rows, max_rows=20)
            if r.related_user_ids:
                self._add_paragraph(f"ID powiązanych użytkowników ({len(r.related_user_ids)}):", bold=True)
                rows = [[str(i+1), str(item)] for i, item in enumerate(sorted(r.related_user_ids))]
                self._add_table(["Lp.", "Wartość"], rows, max_rows=20)

            id_sections = [
                ("E-maile", r.emails), ("Numery telefonów", r.phones),
                ("Adresy IP", r.ips), ("Adresy portfeli (krypto)", r.wallet_addresses),
                ("TXID (hash transakcji)", r.txids), ("BIN karty", r.card_bins),
                ("Ostatnie 4 cyfry karty", r.card_last4), ("IBAN", r.ibans),
                ("Numery kont", r.account_numbers), ("ID urządzeń", r.device_ids),
                ("ID Fvideo", r.fvideo_ids), ("UUID BNC", r.bnc_uuids),
                ("ID zamówień", r.order_ids), ("ID kontrahentów", r.counterparty_ids),
                ("ID transakcji", r.transaction_ids), ("Imiona i nazwiska", r.names),
                ("Narodowości", r.nationalities), ("Numery dokumentów tożsamości", r.id_numbers),
                ("Lokalizacje GEO", r.geolocations), ("Przeglądarki / User Agent", r.browsers),
            ]
            for title, items in id_sections:
                if items:
                    self._add_paragraph(f"{title} ({len(items)}):", bold=True)
                    rows = [[str(i+1), str(item)] for i, item in enumerate(sorted(items))]
                    self._add_table(["Lp.", "Wartość"], rows, max_rows=20)

        self.doc.save(self.output_path)
