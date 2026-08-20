# parsers/htx_parser.py
"""Parser raportów użytkownika HTX w formacie .xlsx."""

from pathlib import Path
from models.schemas import ExtractedIdentifiers


class HTXReportParser:
    """Placeholder — rozbudowa po dostarczeniu przykładowego pliku HTX."""

    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.identifiers = ExtractedIdentifiers(
            source_file=self.file_path.name,
            exchange="htx"
        )

    def parse_all(self) -> ExtractedIdentifiers:
        print("[!] Parser HTX nie jest jeszcze zaimplementowany.")
        print("    Dostarcz przykładowy plik .xlsx z HTX, aby dodać obsługę.")
        return self.identifiers
