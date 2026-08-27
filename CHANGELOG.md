# Changelog

Wszystkie istotne zmiany w projekcie Krypto.

---

## [1.2.1] – 2026-08-27

### Naprawiono (HTX)
- **Dopasowanie folderów ze zdjęciami** — folder ze zdjęciami może mieć nazwę będącą prefiksem UID (krótszą o 1+ znaków, np. UID `448780476` → folder `44878047`). Parser szuka najdłuższego pasującego prefiksu.
- **Portfele w jednej tabeli** — adresy portfeli wyświetlane są w tej samej tabeli co dane użytkownika (zamiast osobnej tabeli pod danymi).

---

## [1.2.0] – 2026-08-27

### Dodano (HTX)
- **Skanowanie zdjęć certyfikowanych** — parser HTX automatycznie wykrywa katalog `certified_photos/` obok pliku `.xlsx` i przypisuje zdjęcia do UID użytkowników.
- **Pole `certified_photos`** w `ExtractedIdentifiers` — mapa `UID → lista ścieżek do zdjęć`.
- **Osadzanie zdjęć w raporcie DOCX** — zdjęcia certyfikowane wyświetlane są w tabeli (max 3 kolumny) bezpośrednio pod danymi użytkownika.

### Zmieniono (HTX / Reporter)
- **Nowa struktura raportu DOCX dla HTX** — karty profilowe per UID: dane → zdjęcia → portfele → transakcje.
- **Sekcja 4 (Pełna lista identyfikatorów)** rozszerzona o tabelę ze zdjęciami certyfikowanymi.

### Naprawiono
- `wallet_addresses_by_user` inicjalizowany również w parserze HTX.
- `certified_photos` inicjalizowane pustą listą dla każdego UID z `register_1`.

---

## [1.1.0] – baseline

### Dodano
- Parser raportów **Binance** (16 arkuszy).
- Parser raportów **HTX** — arkusz `register_1`.
- Ekstrakcja identyfikatorów, bilans przepływów, porównanie między raportami.
- Generator raportu DOCX i wyjście JSON.
