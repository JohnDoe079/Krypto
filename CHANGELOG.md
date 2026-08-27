# Changelog

Wszystkie istotne zmiany w projekcie Krypto.

Format oparty na [Keep a Changelog](https://keepachangelog.com/pl/1.0.0/).

---

## [1.2.0] – 2026-08-27

### Dodano (HTX)
- **Skanowanie zdjęć certyfikowanych** — parser HTX automatycznie wykrywa katalog `certified_photos/` obok pliku `.xlsx` i przypisuje zdjęcia do UID użytkowników. Katalogi wewnątrz muszą mieć numeryczne nazwy = UID.
- **Pole `certified_photos`** w `ExtractedIdentifiers` — mapa `UID → lista ścieżek do zdjęć`.
- **Osadzanie zdjęć w raporcie DOCX** — zdjęcia certyfikowane wyświetlane są w tabeli (max 3 kolumny) bezpośrednio pod danymi użytkownika.

### Zmieniono (HTX / Reporter)
- **Nowa struktura raportu DOCX dla HTX** — zamiast płaskiej tabeli `register_1`, raport renderuje teraz **karty profilowe per UID**:
  1. Dane użytkownika (UID, data utworzenia, imię i nazwisko, email, telefon, numer dokumentu, kraj)
  2. Zdjęcia certyfikowane (osadzone w komórkach tabeli)
  3. Przypisane adresy portfeli dla tego UID
  4. Globalne podsumowanie portfeli
  5. Transakcje (gdy zostaną dodane kolejne arkusze HTX)
- **Sekcja 4 (Pełna lista identyfikatorów)** rozszerzona o tabelę ze zdjęciami certyfikowanymi (UID, nazwa pliku, ścieżka).

### Naprawiono
- `wallet_addresses_by_user` inicjalizowany również w parserze HTX (zapobiega KeyError przy braku UID).
- `certified_photos` inicjalizowane pustą listą dla każdego UID z `register_1` (nawet jeśli brak fizycznych plików).

---

## [1.1.0] – baseline

### Dodano
- Parser raportów **Binance** (16 arkuszy: Customer Info, KYC, Assets, Spot/Funding logs, Deposit/Withdrawal, Fiat, Pay, P2P, OTC, Access Logs, Order History, Devices).
- Parser raportów **HTX** — arkusz `register_1` (dane rejestracyjne / KYC).
- Ekstrakcja identyfikatorów: UID, email, telefon, IP, portfele, TXID, karty, IBAN, urządzenia, geolokalizacje.
- Bilans przepływów per waluta z wykrywaniem duplikatów fiat.
- Porównanie między raportami (wspólne / unikalne identyfikatory).
- Generator raportu DOCX ze spisem treści, tabelami i osadzonymi obrazkami KYC.
- Wyjście JSON: `parsed_report.json` + `parsed_report_comparison.json`.
