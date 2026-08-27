# Changelog

Wszystkie istotne zmiany w projekcie Krypto.

---

## [1.3.4] – 2026-08-27 (hotfix)

### Zmieniono (HTX)
- **Wycofano parsowanie nowych arkuszy HTX** (trade_*, login_1, DeviceFP_1, balance_1, deposit&withdraw&transfer_1).
- Parser HTX przetwarza teraz **TYLKO `register_1`** — dane użytkownika, zdjęcia certyfikowane, adresy portfeli.
- Pozostałe arkusze HTX są pomijane (wyświetlany jest komunikat w logu).
- Raport DOCX renderuje tylko profil użytkownika HTX (bez sekcji logowania, urządzeń i transakcji).
- Pełne dane z pozostałych arkuszy będą dodane w przyszłych wersjach po dokładnej analizie struktury.

---

## [1.3.3] – 2026-08-27 (hotfix)

### Zmieniono (HTX / Reporter)
- **Przywrócone tabele transakcyjne w DOCX** — sekcje logowania, urządzeń i transakcji renderują teraz pełne tabele, ale z limitem **50 wierszy** per typ.
- Dla 1-3 użytkowników (standardowy przypadek) plik DOCX będzie czytelny i szybki.
- Dla plików testowych z wieloma użytkownikami (np. 31) — nadal chroni przed gigantycznym plikiem.
- Przy przekroczeniu limitu 50 wierszy wyświetlany jest komunikat z liczbą pominiętych rekordów.

---

## [1.3.2] – 2026-08-27 (hotfix)

### Naprawiono
- **Zawieszanie LibreOffice / gigantyczny plik DOCX** — sekcje HTX (loginy, urządzenia, transakcje) renderowane są teraz jako SKRÓT w DOCX:
  - Loginy: tylko unikalne IP + lokalizacje + przeglądarki (max 30 wierszy)
  - Urządzenia: tylko liczba + lista ID (max 20 wierszy)
  - Transakcje: tylko liczba per typ + bilans per waluta (bez pełnych tabel czasowych)
  - Pełne dane logowania, urządzeń i transakcji dostępne w pliku JSON.

---

## [1.3.1] – 2026-08-27 (hotfix)

### Naprawiono
- **Błąd składni w reporter.py** — usunięty zduplikowany, niepełny nagłówek tabeli podsumowania powodujący `SyntaxError` przy imporcie.

---

## [1.3.0] – 2026-08-27

### Dodano (HTX)
- **Podmapowanie wszystkich arkuszy HTX**:
  - `register_1` — dane rejestracyjne / KYC (już było)
  - `balance_1` — salda walut
  - `login_1` — historia logowania (IP, geolokalizacja, przeglądarka)
  - `DeviceFP_1` — urządzenia / odciski palców
  - `deposit&withdraw&transfer_1` — depozyty, wypłaty, transfery
  - `trade_YYYY-MM-DD_1` — transakcje handlowe (dynamiczne nazwy, parsowane automatycznie)
- **Generyczne wykrywanie kolumn** — parser sam rozpoznaje kolumny po nazwach (time, currency, amount, type, status, txid, address, ip, device_id itp.), więc działa nawet jeśli nazwy kolumn są nieco inne.
- **Nowe sekcje w raporcie DOCX dla HTX**:
  - Historia logowania (IP, lokalizacje, przeglądarki)
  - Urządzenia zatwierdzone (Device Fingerprints)
  - Transakcje HTX (depozyty, wypłaty, trade) z bilansem per waluta
- **Wyciszenie warningu openpyxl** — `Workbook contains no default style` nie pojawia się już w logach.

### Naprawiono (HTX)
- **Notacja naukowa salda** — wartości takie jak `5e-05` formatowane są teraz jako `0.00005`.
- **Tabela podsumowania per UID** — dla HTX każdy użytkownik ma osobny wiersz z własnymi danymi (imię, email, telefon, saldo, liczba portfeli).
- **Numer dokumentu bez .0** — parser HTX usuwa suffix `.0` z wartości numerycznych.

---

## [1.2.2] – 2026-08-27

### Naprawiono (HTX)
- Tabela podsumowania — osobny wiersz per UID.
- Numer dokumentu bez `.0`.
- Struktura karty użytkownika: 3 osobne części (dane → zdjęcia → portfele).

---

## [1.2.1] – 2026-08-27

### Naprawiono (HTX)
- Dopasowanie folderów ze zdjęciami (prefiks UID).
- Portfele w jednej tabeli.

---

## [1.2.0] – 2026-08-27

### Dodano (HTX)
- Skanowanie zdjęć certyfikowanych z katalogu `certified_photos/`.
- Pole `certified_photos` w `ExtractedIdentifiers`.
- Osadzanie zdjęć w raporcie DOCX.
- Nowa struktura raportu DOCX dla HTX — karty profilowe per UID.

---

## [1.1.0] – baseline

### Dodano
- Parser raportów **Binance** (16 arkuszy).
- Parser raportów **HTX** — arkusz `register_1`.
- Ekstrakcja identyfikatorów, bilans przepływów, porównanie między raportami.
- Generator raportu DOCX i wyjście JSON.
