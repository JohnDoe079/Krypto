# Changelog

Wszystkie istotne zmiany w projekcie Krypto.

---

## [1.2.2] – 2026-08-27

### Naprawiono (HTX)
- **Tabela podsumowania** — dla HTX (wiele użytkowników per plik) każdy UID ma osobny wiersz w tabeli podsumowania zamiast grupowania wszystkich w jednym.
- **Numer dokumentu bez .0** — parser HTX usuwa suffix `.0` z wartości numerycznych pochodzących z Excela (np. `35334118.0` → `35334118`).
- **Struktura karty użytkownika** — podział na 3 osobne części:
  1. Tabela danych osobowych (bez portfeli)
  2. Zdjęcia certyfikowane (tabela z obrazkami)
  3. Osobna tabela z adresami portfeli (na końcu)

---

## [1.2.1] – 2026-08-27

### Naprawiono (HTX)
- **Dopasowanie folderów ze zdjęciami** — folder ze zdjęciami może mieć nazwę będącą prefiksem UID (krótszą o 1+ znaków, np. UID `448780476` → folder `44878047`).
- **Portfele w jednej tabeli** — adresy portfeli wyświetlane w tej samej tabeli co dane użytkownika.

---

## [1.2.0] – 2026-08-27

### Dodano (HTX)
- **Skanowanie zdjęć certyfikowanych** — parser HTX automatycznie wykrywa katalog `certified_photos/` obok pliku `.xlsx` i przypisuje zdjęcia do UID użytkowników.
- **Pole `certified_photos`** w `ExtractedIdentifiers`.
- **Osadzanie zdjęć w raporcie DOCX**.

### Zmieniono (HTX / Reporter)
- **Nowa struktura raportu DOCX dla HTX** — karty profilowe per UID.

---

## [1.1.0] – baseline

### Dodano
- Parser raportów **Binance** (16 arkuszy).
- Parser raportów **HTX** — arkusz `register_1`.
- Ekstrakcja identyfikatorów, bilans przepływów, porównanie między raportami.
- Generator raportu DOCX i wyjście JSON.
