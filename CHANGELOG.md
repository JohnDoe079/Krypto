# Changelog

Wszystkie istotne zmiany w projekcie Krypto.

---

## [1.4.12] – 2026-08-31

### Poprawiono (formatowanie DOCX)
- **Odstęp przed "Unikalne adresy IP z zakresem czasowym:"** — dodano pusty akapit przed nagłówkiem, aby był odstęp od poprzedniej tabeli "Historia logowania". Nagłówek i tabela unikalnych IP pozostają razem (bez odstępu między nimi).
- Zasada globalna: nagłówek sekcji/tabeli ma być bezpośrednio nad tabelą (bez odstępu), ale ma być odstęp od poprzedniej treści/tabeli.

---

## [1.4.11] – 2026-08-31

### Poprawiono (formatowanie DOCX)
- **Usunięto odstęp przed "Unikalne adresy IP z zakresem czasowym:"** — nagłówek jest teraz bezpośrednio nad tabelą.

---

## [1.4.10] – 2026-08-31

### Poprawiono (formatowanie DOCX)
- **Odstępy przed wszystkimi nagłówkami tabel** — dodano pusty akapit przed każdym nagłówkiem sekcji/tabeli w raporcie DOCX.

---

## [1.4.9] – 2026-08-31

### Poprawiono (formatowanie DOCX)
- **Odstępy przed nagłówkami tabel** — dodano pusty akapit przed sekcjami portfeli i unikalnych IP.

---

## [1.4.8] – 2026-08-31

### Dodano (HTX)
- **Obsługa arkusza `login_1`** — historia logowań per UID.
- **Wykrywanie współdzielonych IP** w sekcji porównań.

---

## [1.4.7] – 2026-08-31

### Zmieniono (HTX / Reporter)
- **Salda szczegółowe w jednej komórce**.
- **Kolumna "Źródło" we wszystkich tabelach HTX**.

---

## [1.4.1] – 2026-08-31 (hotfix)

### Naprawiono
- **Salda z `balance_1` nie były wyświetlane**.

---

## [1.4.0] – 2026-08-31

### Dodano (HTX)
- **Parsowanie arkusza `balance_1`**.

---

## [1.1.0] – baseline

### Dodano
- Parser raportów **Binance** (16 arkuszy).
- Parser raportów **HTX** — arkusz `register_1`.
