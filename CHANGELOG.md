# Changelog

Wszystkie istotne zmiany w projekcie Krypto.

---

## [1.4.9] – 2026-08-31

### Poprawiono (formatowanie DOCX)
- **Odstępy przed nagłówkami tabel** — dodano pusty akapit przed sekcjami:
  - "Adresy portfeli kryptowalutowych:" (po zdjęciach certyfikowanych)
  - "Unikalne adresy IP z zakresem czasowym:" (po tabeli szczegółowej logowań)
- **Usunięto zbędny enter** między nagłówkiem "Zdjęcia certyfikowane" a tabelą zdjęć — nagłówek i tabela są teraz bezpośrednio po sobie.
- Zasada stosowana globalnie: przed każdym nagłówkiem sekcji/tabeli jest odstęp od poprzedniej treści, ale nie ma pustego akapitu między nagłówkiem a właściwą tabelą.

---

## [1.4.8] – 2026-08-31

### Dodano (HTX)
- **Obsługa arkusza `login_1`** — historia logowań per UID.
- **Wyświetlanie logowań w profilu użytkownika HTX** — szczegółowa tabela + unikalne IP z zakresem czasowym.
- **Wykrywanie współdzielonych IP** w sekcji porównań (sekcja 3 raportu DOCX).

---

## [1.4.7] – 2026-08-31

### Zmieniono (HTX / Reporter)
- **Salda szczegółowe w jednej komórce** — zamiast rozbijania sald `balance_1` na osobne wiersze.
- **Kolumna "Źródło" we wszystkich tabelach HTX**.
- **Usunięto zduplikowaną definicję** `_render_htx_logins()`.

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
