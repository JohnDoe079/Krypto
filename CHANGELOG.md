# Changelog

Wszystkie istotne zmiany w projekcie Krypto.

---

## [1.4.7] – 2026-08-31

### Zmieniono (HTX / Reporter)
- **Salda szczegółowe w jednej komórce** — zamiast rozbijania sald `balance_1` na osobne wiersze (`Saldo btc`, `Saldo usdt` itp.), wszystkie salda wyświetlane są w jednej komórce pod etykietą **"Salda szczegółowe"**, każde w nowej linii (`\n`).
- **Usunięto dublowanie sald** — wiersze `Saldo <waluta>` zniknęły z tabeli profilu HTX.
- **Kolumna "Źródło" we wszystkich tabelach HTX** — każda tabela w sekcji profilu HTX ma teraz dodatkową kolumnę z nazwą arkusza źródłowego (`register_1`, `balance_1`, `certified_photos`).
- **Usunięto zduplikowaną definicję** `_render_htx_logins()` z `reporter.py`.
- **Zaktualizowano README** — `balance_1` dodano do listy obsługiwanych arkuszy HTX.

---

## [1.4.1] – 2026-08-31 (hotfix)

### Naprawiono
- **Salda z `balance_1` nie były wyświetlane** — metoda `_parse_balance_1` istniała, ale nie była wywoływana w `_parse_sheet`. Dodano wywołanie, salda teraz poprawnie mergują się z profilem użytkownika.
- **Uwaga o `__pycache__`** — folder `__pycache__` generuje się automatycznie przy pierwszym uruchomieniu Pythona, nie jest potrzebny w paczce ZIP.

---

## [1.4.0] – 2026-08-31

### Dodano (HTX)
- **Parsowanie arkusza `balance_1`** — salda per waluta per UID.
- **Salda wyświetlane w tabeli użytkownika** — po danych osobowych pojawia się wiersz z listą wszystkich walut i ich sald.

---

## [1.1.0] – baseline

### Dodano
- Parser raportów **Binance** (16 arkuszy).
- Parser raportów **HTX** — arkusz `register_1`.
- Ekstrakcja identyfikatorów, bilans przepływów, porównanie między raportami.
- Generator raportu DOCX i wyjście JSON.
