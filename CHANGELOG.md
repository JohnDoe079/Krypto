# Changelog

Wszystkie istotne zmiany w projekcie Krypto.

---

## [1.4.8] – 2026-08-31

### Dodano (HTX)
- **Obsługa arkusza `login_1`** — historia logowań per UID:
  - Parsowanie kolumn: `uid`, `login_time`, `login_terminal`, `ip`
  - Wykrywanie zakresu czasowego logowań per arkusz
  - Zapis IP do globalnego zbioru `ips`
- **Wyświetlanie logowań w profilu użytkownika HTX**:
  - Tabela szczegółowa logowań (max 50 wierszy): Czas, Terminal, IP, Źródło
  - Tabela unikalnych IP z zakresem czasowym: Adres IP, Liczba logowań, Pierwsze/Ostatnie logowanie, Źródło
- **Wykrywanie współdzielonych IP** w sekcji porównań (sekcja 3 raportu DOCX):
  - Osobna tabelka pokazująca adresy IP używane przez więcej niż jednego użytkownika
  - Wskazanie pliku i UID dla każdego wspólnego IP
  - Zakres czasowy wspólnego użycia IP

### Zmieniono
- `config.py` — dodano `login_1` do `HTX_SHEETS`
- `schemas.py` — dodano pole `login_records: Dict[str, List[Dict]]` per UID
- `matcher.py` — dodano metodę `_compare_htx_login_ips()` wykrywającą wspólne IP w logowaniach

---

## [1.4.7] – 2026-08-31

### Zmieniono (HTX / Reporter)
- **Salda szczegółowe w jednej komórce** — zamiast rozbijania sald `balance_1` na osobne wiersze, wszystkie salda wyświetlane są w jednej komórce pod etykietą **"Salda szczegółowe"**, każde w nowej linii.
- **Usunięto dublowanie sald** — wiersze `Saldo <waluta>` zniknęły z tabeli profilu HTX.
- **Kolumna "Źródło" we wszystkich tabelach HTX** — każda tabela w sekcji profilu HTX ma teraz dodatkową kolumnę z nazwą arkusza źródłowego.
- **Usunięto zduplikowaną definicję** `_render_htx_logins()` z `reporter.py`.
- **Zaktualizowano README** — `balance_1` dodano do listy obsługiwanych arkuszy HTX.

---

## [1.4.1] – 2026-08-31 (hotfix)

### Naprawiono
- **Salda z `balance_1` nie były wyświetlane** — metoda `_parse_balance_1` istniała, ale nie była wywoływana w `_parse_sheet`. Dodano wywołanie.

---

## [1.4.0] – 2026-08-31

### Dodano (HTX)
- **Parsowanie arkusza `balance_1`** — salda per waluta per UID.

---

## [1.1.0] – baseline

### Dodano
- Parser raportów **Binance** (16 arkuszy).
- Parser raportów **HTX** — arkusz `register_1`.
- Ekstrakcja identyfikatorów, bilans przepływów, porównanie między raportami.
- Generator raportu DOCX i wyjście JSON.
