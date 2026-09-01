# Changelog

Wszystkie istotne zmiany w projekcie Krypto.

---

## [1.4.13] – 2026-08-31

### Dodano (HTX)
- **Obsługa arkuszy `trade_*_1`** — transakcje handlowe HTX:
  - Dynamiczne wykrywanie arkuszy: wzór `trade_*_1` (np. `trade_2024-04_1`, `trade_2024-05_1`)
  - Parsowanie kolumn: `order_type`, `price`, `volume`, `amount`, `uid`, `order_id`, `created_time`, `symbol`, `order_side`
  - Inteligentne rozpoznawanie par walutowych z `symbol` (np. `btcusdt` → base=BTC, quote=USDT)
  - Znane quote currencies: USDT, USDC, BUSD, HUSD, BTC, ETH, TRX
- **Wyświetlanie transakcji handlowych w profilu użytkownika HTX**:
  - Tabela szczegółowa (max 50 wierszy): Czas, Symbol, Strona, Typ, Cena, Wolumen, Wartość, Order ID, Źródło
  - Podsumowanie per waluta: Kupno/Sprzedaż/Netto (volume + amount), liczba transakcji
  - Sortowanie po czasie

---

## [1.4.12] – 2026-08-31

### Poprawiono (formatowanie DOCX)
- **Ujednolicenie nagłówków tabel** — wszystkie nagłówki tabel zmienione z `_add_paragraph(..., bold=True)` na `_add_heading(..., level=4)`.
- Nagłówek `level=4` ma wbudowany odstęp przed (`space_before`), dzięki czemu każda sekcja tabeli ma odstęp od poprzedniej treści, ale nagłówek jest bezpośrednio nad tabelą (bez pustego wiersza między).

---

## [1.4.11] – 2026-08-31

### Poprawiono (formatowanie DOCX)
- **Usunięto odstęp przed "Unikalne adresy IP z zakresem czasowym:"**.

---

## [1.4.10] – 2026-08-31

### Poprawiono (formatowanie DOCX)
- **Odstępy przed wszystkimi nagłówkami tabel**.

---

## [1.4.9] – 2026-08-31

### Poprawiono (formatowanie DOCX)
- **Odstępy przed nagłówkami tabel**.

---

## [1.4.8] – 2026-08-31

### Dodano (HTX)
- **Obsługa arkusza `login_1`**.
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
