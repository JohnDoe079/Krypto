# Krypto v1.4.7

**Parser i komparator raportów giełdowych kryptowalutowych**

Automatyczne skanowanie, parsowanie i analiza raportów użytkownika z giełd **Binance** oraz **HTX** (Huobi) w formacie `.xlsx`. Narzędzie ekstrahuje identyfikatory, salda, transakcje i generuje szczegółowy raport porównawczy w formacie DOCX oraz JSON.

---

## 📁 Struktura repozytorium

```
Krypto/
├── main.py                 # Główny skrypt uruchamiający analizę
├── config.py               # Konfiguracja arkuszy i katalogów danych
├── matcher.py              # Moduł porównujący identyfikatory między raportami
├── reporter.py             # Generator raportów DOCX
├── version.py              # Wersja projektu (v1.4.7)
├── requirements.txt        # Zależności Pythona
├── CHANGELOG.md            # Historia zmian
├── models/
│   └── schemas.py          # Modele danych i funkcje pomocnicze ekstrakcji
└── parsers/
    ├── __init__.py         # Eksport parserów
    ├── binance_parser.py   # Parser raportów Binance (.xlsx)
    └── htx_parser.py       # Parser raportów HTX (.xlsx)
```

---

## ⚙️ Wymagania

- Python 3.9+
- Zależności (z `requirements.txt`):
  - `pandas` – przetwarzanie danych z Excela
  - `openpyxl` – obsługa plików `.xlsx`
  - `python-docx` – generowanie raportów Word

Instalacja zależności:
```bash
pip install -r requirements.txt
```

---

## 🚀 Użycie

### 1. Przygotuj dane

Umieść pliki `.xlsx` z raportami giełdowymi w odpowiednich katalogach:
```
data/
├── binance/          # Raporty Binance
└── htx/              # Raporty HTX
```

### 2. Uruchom analizę

```bash
# Pełna analiza + raport DOCX
python main.py

# Tylko JSON, bez generowania DOCX
python main.py --no-docx

# Szczegółowe logi
python main.py -v
```

### 3. Wyjście

| Plik | Opis |
|------|------|
| `parsed_report.json` | Zestawienie sparsowanych danych ze wszystkich plików |
| `parsed_report_comparison.json` | Wyniki porównania identyfikatorów między raportami |
| `Raport_Analiza.docx` | Sformatowany raport Word z analizą |

---

## 📊 Funkcjonalność

### Parsowane arkusze (Binance)
- **Customer Information** – dane użytkownika, KYC, kontakt
- **KYC Documents** – dokumenty tożsamości (z wyciąganiem obrazków)
- **Assets Overview** – salda walutowe per portfel (Spot, Funding, Futures, Earn, Margin, Pool)
- **Spot Asset Log** – historia ruchów na portfelu Spot
- **Funding Asset Log** – historia ruchów na portfelu Funding
- **Deposit History / Withdrawal History** – wpłaty i wypłaty (adresy portfeli, TXID) + sumowanie do bilansu
- **Fiat Deposit / Fiat Trades** – dane fiat (karty, IBAN, e-maile)
- **Binance Pay / P2P / OTC Trading** – transakcje peer-to-peer
- **Access Logs / Approved Devices** – logi IP, urządzenia, geolokalizacja
- **Order History** – historia zleceń

### Parsowane arkusze (HTX)
- **register_1** – dane rejestracyjne / KYC HTX:
  - UID użytkownika, imię i nazwisko, e-mail, telefon
  - Numer dokumentu (idcard), kraj rejestracji, data utworzenia konta
  - **Adres portfela (user_address)** — wykazywany per UID w raporcie
  - Dane płatnicze: bankcard, alipay, wechat
- **balance_1** – salda walutowe per UID:
  - Waluta i saldo w jednej komórce (po enterach)
  - Źródło arkusza oznaczone w kolumnie "Źródło"

### Ekstrahowane identyfikatory
- ID użytkownika (właściciel i powiązani)
- E-maile, numery telefonów
- Adresy IP, geolokalizacje, przeglądarki
- **Adresy portfeli kryptowalutowych (per UID dla HTX)**
- TXID (hash transakcji blockchain)
- BIN karty, ostatnie 4 cyfry, IBAN, numery kont
- ID urządzeń, zamówień, kontrahentów
- Imiona, nazwiska, narodowości, numery dokumentów

### Bilans przepływów
- **Per arkusz**: Spot, Funding, Deposit, Withdrawal — każdy osobno z podsumowaniem przychodów/rozchodów
- **Łączny bilans**: automatyczne zsumowanie wszystkich źródeł per waluta + porównanie z Assets Overview
- Wykrywanie ujemnych sald i transakcji pending

### Porównanie raportów
- Automatyczne wykrywanie **wspólnych identyfikatorów** między raportami (potencjalne powiązania)
- Lista **unikalnych identyfikatorów** per plik
- Kontekst czasowy dla wspólnych danych

### Raport DOCX
- Strona tytułowa
- Spis treści
- Podsumowanie wszystkich plików (tabela z kolumną Źródło dla HTX)
- Szczegółowa analiza każdego raportu:
  - Dane KYC i basic info (w tym HTX Register z saldami szczegółowymi)
  - Salda walutowe z podziałem na portfele
  - Logi Spot/Funding/Deposit/Withdrawal z podsumowaniem przepływów per waluta
  - **Pełny bilans łączony** — suma ze wszystkich źródeł z porównaniem do Assets Overview
  - Wykrywanie ujemnych sald i transakcji pending
- Porównanie między raportami
- Pełna lista identyfikatorów (w tym **adresy portfeli per UID** dla HTX)

---

## 🛠️ Parametry CLI

```
python main.py [-h] [-o OUTPUT] [-r REPORT] [--no-docx] [-v]

  -o, --output     Plik wyjściowy JSON (domyślnie: parsed_report.json)
  -r, --report     Plik raportu DOCX (domyślnie: Raport_Analiza.docx)
  --no-docx        Nie generuj raportu DOCX
  -v, --verbose    Szczegółowe logi w konsoli
```

---

## 📝 Uwagi

- Parser **HTX** obsługuje arkusze **register_1** (dane rejestracyjne / KYC) oraz **balance_1** (salda walutowe). Kolejne arkusze HTX będą dodawane iteracyjnie.
- Transakcje oznaczone jako *pending* / *processing* / *initiated* są pomijane w podsumowaniu przepływów, ale widoczne w logach.
- Status **"Completed"** w Deposit/Withdrawal History jest traktowany jako potwierdzony (nie pending).
- Obrazki z arkusza **KYC Documents** są zapisywane do `output/images/`.
- Różnica między pełnym bilansem przepływów a saldem z **Assets Overview** może wynikać z transferów między portfelami (Spot ↔ Funding) lub środków w innych produktach (Futures, Earn, Margin, Pool).

---

## 👤 Autor

**JohnDoe079**

---

*Repozytorium: [github.com/JohnDoe079/Krypto](https://github.com/JohnDoe079/Krypto)*
