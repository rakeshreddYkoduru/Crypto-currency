# Crypto Coin Comparison Dashboard

## Project Overview
An interactive Excel dashboard that allows investors to compare two cryptocurrency coins side-by-side. Enter any two coin names to instantly view and compare their Symbol, Price, Volume, Market Capital, and Circulating Supply — along with three KPI cards showing the difference between both coins.

---

## Project Structure

```
crypto_comparison_project/
│
├── Crypto_Comparison_Dashboard.xlsx   ← Main deliverable (open this)
├── crypto_dataset.xlsx                ← Source data (200 coins)
├── build_crypto_comparison.py         ← Python script that built the dashboard
└── README.md                          ← This file
```

---

## How to Use the Dashboard

1. **Open** `Crypto_Comparison_Dashboard.xlsx`
2. Go to the **Crypto Comparison** sheet (first tab)
3. **Type a coin name** in cell `C7` (Coin Name 1) — e.g. `Bitcoin`
4. **Type a coin name** in cell `E7` (Coin Name 2) — e.g. `Ethereum`
5. All fields and KPIs update **automatically**

---

## Features

### Input Fields
| Cell | Purpose         |
|------|-----------------|
| C7   | Coin Name 1 input |
| E7   | Coin Name 2 input |

### Data Validation Rules (applied to C7 & E7)
- Coin name length must be **greater than 2** characters
- Coin name length must be **less than 11** characters
- **No digits** (0–9) allowed in the name
- Shows an error popup if the rule is violated
- Shows an input prompt to guide the investor

### Auto-Populated Specification Table
| Field              | Source Column in Data Sheet |
|--------------------|-----------------------------|
| Symbol             | Column B                    |
| Price (USD)        | Column C                    |
| Volume (24h)       | Column P                    |
| Market Capital     | Column N                    |
| Circulating Supply | Column O                    |

All values are pulled via `VLOOKUP` from the **Crypto Data** sheet.

### KPI Cards (Difference = Coin 1 − Coin 2)
| KPI                          | Format      |
|------------------------------|-------------|
| 📦 Volume Difference          | USD ($)     |
| 🔄 Circulating Supply Difference | Whole number |
| 💰 Market Capital Difference  | USD ($)     |

---

## Rebuilding the Dashboard

If you want to regenerate the dashboard from scratch using the Python script:

### Requirements
```
pip install openpyxl pandas
```

### Run
```bash
python build_crypto_comparison.py
```

This will overwrite `Crypto_Comparison_Dashboard.xlsx` with a freshly built file.

---

## Dataset
- **Source file:** `crypto_dataset.xlsx`
- **Sheet:** `Crypto Data`
- **Coins:** 200 cryptocurrencies
- **Columns:** Coin Name, Symbol, Price (USD), 1h/24h/7d Change, Market Cap, Circulating Supply, Volume (24h), and more

---

## Notes
- If a coin name is not found in the dataset, cells show `Not Found`
- The **Crypto Data** sheet can be refreshed/replaced with updated data — formulas will auto-adjust
- Freeze panes are set at row 8 so headers stay visible while scrolling
