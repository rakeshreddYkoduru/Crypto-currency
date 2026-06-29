# Crypto Liquidity Dashboard — Project Folder

## Overview
An interactive Excel dashboard that helps investors analyze **cryptocurrency liquidity** (Volume 24h) across price range categories. A dropdown slicer lets investors filter by price range — the pie charts and breakdown tables show the **Top 5 coins by Volume** plus an **"Others"** segment for the full picture.

---

## Project Structure

```
crypto_liquidity_project/
│
├── Crypto_Liquidity_Dashboard.xlsx   ← Main deliverable (open this)
├── crypto_dataset.xlsx                ← Source data (200 coins)
├── build_liquidity_dashboard.py       ← Python script that built the dashboard
└── README.md                          ← This file
```

---

## How to Use the Dashboard

1. **Open** `Crypto_Liquidity_Dashboard.xlsx`
2. Go to the **Liquidity Dashboard** sheet (first tab, navy color)
3. **Use the dropdown at cell E6** to select a price range:
   - `All Coins` → shows top 5 by volume across all 200 coins
   - `0 to 50 (USD)` → filters to coins priced ≤ $50 (193 coins)
   - `Greater than 50 (USD)` → filters to coins priced > $50 (7 coins)
4. **Reference the 3 pie charts** to compare liquidity distribution visually
5. **Read the breakdown tables** below the charts for exact volume figures and percentage shares

---

## Dashboard Features

### 🔽 Price Range Slicer (Cell E6)
| Selection              | Coins Included | Description                          |
|------------------------|----------------|--------------------------------------|
| All Coins              | 200            | Entire dataset, all price ranges     |
| 0 to 50 (USD)          | 193            | Coins with price ≤ $50               |
| Greater than 50 (USD)  | 7              | Coins with price > $50               |

### 📊 Three Pie Charts
Each pie chart shows:
- **Top 5 coins** ranked by Volume (24h) in that category
- **Others** — combined volume of all remaining coins in the category
- **Percentage labels** on each slice for quick relative comparison

| Chart Title                       | Source Category           |
|-----------------------------------|---------------------------|
| 🌐 All Coins — Top 5 by Volume    | All 200 coins             |
| 💚 Price 0–50 USD — Top 5 Volume  | 0 to 50 (USD) category    |
| 🔵 Price >50 USD — Top 5 Volume   | Greater than 50 category  |

### 📋 Volume Breakdown Tables
Three tables (one per category) showing:
- Rank (1–5 + Others)
- Coin Name
- Volume (24h) in USD (Billions)
- Share % of total category volume

---

## Pre-computed Results

### All Coins — Top 5 by Volume (24h)
| Rank | Coin          | Volume (24h)    |
|------|---------------|-----------------|
| 1    | Tether        | $125,000,000,000 |
| 2    | Bitcoin       | $42,000,000,000  |
| 3    | Ethereum      | $18,000,000,000  |
| 4    | Tornado Cash  | $10,856,830,000  |
| 5    | Comtech Gold  | $9,600,629,000   |
| —    | Others        | $142,219,963,300 |

### 0 to 50 USD — Top 5 by Volume (24h)
| Rank | Coin          | Volume (24h)    |
|------|---------------|-----------------|
| 1    | Tether        | $125,000,000,000 |
| 2    | Tornado Cash  | $10,856,830,000  |
| 3    | USDC          | $9,000,000,000   |
| 4    | TrueUSD       | $8,409,770,000   |
| 5    | AVA (Travala) | $6,942,152,000   |
| —    | Others        | $104,210,778,896 |

### Greater than 50 USD — Top 5 by Volume (24h)
| Rank | Coin             | Volume (24h)   |
|------|------------------|----------------|
| 1    | Bitcoin          | $42,000,000,000 |
| 2    | Ethereum         | $18,000,000,000 |
| 3    | Comtech Gold     | $9,600,629,000  |
| 4    | DeFi Pulse Index | $5,582,320,000  |
| 5    | Solana           | $4,300,000,000  |
| —    | Others           | $3,774,942,188  |

---

## Sheet Structure

| Sheet Name           | Visibility | Purpose                                      |
|----------------------|------------|----------------------------------------------|
| Liquidity Dashboard  | Visible    | Main investor-facing dashboard               |
| Crypto Data          | Visible    | Source dataset (200 coins, 18 columns)       |
| Helper_TopN          | Hidden     | Pre-computed top-5+Others tables for charts  |

---

## Rebuild Instructions

### Requirements
```bash
pip install openpyxl pandas
```

### Run
```bash
python build_liquidity_dashboard.py
```

Regenerates `Crypto_Liquidity_Dashboard.xlsx` from `crypto_dataset.xlsx`.

---

## Key Metric: Volume (24h) as Liquidity Proxy
Volume (24h) measures how much of a coin was traded in the last 24 hours. High volume = high liquidity = easier for investors to buy/sell without large price impact.

- **High liquidity coins**: Tether, Bitcoin, Ethereum (volumes in billions)
- **Low liquidity coins**: Many niche altcoins (volumes in thousands)
