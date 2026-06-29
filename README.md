# Crypto Market Analysis Dashboard

## Project Overview

The **Crypto Market Analysis Dashboard** is an interactive Microsoft Excel project developed to help investors analyze cryptocurrency market data and make informed investment decisions. The project uses data from the top 200 cryptocurrencies and provides interactive dashboards, charts, KPIs, slicers, and VBA automation to visualize market trends, liquidity, and price movements.

This project was completed as part of a Data Analytics internship by implementing multiple analytical tasks on a single dataset.

---

## Dataset

**Source:** CoinMarketCap

**Dataset Size:** Top 200 Cryptocurrency Coins

### Dataset Attributes

* Coin Name
* Symbol
* Current Price
* 1 Hour Change (%)
* 24 Hour Change (%)
* 7 Day Change (%)
* Volume (24 Hours)
* Market Capitalization
* Circulating Supply

---

## Project Objectives

* Analyze cryptocurrency price movements.
* Compare previous and current prices.
* Measure liquidity using trading volume.
* Compare two cryptocurrencies using key metrics.
* Build interactive dashboards for investors.
* Implement VBA automation for dashboard visibility.
* Provide dynamic filtering using slicers and data validation.

---

# Dashboard Features

## Task 1 – Price Trend Analysis

* Filter cryptocurrencies priced between **$0 and $5**.
* Display the **Top 10 coins** based on previous 1-hour price.
* Compare:

  * Current Price
  * Previous 24-Hour Price
  * Previous 7-Day Price
* Interactive charts and slicers.

---

## Task 2 – Price Change Analysis

* Calculate previous one-hour prices.
* Calculate price increase.
* Display Top 10 coins based on price change.
* Price range slicer:

  * Price ≤ $10
  * Price ≥ $10
* Summary table showing coin symbols and price changes.

---

## Task 3 – Liquidity Dashboard

* Filter coins beginning with:

  * A, E, I, O, U
  * B
  * C
  * D
* Display Top 10 cryptocurrencies based on **24-Hour Volume**.
* VBA automation:

  * Dashboard visible only between **9:00 AM and 5:00 PM**.
  * Displays:

```
Please open in working hours (9 AM to 5 PM)
```

outside working hours.

---

## Task 4 – Cryptocurrency Comparison

Compare two cryptocurrencies by entering their names.

Displays:

* Symbol
* Price
* Volume
* Market Capitalization
* Circulating Supply

### Data Validation

* Coin name length greater than 2 characters.
* Coin name length less than 11 characters.
* Numbers are not allowed.

### KPI Cards

* Volume Difference
* Market Capital Difference
* Circulating Supply Difference

---

## Task 5 – Liquidity Distribution

* Price category slicer:

  * $0–50
  * Greater than $50
* Dynamic pie chart showing:

  * Top 5 cryptocurrencies by Volume (24 Hours)
  * Remaining coins grouped as **Others**

---

# Tools & Technologies

* Microsoft Excel
* Pivot Tables
* Pivot Charts
* Slicers
* Conditional Formatting
* Data Validation
* VBA (Visual Basic for Applications)
* Excel Formulas

---

# Repository Structure

```
Crypto-Market-Analysis/
│
├── README.md
├── Crypto_Market_Analysis.xlsx
│
├── dataset/
│   ├── crypto_market_data.xlsx
│   └── cleaned_crypto_market_data.xlsx
│
├── tasks/
│   ├── Task_1.md
│   ├── Task_2.md
│   ├── Task_3.md
│   ├── Task_4.md
│   └── Task_5.md
│
├── screenshots/
│   ├── Task_1/
│   ├── Task_2/
│   ├── Task_3/
│   ├── Task_4/
│   └── Task_5/
│
├── documentation/
│   ├── Project_Report.pdf
│   └── Dashboard_Guide.pdf
│
└── assets/
```

---

# Screenshots

The repository contains screenshots for each task, including:

* Dashboard Overview
* Interactive Charts
* Slicers
* KPI Cards
* Data Validation
* VBA Working Hours Window
* Pie Chart
* Coin Comparison Dashboard

---

# How to Use

1. Download the repository.
2. Open **Crypto_Market_Analysis.xlsx**.
3. Enable Macros when prompted.
4. Use slicers and filters to interact with the dashboards.
5. Enter valid cryptocurrency names in the comparison dashboard to compare two coins.
6. Review charts, KPIs, and liquidity insights.

---

# Key Features

* Interactive Excel Dashboards
* Dynamic Charts
* KPI Cards
* Pivot Tables
* Pivot Charts
* VBA Automation
* Data Validation
* Dynamic Slicers
* Investor-Friendly Visualizations

---

# Future Enhancements

* Live CoinMarketCap API integration
* Automatic data refresh
* Power BI version
* Machine Learning-based price prediction
* Portfolio tracking dashboard

---

# Author

**Name:** Rakesh Reddy

**Domain:** Data Analytics Internship

**Project:** Crypto Market Analysis Dashboard

---

## License

This project was developed for educational and internship purposes.
