# 🇨🇦📊 Canada Job Market Analysis
### Python + SQL + Power BI | Real Statistics Canada Data

![Python](https://img.shields.io/badge/Python-3.8+-blue?style=flat-square&logo=python&logoColor=white)
![SQL](https://img.shields.io/badge/SQL-SQLite-4479A1?style=flat-square&logo=sqlite&logoColor=white)
![Power BI](https://img.shields.io/badge/Power%20BI-Dashboard-F2C811?style=flat-square&logo=powerbi&logoColor=black)
![Data](https://img.shields.io/badge/Data-Statistics%20Canada-red?style=flat-square)
![Status](https://img.shields.io/badge/Status-Complete-brightgreen?style=flat-square)

---

## 📌 Project Overview

An end-to-end data analytics project analysing Canada's labour market from 2019 to 2024 — covering the full COVID-19 employment crash and recovery story across 17 industries.

> **Data Source:** Statistics Canada — Table 14-10-0202-01 (Labour Force Survey) — Real government open data ✅

🔗 **[View Live Power BI Dashboard](https://app.powerbi.com/view?r=eyJrIjoiYzYyYzE2YTAtN2ZjNC00MjFlLWIzOWMtYmE5Y2RhMTRhZTNhIiwidCI6ImQ0MWZkYWIxLTdlMTUtNGNmZC1iNWZhLTcyMDBlNTRkZWI2YiJ9)**

---

## 🎯 Problem Statement

Canada's labour market went through one of the most dramatic shifts in modern history between 2019 and 2024. This project answers:
- Which industries were hit hardest by COVID-19?
- How fast did each industry recover?
- Which sectors surpassed pre-COVID employment levels?
- What are the national employment trends month by month?

---

## 📂 Dataset

| Detail | Info |
|---|---|
| Source | Statistics Canada, Table 14-10-0202-01 |
| Period | January 2019 → December 2024 (6 years) |
| Frequency | Monthly |
| Industries | 17 NAICS categories |
| Metric | Employment (persons in thousands) |
| Raw rows | 2,016 |
| Cleaned rows | 1,224 |

---

## 🐍 Python Pipeline

Built an end-to-end Python script using `pandas`, `sqlite3` and `openpyxl`:

```
Raw CSV (Statistics Canada)
        ↓
pandas — load, clean, parse dates
        ↓
sqlite3 — load into in-memory SQL database
        ↓
5 SQL queries — trend, ranking, COVID impact, YoY, quarterly
        ↓
openpyxl — export to formatted Excel workbook
        ↓
Power BI — connect Excel, build 5-page dashboard
```

**Cleaning steps:**
- Removed NAICS bracket codes from industry names
- Derived Year, Month, Quarter from REF_DATE
- Filtered to top-level industry categories only
- Handled null values and type conversions

---

## 🗄️ SQL Queries (5)

| Query | Purpose | Key Technique |
|---|---|---|
| Q1 — National Trend | Monthly employment 2019–2024 | GROUP BY, ORDER BY |
| Q2 — Industry Ranking | Industries by employment size | AVG, RANK() OVER |
| Q3 — COVID Impact | Drop % and recovery status | CASE WHEN, CTEs, self-JOIN |
| Q4 — YoY Growth | Annual % change per industry | Self-JOIN on yearly averages |
| Q5 — Quarterly Trend | Seasonality by quarter | GROUP BY Year + Quarter |

---

## 📊 Power BI Dashboard — 5 Pages

| Page | Visual | Insight |
|---|---|---|
| 📈 National Trend | Line chart + KPI cards | COVID crash in April 2020 clearly visible |
| 🏭 Industry Analysis | Bar chart + Treemap | Healthcare = largest employer nationally |
| 💉 COVID Impact | Clustered column + table | Accommodation hit hardest (-40%+) |
| 📉 YoY Growth | Line chart + conditional matrix | 2020 = red across all industries |
| 📅 Quarterly Trend | Line chart + matrix | Seasonality in Agriculture, Accommodation |

---

## 💡 Key Findings

- 📉 April 2020 — Canada lost ~3,500K jobs in a single month — sharpest drop on record
- 🍽️ Accommodation and Food Services hit hardest — down 40%+ at peak COVID impact
- 🏥 Healthcare and Public Administration most resilient — barely affected in 2020
- 📈 Most industries fully surpassed pre-COVID employment levels by 2023
- 🏆 National employment reached an all-time high in 2024

---

## 🗂️ Project Structure

```
Canada-Job-Market-Analysis/
├── data/
│   ├── 1410002201_databaseLoadingData.csv
│   └── canada_employment_clean.csv
├── output/
│   └── Canada_Employment_Analysis.xlsx
├── analysis.py
└── README.md
```

---

## ⚙️ How to Run

```bash
git clone https://github.com/ruchithayathiraj/Canada-Job-Market-Analysis
pip install pandas openpyxl
python analysis.py
```

---

## 🛠️ Tools Used

Python · pandas · sqlite3 · openpyxl · SQL · Power BI · Statistics Canada Open Data

---

## 👩‍💻 Author

**Ruchitha Yathirajulu** — Business Analyst | Data Analyst | Ottawa, Canada

[LinkedIn](https://www.linkedin.com/in/ruchitha-yathirajulu-b87555191/) | [Portfolio](https://ruchithayathiraj.github.io/My-Portfolio) | yathirajuluruchitha@gmail.com
