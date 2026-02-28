# Mobile App Funnel & Conversion Optimization — Excel Analytics Dashboard

**Author:** Sai Keerthana Malothu  
**Tools:** Python (pandas, numpy) | Excel 365 (Power Query, Pivot Tables, DAX-style formulas, Slicers) | Jupyter Notebook  
**Domain:** Product Analytics | Mobile App | Funnel Analysis | Acquisition Efficiency  
**Dataset:** Synthetic dataset simulating a real-world mobile app with 5,000 users across 6 acquisition campaigns

---

## Project Overview

This end-to-end analytics project analyses a mobile app funnel across 6 acquisition campaigns, 5 age segments, 2 devices, and 8 countries. The goal was to identify where users drop off, which campaigns deliver the best retained user value, and what interventions will improve the overall 3.9% Day 7 activation rate.

The project delivers a **fully interactive Excel BI dashboard** — built without Power BI — demonstrating advanced Excel analytics capability including Power Query, Pivot Tables, dynamic slicers, and cross-sheet formula references.

---

## Business Problem

A mobile app company is losing 96% of users between install and Day 7 activation — far below the 25% industry target. The product and marketing teams need to know:

- Where in the funnel are users dropping off?
- Which campaigns deliver the lowest cost per **retained** user?
- Which user segments convert best and should be prioritised?
- What are the top 3 actions to improve conversion immediately?

---

## Key Finding — Funnel Results

| Stage | Users | % of Install | Drop-off | vs Target |
|---|---|---|---|---|
| Install | 5,000 | 100% | — | 100% |
| Signup | 3,831 | 76.6% | 1,169 | ⚠ -8.4pp |
| Onboarding | 2,407 | 48.1% | 1,424 |  -21.9pp |
| First Action | 1,146 | 22.9% | 1,261 |  -22.1pp |
| Day 1 Active | 515 | 10.3% | 631 |  -29.7pp |
| Day 7 Active | 193 | 3.9% | 322 |  -21.1pp |

**96.1% of users who install never become active by Day 7.**  
**Biggest leak: Signup → Onboarding — 1,424 users lost in the first 48 hours.**

---

## Campaign Performance — CPA = Cost ÷ Day 7 Retained Users

| Campaign | Users | Day 7 Active | Conv Rate | Spend | CPA / Retained User | Decision |
|---|---|---|---|---|---|---|
| App Store Optimisation | 894 | 69 | 7.7% | $0 | Free | SCALE |
| Email Referral | 480 | 28 | 5.8% | $5,000 | $179 | SCALE |
| Google Ads — Search | 1,326 | 67 | 5.1% | $45,000 | $672 | GROW |
| Influencer Partnership | 564 | 13 | 2.3% | $25,000 | $1,923 | REVIEW |
| Meta Ads — Social | 1,089 | 10 | 0.9% | $38,000 | $3,800 | CUT |
| TikTok Ads | 647 | 6 | 0.9% | $32,000 | $5,333 | CUT |

**CPA defined as Campaign Spend ÷ Day 7 Active Users — the strongest acquisition metric because it ties spend directly to retained, active users who generate ongoing value.**

**TikTok CPA ($5,333) is 30x Email Referral ($179). Reallocating $32,000 TikTok budget to Email/Google would acquire ~179 additional retained users at same spend.**

---

## Top Segments

| Segment | Type | Users | Day 7 Conv | vs Average | Action |
|---|---|---|---|---|---|
| App Store Optimisation | Channel | 894 | 7.7% | +3.8% | Double Down |
| 25-34 Age Group | Age | 1,521 | 5.9% | +2.0% | Double Down |
| Email Referral | Channel | 480 | 5.8% | +1.9% | Double Down |
| iOS Users | Device | 2,903 | 4.9% | +1.0% | Lead Platform |
| Android Users | Device | 2,097 | 2.4% | -1.5% | Investigate |
| TikTok Users | Channel | 647 | 0.9% | -3.0% | Stop |

---

## Top 3 Recommendations

1. **Simplify signup to 3 fields maximum** — target +8pp signup rate — Week 1-2, Low effort
2. **Redesign onboarding to 3 steps** — target +15pp completion — Month 1, Medium effort
3. **Pause TikTok, reallocate $32K to Google + Email Referral** — recover ~$31,862 in wasted spend — Week 1, Low effort

---

## Excel Dashboard — 7 Sheets

### Dashboard — Interactive Overview
https://github.com/saikeerthanamalothu-dot/mobile-funnel-analytics/blob/main/%20dashboard_screenshot.png

**Features:**
- 5 dynamic KPI cards — update live when slicers are applied
- 4 interactive slicers — filter by Device, Age Group, Campaign, Country
- Dynamic insight bar — shows filtered metrics in real time
- Funnel table with RAG status (FIX NOW / CRITICAL / MONITOR)
- Campaign table with CPA and SCALE/REVIEW/CUT decisions
- Segment table with DOUBLE DOWN / DEPRIORITISE actions
- Benchmark gap table vs industry targets

### Funnel Analysis
Detailed funnel breakdown by device and age group with benchmark gaps and recommendations

### Campaign Analysis
Full campaign metrics, CPA analysis, reallocation modelling and decision rationale

### Calculations
Source data for all charts — funnel bar chart, campaign conversion chart, benchmark comparison chart. Benchmark targets are editable yellow cells.

### README Sheet
Full metric definitions, data sources, benchmark citations, and workbook guide

### data_users / data_events / data_campaigns
Raw data tables (hidden) — source for all pivot tables and calculations

---

## Excel Skills Demonstrated

| Skill | Where Used |
|---|---|
| Power Query | Loading and transforming all 3 data tables |
| Pivot Tables | 4 pivot tables — by device, age, campaign, country |
| Slicers | 4 connected slicers filtering all pivot tables simultaneously |
| Dynamic formulas | KPI cards using INDEX/MATCH referencing pivot Grand Totals |
| XLOOKUP | Joining Campaign_Name from campaigns table to users table |
| Cross-sheet references | Dashboard KPIs pulling live from Pivot_Data sheet |
| Conditional formatting | RAG status, performance indicators |
| Power Query relationships | tbl_users and tbl_campaigns named tables |
| Chart building | Funnel, campaign comparison, benchmark charts |
| Dashboard design | Visual hierarchy, colour coding, decision-oriented layout |

---

## Project Workflow

```
Python (generate data) → Excel Power Query (load + clean) → 
Pivot Tables (aggregate) → Slicers (interactivity) → 
Dashboard (visualise) → Recommendations (decisions)
```

### Notebook — 01_generate_data.ipynb
- Generates 3 related tables: dim_users (5,000 rows), fact_events (13,000+ rows), dim_campaigns (6 rows)
- Conversion probabilities vary by campaign, device, and age group to create realistic patterns
- Exports all files as Excel-ready .xlsx files

---

## Repository Structure

```
mobile-funnel-analytics/
├── README.md
├── 01_generate_data.ipynb         ← Python data generation
├── mobile_funnel_dashboard.xlsx   ← Full interactive Excel dashboard
├── dim_users.xlsx                 ← Raw user data (5,000 rows)
├── dim_campaigns.xlsx             ← Campaign data (6 campaigns)
├── fact_events.xlsx               ← Event-level data (13,000+ rows)
├── funnel_summary.xlsx            ← Funnel summary table
└── campaign_performance.xlsx      ← Campaign metrics table
```

---

## How to Use

1. Clone the repository
```bash
git clone https://github.com/saikeerthanamalothu-dot/mobile-funnel-analytics.git
cd mobile-funnel-analytics
```

2. Install Python dependencies
```bash
pip install pandas numpy openpyxl jupyter
```

3. Run the data generation notebook
```bash
jupyter notebook 01_generate_data.ipynb
```

4. Open `mobile_funnel_dashboard.xlsx` in Excel 365 — all pivot tables, slicers and formulas work immediately

5. Use the slicers on the Dashboard sheet to filter by Device, Age Group, Campaign or Country — all KPI cards and the insight bar update dynamically

---

## Benchmark Sources

| Benchmark | Source |
|---|---|
| Signup Rate (85% target) | AppsFlyer Mobile Benchmarks 2024 |
| Onboarding Rate (70% target) | Adjust App Trends Report 2024 |
| Day 1 Retention (40% target) | Liftoff Mobile Report 2024 |
| Day 7 Retention (25% target) | AppsFlyer Mobile Benchmarks 2024 |

*Benchmarks represent typical ranges for consumer utility/productivity apps. Treat as directional targets. All benchmark values are editable in the Calculations sheet.*

---

## How This Relates to Real-World Work

This project mirrors what a BI Analyst or Product Analyst would deliver in a mobile or SaaS company:

- **The funnel analysis** is what the product team uses to prioritise engineering work
- **The CPA table** is what the marketing team uses to allocate budget
- **The segment analysis** is what informs targeting strategy
- **The interactive dashboard** is what stakeholders use weekly to monitor KPIs
- **The benchmark gap table** is what gets presented to the VP of Product or CMO

The goal was not just to analyse data but to deliver a complete decision-support tool that a non-technical stakeholder could use immediately.

---

## About

This is Project 3 in a data analytics portfolio demonstrating end-to-end capability across Python, SQL, Excel BI, Power BI, machine learning, and business analysis. The project specifically showcases modern Excel BI — Power Query, Pivot Tables, and dynamic slicers — as an alternative to Power BI for organisations that work primarily in Excel.

**Project 1:** [Spend Analysis & Cost Optimization Dashboard](https://github.com/saikeerthanamalothu-dot/Spend-Analysis-Cost-Optimization-Dashboard)  
**Project 2:** [SaaS Customer Churn Prediction & Retention Strategy](https://github.com/saikeerthanamalothu-dot/saas-churn-prediction)

---

*Sai Keerthana Malothu | MSc Business Analytics*
