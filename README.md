# Global Campaign Performance Analysis

## Project Overview
This project turns a messy marketing dataset of over 3,000 rows into a clean, interactive Excel dashboard. 

The goal was to fix broken data exports, link separate tables using lookups, and build a tracking panel to evaluate campaign sales, website traffic, and manager performance.

---

## Technical Workflow & Steps

### 1. Data Cleaning
Before building any charts, the raw data was cleaned using standard Excel formulas:
* **Fixed Text Spacing:** Used `TRIM` and `PROPER` to remove hidden trailing spaces and fix random capitalization errors in the Audience and Placement columns.
* **Cleared System Errors:** Found missing data placeholders (`-999` and `ERROR_VAL`) and replaced them with standard zeros (`0`) using `IFERROR` and Find & Replace shortcuts so they would not break dashboard averages.
* **Standardized Categories:** Simplified irregular text values (like changing `Search_Top` to a clean `Search` tag) to create neat, uniform categories for filtering.

### 2. Table Linking via XLOOKUP
Instead of copying and pasting data across sheets, the separate tables were connected using clean database logic:
* **Campaign Metadata Mapping:** Used `XLOOKUP` to match transaction codes directly to a secondary tab to pull back campaign names and employee managers.
* **Region Directory Joins:** Used a second `XLOOKUP` to connect the main sheet to a regional reference tab to pull back executive director names and target metrics.

### 3. Performance Tier Logic
Created a performance engine using an `IFS` formula to automatically classify every row based on profit margins and sales volumes:
```excel
=IFERROR(IFS((M2-K2)/M2<=0, "Unprofitable", AND((M2-K2)/M2>0.75, L2>1000), "Executive Tier", TRUE, "Standard Production"), "Unprofitable")
```

---

## Dashboard Features & Core Metrics

The final tab features a custom green canvas layout containing 5 pyramid KPI blocks and a 4-chart reporting grid.

### The 5 KPI Blocks
* **Total Global Return on Ad Spend (457%):** Tracks total investment efficiency (`Total Revenue / Total Spend`).
* **Executive Tier Campaign Volume (302):** Counts the total number of elite, high-margin campaigns.
* **Global Conversion Rate % (8.74%):** Measures website and checkout efficiency (`Conversions / Clicks`).
* **Ad Click-Through Rate % (4.35%):** Measures creative ad engagement (`Clicks / Impressions`).
* **The Regional Director Leaderboard:** Instantly ranks your top territory managers (Sarah Jenkins, David Ross, Michael Chang) by total generated revenue.

### The 4 Charts
* **Channel Traffic vs. Conversions (Combo Chart):** Plots click bars against conversion lines to highlight customer drop-off points across channels.
* **Sales Share by Ad Placement (Pie Chart):** Shows the percentage share of sales driven by screen layouts to identify wasted budget.
* **Spend vs. Revenue Over Time (Line Chart):** A dual-line timeline chart showing the true gap between daily ad spend and daily incoming revenue.
* **Revenue by Audience Segment (Bar Chart):** A column chart sorting sales performance by target customer profiles (B2B, Consumer, Enterprise, SMB).

---

## Strategic Insights & Recommendations

Based on the dashboard charts and metrics, the marketing agency should focus on these four key areas:

* **Put More Budget into Paid Search and Social Media:** The *Channel Traffic vs. Conversions* chart proves that these two channels drive the highest number of clicks and sales. The agency should move money here and away from lower-performing channels like Display and Affiliate.
* **Fix Landing Pages for Email and Affiliate Channels:** The *Channel Traffic vs. Conversions* chart shows high click volumes but flat sales lines for Email and Affiliate. This means people click the ads but leave without buying. The agency needs to check the website checkout flow and copy for these channels to stop wasting money.
* **Focus on Business Customers (Enterprise and SMB):** The *Revenue by Audience Segment* chart clearly identifies `Enterprise`, `SMB`, and `B2B` as the primary revenue drivers. Shifting ad targeting away from individual `Consumer` profiles will bring in higher-value sales.
* **Copy Sarah Jenkins' Strategy:** Sarah Jenkins holds the top spot on the *Regional Director Leaderboard*. The agency should have a meeting so David Ross and Michael Chang can learn and copy her exact budget and placement strategies in their own underperforming regions.

---

## How to Interact with the Dashboard
1. Open the `.xlsx` workbook and go to the main dashboard tab.
2. Go to **View** on the top Excel ribbon and uncheck **Gridlines** to view the clean green background layout.
3. Click any button on the **Target_Region** or **Audience_Segment** Master Slicers on the side of the menu to filter all data instantly.
