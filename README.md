# Vrinda Store: 2022 Sales Analysis

An end-to-end analysis of 31,047 online order lines for **Vrinda Store**, an Indian fashion retailer selling through marketplaces such as Amazon, Myntra and Flipkart. The project cleans the raw data, builds an Excel dashboard, and answers the questions a business owner would actually ask.

**Tools:** Excel (Pivot Tables, Slicers, Dashboard) · Python (pandas, matplotlib)

---

## 1. Business problem

The store owner (the client) wants to know:

> *"Who is buying from us, what are they buying, where, and through which channel, and what should we do next to grow in 2023?"*

## 2. Dataset

| Item | Detail |
|---|---|
| Rows | 31,047 order lines (28,471 unique orders) |
| Period | Calendar year 2022 |
| Columns | Order ID, Customer ID, Gender, Age, Age group, Date, Month, Status, Channel, SKU, Category, Size, Qty, Amount, City, State, B2B |
| Currency | INR (all rows) |

## 3. Headline numbers

| Metric | Value |
|---|---|
| Total revenue | **₹21.18 million** |
| Unique orders | 28,471 |
| Average value per order line | ₹682 |
| Average value per full order | ₹744 |
| Delivered successfully | 92.3% |
| Revenue lost to cancel / refund / return | ₹1.47 million (6.9%) |

---

## 4. Client questions and analyst answers

The format below is a role-play: the **client** asks, the **analyst** answers using the data.

### Q1. Which customers bring in the most money?

**Women: 64% of revenue (₹13.6M)**, men 36% (₹7.6M). By age, adults (30–49) bring in 50%, ages 18–29 bring 30%, and seniors (50+) bring 20%.

![Age and gender](images/06_age_gender.png)

**Insight:** Women aged 30–49 are the single biggest group (32% of all revenue). Marketing should be built around them first.

### Q2. What is selling best?

**"Set" products make up almost half of revenue (49.6%)**, followed by Kurta (23.4%) and Western Dress (14.9%). These three categories together give 88% of sales.

![Category revenue](images/03_category_revenue.png)

**Insight:** Kurta sells in very high volume (10,446 lines) but at a low price (avg ₹475), while Sets average ₹848. Selling a Kurta together with a matching Set or Top (a bundle) is the easiest way to raise basket value.

### Q3. Which sales channel should we focus on?

**Amazon, Myntra and Flipkart give 80% of revenue** (35.5%, 23.3%, 21.6%).

![Channel share](images/02_channel_share.png)

**Insight:** The store depends heavily on three marketplaces. If one changes its fees or rules, the business feels it immediately. Ajio (6.3%) has the highest average order value among the big channels (₹690), so it is a good second-tier growth channel. Meesho has the lowest (₹664).

### Q4. Where in India are our customers?

**Maharashtra (14.1%), Karnataka (12.5%) and Uttar Pradesh (9.9%)** lead. The top 5 states give about 52% of revenue. **Bengaluru is the #1 city** (₹1.67M), followed by Hyderabad and New Delhi.

![Top states](images/04_top_states.png)

**Insight:** The five southern states (Karnataka, Telangana, Tamil Nadu, Kerala, Andhra Pradesh) together give about 38% of revenue. Regional ads in Kannada, Telugu and Tamil are a low-cost opportunity.

### Q5. How did sales change through the year?

Sales peaked in **March (₹1.93M)** and then drifted down to **December (₹1.62M)**, a drop of about 16%. The second half of the year earned 7.5% less than the first half.

![Monthly revenue](images/01_monthly_revenue.png)

**Insight:** There is no festive-season rise (Oct–Dec) in this data. That is unusual for Indian fashion and is worth investigating (see the data warning in Section 6).

### Q6. How many orders go wrong?

**7.7% of order lines are not delivered**: returned 3.4%, cancelled 2.7%, refunded 1.7%.

![Order status](images/05_order_status.png)

**Insight:** Western Dress has the highest problem rate among big categories (**11.7%** not delivered) while Set has the lowest (5.0%). Checking size charts, photos and packaging for Western Dress could recover real money.

### Q7. Is B2B (business) selling worth it?

B2B is only **0.6%** of order lines (186 lines, ₹126K) with the same average value as normal customers. It is not a priority right now.

---

## 5. Recommendations

1. **Target women aged 30–49** in campaigns. They are the largest revenue group.
2. **Bundle Kurta + Set/Top** to lift the low ₹475 Kurta basket.
3. **Reduce Western Dress returns** (11.7%) with better size guides and photos.
4. **Reduce dependence on three marketplaces:** grow Ajio and the store's own website.
5. **Run regional campaigns in South India** (about 38% of revenue) and in Bengaluru and Hyderabad.
6. **Plan a stronger Q4 promotion.** Sales fell from March to December with no festive spike.

---

## 6. Data quality problems found and fixed

Cleaning is part of the work, so these are documented honestly.

| Problem | What was wrong | Fix |
|---|---|---|
| Corrupted text | A find-and-replace (M → "Men", W → "Women") damaged other columns: `AMenazon`, `Menyntra`, `MenAHARASHTRA`, `WomenEST BENGAL`, `Womenestern Dress`, and size "Men". 76 cells in total. | Reversed in `analysis.py` |
| Wrong column name | `AMenount` instead of `Amount` | Renamed |
| Inconsistent spelling | `kurta` vs `Kurta`, `Delhi` / `delhi` / `New Delhi`, `teenager` vs `Teenager` | Standardised |
| Misleading label | "Teenager" is used for ages 18–29 (no one is under 18) | Suggest renaming to "Young Adult" |

### Warnings the reader should know

- **Dates look wrong.** Every month contains only orders from the **4th, 5th and 6th**. A full year of real orders would cover every day. Day and month may have been swapped when the file was imported, or the data is a sample. The monthly trend in Q5 should be treated with caution until the dates are verified.
- **Gender looks odd.** 3,606 "Men" orders are Western Dress and 5,138 are Sets, in a catalogue that is mostly women's wear. The Gender column may describe something other than the wearer (for example, the account holder).
- **Customer ID is not a real customer ID.** It is copied from the middle digits of the Order ID, so repeat-customer analysis is not possible.

---

## 7. Dashboard note (what to fix in the Excel file)

The pivot tables in the workbook are **filtered to January only** by a month slicer. That is why the dashboard shows ₹1.82M revenue and 2,702 orders instead of the full-year ₹21.18M. Clear the slicer (or add a "Select all" view) so the first screen shows the whole year.

## 8. How to run

```bash
pip install pandas openpyxl matplotlib
python analysis.py
```

Put the Excel file in `data/` first. The script cleans the data, prints the tables, saves `data/cleaned_vrinda_store.csv`, and rebuilds all charts in `images/`.

## 9. Repository structure

```
vrinda-store-analysis/
├── README.md
├── analysis.py
├── data/
│   ├── Vrinda_Store_Data_Analysis.xlsx
│   └── cleaned_vrinda_store.csv
└── images/            (6 charts)
```

## 10. What I learned

- How to clean messy real-world data before trusting any number.
- How to turn pivot tables into business answers and recommendations.
- Why a dashboard must be checked for hidden filters.
