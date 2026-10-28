"""
Vrinda Store - Annual Sales Analysis 2022
Steps: 1) load  2) clean  3) analyse  4) save charts
Run:  python analysis.py
Needs: pandas, openpyxl, matplotlib
"""
import warnings; warnings.filterwarnings("ignore")
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

SRC = "data/Vrinda_Store_Data_Analysis.xlsx"
MONTHS = ["January","February","March","April","May","June","July",
          "August","September","October","November","December"]
COLOR = "#4C72B0"

# ---------- 1) LOAD ----------
df = pd.read_excel(SRC, sheet_name="Vrinda Store")
df = df.rename(columns={"AMenount": "Amount", "Channel ": "Channel"})

# ---------- 2) CLEAN ----------
# A find-and-replace (M -> Men, W -> Women) damaged text in other columns,
# e.g. "AMenazon", "MenAHARASHTRA", "Womenestern Dress". Reverse it.
fixed = 0
for col in ["Channel", "Category", "Size", "SKU", "ship-city", "ship-state"]:
    before = df[col].copy()
    df[col] = df[col].str.replace("Women", "W", regex=False).str.replace("Men", "M", regex=False)
    fixed += (before != df[col]).sum()
print(f"Text cells repaired: {fixed}")

# "Amazon" lost a lowercase m (-> "AMazon"), so match channel names by lowercase
known = {c.lower(): c for c in ["Amazon", "Myntra", "Flipkart", "Ajio", "Nalli", "Meesho", "Others"]}
df["Channel"] = df["Channel"].str.strip().str.lower().map(known)
df["Category"] = df["Category"].str.strip().str.title()
df["ship-state"] = df["ship-state"].str.upper().str.strip().replace({"NEW DELHI": "DELHI"})
df["Age group"] = df["Age group"].str.title()
df.to_csv("data/cleaned_vrinda_store.csv", index=False)

# ---------- 3) ANALYSE ----------
total = df["Amount"].sum()
orders = df["Order ID"].nunique()
print(f"Revenue: Rs {total:,.0f} | Orders: {orders:,} | Lines: {len(df):,}")

def table(col):
    t = df.groupby(col)["Amount"].agg(revenue="sum", lines="count", avg_value="mean")
    t["share_%"] = (t["revenue"] / total * 100).round(1)
    return t.sort_values("revenue", ascending=False)

monthly = df.groupby("Months")["Amount"].sum().reindex(MONTHS)
by_channel, by_cat = table("Channel"), table("Category")
by_state, by_gender, by_age = table("ship-state"), table("Gender"), table("Age group")
status = df["Status"].value_counts(normalize=True) * 100
not_delivered = df.assign(bad=df["Status"] != "Delivered").groupby("Category")["bad"].mean() * 100

for name, t in [("CHANNEL", by_channel), ("CATEGORY", by_cat), ("STATE (top 10)", by_state.head(10)),
                ("GENDER", by_gender), ("AGE GROUP", by_age)]:
    print(f"\n{name}\n{t.round(1)}")
print("\nORDER STATUS %\n", status.round(1))

# ---------- 4) CHARTS ----------
def clean_axes(ax, title):
    ax.set_title(title, fontsize=13, fontweight="bold", loc="left")
    ax.spines[["top", "right"]].set_visible(False)

fig, ax = plt.subplots(figsize=(9, 4.5))
ax.plot(MONTHS, monthly / 1e6, marker="o", color=COLOR, lw=2.5)
ax.set_ylabel("Revenue (Rs million)"); ax.set_ylim(1.4, 2.1)
plt.xticks(rotation=45, ha="right"); clean_axes(ax, "Monthly revenue, 2022")
plt.tight_layout(); plt.savefig("images/01_monthly_revenue.png", dpi=150); plt.close()

fig, ax = plt.subplots(figsize=(8, 4.5))
c = by_channel["share_%"].sort_values()
ax.barh(c.index, c.values, color=COLOR)
for i, v in enumerate(c.values): ax.text(v + .4, i, f"{v}%", va="center")
clean_axes(ax, "Share of revenue by sales channel"); ax.set_xlabel("% of revenue")
plt.tight_layout(); plt.savefig("images/02_channel_share.png", dpi=150); plt.close()

fig, ax = plt.subplots(figsize=(8, 4.5))
k = by_cat.sort_values("revenue")
ax.barh(k.index, k["revenue"] / 1e6, color=COLOR)
for i, (r, a) in enumerate(zip(k["revenue"], k["avg_value"])):
    ax.text(r / 1e6 + .1, i, f"avg Rs {a:,.0f}", va="center", fontsize=9)
ax.set_xlim(0, 12.5); ax.set_xlabel("Revenue (Rs million)")
clean_axes(ax, "Revenue by category (with average order-line value)")
plt.tight_layout(); plt.savefig("images/03_category_revenue.png", dpi=150); plt.close()

fig, ax = plt.subplots(figsize=(8, 4.5))
s = by_state.head(10).sort_values("revenue")
ax.barh(s.index, s["revenue"] / 1e6, color=COLOR); ax.set_xlabel("Revenue (Rs million)")
clean_axes(ax, "Top 10 states by revenue")
plt.tight_layout(); plt.savefig("images/04_top_states.png", dpi=150); plt.close()

fig, ax = plt.subplots(figsize=(7, 4))
ax.bar(status.index, status.values, color=[COLOR, "#DD8452", "#C44E52", "#8172B3"])
for i, v in enumerate(status.values): ax.text(i, v + 1, f"{v:.1f}%", ha="center")
ax.set_ylim(0, 105); clean_axes(ax, "Order status")
plt.tight_layout(); plt.savefig("images/05_order_status.png", dpi=150); plt.close()

fig, ax = plt.subplots(figsize=(7, 4))
g = df.pivot_table(index="Age group", columns="Gender", values="Amount", aggfunc="sum") / total * 100
g.reindex(["Teenager", "Adult", "Senior"]).plot(kind="bar", ax=ax, color=["#DD8452", COLOR], rot=0)
ax.set_ylabel("% of total revenue"); clean_axes(ax, "Revenue by age group and gender")
plt.tight_layout(); plt.savefig("images/06_age_gender.png", dpi=150); plt.close()
print("\nDone. Charts saved in images/")
