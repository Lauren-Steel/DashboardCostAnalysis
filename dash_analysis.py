import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# === Load Excel files ===
dashboard_df = pd.read_excel("dashboard_catalog.xlsx")
effort_levels_df = pd.read_excel("effort_levels.xlsx")
hourly_rates_df = pd.read_excel("hourly_rates.xlsx")

# === Normalize column names ===
dashboard_df.columns = dashboard_df.columns.str.strip().str.lower()
effort_levels_df.columns = effort_levels_df.columns.str.strip().str.lower()
hourly_rates_df.columns = hourly_rates_df.columns.str.strip().str.lower()

# === Rename and clean dashboard data ===
dashboard_df = dashboard_df.dropna(subset=["dashboard name", "estimated conversion level"])
dashboard_df = dashboard_df.rename(columns={
    "dashboard name": "dashboard",
    "data visualization tool": "tool",
    "estimated conversion level": "conversion_level",
    "available data pages": "pages"
})
dashboard_df["conversion_level"] = dashboard_df["conversion_level"].str.strip().str.lower()
dashboard_df["status"] = dashboard_df["status"].astype(str).str.strip().str.lower()
dashboard_df["pages"] = pd.to_numeric(dashboard_df["pages"], errors="coerce")

# === Map effort levels ===
effort_levels_df["conversion level"] = effort_levels_df["conversion level"].str.strip().str.lower()
effort_lookup = effort_levels_df.set_index("conversion level")
dashboard_df["base_hours"] = dashboard_df["conversion_level"].map(effort_lookup["conversion rate hrs"])
dashboard_df["junior_pct"] = dashboard_df["conversion_level"].map(effort_lookup["junior analyst efforts"])
dashboard_df["manager_pct"] = dashboard_df["conversion_level"].map(effort_lookup["manager efforts"])

# === Page multiplier ===
def page_multiplier(pages):
    if pd.isna(pages):
        return 1.0
    pages = min(int(pages), 10)
    scale = {
        1: 1.00, 2: 1.05, 3: 1.10, 4: 1.15, 5: 1.20,
        6: 1.25, 7: 1.30, 8: 1.35, 9: 1.40, 10: 1.45
    }
    return scale.get(pages, 1.0)

dashboard_df["page_factor"] = dashboard_df["pages"].apply(page_multiplier)

# === Status multiplier ===
status_multiplier = {
    "active": 1.0,
    "in progress": 1.1,
    "broken": 1.3
}
dashboard_df["status_factor"] = dashboard_df["status"].map(status_multiplier)

# === Adjusted hours and cost calculation ===
dashboard_df["adjusted_hours"] = (
    dashboard_df["base_hours"] *
    dashboard_df["page_factor"] *
    dashboard_df["status_factor"]
)

dashboard_df["jr_hours"] = dashboard_df["adjusted_hours"] * (dashboard_df["junior_pct"] / 100)
dashboard_df["mgr_hours"] = dashboard_df["adjusted_hours"] * (dashboard_df["manager_pct"] / 100)

# === Apply hourly rates from file ===
rates = hourly_rates_df.set_index("position")["hourly rates"].to_dict()
jr_rate = rates.get("Junior Analyst", 0)
mgr_rate = rates.get("Manager", 0)

dashboard_df["jr_cost"] = dashboard_df["jr_hours"] * jr_rate
dashboard_df["mgr_cost"] = dashboard_df["mgr_hours"] * mgr_rate
dashboard_df["total_cost"] = dashboard_df["jr_cost"] + dashboard_df["mgr_cost"]

# === Totals ===
total_hours = dashboard_df["adjusted_hours"].sum()
total_jr_cost = dashboard_df["jr_cost"].sum()
total_mgr_cost = dashboard_df["mgr_cost"].sum()

# === Contractor calculation using fixed rate of $43/hour ===
contractor_rate = 43
contractor_total_cost = total_hours * contractor_rate

# === Combined Bar Chart: Option A vs Option B ===
scenarios = ["Option A", "Option B"]
roles = ["Junior Analyst", "Manager", "Contractor"]

cost_data = {
    "Option A": [total_jr_cost, total_mgr_cost, 0],
    "Option B": [0, 0, contractor_total_cost]
}
cost_df = pd.DataFrame(cost_data, index=roles)

x = np.array([0, 0.6])
width = 0.5

fig, ax = plt.subplots(figsize=(10, 6))

# Plot stacked bar (Option A)
bar1 = ax.bar(x[0], cost_df.loc["Junior Analyst", "Option A"], width, label="Junior Analyst", color="#99ccff")
bar2 = ax.bar(x[0], cost_df.loc["Manager", "Option A"], width, bottom=cost_df.loc["Junior Analyst", "Option A"], label="Manager", color="#336699")

# Plot single bar (Option B)
bar3 = ax.bar(x[1], cost_df.loc["Contractor", "Option B"], width, label="Contractor", color="#66cc99")

# Add cost labels above each bar
def add_label(x_pos, total):
    ax.text(x_pos, total + 300, f"${total:,.0f}", ha='center', va='bottom', fontsize=10, fontweight='bold')

add_label(x[0], total_jr_cost + total_mgr_cost)
add_label(x[1], contractor_total_cost)

# Styling
ax.set_title("Total Dashboard Conversion Cost: Option A vs Option B")
ax.set_ylabel("Total Cost ($)")
ax.set_xticks(x)
ax.set_xticklabels(scenarios)
ax.set_ylim(0, max(total_jr_cost + total_mgr_cost, contractor_total_cost) * 1.15)
ax.legend(title="Role")
plt.tight_layout()
plt.savefig("combined_conversion_cost_comparison.png")
plt.show()

# === Heatmap: Cost per dashboard ===
heatmap_df = dashboard_df.set_index("dashboard")[["total_cost"]].sort_values("total_cost", ascending=False)

plt.figure(figsize=(10, 8))
sns.heatmap(
    heatmap_df,
    annot=True,
    cmap="Blues",
    fmt=".0f",
    linewidths=0.5,
    linecolor='gray',
    vmin=0
)
plt.title("Per-Dashboard Cost to Convert to Power BI")
plt.xlabel("Total Cost ($)")
plt.ylabel("Dashboard")
plt.tight_layout()
plt.savefig("per_dashboard_conversion_cost_heatmap.png")
plt.show()

# === Heatmap: Hours per dashboard ===
hours_df = dashboard_df.set_index("dashboard")[["adjusted_hours"]].sort_values("adjusted_hours", ascending=False)

plt.figure(figsize=(10, 8))
sns.heatmap(
    hours_df,
    annot=True,
    cmap="PuBuGn",
    fmt=".1f",
    linewidths=0.5,
    linecolor='gray',
    vmin=0
)
plt.title("Estimated Hours to Convert Each Dashboard to Power BI")
plt.xlabel("Total Hours")
plt.ylabel("Dashboard")
plt.tight_layout()
plt.savefig("per_dashboard_conversion_hours_heatmap.png")
plt.show()

# === Print total hours and cost summary ===
print(f"\n🕒 Total Hours Required for All Dashboard Conversions: {total_hours:.1f} hours")
print(f"💰 Total Cost if Contractor Did All Work (@ $43/hr): ${contractor_total_cost:,.2f}")
print(f"💼 Total Cost if Split Between Roles: ${total_jr_cost + total_mgr_cost:,.2f}")

# === Optional: Export detailed breakdown ===
# dashboard_df.to_csv("dashboard_conversion_costs_detailed.csv", index=False)

"""
SUMMARY:
- Contractor hourly rate set to $43/hr
- Combined bar chart: Option A (split) vs Option B (contractor)
- Cost labels, full legend, heatmaps for cost + time
- Detailed summary printed
"""



