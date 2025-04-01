import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# === Load Excel files ===
dashboard_df = pd.read_excel("dashboard_catalog.xlsx")
effort_levels_df = pd.read_excel("effort_levels.xlsx")
hourly_rates_df = pd.read_excel("hourly_rates.xlsx")

# === Normalize column names ===
dashboard_df.columns = dashboard_df.columns.str.strip().str.lower()
effort_levels_df.columns = effort_levels_df.columns.str.strip().str.lower()
hourly_rates_df.columns = hourly_rates_df.columns.str.strip().str.lower()

# === Rename key columns in dashboard catalog ===
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

# === Map baseline hours from effort_levels.xlsx ===
effort_levels_df["conversion level"] = effort_levels_df["conversion level"].str.strip().str.lower()
effort_lookup = effort_levels_df.set_index("conversion level")
dashboard_df["base_hours"] = dashboard_df["conversion_level"].map(effort_lookup["conversion rate hrs"])
dashboard_df["junior_pct"] = dashboard_df["conversion_level"].map(effort_lookup["junior analyst efforts"])
dashboard_df["manager_pct"] = dashboard_df["conversion_level"].map(effort_lookup["manager efforts"])

# === Page multiplier function ===
def page_multiplier(pages):
    if pd.isna(pages):
        return 1.0
    pages = min(int(pages), 10)
    scale = {
        1: 1.00,
        2: 1.05,
        3: 1.10,
        4: 1.15,
        5: 1.20,
        6: 1.25,
        7: 1.30,
        8: 1.35,
        9: 1.40,
        10: 1.45
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

# === Final adjusted total hours ===
dashboard_df["adjusted_hours"] = (
    dashboard_df["base_hours"] *
    dashboard_df["page_factor"] *
    dashboard_df["status_factor"]
)

# === Role-specific hours ===
dashboard_df["jr_hours"] = dashboard_df["adjusted_hours"] * (dashboard_df["junior_pct"] / 100)
dashboard_df["mgr_hours"] = dashboard_df["adjusted_hours"] * (dashboard_df["manager_pct"] / 100)

# === Apply hourly rates ===
rates = hourly_rates_df.set_index("position")["hourly rates"].to_dict()
jr_rate = rates.get("Junior Analyst", 0)
mgr_rate = rates.get("Manager", 0)

dashboard_df["jr_cost"] = dashboard_df["jr_hours"] * jr_rate
dashboard_df["mgr_cost"] = dashboard_df["mgr_hours"] * mgr_rate
dashboard_df["total_cost"] = dashboard_df["jr_cost"] + dashboard_df["mgr_cost"]

# === Stacked Bar Chart ===
total_jr_cost = dashboard_df["jr_cost"].sum()
total_mgr_cost = dashboard_df["mgr_cost"].sum()

stacked_cost = pd.DataFrame({
    "Junior Analyst": [total_jr_cost],
    "Manager": [total_mgr_cost]
}, index=["Total Conversion Cost"])

colors = ["#99ccff", "#336699"]
stacked_cost.plot(
    kind='bar',
    stacked=True,
    figsize=(8, 6),
    color=colors
)
plt.title("Total Dashboard Conversion Cost to Power BI (by Role)")
plt.ylabel("Total Cost ($)")
plt.xlabel("")
plt.xticks(rotation=0)  # This makes the x-axis label horizontal
plt.legend(title="Role")
plt.tight_layout()
plt.savefig("total_conversion_cost_by_role.png")
plt.show()

# === Calculate total hours for use in both scenarios ===
total_hours = dashboard_df["adjusted_hours"].sum()

# === Stacked Bar Chart (All work done by Junior Analyst / Contractor) ===
contractor_total_cost = total_hours * jr_rate

contractor_cost_df = pd.DataFrame({
    "Contractor": [contractor_total_cost]
}, index=["Total Conversion Cost"])

contractor_colors = ["#99ccff"]
contractor_cost_df.plot(
    kind='bar',
    stacked=True,
    figsize=(8, 6),
    color=contractor_colors
)
plt.title("Total Conversion Cost if Done Entirely by Contractor")
plt.ylabel("Total Cost ($)")
plt.xlabel("")
plt.xticks(rotation=0)  # Horizontal x-axis label
plt.legend(title="Role")
plt.tight_layout()
plt.savefig("total_conversion_cost_contractor_only.png")
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

# === NEW Heatmap: Hours per dashboard ===
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

# === Print total hours and full analyst cost scenario ===
total_hours = dashboard_df["adjusted_hours"].sum()
all_jr_cost = total_hours * jr_rate

print(f"\n🕒 Total Hours Required for All Dashboard Conversions: {total_hours:.1f} hours")
print(f"💰 Total Cost if Junior Analyst Did All Work: ${all_jr_cost:,.2f}")
print(f"💼 Total Cost if Split Between Roles (Based on Effort): ${total_jr_cost + total_mgr_cost:,.2f}")


# === Optional: Export detailed cost matrix ===
# dashboard_df.to_csv("dashboard_conversion_costs_detailed.csv", index=False)

"""
UPDATED SUMMARY:
- Adds a heatmap showing hours per dashboard
- Prints total hours required for all conversions
- Calculates and prints cost if Junior Analyst handled 100% of the work
- Continues to output:
  - Stacked bar chart (cost by role)
  - Cost heatmap (per dashboard)
  - Optional CSV export
"""
