# biomass_forecast_policy_separate_plots.py
# Monte Carlo forecasts with policy-based caps
# + National aggregation added
# + National plots added
# + National sheet added to Excel
# + Distribution plots added (county & national)
# + Distribution numbers added to Excel

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import os
import re

plt.ioff()   # SPEED: disable interactive plotting

# -----------------------------
# Settings (UNCHANGED)
# -----------------------------
INPUT_XLSX = "historical_biomass.xlsx"
OUTPUT_XLSX = "biomass_forecast_all_counties_policy_separate.xlsx"
PLOTS_DIR = "feedstock_uncertainty_points_clouds_policy"
NATIONAL_PLOTS_DIR = "national_feedstock_uncertainty"
DIST_PLOTS_DIR = "distribution_plots"
NATIONAL_DIST_PLOTS_DIR = "national_distribution_plots"

os.makedirs(PLOTS_DIR, exist_ok=True)
os.makedirs(NATIONAL_PLOTS_DIR, exist_ok=True)
os.makedirs(DIST_PLOTS_DIR, exist_ok=True)
os.makedirs(NATIONAL_DIST_PLOTS_DIR, exist_ok=True)

np.random.seed(42)

# -----------------------------
# Safe filename
# -----------------------------
def safe_filename(text):

    # remove problematic characters
    text = re.sub(r'[\\/*?:"<>|()]+', "", text)

    # replace spaces with underscore
    text = text.replace(" ", "_")

    # remove double underscores
    text = re.sub("_+", "_", text)

    return text.strip("_")


# -----------------------------
# Deterministic county drift
# -----------------------------
def get_county_drift(county, feed):
    seed = abs(hash(county + feed)) % 10_000
    rng = np.random.RandomState(seed)
    return rng.uniform(-0.002, 0.002)

# -----------------------------
# Feedstocks, sectors, counties
# -----------------------------
feedstocks = [
    "Wheat straw","Wheat husk","Wheat bran","Barley straw",
    "Willow leaves","Willow bark","Willow woody stems",
    "Hemp straw","Sitka spruce (wood)","Sitka spruce (needles)",
    "Sitka spruce (branch)","Sitka spruce (bark)",
    "Prawn and shrimp","Crab","Seaweed (Ascophyllum nodosum)",
    "Miscanthus","Pig slurry","Sheep manure","Cattle slurry",
    "Chicken manure","Bull slurry","Other cow slurry",
    "Pasture grass (Perennial ryegrass)","Silage grass"
]


# -----------------------------
# Feedstock economic parameters
# -----------------------------
feedstock_econ = {
    "Wheat straw": {"price":36, "import":599, "export":316},
    "Wheat husk": {"price":0, "import":599, "export":316},
    "Wheat bran": {"price":0, "import":599, "export":316},
    "Barley straw": {"price":50, "import":146, "export":187},
    "Willow leaves": {"price":30, "import":0, "export":0},
    "Willow bark": {"price":30, "import":0, "export":0},
    "Willow woody stems": {"price":28, "import":0, "export":0},
    "Hemp straw": {"price":200, "import":0, "export":0},
    "Sitka spruce (wood)": {"price":38, "import":0, "export":0},
    "Sitka spruce (needles)": {"price":55, "import":0, "export":0},
    "Sitka spruce (branch)": {"price":85, "import":0, "export":0},
    "Sitka spruce (bark)": {"price":25, "import":0, "export":0},
    "Prawn and shrimp": {"price":7300, "import":4188, "export":7930},
    "Crab": {"price":6000, "import":2350, "export":7840},
    "Seaweed (Ascophyllum nodosum)": {"price":800, "import":0, "export":0},
    "Miscanthus": {"price":90, "import":0, "export":0},
    "Pig slurry": {"price":32, "import":83, "export":214},
    "Sheep manure": {"price":0, "import":6, "export":60},
    "Cattle slurry": {"price":7, "import":43, "export":549},
    "Chicken manure": {"price":33, "import":155, "export":88},
    "Bull slurry": {"price":7, "import":43, "export":549},
    "Other cow slurry": {"price":7, "import":43, "export":549},
    "Pasture grass (Perennial ryegrass)": {"price":60, "import":0, "export":0},
    "Silage grass": {"price":150, "import":0, "export":0}
}

# -----------------------------
# Calorific values (MJ/kg)
# -----------------------------
calorific_values = {
    "Wheat straw":14.4,
    "Wheat husk":17.6,
    "Wheat bran":16.4,
    "Barley straw":14.7,
    "Willow leaves":19.4,
    "Willow bark":18.0,
    "Willow woody stems":17.8,
    "Hemp straw":18.3,
    "Sitka spruce (wood)":17.8,
    "Sitka spruce (needles)":19.2,
    "Sitka spruce (branch)":19.4,
    "Sitka spruce (bark)":19.1,
    "Prawn and shrimp":4.70,
    "Crab":5.19,
    "Seaweed (Ascophyllum nodosum)":12.5,
    "Miscanthus":19.1,
    "Pig slurry":17.8,
    "Sheep manure":16.0,
    "Cattle slurry":11.3,
    "Chicken manure":14.5,
    "Bull slurry":11.3,
    "Other cow slurry":11.3,
    "Pasture grass (Perennial ryegrass)":19.0,
    "Silage grass":11.0
}

# -----------------------------
# Biogas composition (%)
# -----------------------------
biogas_composition = {

    "Wheat straw": {"protein":4.5, "carb":66.7, "fat":1.5},
    "Wheat husk": {"protein":6.0, "carb":60.0, "fat":5.0},
    "Wheat bran": {"protein":15.3, "carb":80.7, "fat":3.3},
    "Barley straw": {"protein":3.8, "carb":66.6, "fat":1.4},

    "Willow leaves": {"protein":11.7, "carb":74.0, "fat":3.7},
    "Willow bark": {"protein":3.0, "carb":45.1, "fat":1.5},
    "Willow woody stems": {"protein":1.0, "carb":52.6, "fat":0.3},

    "Hemp straw": {"protein":6.9, "carb":5.3, "fat":1.2},

    "Prawn and shrimp": {"protein":24.3, "carb":1.5, "fat":0.8},
    "Crab": {"protein":19.5, "carb":0.1, "fat":5.1},

    "Miscanthus": {"protein":15.0, "carb":52.6, "fat":0.5},

    "Pig slurry": {"protein":15.1, "carb":24.76, "fat":2.0},
    "Sheep manure": {"protein":5.0, "carb":40.0, "fat":1.5},

    "Cattle slurry": {"protein":11.7, "carb":39.3, "fat":3.8},
    "Chicken manure": {"protein":46.9, "carb":27.0, "fat":2.6},

    "Bull slurry": {"protein":11.7, "carb":39.3, "fat":3.8},
    "Other cow slurry": {"protein":11.7, "carb":39.3, "fat":3.8},

    "Pasture grass (Perennial ryegrass)": {"protein":15.1, "carb":16.2, "fat":2.8},
    "Silage grass": {"protein":13.5, "carb":7.0, "fat":4.5}
}


# -----------------------------
# Methane potential function
# -----------------------------
def methane_yield(feed):

    if feed not in biogas_composition:
        return None

    comp = biogas_composition[feed]

    P = comp["protein"] / 100
    C = comp["carb"] / 100
    F = comp["fat"] / 100

    # m3 CH4 per tonne
    bmp = (496 * P) + (415 * C) + (1014 * F)

    return bmp


sectors = {
    "Agriculture": ["Wheat straw","Wheat husk","Wheat bran","Barley straw",
                    "Hemp straw","Miscanthus","Pig slurry","Sheep manure",
                    "Cattle slurry","Chicken manure","Bull slurry","Other cow slurry",
                    "Pasture grass (Perennial ryegrass)","Silage grass"],
    "Forestry": ["Willow leaves","Willow bark","Willow woody stems",
                 "Sitka spruce (wood)","Sitka spruce (needles)",
                 "Sitka spruce (branch)","Sitka spruce (bark)"],
    "Marine": ["Prawn and shrimp","Crab","Seaweed (Ascophyllum nodosum)"]
}

# -----------------------------
# Sector finder function
# -----------------------------
def find_sector(feed):
    for sec, members in sectors.items():
        if feed in members:
            return sec
    return "Unknown"

counties = [
    "Carlow","Cavan","Clare","Cork","Donegal","Dublin","Galway","Kerry",
    "Kildare","Kilkenny","Laois","Leitrim","Limerick","Longford","Louth",
    "Mayo","Meath","Monaghan","Offaly","Roscommon","Sligo","Tipperary",
    "Waterford","Westmeath","Wexford","Wicklow"
]

max_expansion_factor = {f: 1.5 for f in feedstocks}



# -----------------------------
# Read data
# -----------------------------
years = [2020, 2022]
data = []

for year in years:
    df_year = pd.read_excel(INPUT_XLSX, sheet_name=str(year))
    for _, row in df_year.iterrows():
        county = row['County']
        for f in feedstocks:
            data.append([county, f, year, row[f]])

df_obs = pd.DataFrame(data, columns=["County","Feedstock","Year","Quantity"])



# -----------------------------
# County-to-national ratio for 2020
# -----------------------------
national_2020 = df_obs[df_obs["Year"]==2020].groupby("Feedstock")["Quantity"].sum()
county_frac_2020 = df_obs[df_obs["Year"]==2020].set_index(["County","Feedstock"])
county_frac_2020["Fraction"] = county_frac_2020["Quantity"] / county_frac_2020.index.get_level_values("Feedstock").map(national_2020)



# -----------------------------
# ML-based forecasts (Random Forest / Gradient Boosting)
# -----------------------------
from sklearn.ensemble import GradientBoostingRegressor

forecast_years = [2030, 2040, 2050]
forecast_rows = []


forecast_rows = []

for feed in feedstocks:

    df_feed = df_obs[df_obs["Feedstock"] == feed]

    ml_records = []

    for county, grp in df_feed.groupby("County"):

        q2020 = float(grp.loc[grp["Year"]==2020, "Quantity"].values[0])
        q2022 = float(grp.loc[grp["Year"]==2022, "Quantity"].values[0])

        if q2020 <= 0:
            growth = 1.0
        else:
            growth = q2022 / q2020

        ml_records.append({
            "County": county,
            "Q2020": q2020,
            "Growth": growth
        })

    ml_df = pd.DataFrame(ml_records)

    # Features
    X = ml_df[["Q2020"]]
    y = np.log(ml_df["Growth"])

    # Train model ONLY for this feedstock
    model = GradientBoostingRegressor(
        n_estimators=100,
        max_depth=3,
        random_state=42
    )

    model.fit(X, y)

    # Predict growth for national level
    national_2022 = df_feed[df_feed["Year"]==2022]["Quantity"].sum()
    
    national_2020 = df_feed[df_feed["Year"]==2020]["Quantity"].sum()
    national_growth = national_2022 / national_2020

    X_pred = pd.DataFrame([{"Q2020": national_2022}])

    log_growth_pred = model.predict(X_pred)[0]

    growth_pred_ml = np.exp(log_growth_pred)

    # shrink towards national observed growth
    growth_pred = 0.7 * growth_pred_ml + 0.3 * national_growth

    # Create national forecast
    for fy in forecast_years:

        years_passed = fy - 2022
  
        # growth gradually slows over time
        SATURATION_TIME = 50
        decay_rate = 1 / SATURATION_TIME
        effective_growth = 1 + (growth_pred - 1) * np.exp(-decay_rate * years_passed)
        national_base = national_2022 * (effective_growth ** (years_passed / 2))        


        # distribute to counties using 2020 share
        for county in counties:

            base_frac = county_frac_2020.loc[(county, feed), "Fraction"]
            delta = get_county_drift(county, feed)
            years_since_2020 = fy - 2020
            frac_t = base_frac * (1 + delta * years_since_2020)
            county_value = national_base * frac_t


            forecast_rows.append([
                county,
                feed,
                fy,
                county_value,
                growth_pred
            ])

forecast_df = pd.DataFrame(
    forecast_rows,
    columns=["County","Feedstock","Year","BaseForecast","PredictedGrowth"]
)

forecast_df["Sector"] = forecast_df["Feedstock"].apply(find_sector)





# -----------------------------
# Scenarios
# -----------------------------
scenarios = {
    "HighPolicy": {"mean_mult": 1.15, "vol_amp": 1.05, "color":"green"},
    "Baseline":   {"mean_mult": 1.00, "vol_amp": 1.00, "color":"blue"},
    "LowPolicy":  {"mean_mult": 0.85, "vol_amp": 1.00, "color":"red"}
}


# -----------------------------
# Policy economic parameters
# -----------------------------
alpha_cost = 0.35
beta_import = 0.25
gamma_export = 0.15

max_price = max(v["price"] for v in feedstock_econ.values())

sector_uncertainty = {"Agriculture":0.08,"Forestry":0.05,"Marine":0.12}
WITHIN_SECTOR_CORR = 0.30
N_ITER = 1000


# -----------------------------
# Mobilization Cost Adjustment
# -----------------------------
def mobilization_factor(feed, base_quantity):

    econ = feedstock_econ.get(feed, {"price":0,"import":0,"export":0})

    price = econ["price"]
    imp = econ["import"]
    exp = econ["export"]

    if base_quantity <= 0:
        production = 1
    else:
        production = base_quantity

    cost_penalty = 1 - alpha_cost * (price / max_price)

    import_penalty = 1 - beta_import * (imp / (imp + production + 1e-6))

    export_bonus = 1 + gamma_export * (exp / (production + 1e-6))

    factor = cost_penalty * import_penalty * export_bonus

    return max(0.2, min(1.5, factor))

# -----------------------------
# Monte Carlo
# -----------------------------
sim_records = []
sim_distributions = {}

grouped = forecast_df.groupby(["County","Sector","Year"])

for (county, sector, year), group in grouped:
    feeds = group["Feedstock"].tolist()
    n_feeds = len(feeds)

    corr = np.full((n_feeds, n_feeds), WITHIN_SECTOR_CORR)
    np.fill_diagonal(corr, 1.0)
    L = np.linalg.cholesky(corr)

    base_unc = sector_uncertainty.get(sector, 0.08)

    for scenario, svals in scenarios.items():
        mean_mult = svals["mean_mult"]
        vol_amp = svals["vol_amp"]

        # Feedstock-level correlated shocks
        Z = np.random.normal(size=(N_ITER, n_feeds))
        correlated = Z @ L.T

        # National climate/policy shock affecting all counties
        national_shock = np.random.normal(0, 0.04, size=N_ITER)

        for idx, row in group.reset_index(drop=True).iterrows():
            feed = row["Feedstock"]
            
            
            base_raw = float(row["BaseForecast"])
            mob_factor = mobilization_factor(feed, base_raw)
            base = base_raw * mob_factor
            

            sim_factors = np.exp(
               correlated[:, idx] * base_unc * vol_amp
               + national_shock
            )
            
            sim_values = base * mean_mult * sim_factors

            neg_shock = np.random.binomial(1, 0.10, size=N_ITER)
            sim_values *= (1 - neg_shock * np.random.uniform(0.05,0.20,size=N_ITER))

            sim_values = np.minimum(
                sim_values,
                base * mean_mult * max_expansion_factor[feed]
            )

            sim_records.append({
                "County": county,
                "Feedstock": feed,
                "Sector": sector,
                "Year": year,
                "Scenario": scenario,
                "Mean": float(np.mean(sim_values)),
                "P10": float(np.percentile(sim_values, 10)),
                "P90": float(np.percentile(sim_values, 90))
            })

            key = (county, feed, scenario)
            sim_distributions.setdefault(key, []).append(sim_values)

sim_df = pd.DataFrame(sim_records)

# flatten distributions
for k in sim_distributions:
    sim_distributions[k] = np.concatenate(sim_distributions[k])

# -----------------------------
# Add Observed
# -----------------------------
obs_records = []
for _, row in df_obs.iterrows():
    obs_records.append({
        "County": row["County"],
        "Feedstock": row["Feedstock"],
        "Sector": find_sector(row["Feedstock"]),
        "Year": row["Year"],
        "Scenario": "Observed",
        "Mean": row["Quantity"],
        "P10": row["Quantity"],
        "P90": row["Quantity"]
    })

obs_df = pd.DataFrame(obs_records)
sim_df = pd.concat([sim_df, obs_df], ignore_index=True)

# =====================================================
# NATIONAL AGGREGATION
# =====================================================
national_df = (
    sim_df
    .groupby(["Feedstock","Sector","Year","Scenario"], as_index=False)
    [["Mean","P10","P90"]]
    .sum()
)

national_df["County"] = "National"


# =====================================================
# Save Excel
# =====================================================
with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
    sim_df.to_excel(writer, sheet_name="CountyLevel", index=False)
    national_df.to_excel(writer, sheet_name="NationalLevel", index=False)

print(f"✅ National plots saved in folder: {NATIONAL_PLOTS_DIR}")

# =====================================================
# NATIONAL PLOTS (UNCHANGED)
# =====================================================
for feed in feedstocks:

    plt.figure(figsize=(8,6))

    obs_nat = national_df[
        (national_df["Feedstock"]==feed) &
        (national_df["Scenario"]=="Observed")
    ]

    plt.scatter(obs_nat["Year"], obs_nat["Mean"], color='black', s=80, label="Observed", zorder=5)

    for scenario, svals in scenarios.items():

        tmp = national_df[
            (national_df["Feedstock"]==feed) &
            (national_df["Scenario"]==scenario)
        ].sort_values("Year")

        plt.scatter(tmp["Year"], tmp["Mean"], color=svals["color"], s=90, label=scenario, zorder=4)

        yerr_lower = tmp["Mean"] - tmp["P10"]
        yerr_upper = tmp["P90"] - tmp["Mean"]

        plt.errorbar(tmp["Year"], tmp["Mean"],
                     yerr=[yerr_lower, yerr_upper],
                     fmt='none',
                     ecolor=svals["color"],
                     elinewidth=4,
                     alpha=0.35,
                     capsize=0,
                     zorder=3)

    plt.title(f"National - {feed}", fontsize=14, fontweight='bold')
    plt.xlabel("Year", fontsize=12)
    plt.ylabel("Biomass availability (tonnes)", fontsize=12)
    plt.grid(True, linestyle='-', linewidth=0.8, alpha=0.6)
    plt.legend(fontsize=11)
    plt.tight_layout()

    fname = f"National_{safe_filename(feed)}_uncertainty.png"
    plt.savefig(os.path.join(NATIONAL_PLOTS_DIR, fname), dpi=300)
    plt.close()
    

# -----------------------------
# Publication-style Plots
# -----------------------------
for county in counties:
    for feed in feedstocks:
        plt.figure(figsize=(8,6))

        # Observed (2020 & 2022 only)
        obs_val = sim_df[
            (sim_df["County"]==county) & (sim_df["Feedstock"]==feed) & (sim_df["Scenario"]=="Observed")
        ]

        plt.scatter(
            obs_val["Year"],
            obs_val["Mean"],
            color='black',
            s=80,
            label="Observed",
            zorder=5
        )

        # Forecast scenarios (2030+ only)
        for scenario, svals in scenarios.items():
            tmp = sim_df[
                (sim_df["County"]==county) & (sim_df["Feedstock"]==feed) & (sim_df["Scenario"]==scenario)
            ].sort_values("Year")

            plt.scatter(
                tmp["Year"],
                tmp["Mean"],
                color=svals["color"],
                s=90,
                label=scenario,
                zorder=4
            )

            yerr_lower = tmp["Mean"] - tmp["P10"]
            yerr_upper = tmp["P90"] - tmp["Mean"]

            plt.errorbar(
                tmp["Year"],
                tmp["Mean"],
                yerr=[yerr_lower, yerr_upper],
                fmt='none',
                ecolor=svals["color"],
                elinewidth=4,
                alpha=0.35,
                capsize=0,
                zorder=3
            )

        plt.title(f"{county} - {feed}", fontsize=14, fontweight='bold')
        plt.xlabel("Year", fontsize=12)
        plt.ylabel("Biomass availability (tonnes)", fontsize=12)
        plt.grid(True, linestyle='-', linewidth=0.8, alpha=0.6)
        plt.legend(fontsize=11)
        plt.tight_layout()

        safe_feed = safe_filename(feed)
        safe_county = safe_filename(county)
        fname = f"{safe_feed}_{safe_county}_uncertainty.png"

        plt.savefig(os.path.join(PLOTS_DIR, fname), dpi=300)
        plt.close()

print(f"✅ County Plots saved in folder: {PLOTS_DIR}")


# =====================================================
# Distribution Plots (FAST VERSION)
# =====================================================
dist_summary_records = []

for county in counties:
    for feed in feedstocks:

        plt.figure(figsize=(8,6))

        for scenario, svals in scenarios.items():

            key = (county, feed, scenario)

            if key not in sim_distributions:
                continue

            values = sim_distributions[key]

            plt.hist(values, bins=50, alpha=0.5, color=svals["color"], label=scenario)

            dist_summary_records.append({
                "County": county,
                "Feedstock": feed,
                "Scenario": scenario,
                "Mean": values.mean(),
                "P10": np.percentile(values,10),
                "P90": np.percentile(values,90),
                "Min": values.min(),
                "Max": values.max()
            })

        plt.title(f"Distribution - {county} - {feed}", fontsize=14, fontweight='bold')
        plt.xlabel("Biomass Output (tonnes)", fontsize=12)
        plt.ylabel("Frequency", fontsize=12)
        plt.grid(True, linestyle='-', linewidth=0.8, alpha=0.6)
        plt.legend(fontsize=11)
        plt.tight_layout()

        fname = f"Dist_{safe_filename(feed)}_{safe_filename(county)}.png"
        plt.savefig(os.path.join(DIST_PLOTS_DIR, fname), dpi=300)
        plt.close()

# Save distribution stats
dist_summary_df = pd.DataFrame(dist_summary_records)

with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    dist_summary_df.to_excel(writer, sheet_name="DistributionStats", index=False)

print(f"✅ County distribution plots saved in folder: {DIST_PLOTS_DIR}")
print(f"✅ County distribution stats saved to Excel sheet: DistributionStats")


# =====================================================
# NATIONAL DISTRIBUTION PLOTS (NEW SECTION)
# =====================================================

national_dist_records = []

for feed in feedstocks:

    plt.figure(figsize=(8,6))

    for scenario, svals in scenarios.items():

        national_values = []

        for county in counties:

            key = (county, feed, scenario)

            if key in sim_distributions:
                national_values.append(sim_distributions[key])

        if len(national_values) == 0:
            continue

        # Sum counties elementwise to get national simulation
        national_values = np.sum(np.vstack(national_values), axis=0)

        plt.hist(
            national_values,
            bins=50,
            alpha=0.5,
            color=svals["color"],
            label=scenario
        )

        national_dist_records.append({
            "County": "National",
            "Feedstock": feed,
            "Scenario": scenario,
            "Mean": national_values.mean(),
            "P10": np.percentile(national_values,10),
            "P90": np.percentile(national_values,90),
            "Min": national_values.min(),
            "Max": national_values.max()
        })

    plt.title(f"Distribution - National - {feed}", fontsize=14, fontweight='bold')
    plt.xlabel("Biomass Output (tonnes)", fontsize=12)
    plt.ylabel("Frequency", fontsize=12)
    plt.grid(True, linestyle='-', linewidth=0.8, alpha=0.6)
    plt.legend(fontsize=11)
    plt.tight_layout()

    safe_feed = safe_filename(feed)
    fname = f"Dist_National_{safe_feed}.png"

    plt.savefig(os.path.join(NATIONAL_DIST_PLOTS_DIR, fname), dpi=300)
    plt.close()


# =====================================================
# ENERGY POTENTIAL CALCULATIONS
# =====================================================

energy_records = []
sector_energy_records = []

for _, row in sim_df.iterrows():

    feed = row["Feedstock"]
    county = row["County"]
    sector = row["Sector"]
    year = row["Year"]
    scenario = row["Scenario"]

    quantity_ktonnes = row["Mean"]

    if feed not in calorific_values:
        continue

    cv = calorific_values[feed]

    # Convert units
    # ktonnes -> tonnes -> kg
    kg = quantity_ktonnes * 1e6

    energy_MJ = kg * cv
    energy_TJ = energy_MJ / 1e6

    energy_records.append({
        "County":county,
        "Feedstock":feed,
        "Sector":sector,
        "Year":year,
        "Scenario":scenario,
        "Energy_TJ":energy_TJ
    })

energy_df = pd.DataFrame(energy_records)

# Sector aggregation per county

sector_energy_df = (
    energy_df
    .groupby(["County","Sector","Year","Scenario"], as_index=False)
    ["Energy_TJ"]
    .sum()
)


# =====================================================
# BIOGAS POTENTIAL CALCULATIONS
# =====================================================

biogas_records = []

for _, row in sim_df.iterrows():

    feed = row["Feedstock"]
    county = row["County"]
    sector = row["Sector"]
    year = row["Year"]
    scenario = row["Scenario"]

    if feed not in biogas_composition:
        continue

    quantity_ktonnes = row["Mean"]

    # convert ktonnes -> tonnes
    tonnes = quantity_ktonnes * 1000

    bmp = methane_yield(feed)

    if bmp is None:
        continue

    methane_m3 = tonnes * bmp

    # energy conversion
    methane_energy_MJ = methane_m3 * 35.8
    methane_energy_TJ = methane_energy_MJ / 1e6

    biogas_records.append({
        "County":county,
        "Feedstock":feed,
        "Sector":sector,
        "Year":year,
        "Scenario":scenario,
        "Methane_m3":methane_m3,
        "Methane_Energy_TJ":methane_energy_TJ
    })

biogas_df = pd.DataFrame(biogas_records)

biogas_sector_df = (
    biogas_df
    .groupby(["County","Sector","Year","Scenario"], as_index=False)
    [["Methane_m3","Methane_Energy_TJ"]]
    .sum()
)


# Save national distribution stats
national_dist_df = pd.DataFrame(national_dist_records)

with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    
    national_dist_df.to_excel(writer, sheet_name="NationalDistributionStats", index=False)
    
    energy_df.to_excel(writer, sheet_name="FeedstockEnergy", index=False)
    
    sector_energy_df.to_excel(writer, sheet_name="SectorEnergyCounty", index=False)
    
    biogas_df.to_excel(writer, sheet_name="BiogasFeedstock", index=False)

    biogas_sector_df.to_excel(writer, sheet_name="BiogasSectorCounty", index=False)

print(f"✅ National distribution plots saved in folder: {NATIONAL_DIST_PLOTS_DIR}")
print("✅ National distribution stats saved to Excel sheet: NationalDistributionStats")