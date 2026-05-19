# -*- coding: utf-8 -*-
"""
Warehouse bezetting checker — reconstructeer per dag waar elk pallet staat.

Doel:
    Visualiseer de dagelijkse bezetting per hal. Nuttig voor:
    - Capaciteitsplanning (welke hallen lopen vol?)
    - Seizoenspatronen (met jaardata)
    - Validatie van storage_days output

Werkwijze:
    - Elke scan is een "pallet staat nu op locatie X" event
    - Per dag per pallet: laatst bekende locatie
    - Forward-fill: als geen scan op een dag → locatie van vorige dag behouden
    - Locatiecode → halnummer mapping (bijv. 6F-228 → HAL6)
    - Machine-locaties (F1_IN etc.) en exits (VERZEND, KWIJT) niet meerekenen

Bekende beperking:
    Met kort datavenster (1 week) zijn veel pallets "vastgeprikt" op hun
    laatste locatie. Bijv. 500+ pallets op HAL5 (docks) = pallets die
    binnenkwamen maar waarvan de vervolgscan buiten de meetperiode valt.
    Met jaardata verdwijnt dit grotendeels.

Capaciteiten:
    Alleen ingevuld waar bekend uit stellingindeling van de klant.
    HAL2, HAL3, HAL5, HAL6 capaciteit is onbekend.
    HAL3 en HAL5 zijn geen opsaglocaties (staging/docks).
"""

from pathlib import Path
import pandas as pd
import numpy as np
import re
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# =================== CONFIGURATIE ===================
BASE_DIR = Path("/Users/kars/Projects/activitybasedcosting")
INPUT_XLSX = BASE_DIR / "data" / "bronze" / "Copy of Scanbewegingen 11-01 tm 18-01.xlsx"

COL_DRAGER = "Dragernr."
COL_LOC = "Opslaglocatie"
COL_DATE = "Registratiedatum"
COL_TIME = "Registratietijd"

# Locaties die NIET als "in warehouse" tellen
MACHINE_LOCS = {"F1_IN", "F2_IN", "F3_IN", "F4_IN", "F5_IN", "P1_IN", "P2_IN", "A1_IN"}
EXIT_LOCS = {"VERZEND", "KWIJT"}

# Bekende capaciteiten (palletplekken, NIET × verdiepingen — platte telling per stelling)
CAPACITEIT = {
    "HAL1": 60,    # A(24) + R(36)
    "HAL4": 39,    # B(39)
    "HAL7": 153,   # A(45) + B+C(108)
    "HAL8": 153,   # zelfde als HAL7
}

# =================== HELPERS ===================
def parse_timestamp(df: pd.DataFrame) -> pd.Series:
    """Zelfde timestamp-parser als warehouse_model.py."""
    d = pd.to_datetime(df[COL_DATE], errors="coerce", dayfirst=True).dt.strftime("%Y-%m-%d")
    raw = df[COL_TIME].astype(str).str.strip().str.replace(",", ".", regex=False)
    def fix_time(s: str) -> str:
        m = re.match(r"^(\d{1,2}):(\d{2}):(\d{2})(?:\.(\d{1,6}))?$", s)
        if not m: return s
        hh, mm, ss, frac = m.groups()
        hh = hh.zfill(2)
        frac = "000" if not frac else (frac + "000")[:3]
        return f"{hh}:{mm}:{ss}.{frac}"
    fixed = raw.apply(fix_time)
    return pd.to_datetime(d + " " + fixed, errors="coerce")

def extract_hal(loc: str) -> str:
    """
    Map locatiecode naar halnaam.
    
    Patronen:
    - HAL1, HAL5          → HAL1, HAL5
    - 6F-228, 2R-009      → HAL6, HAL2 (eerste cijfer = halnummer)
    - 7-1-123             → HAL7
    - F1_IN, A1_OUT       → MACHINE
    - DOK2                → DOK2
    - VERZEND, KWIJT      → EXIT
    - Onbekend            → OVERIG
    """
    s = str(loc).strip().upper()
    if not s or s == "NAN": return "ONBEKEND"
    if s in MACHINE_LOCS: return "MACHINE"
    m = re.match(r"^[A-Z]\d+_(IN|OUT)$", s)
    if m: return "MACHINE"
    if s in EXIT_LOCS: return "EXIT"
    m = re.match(r"^HAL(\d+)$", s)
    if m: return f"HAL{m.group(1)}"
    m = re.match(r"^DOK(\d+)$", s)
    if m: return f"DOK{m.group(1)}"
    m = re.match(r"^(\d+)[A-Z]-\d+$", s)
    if m: return f"HAL{m.group(1)}"
    m = re.match(r"^(\d+)-\d+-\d+$", s)
    if m: return f"HAL{m.group(1)}"
    return "OVERIG"

def is_in_warehouse(hal: str) -> bool:
    """True als het pallet fysiek in het warehouse staat (niet in machine/verzonden)."""
    return hal not in {"MACHINE", "EXIT", "ONBEKEND"}

# =================== MAIN ===================
def main():
    print("Laden van scanbewegingen...")
    df = pd.read_excel(INPUT_XLSX, engine="openpyxl")
    df["Timestamp"] = parse_timestamp(df)
    df["Loc"] = df[COL_LOC].fillna("").astype(str).str.strip().str.upper()
    df["Drager"] = df[COL_DRAGER].fillna("").astype(str).str.strip()
    df = df[df["Timestamp"].notna() & (df["Drager"] != "") & (df["Loc"] != "")].copy()
    df = df.sort_values(["Drager", "Timestamp"]).reset_index(drop=True)
    print(f"Scans geladen: {len(df):,} | Unieke pallets: {df['Drager'].nunique():,}")

    events = df[["Drager", "Timestamp", "Loc"]].copy()
    events["Dag"] = events["Timestamp"].dt.date

    min_dag = events["Dag"].min()
    max_dag = events["Dag"].max()
    alle_dagen = pd.date_range(min_dag, max_dag, freq="d").date
    print(f"Periode: {min_dag} t/m {max_dag} ({len(alle_dagen)} dagen)")

    # Per pallet per dag: laatst bekende locatie
    last_per_dag = (
        events.sort_values(["Drager", "Timestamp"])
        .groupby(["Drager", "Dag"]).agg(Loc=("Loc", "last")).reset_index()
    )

    # Volledig grid (pallet × dag) met forward-fill
    print("Bezetting reconstructeren per dag...")
    all_pallets = sorted(events["Drager"].unique())
    idx = pd.MultiIndex.from_product([all_pallets, alle_dagen], names=["Drager", "Dag"])
    grid = pd.DataFrame(index=idx).reset_index()
    grid = grid.merge(last_per_dag, on=["Drager", "Dag"], how="left")
    grid = grid.sort_values(["Drager", "Dag"])
    grid["Loc"] = grid.groupby("Drager")["Loc"].ffill()

    # Verwijder rijen vóór eerste scan (pallet was nog niet in warehouse)
    first_scan = events.groupby("Drager")["Dag"].min().reset_index().rename(columns={"Dag": "EersteDag"})
    grid = grid.merge(first_scan, on="Drager", how="left")
    grid = grid[grid["Dag"] >= grid["EersteDag"]].drop(columns=["EersteDag"])
    grid = grid[grid["Loc"].notna() & (grid["Loc"] != "")].copy()

    # Hal toewijzen
    grid["Hal"] = grid["Loc"].apply(extract_hal)
    grid["InWarehouse"] = grid["Hal"].apply(is_in_warehouse)

    # Debug: welke locaties vallen in OVERIG?
    overig = grid[grid["Hal"] == "OVERIG"]["Loc"].value_counts()
    if not overig.empty:
        print(f"\nOVERIG locaties:")
        for loc, cnt in overig.items():
            print(f"  {loc:<25} {cnt}")

    # --- Bezetting per hal per dag ---
    wh = grid[grid["InWarehouse"]].copy()
    bezetting_dag = wh.groupby(["Dag", "Hal"]).agg(Pallets=("Drager", "nunique")).reset_index()
    totaal_dag = wh.groupby("Dag").agg(Pallets=("Drager", "nunique")).reset_index()

    # --- Piekbezetting per hal ---
    print("\n" + "=" * 70)
    print(f"{'Hal':<15} {'Gem.':>8} {'Piek':>8} {'Piek dag':>12} {'Capaciteit':>11} {'Piek %':>8}")
    print("=" * 70)
    for hal in sorted(bezetting_dag["Hal"].unique()):
        sub = bezetting_dag[bezetting_dag["Hal"] == hal]
        gem = sub["Pallets"].mean()
        piek_idx = sub["Pallets"].idxmax()
        piek = sub.loc[piek_idx, "Pallets"]
        piek_dag = sub.loc[piek_idx, "Dag"]
        cap = CAPACITEIT.get(hal, None)
        cap_str = str(cap) if cap else "?"
        piek_pct = f"{100 * piek / cap:.0f}%" if cap else "?"
        print(f"{hal:<15} {gem:>8.0f} {piek:>8} {str(piek_dag):>12} {cap_str:>11} {piek_pct:>8}")
    print("-" * 70)
    print(f"{'TOTAAL':<15} {totaal_dag['Pallets'].mean():>8.0f} {totaal_dag['Pallets'].max():>8} {str(totaal_dag.loc[totaal_dag['Pallets'].idxmax(), 'Dag']):>12}")

    # --- Dagelijks detail (ASCII bar chart) ---
    print(f"\n\nDagelijkse totale bezetting:")
    print("=" * 40)
    for _, r in totaal_dag.iterrows():
        bar = "█" * (r["Pallets"] // 20)
        print(f"  {r['Dag']}  {r['Pallets']:>6}  {bar}")

    # --- Plots ---
    fig, axes = plt.subplots(2, 1, figsize=(14, 10))
    fig.suptitle("Warehouse bezetting over tijd", fontsize=14, fontweight="bold")

    ax1 = axes[0]
    ax1.bar(totaal_dag["Dag"], totaal_dag["Pallets"], color="#EC7A00", edgecolor="white")
    ax1.set_ylabel("Pallets in warehouse")
    ax1.set_title("Totale bezetting per dag")
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%d-%b"))

    ax2 = axes[1]
    pivot = bezetting_dag.pivot_table(index="Dag", columns="Hal", values="Pallets", fill_value=0)
    hal_order = pivot.mean().sort_values(ascending=True).index.tolist()
    pivot = pivot[hal_order]
    colors = plt.cm.tab20(np.linspace(0, 1, len(hal_order)))
    pivot.plot(kind="bar", stacked=True, ax=ax2, color=colors, edgecolor="white", linewidth=0.3)
    ax2.set_ylabel("Pallets")
    ax2.set_title("Bezetting per hal per dag (gestapeld)")
    ax2.legend(bbox_to_anchor=(1.01, 1), loc="upper left", fontsize=8)
    ax2.set_xticklabels([d.strftime("%d-%b") for d in pivot.index], rotation=45)

    plt.tight_layout()
    out_path = BASE_DIR / "data" / "bezetting_historisch.png"
    plt.savefig(out_path, dpi=150, bbox_inches="tight")
    print(f"\n📊 Plot opgeslagen: {out_path}")
    plt.show()

if __name__ == "__main__":
    main()