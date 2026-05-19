# -*- coding: utf-8 -*-
"""
Product Groups extractor — unieke productomschrijvingen uit scanbewegingen.

Doel:
    Haal alle unieke productomschrijvingen uit de brondata en tel hoeveel
    unieke pallets er per product zijn. Output dient als basis voor de
    product_groups.xlsx mapping (Omschrijving → ProductGroup + Bedrijf)
    die de klant invult en die de hoofdpipeline gebruikt voor verrijking.

Output: GOLD/product_groups.xlsx met twee tabs:
    - unique_oms:    lijst van unieke omschrijvingen (voor klant om in te vullen)
    - counts_by_oms: telling unieke pallets per omschrijving (voor prioritering)

Aanpak:
    - Per pallet (Dragernr.) de laatste niet-lege omschrijving pakken
    - Een pallet telt maximaal 1x mee per omschrijving
    - Casing behouden (bewust, voor klantcommunicatie)
"""

from pathlib import Path
import numpy as np
import pandas as pd

# =================== CONFIGURATIE ===================
BASE_DIR     = Path("/Users/kars/Projects/activitybasedcosting")
BRONZE_DIR   = BASE_DIR / "data" / "bronze"
GOLD_DIR     = BASE_DIR / "data" / "gold"

INPUT_XLSX   = BRONZE_DIR / "Copy of Scanbewegingen 11-01 tm 18-01.xlsx"
OUTPUT_XLSX  = GOLD_DIR / "product_groups.xlsx"

UNIQUE_SHEET = "unique_oms"
COUNT_SHEET  = "counts_by_oms"

COL_DRAGER = "Dragernr."
COL_OMS    = "Omschrijving"

# =================== HELPERS ===================
def norm_text(s: pd.Series) -> pd.Series:
    """Trim whitespace, collapse meerdere spaties. Behoud hoofdletters."""
    return s.fillna("").astype(str).str.strip().str.replace(r"\s+", " ", regex=True)

# =================== MAIN ===================
def main():
    GOLD_DIR.mkdir(parents=True, exist_ok=True)
    if not INPUT_XLSX.exists():
        raise FileNotFoundError(f"Bronze input niet gevonden: {INPUT_XLSX}")

    df = pd.read_excel(INPUT_XLSX, engine="openpyxl", usecols=lambda c: c in {COL_DRAGER, COL_OMS})
    for c in (COL_DRAGER, COL_OMS):
        if c not in df.columns:
            raise KeyError(f"Verplichte kolom ontbreekt: {c}")

    df["OmschrijvingN"] = norm_text(df[COL_OMS])
    df = df.reset_index(drop=True)
    df["_row"] = np.arange(len(df), dtype=int)

    # Per pallet: laatste niet-lege omschrijving (representatief)
    df["OmsNonEmpty"] = df["OmschrijvingN"].replace("", pd.NA)
    last_per_pallet = (
        df.sort_values("_row")
          .groupby(COL_DRAGER, sort=False)["OmsNonEmpty"]
          .agg(lambda s: s.dropna().iloc[-1] if s.dropna().size else pd.NA)
          .reset_index()
          .rename(columns={"OmsNonEmpty": "Omschrijving"})
    )
    last_per_pallet = last_per_pallet.dropna(subset=["Omschrijving"])

    # Unieke omschrijvingen (voor klant om ProductGroup/Bedrijf bij te vullen)
    unique_oms = (
        last_per_pallet[["Omschrijving"]]
        .drop_duplicates(ignore_index=True)
        .assign(_key=lambda d: d["Omschrijving"].str.casefold())
        .sort_values("_key").drop(columns="_key").reset_index(drop=True)
    )

    # Aantallen unieke pallets per omschrijving (voor prioritering)
    counts = (
        last_per_pallet
        .groupby("Omschrijving", as_index=False, sort=False)[COL_DRAGER].nunique()
        .rename(columns={COL_DRAGER: "Unieke_pallets"})
        .assign(_key=lambda d: d["Omschrijving"].str.casefold())
        .sort_values(["Unieke_pallets", "_key"], ascending=[False, True])
        .drop(columns="_key").reset_index(drop=True)
    )

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl", mode="w") as xw:
        unique_oms.to_excel(xw, index=False, sheet_name=UNIQUE_SHEET)
        counts.to_excel(xw, index=False, sheet_name=COUNT_SHEET)

    print(f"✅ Unieke omschrijvingen: {len(unique_oms)}")
    print(f"✅ Telling per omschrijving: {len(counts)}")
    print(f"📄 Output: {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()