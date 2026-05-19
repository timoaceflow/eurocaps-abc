# -*- coding: utf-8 -*-
"""
Shift-uren checker — schat totale "ingeklokte" uren per user.

Doel:
    Bepaal hoeveel uur elke warehouse-medewerker aanwezig was, op basis van
    eerste en laatste scan per shift. Dit getal gebruik je als totaal betaalde
    uren in het ABC-model (de kostenbasis die je verdeelt over activiteiten).

Werkwijze:
    - Sorteer scans per user op tijd
    - Als er een gap > SHIFT_BREAK_SEC tussen twee scans zit → nieuwe shift
    - Shiftduur = laatste scan - eerste scan per shift
    - EURO-CAPS users worden overgeslagen (administratieve scans)
    - Twee rijen = 1 scan (FROM + TO), dus scantelling wordt gehalveerd

Belangrijk:
    Dit is een BENADERING — geen echte klokdata. De drempel van 4 uur is
    gekozen zodat pauzes en niet-scan werkzaamheden (2-3 uur) binnen een
    shift vallen, maar echte shiftwisselingen (>4 uur) worden gesplitst.
    
    Vergelijking met de batch pipeline:
    - Batch pipeline: ~85-100 directe uren (tijd besteed aan scanbare taken)
    - Dit script: ~200-225 totale uren (ingeklokte/aanwezige tijd)
    - Verschil: indirecte tijd (looptijd, wachttijd, pauze, overleg)
    - Productiviteitsratio: ~44% — normaal voor warehouse zonder route-optimalisatie
"""

from pathlib import Path
import pandas as pd
import re

# =================== CONFIGURATIE ===================
INPUT_XLSX      = Path("/Users/kars/Projects/activitybasedcosting/data/bronze/Copy of Scanbewegingen 11-01 tm 18-01.xlsx")
SHIFT_BREAK_SEC = 4 * 3600  # Gap > 4 uur → nieuwe shift (gekozen op basis van typische shiftwisseling)

COL_USER = "Gebruikers-ID"
COL_DATE = "Registratiedatum"
COL_TIME = "Registratietijd"

EXCLUDE_USER_PREFIXES = ["EURO-CAPS"]

# =================== HELPERS ===================
def parse_timestamp(df: pd.DataFrame) -> pd.Series:
    """Combineer datum + tijd kolom tot Timestamp (zelfde logica als warehouse_model.py)."""
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

def fmt_hours(seconds: float) -> str:
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    return f"{h}u{m:02d}m"

def fmt_mmss(seconds: float) -> str:
    s = int(round(seconds))
    m, s = divmod(s, 60)
    return f"{m:02d}:{s:02d}"

# =================== MAIN ===================
def main():
    df = pd.read_excel(INPUT_XLSX, engine="openpyxl")
    df["Timestamp"] = parse_timestamp(df)
    df["User"] = df[COL_USER].fillna("").astype(str).str.strip()

    df = df[df["Timestamp"].notna() & (df["User"] != "")].copy()
    if EXCLUDE_USER_PREFIXES:
        prefixes = tuple(p.upper() for p in EXCLUDE_USER_PREFIXES)
        df = df[~df["User"].str.upper().str.startswith(prefixes)]

    df = df.sort_values(["User", "Timestamp"]).reset_index(drop=True)

    # Shifts detecteren: nieuwe shift als gap > drempel
    df["PrevTs"] = df.groupby("User")["Timestamp"].shift(1)
    df["GapSec"] = (df["Timestamp"] - df["PrevTs"]).dt.total_seconds()
    df["NewShift"] = (df["GapSec"].isna()) | (df["GapSec"] > SHIFT_BREAK_SEC)
    df["ShiftId"] = df.groupby("User")["NewShift"].cumsum()

    shifts = (
        df.groupby(["User", "ShiftId"], sort=False)
        .agg(ShiftStart=("Timestamp", "min"), ShiftEnd=("Timestamp", "max"), Scans=("Timestamp", "count"))
        .reset_index()
    )
    shifts["Duration_s"] = (shifts["ShiftEnd"] - shifts["ShiftStart"]).dt.total_seconds()
    shifts = shifts.sort_values(["User", "ShiftStart"]).reset_index(drop=True)

    # Twee rijen = 1 echte scan (FROM + TO locatie)
    shifts["Scans_real"] = (shifts["Scans"] / 2).astype(int)

    # --- Output: detail per user per shift ---
    print("=" * 100)
    print(f"{'User':<15} {'Shift':<6} {'Start':>20} {'Eind':>20} {'Duur':>10} {'Scans':>6} {'Tijd/scan':>10}")
    print("=" * 100)

    user_totals = []
    for user, grp in shifts.groupby("User", sort=True):
        total_s = 0
        total_scans = 0
        for i, (_, r) in enumerate(grp.iterrows(), 1):
            start_str = r["ShiftStart"].strftime("%Y-%m-%d %H:%M")
            end_str = r["ShiftEnd"].strftime("%Y-%m-%d %H:%M")
            dur_str = fmt_hours(r["Duration_s"])
            scans = int(r["Scans_real"])
            per_scan = fmt_mmss(r["Duration_s"] / scans) if scans > 0 else "--:--"
            print(f"{user:<15} {i:<6} {start_str:>20} {end_str:>20} {dur_str:>10} {scans:>6} {per_scan:>10}")
            total_s += r["Duration_s"]
            total_scans += scans
        total_str = fmt_hours(total_s)
        total_per_scan = fmt_mmss(total_s / total_scans) if total_scans > 0 else "--:--"
        print(f"{'':<15} {'TOTAAL':<6} {'':>20} {'':>20} {total_str:>10} {total_scans:>6} {total_per_scan:>10}")
        print("-" * 100)
        user_totals.append({"User": user, "Shifts": len(grp), "Totaal_s": total_s, "Totaal": total_str, "Scans": total_scans, "PerScan": total_per_scan})

    # --- Output: samenvatting ---
    print("\n" + "=" * 70)
    print(f"{'User':<15} {'Shifts':>7} {'Totaal uren':>15} {'Scans':>8} {'Tijd/scan':>10}")
    print("=" * 70)
    for t in sorted(user_totals, key=lambda x: x["Totaal_s"], reverse=True):
        print(f"{t['User']:<15} {t['Shifts']:>7} {t['Totaal']:>15} {t['Scans']:>8} {t['PerScan']:>10}")
    print("=" * 70)
    grand_total = sum(t["Totaal_s"] for t in user_totals)
    grand_scans = sum(t["Scans"] for t in user_totals)
    grand_per_scan = fmt_mmss(grand_total / grand_scans) if grand_scans > 0 else "--:--"
    print(f"{'GRAND TOTAL':<15} {'':>7} {fmt_hours(grand_total):>15} {grand_scans:>8} {grand_per_scan:>10}")

if __name__ == "__main__":
    main()