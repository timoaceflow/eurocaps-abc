# -*- coding: utf-8 -*-
"""
===================================================================================
WAREHOUSE MODEL EXPORTER — Activity Based Costing Pipeline
===================================================================================

Doel:
    Dit script verwerkt ruwe scanbewegingen uit een warehouse (EuroCaps) tot een
    gestructureerd datamodel voor Activity Based Costing (ABC) in Power BI.
    Het produceert één Excel workbook (warehouse_model.xlsx) met meerdere tabs
    die als databron dienen voor het Power BI rapport "EuroCaps - ABC".

Twee kernmetingen:
    1. PICK/HANDLING-TIJD PER PALLET — Hoe lang duurt het om één pallet te verplaatsen,
       per type activiteit? Gemeten via batches van micropair-events (sec/pallet).
    2. STORAGE-DUUR PER PALLET — Hoe lang staat een pallet in de stelling?
       Gemeten via een episode-state-machine (dagen).

Databron:
    - Ruwe scanbewegingen uit Business Central (Excel export)
    - Elke fysieke palletverplaatsing genereert 2 rijen: FROM-locatie en TO-locatie,
      ~0.03ms na elkaar (een "micropair")
    - Administratieve scans (EURO-CAPS users) genereren soms singletons (1 rij)

Output tabs in warehouse_model.xlsx:
    - storage_days:      Één rij per uniek pallet (Dragernr.) met StorageTime, PickTime,
                         StorageDays, CurrentLocation, HasStorage, HasPick
    - batches:           Één rij per batch (groep opeenvolgende verplaatsingen door één user)
                         met Duration_s, Sec per pallet, Activity label
    - batch_pallets:     Bridge-tabel: BatchId ↔ Dragernr., verrijkt met ProductGroup,
                         Bedrijf, Activity. Bevat ook pallets zonder batchactiviteit
                         (EURO-CAPS-only) voor complete bridge in Power BI
    - activity_segments: Lineaire tijdlijn per user met Work- en Idle-segmenten
    - activity_segments_qa: Validatietabel (Work-som = Batch-som, Idle-som = Expected)
    - kosten_parameters: Parametertabel voor kostenberekeningen in Power BI

Power BI relaties:
    - batches[BatchId] (1) → batch_pallets[BatchId] (*)
    - storage_days[Dragernr.] (1) → batch_pallets[Dragernr.] (*) — BIDIRECTIONEEL
    - batches[BatchId] (1) → activity_segments[BatchId] (*)
    - DimDate[Date] (1) → batches[BatchDate] (*)

Versiegeschiedenis:
    v2.9  — Eerste versie met activity segments
    v2.10 — Interne Relocatie als 5e activiteit (Magazijndagboek + TRANSFER)
    v2.11 — Batch-scan correctie voor snelle scans (<20s avg gap)
    v2.12 — Complete bridge: ontbrekende pallets aangevuld in batch_pallets
            + OPSLAG-NB toegevoegd als pick-trigger locatie

Auteur: Kars @ Aceflow
===================================================================================
"""

import os
from pathlib import Path
import pandas as pd
import numpy as np
import re
import hashlib
from dotenv import load_dotenv
load_dotenv()

# ===================================================================================
# PADEN — Pas INPUT_XLSX aan als je een ander bronbestand wilt gebruiken.
# De mappenstructuur volgt bronze/silver/gold patroon:
#   bronze = ruwe input (scanbewegingen, product_groups)
#   silver = tussenresultaten (optioneel, voor QA)
#   gold   = definitieve output voor Power BI
# ===================================================================================
# BRONZE_DIR = Path(os.environ["BRONZE_DIR"])
# SILVER_DIR = Path(os.environ["SILVER_DIR"])
# GOLD_DIR = Path(os.environ["GOLD_DIR"])

SHAREPOINT_LOC      = Path(os.environ["SHAREPOINT_LOC"])
INPUT_XLSX          = SHAREPOINT_LOC / Path(os.environ["INPUT_FILE"])
PRODUCT_GROUPS_XLSX = SHAREPOINT_LOC / Path(os.environ["PRODUCT_GROUPS_FILE"])  # Mapping: Omschrijving → ProductGroup + Bedrijf
OUTPUT_XLSX         = SHAREPOINT_LOC   / Path(os.environ["OUTPUT_FILE"])

# Optioneel: exporteer silver-events voor handmatige QA/debugging
SAVE_SILVER_EVENTS  = False
SILVER_EVENTS_XLSX  = SHAREPOINT_LOC / "silver_storage_events.xlsx"
SILVER_EVENTS_SHEET = "silver_storage"

# Sheet-namen in de output Excel
SHEET_STORAGE_DAYS      = "storage_days"
SHEET_BATCHES           = "batches"
SHEET_BATCH_PALLETS     = "batch_pallets"
SHEET_ACTIVITY_SEGMENTS = "activity_segments"
SHEET_ACTIVITY_QA       = "activity_segments_qa"
SHEET_KOSTEN_PARAMS     = "kosten_parameters"

# ===================================================================================
# KOLOMNAMEN — Deze moeten exact matchen met de kolomnamen in het bronbestand.
# Als de klant de export wijzigt, moeten deze mee aangepast worden.
# ===================================================================================
COL_DRAGER   = "Dragernr."         # Unieke pallet-ID (barcode/SSCC)
COL_USER     = "Gebruikers-ID"     # Wie de scan uitvoert (bijv. MALAE, EURO-CAPS\YWISMAN)
COL_LOC      = "Opslaglocatie"     # Locatiecode (bijv. HAL5, 6F-228, F1_IN)
COL_ZONE     = "Zone"
COL_DATE     = "Registratiedatum"  # Datum van scan (DD-MM-YYYY formaat)
COL_TIME     = "Registratietijd"   # Tijd van scan (HH:MM:SS.mmm formaat)
COL_DOCMAG   = "Mag.-documentsoort"  # Bepaalt Activity: Ontvangst, Verzending, Magazijndagboek, 9
COL_DOCSOORT = "Documentsoort"
COL_ACTIE    = "Actiesoort"
COL_OMS      = "Omschrijving"      # Productomschrijving, gebruikt voor ProductGroup join
COL_REFDOC   = "Referentiedocument"  # Bijv. "Opslag", "Pick" — helpt bij classificatie
COL_DAGBOEK  = "Dagboeksjabloon"   # Bijv. "TRANSFER", "STANDAARD" — onderscheidt Interne Relocatie
COL_MAGDOCNO = "Mag.-documentnr."  # Verzendingsnummer (GroupKey voor Outbound batches)
COL_ORDER    = "Bronnr."           # Ordernummer (GroupKey voor Inbound batches)
COL_LOT      = "Lotnr."

# ===================================================================================
# PARAMETERS — STORAGE DAYS
#
# De storage pipeline bepaalt per pallet:
#   - StorageTime: wanneer het pallet de stelling in ging
#   - PickTime: wanneer het pallet de stelling weer verliet
#   - StorageDays: verschil in dagen
#
# Events worden gedetecteerd via "micropairs": twee opeenvolgende rijen voor
# hetzelfde pallet, <1 seconde apart, met verschillende locaties.
# Singletons (losse rijen zonder pair) worden ook meegenomen voor completere
# episode-detectie — belangrijk voor administratieve scans (EURO-CAPS).
# ===================================================================================
MICROPAIR_MAX_SEC           = 1.0    # Max tijdsverschil tussen twee rijen om als micropair te gelden
USE_FALLBACK_EVENTS         = False
MICROPAIR_REQUIRE_SAME_USER = False  # Micropairs hoeven niet door dezelfde user gescand te zijn
ENABLE_SINGLETON_EVENTS     = True   # Ook losse rijen (geen pair) meenemen als events

# Locaties die een PICK-event triggeren (pallet verlaat de storage):
# - F*_IN/P*_IN/A*_IN: machine-ingangen (pallet gaat productielijn in)
# - KWIJT: pallet is kwijtgeraakt / afgeboekt
# - OPSLAG-NB: "Opslag Niet Bepaald" — pallet gaat naar externe opslag, niet meer ons probleem
MACHINE_IN_LOCS = {"F1_IN","F2_IN","F3_IN","F4_IN","F5_IN","P1_IN","P2_IN","A1_IN","KWIJT","OPSLAG-NB"}

# Locaties die NIET als opslaglocatie gelden (geen storage-event triggeren)
NON_STORAGE_TOKENS = {"VERZEND","KWIJT","ONTVANGST","OPSLAG-NB"}

# ===================================================================================
# PARAMETERS — PICK/BATCHES
#
# De batch pipeline groepeert micropair-events tot "batches": een aaneengesloten reeks
# verplaatsingen door één user, voor één documentgroep/order.
#
# EURO-CAPS users worden UITGESLOTEN uit de batch pipeline — hun scans zijn
# administratief (soms 100 scans in <1 seconde) en zouden de picktijd per pallet
# onbetrouwbaar maken. Ze worden WEL meegenomen in de storage pipeline.
#
# Belangrijke business rules voor batchtijden:
#
# 1. IDLE-CHAINING (≤25 min):
#    Als er een volgende batch start binnen 25 min na het laatste event van de
#    huidige batch, dan wordt de starttijd van die volgende batch als eindtijd
#    van de huidige batch gepakt. Rationale: de idle-tijd ertussen is "afrondtijd"
#    (scanner wegleggen, naar volgende locatie rijden).
#
# 2. MACRO-PAUZE SPLIT (>25 min):
#    Als er binnen dezelfde DocGroup/GroupKey een gap >25 min zit, wordt de batch
#    gesplitst. Rationale: medewerker is duidelijk gestopt en later hervat.
#
# 3. BATCH-SCAN CORRECTIE (<20s avg gap, niet Verzending):
#    Als de gemiddelde gap per pallet binnen een batch <20s is, dan heeft de
#    medewerker waarschijnlijk eerst alles gescand en daarna pas verplaatst.
#    In dat geval: BatchEnd = LastEvent + N × DocGroup-gemiddelde (in plaats van
#    LastEvent + 1 × AvgGap, wat te kort zou zijn).
#    Verzending is uitgezonderd omdat daar het scan-en-loop patroon normaal is.
#
# 4. N=1 FALLBACK:
#    Batches met maar 1 pallet krijgen als eindtijd: LastEvent + DocGroup-gemiddelde
#    (berekend uit batches met N≥2). Als dat niet beschikbaar is: globale mediaan.
#    Als dat ook niet: 120 seconden.
#
# Deze drempelwaarden zijn bepaald op basis van een gap-analyse (histogram van
# tijdsgaps tussen opeenvolgende scans). Zie gap_analyse.py voor de onderbouwing.
# Bij 20-25 min zakt de frequentie van gaps significant, wat wijst op een
# natuurlijk breekpunt tussen "werktijd" en "pauze".
# ===================================================================================
MICRO_PAIR_SEC               = 1.0       # Zelfde als MICROPAIR_MAX_SEC, voor batch-events
CHAIN_MAX_IDLE_SEC           = 25 * 60   # Max idle die naar vorige batch wordt gechained
MACRO_PAUSE_SPLIT_SEC        = 25 * 60   # Gap waarboven een batch wordt gesplitst
DOCGROUP_SINGLE_FALLBACK_SEC = 120       # Fallback eindtijd voor N=1 batches zonder DocGroup-gem
BATCH_SCAN_CORRECTION_SEC    = 20        # Drempel voor batch-scan correctie (avg gap < dit)
EXCLUDE_USER_PREFIXES_PICK   = ["EURO-CAPS"]  # Users die NIET in de batch pipeline komen

# ===================================================================================
# PARAMETERS — ACTIVITY LABELS & IDLE SEGMENTEN
#
# Elke batch krijgt een Activity-label op basis van Mag.-documentsoort (DocGroup):
#   - Ontvangst        → Inbound (pallets binnenkomen)
#   - Verzending       → Outbound (pallets vertrekken)
#   - Magazijndagboek  → Klaarzetten (pallets binnen warehouse verplaatsen)
#   - Magazijndagboek + Dagboeksjabloon=TRANSFER → Interne Relocatie (reorganisatie)
#   - Doc9 (=code "9") → Machine-aanvoer (pallet naar productielijn)
#
# Activity segments vormen een lineaire tijdlijn per user:
#   - Work-segmenten: exact de batchblokken (incl. gechained idle ≤25min)
#   - Idle-segmenten: gaps >30min tussen batches (apart, Activity = "Idle/Pauze")
#
# Let op: IDLE_GAP_THRESHOLD_SEC (30min) is bewust hoger dan CHAIN_MAX_IDLE_SEC (25min).
# Gaps van 25-30 min worden niet gechained naar de vorige batch, maar ook niet als
# apart Idle-segment getoond. Ze "verdwijnen" — dit is gewenst, want het zijn korte
# pauzes die niet als productieve of als echte idle-tijd gelden.
# ===================================================================================
IDLE_GAP_THRESHOLD_SEC               = 30 * 60
WRITE_ACTIVITY_TO_BATCH_PALLETS      = True   # Activity ook op de bridge-tabel zetten
WRITE_ACTIVITY_SEGMENTS_QA_SHEET     = True   # QA-tab meeschrijven in output
ACTIVITY_IDLE_LABEL                  = "Idle/Pauze"
ACTIVITY_MAP = {
    "Ontvangst":        "Inbound",
    "Verzending":       "Outbound",
    "Magazijndagboek":  "Klaarzetten",
    "Doc9":             "Machine-aanvoer",
}

DEBUG = True
PIPELINE_VERSION = "v2.12-abc-complete-bridge"

# ===================================================================================
# SHARED HELPERS
# ===================================================================================

def parse_timestamp(df: pd.DataFrame, col_date=COL_DATE, col_time=COL_TIME) -> pd.Series:
    """
    Combineer Registratiedatum + Registratietijd tot één pandas Timestamp.
    
    De brondata heeft soms inconsistente tijdformaten (bijv. "8:05:03" ipv "08:05:03",
    komma's als decimaalscheider). Deze functie normaliseert dat.
    """
    d = pd.to_datetime(df[col_date], errors="coerce", dayfirst=True).dt.strftime("%Y-%m-%d")
    raw = df[col_time].astype(str).str.strip().str.replace(",", ".", regex=False)
    def fix_time(s: str) -> str:
        m = re.match(r"^(\d{1,2}):(\d{2}):(\d{2})(?:\.(\d{1,6}))?$", s)
        if not m: return s
        hh, mm, ss, frac = m.groups()
        hh = hh.zfill(2)
        frac = "000" if not frac else (frac + "000")[:3]
        return f"{hh}:{mm}:{ss}.{frac}"
    fixed = raw.apply(fix_time)
    return pd.to_datetime(d + " " + fixed, errors="coerce")

def mmss(seconds_per_pallet: float) -> str:
    """Formatteer seconden naar mm:ss string (bijv. 112.6 → '01:53')."""
    v = 0.0 if pd.isna(seconds_per_pallet) else float(seconds_per_pallet)
    s = int(np.ceil(v)) if v > 0 else 0
    m, s = divmod(s, 60)
    return f"{m:02d}:{s:02d}"

def make_batch_id(user: str, dg: str, gk: str, start_ts: pd.Timestamp) -> str:
    """Genereer deterministische unieke BatchId als SHA1-hash van user|docgroup|groupkey|starttijd."""
    s = f"{str(user).strip()}|{str(dg).strip()}|{str(gk).strip()}|{pd.Timestamp(start_ts).isoformat()}"
    return hashlib.sha1(s.encode("utf-8")).hexdigest()

def map_activity_from_docgroup(dg: str, gk: str = "") -> str:
    """
    Map DocGroup (+ GroupKey) naar Activity-label.
    
    Speciale case: Magazijndagboek met GroupKey="Magazijndagboek-TRANSFER" wordt
    "Interne Relocatie" — dit zijn pallets die binnen het warehouse worden verplaatst
    (reorganisatie), niet klaargezet voor verzending of machine.
    
    Bekende beperking: sommige Klaarzetten-scans zijn eigenlijk tijdelijke
    herplaatsingen (bijv. naar HAL2 als staging area). Dit is niet automatisch
    te onderscheiden van echte klaarzet-acties, omdat de data geen intentie bevat.
    """
    dg = str(dg).strip()
    if dg == "Magazijndagboek" and str(gk).strip() == "Magazijndagboek-TRANSFER":
        return "Interne Relocatie"
    return ACTIVITY_MAP.get(dg, "Onbekend")

# ===================================================================================
# STORAGE DAYS — Episode State Machine
#
# Per pallet (Dragernr.) wordt één storage-episode bepaald:
#   StorageTime = eerste keer dat het pallet in de stelling wordt geplaatst
#   PickTime    = eerste keer daarna dat het pallet de stelling verlaat
#   StorageDays = verschil in dagen (of t.o.v. as_of als episode nog open is)
#
# Classificatieregels voor events:
#   PUT-AWAY (storage-in): Ontvangst events, of Magazijndagboek met Ref="Opslag"
#   PICK (storage-uit): Verzending, Doc9+Pick, TRANSFER naar machine, of
#                       elke scan naar een MACHINE_IN_LOC (incl. KWIJT, OPSLAG-NB)
#
# State machine logica:
#   - Eerste put-away → zet StorageTime
#   - Pick na StorageTime → zet PickTime
#   - Nieuwe put-away na pick → reset PickTime (pallet gaat weer terug in stelling)
#   - StorageTime ≥ PickTime → verwijder PickTime (data-anomalie)
# ===================================================================================

def norm_text(s: pd.Series) -> pd.Series:
    """Normaliseer tekstveld: strip whitespace, collapse meerdere spaties."""
    return s.fillna("").astype(str).str.strip().str.replace(r"\s+", " ", regex=True)

def norm_loc(s: pd.Series) -> pd.Series:
    """Normaliseer locatiecodes: uppercase + tekst-normalisatie."""
    return norm_text(s).str.upper()

def is_storage_like_loc(loc: str) -> bool:
    """
    Bepaal of een locatiecode een "echte" opslaglocatie is (stelling/hal).
    
    Patronen die als opslag gelden:
    - HAL1, HAL2, etc. (vloerplekken in hallen)
    - 6F-228, 2R-009, 1A-012 (stellinglocaties: hal-rij-plek)
    - 7-1-123 (alternatief stellingformaat)
    
    NIET opslag: machine-locaties (F1_IN etc.), VERZEND, KWIJT, ONTVANGST, OPSLAG-NB
    """
    s = (loc or "").strip().upper()
    if not s: return False
    if s in MACHINE_IN_LOCS:    return False
    if s in NON_STORAGE_TOKENS: return False
    if re.match(r"^HAL\d{1,2}$", s):        return True
    if re.match(r"^\d+[A-Z]-\d{2,4}$", s):  return True
    if re.match(r"^\d+-\d-\d{3}$", s):      return True
    if re.match(r"^\d+[A-Z]-\d{3}$", s):    return True
    return False

def map_doc_group_storage(x: str) -> str:
    """Map Mag.-documentsoort naar genormaliseerde groep voor storage classificatie (lowercase)."""
    s = str(x).strip().lower()
    if not s: return "overig"
    if s in {"ontvangst", "inkoopontvangst"}:          return "ontvangst"
    if s in {"verzending", "verkoopverzending"}:       return "verzending"
    if s in {"magazijndagboek", "artikeldagboek", "artikeldagb.", "transfer"}:
        return "magazijndagboek"
    if s == "9" or s.startswith("9"):                  return "doc9"
    if "ontvang" in s:                                 return "ontvangst"
    if "verzend" in s:                                 return "verzending"
    if "magazijndagboek" in s:                         return "magazijndagboek"
    return "overig"

def storage_load_scans_raw() -> pd.DataFrame:
    """
    Laad ruwe scanbewegingen voor de storage pipeline.
    
    GEEN uitsluitingen hier — alle users (incl. EURO-CAPS) worden meegenomen.
    Dit is belangrijk zodat administratieve scans (bijv. naar OPSLAG-NB of VERZEND)
    wel als pick-event worden gedetecteerd en storage-episodes correct worden afgesloten.
    """
    if not INPUT_XLSX.exists():
        raise FileNotFoundError(f"Bronze input niet gevonden: {INPUT_XLSX}")
    df = pd.read_excel(INPUT_XLSX, engine="openpyxl")
    for c in [COL_DRAGER, COL_LOC, COL_ZONE, COL_DOCMAG, COL_DOCSOORT, COL_ACTIE,
              COL_OMS, COL_REFDOC, COL_DAGBOEK, COL_USER]:
        if c not in df.columns: df[c] = ""
    df["Timestamp"]     = parse_timestamp(df)
    df["LocNorm"]       = norm_loc(df[COL_LOC])
    df["ZoneNorm"]      = norm_text(df[COL_ZONE]).str.upper()
    df["User"]          = norm_text(df.get(COL_USER, ""))
    df["DocGroup"]      = df[COL_DOCMAG].map(map_doc_group_storage)
    df["DocSoortN"]     = norm_text(df[COL_DOCSOORT]).str.lower()
    df["DocMagN"]       = norm_text(df[COL_DOCMAG]).str.lower()
    df["RefDocN"]       = norm_text(df[COL_REFDOC]).str.lower()
    df["DagboekN"]      = norm_text(df[COL_DAGBOEK]).str.upper()
    df["OmschrijvingN"] = norm_text(df.get(COL_OMS, ""))
    df = df.reset_index(drop=True)
    df["RowId"] = np.arange(len(df), dtype=int)
    # Forward-fill locatie per pallet zodat we altijd de laatst bekende locatie hebben
    tmp = df["LocNorm"].replace("", pd.NA)
    df["LocLastKnown"] = tmp.groupby(df[COL_DRAGER], sort=False).ffill()
    return df

def storage_build_micro_events(df: pd.DataFrame) -> pd.DataFrame:
    """
    Detecteer micropairs en singletons per pallet.
    
    Micropair: twee opeenvolgende rijen voor hetzelfde pallet, ≤1s apart,
    met verschillende locaties → dit is één fysieke verplaatsing.
    
    Singleton: een rij die niet deel uitmaakt van een micropair.
    Belangrijk voor administratieve scans (bijv. EURO-CAPS naar OPSLAG-NB)
    die als enkele rij voorkomen en toch een pick-event moeten triggeren.
    """
    def per_group(g: pd.DataFrame) -> pd.DataFrame:
        g = g.sort_values(["RowId"], kind="stable").copy()
        g["TsNext"]       = g["Timestamp"].shift(-1)
        g["LocNext"]      = g["LocNorm"].shift(-1)
        g["UserNext"]     = g["User"].shift(-1)
        g["DocSoortNext"] = g["DocSoortN"].shift(-1)
        g["DocGroupNext"] = g["DocGroup"].shift(-1)
        g["DocMagNext"]   = g["DocMagN"].shift(-1)
        g["RefDocNext"]   = g["RefDocN"].shift(-1)
        g["DagboekNext"]  = g["DagboekN"].shift(-1)
        g["RowIdNext"]    = g["RowId"].shift(-1)
        dt = (g["TsNext"] - g["Timestamp"]).dt.total_seconds()
        to_filled   = g["LocNext"].notna() & (g["LocNext"] != "")
        loc_changed = (g["LocNorm"] != g["LocNext"]) & to_filled
        same_user   = (g["User"] == g["UserNext"])
        base_mask   = same_user if MICROPAIR_REQUIRE_SAME_USER else True
        pair_mask   = base_mask & (dt >= 0) & (dt <= MICROPAIR_MAX_SEC) & loc_changed

        # --- Micropairs ---
        ev_pairs = g.loc[pair_mask, [
            "Timestamp","LocNorm","LocNext","DocSoortN","DocSoortNext",
            "DocGroup","DocGroupNext","DocMagN","DocMagNext","RefDocN","RefDocNext",
            "DagboekN","DagboekNext","OmschrijvingN","RowIdNext","RowId"
        ]].copy()
        if not ev_pairs.empty:
            ev_pairs.rename(columns={
                "Timestamp":"EventTime","LocNorm":"FromLoc","LocNext":"ToLoc",
                "DocSoortN":"DocSoortFrom","DocSoortNext":"DocSoortTo",
                "DocGroup":"DocGroupFrom","DocGroupNext":"DocGroupTo",
                "DocMagN":"DocMagFrom","DocMagNext":"DocMagTo",
                "RefDocN":"RefDocFrom","RefDocNext":"RefDocTo",
                "DagboekN":"DagboekFrom","DagboekNext":"DagboekTo",
                "OmschrijvingN":"Omschrijving","RowIdNext":"EventOrd",
            }, inplace=True)
            ev_pairs[COL_DRAGER] = g[COL_DRAGER].iloc[0]
            ev_pairs["EventOrd"] = ev_pairs["EventOrd"].astype(int)
            ev_pairs["Source"]   = "micro"
        else:
            ev_pairs = pd.DataFrame(columns=[
                COL_DRAGER,"EventTime","EventOrd","FromLoc","ToLoc",
                "DocSoortFrom","DocSoortTo","DocGroupFrom","DocGroupTo",
                "DocMagFrom","DocMagTo","RefDocFrom","RefDocTo",
                "DagboekFrom","DagboekTo","Omschrijving","Source"
            ])

        # --- Singletons: rijen die niet in een micropair zitten ---
        if ENABLE_SINGLETON_EVENTS:
            used_from_mask = pair_mask
            used_to_rowids = g.loc[pair_mask, "RowIdNext"].dropna().astype(int)
            used_to_mask   = g["RowId"].isin(used_to_rowids.values)
            unpaired_mask  = (~used_from_mask) & (~used_to_mask)
            single = g.loc[unpaired_mask, [
                "Timestamp","LocNorm","DocSoortN","DocGroup","DocMagN","RefDocN","DagboekN","OmschrijvingN","RowId"
            ]].copy()
            if not single.empty:
                single.rename(columns={
                    "Timestamp":"EventTime","LocNorm":"FromLoc","DocSoortN":"DocSoortFrom",
                    "DocGroup":"DocGroupFrom","DocMagN":"DocMagFrom","RefDocN":"RefDocFrom",
                    "DagboekN":"DagboekFrom","OmschrijvingN":"Omschrijving","RowId":"EventOrd",
                }, inplace=True)
                # Singletons: FromLoc = ToLoc (we weten alleen de ene locatie)
                single["ToLoc"]      = single["FromLoc"]
                single["DocSoortTo"] = single["DocSoortFrom"]
                single["DocGroupTo"] = single["DocGroupFrom"]
                single["DocMagTo"]   = single["DocMagFrom"]
                single["RefDocTo"]   = single["RefDocFrom"]
                single["DagboekTo"]  = single["DagboekFrom"]
                single[COL_DRAGER]   = g[COL_DRAGER].iloc[0]
                single["EventOrd"]   = single["EventOrd"].astype(int)
                single["Source"]     = "single"
            else:
                single = pd.DataFrame(columns=ev_pairs.columns)
        else:
            single = pd.DataFrame(columns=ev_pairs.columns)
        return pd.concat([ev_pairs, single], ignore_index=True)

    parts  = [per_group(g) for _, g in df.groupby(COL_DRAGER, sort=False)]
    events = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()
    return events

def storage_classify_events(events: pd.DataFrame) -> pd.DataFrame:
    """
    Classificeer events als IsPutAway (storage-in) of IsPick (storage-uit).
    
    Pick heeft voorrang op PutAway: als een event aan beide criteria voldoet,
    wordt het als Pick geclassificeerd (bijv. Ontvangst naar een machine-locatie).
    """
    if events.empty:
        events["IsPutAway"] = []
        events["IsPick"]    = []
        return events

    # Een scan naar een MACHINE_IN_LOC is altijd een pick-trigger
    to_is_machine_in = events["ToLoc"].astype(str).str.upper().isin(MACHINE_IN_LOCS)

    # Put-away triggers: Ontvangst of Magazijndagboek+Opslag
    is_ontvangst = ((events.get("DocGroupTo","") == "ontvangst") |
                    (events.get("DocGroupFrom","") == "ontvangst"))
    is_mag_opslag = (
        ((events.get("DocGroupTo","")   == "magazijndagboek") & (events.get("RefDocTo","")   == "opslag")) |
        ((events.get("DocGroupFrom","") == "magazijndagboek") & (events.get("RefDocFrom","") == "opslag"))
    )

    # Pick triggers: Verzending, Doc9+Pick, TRANSFER naar machine, of elke scan naar machine-loc
    is_verzending = ((events.get("DocGroupTo","") == "verzending") |
                     (events.get("DocGroupFrom","") == "verzending"))
    is_doc9_pick = (
        ((events.get("DocMagTo","")   == "9") & (events.get("RefDocTo","")   == "pick")) |
        ((events.get("DocMagFrom","") == "9") & (events.get("RefDocFrom","") == "pick"))
    )
    is_transfer_machine_in = ((events.get("DocGroupTo","") == "magazijndagboek") &
                              (events.get("DagboekTo","")  == "TRANSFER") &
                              to_is_machine_in)
    is_any_to_machine_in = to_is_machine_in

    putaway_mask = is_ontvangst | is_mag_opslag
    pick_mask    = is_verzending | is_doc9_pick | is_transfer_machine_in | is_any_to_machine_in

    ev = events.copy()
    ev["IsPick"]    = pick_mask
    ev["IsPutAway"] = putaway_mask & (~ev["IsPick"])  # Pick heeft voorrang
    return ev

def storage_summarize(df_scans: pd.DataFrame, ev_all: pd.DataFrame) -> pd.DataFrame:
    """
    Bepaal per pallet de storage-episode via een state machine.
    
    Retourneert één rij per uniek Dragernr. met:
    - StorageTime, PickTime, StorageDays
    - CurrentLocation (laatst bekende locatie)
    - HasStorage, HasPick (booleans voor filtering in Power BI)
    
    Open episodes (HasStorage=True, HasPick=False) krijgen StorageDays berekend
    t.o.v. as_of (= laatste timestamp in de dataset).
    """
    cols = [COL_DRAGER, "Omschrijving", "ProductGroup", "StorageTime", "PickTime",
            "StorageDays", "CurrentLocation", "HasStorage", "HasPick"]
    if ev_all.empty:
        return pd.DataFrame(columns=cols)

    # as_of: referentiedatum voor open episodes (= einde van de dataset)
    as_of = pd.to_datetime(df_scans["Timestamp"].max())

    # Laatst bekende locatie en omschrijving per pallet
    last_state = (
        df_scans.sort_values(["RowId"])
                .groupby(COL_DRAGER, as_index=False, sort=False)
                .agg(Last_location=("LocLastKnown","last"), Omschrijving=("OmschrijvingN","last"))
    )
    ev_sorted = ev_all.sort_values(["EventTime","EventOrd"]).copy()

    rows = []
    for drager, g in ev_sorted.groupby(COL_DRAGER, sort=False):
        g = g.sort_values(["EventTime","EventOrd"]).copy()
        storage_time = pd.NaT
        pick_time    = pd.NaT

        # State machine: loop door events in chronologische volgorde
        for _, ev in g.iterrows():
            # Fallback: Magazijndagboek naar een storage-like locatie (zonder expliciete classificatie)
            storage_like_fallback = (
                ((ev.get("DocGroupTo","") == "magazijndagboek") or (ev.get("DocGroupFrom","") == "magazijndagboek"))
                and is_storage_like_loc(ev.get("ToLoc",""))
                and not bool(ev.get("IsPick", False))
            )
            storage_trigger = bool(ev.get("IsPutAway", False)) or storage_like_fallback
            pick_trigger    = bool(ev.get("IsPick", False))

            if storage_trigger:
                if pd.isna(storage_time):
                    # Eerste storage event → zet StorageTime
                    storage_time = ev["EventTime"]
                    if pd.notna(pick_time): pick_time = pd.NaT  # Reset pick als er al een was
                elif pd.notna(pick_time):
                    # Pallet gaat weer terug in stelling na een pick → reset PickTime
                    if ev["EventTime"] >= pick_time: pick_time = pd.NaT
            elif pick_trigger:
                if pd.notna(storage_time) and pd.isna(pick_time):
                    pick_time = ev["EventTime"]
                elif pd.isna(storage_time) and pd.isna(pick_time):
                    # Pick zonder storage (bijv. pallet kwam vóór de meetperiode binnen)
                    pick_time = ev["EventTime"]

        # Data-anomalie check: StorageTime mag niet na PickTime liggen
        if pd.notna(storage_time) and pd.notna(pick_time) and (storage_time >= pick_time):
            pick_time = pd.NaT

        has_storage = pd.notna(storage_time)
        has_pick    = pd.notna(pick_time)

        # StorageDays berekenen
        if has_storage and has_pick:
            storage_days = max((pick_time - storage_time).total_seconds(), 0.0) / 86400.0
        elif has_storage and not has_pick:
            # Open episode: bereken t.o.v. einde dataset
            storage_days = max((as_of - storage_time).total_seconds(), 0.0) / 86400.0
        else:
            storage_days = np.nan

        # Laatst bekende locatie bepalen
        ls = last_state[last_state[COL_DRAGER] == drager]
        curr_loc = ls["Last_location"].iloc[0] if not ls.empty else ""
        if (pd.isna(curr_loc)) or (str(curr_loc).strip() == ""):
            to_last = g["ToLoc"].replace("", pd.NA).ffill().dropna()
            curr_loc = to_last.iloc[-1] if len(to_last) > 0 else ""
        oms = ls["Omschrijving"].iloc[0] if not ls.empty else ""

        rows.append({
            COL_DRAGER: drager, "Omschrijving": oms, "ProductGroup": "",
            "StorageTime": storage_time, "PickTime": pick_time,
            "StorageDays": round(storage_days, 2) if pd.notna(storage_days) else np.nan,
            "CurrentLocation": curr_loc, "HasPick": bool(has_pick), "HasStorage": bool(has_storage),
        })

    out = pd.DataFrame(rows)
    sort_key = out["StorageTime"].combine_first(out["PickTime"])
    out = (out[cols].assign(_k=sort_key)
           .sort_values(["_k", COL_DRAGER], na_position="last")
           .drop(columns=["_k"]).reset_index(drop=True))
    return out

# ===================================================================================
# PICK/BATCHES — Batch Detection Pipeline
#
# Stappen:
# 1. Laad scans, sluit EURO-CAPS users uit
# 2. Detecteer micropair-events (FROM→TO verplaatsingen)
# 3. Groepeer events tot batches (per User, DocGroup, GroupKey)
# 4. Bepaal BatchEnd via idle-chaining of basisregel
# 5. Bereken Duration_s en Sec per pallet
# 6. Bouw bridge-tabel (BatchId ↔ Dragernr.)
# 7. Ken Activity-labels toe
# ===================================================================================

def map_doc_group_pick(x: str) -> str:
    """Map Mag.-documentsoort naar genormaliseerde groep voor batch pipeline (Capitalized)."""
    s = str(x).strip().lower()
    if not s: return "Overig"
    if s in {"ontvangst", "inkoopontvangst"}:                 return "Ontvangst"
    if s in {"verzending", "verkoopverzending"}:              return "Verzending"
    if s in {"magazijndagboek", "artikeldagboek", "transfer",
             "artikeldagb.", "magazijnboek"}:                 return "Magazijndagboek"
    if s.startswith("9"):                                     return "Doc9"
    if "ontvang" in s:                                        return "Ontvangst"
    if "verzend" in s:                                        return "Verzending"
    if "magazijndagboek" in s or "artikeldagb" in s:          return "Magazijndagboek"
    return "Overig"

def group_key_for_pick(row) -> str:
    """
    Bepaal de GroupKey per scanrij — dit bepaalt batch-grenzen.
    
    - Ontvangst: gegroepeerd op ordernummer (Bronnr.)
    - Verzending: gegroepeerd op verzendingsnummer (Mag.-documentnr.)
    - Magazijndagboek: vast "Magazijndagboek" of "Magazijndagboek-TRANSFER"
      (TRANSFER wordt apart gegroepeerd → Interne Relocatie activiteit)
    - Doc9: vast "9"
    """
    dg = row["DocGroup"]
    if dg == "Ontvangst":      return str(row.get(COL_ORDER, "") or "")
    if dg == "Verzending":     return str(row.get(COL_MAGDOCNO, "") or "")
    if dg == "Magazijndagboek":
        dagboek = str(row.get("DagboekN", "") or "").strip().upper()
        if dagboek == "TRANSFER": return "Magazijndagboek-TRANSFER"
        return "Magazijndagboek"
    if dg == "Doc9":           return "9"
    return ""

def pick_build_events(scans_pick: pd.DataFrame) -> pd.DataFrame:
    """
    Detecteer micropair-events uit de (gefilterde) scans.
    
    Alleen micropairs worden hier meegenomen (geen singletons) — dit zijn de
    "echte" palletverplaatsingen door warehouse-medewerkers.
    """
    df = scans_pick.sort_values(["RowId"]).reset_index(drop=True)
    g  = df.groupby(COL_DRAGER, sort=False)
    next_t   = g["Timestamp"].shift(-1)
    next_loc = g[COL_LOC].shift(-1)
    dt_sec   = (next_t - df["Timestamp"]).dt.total_seconds()
    norm = lambda s: s.fillna("").astype(str).str.upper().str.replace(r"\s+", " ", regex=True).str.strip()
    loc_changed = norm(df[COL_LOC]) != norm(next_loc)
    cond_time = next_t.notna() & (dt_sec >= 0) & (dt_sec <= MICRO_PAIR_SEC)
    is_pair   = cond_time & loc_changed
    ev = df.loc[is_pair, [COL_DRAGER, "User", "Timestamp", "DocGroup", "GroupKey", "RowId"]].copy()
    ev.rename(columns={"Timestamp": "EventTime", "RowId": "EventRowId"}, inplace=True)
    return ev.sort_values(["User","EventRowId"]).reset_index(drop=True)

def pick_build_batches(events: pd.DataFrame) -> pd.DataFrame:
    """
    Groepeer events tot batches per User.
    
    Een batch eindigt wanneer:
    - DocGroup of GroupKey verandert (andere taak/order), of
    - Gap > MACRO_PAUSE_SPLIT_SEC binnen dezelfde DocGroup/GroupKey (pauze)
    """
    rows = []
    for user, ue in events.sort_values(["User","EventRowId"]).groupby("User", sort=False):
        ue = ue.reset_index(drop=True)
        if ue.empty: continue
        cur_dg, cur_gk, start_i, prev_ts = ue.loc[0,"DocGroup"], ue.loc[0,"GroupKey"], 0, ue.loc[0,"EventTime"]
        for i in range(1, len(ue)):
            dg, gk, ts = ue.loc[i,"DocGroup"], ue.loc[i,"GroupKey"], ue.loc[i,"EventTime"]
            changed     = (dg != cur_dg) or (gk != cur_gk)
            gap         = (ts - prev_ts) if (pd.notna(ts) and pd.notna(prev_ts)) else pd.Timedelta(0)
            macro_break = (not changed) and (gap > pd.Timedelta(seconds=MACRO_PAUSE_SPLIT_SEC))
            if changed or macro_break:
                seg = ue.loc[start_i:i-1]
                rows.append({"User":user,"DocGroup":cur_dg,"GroupKey":cur_gk,
                             "BatchStart":seg["EventTime"].iloc[0],"LastEvent":seg["EventTime"].iloc[-1],
                             "Aantal Pallets":int(len(seg))})
                start_i, cur_dg, cur_gk = i, dg, gk
            prev_ts = ts
        seg = ue.loc[start_i:len(ue)-1]
        rows.append({"User":user,"DocGroup":cur_dg,"GroupKey":cur_gk,
                     "BatchStart":seg["EventTime"].iloc[0],"LastEvent":seg["EventTime"].iloc[-1],
                     "Aantal Pallets":int(len(seg))})
    return pd.DataFrame(rows).sort_values(["User","BatchStart"]).reset_index(drop=True)

def pick_finalize_batches(batches: pd.DataFrame, events: pd.DataFrame, scans_pick: pd.DataFrame):
    """
    Finaliseer batches: bepaal BatchEnd, bereken tijden, bouw bridge.
    
    BatchEnd wordt bepaald door (in volgorde van prioriteit):
    1. Idle-chaining: als volgende batch start binnen CHAIN_MAX_IDLE_SEC → BatchEnd = NextBatchStart
    2. Batch-scan correctie: als avg gap < BATCH_SCAN_CORRECTION_SEC en niet Verzending
       → BatchEnd = LastEvent + N × DocGroup-gemiddelde
    3. Standaard: BatchEnd = LastEvent + 1 × AvgEventGap
    4. N=1 fallback: BatchEnd = LastEvent + DocGroup-gemiddelde (of globale mediaan, of 120s)
    
    Retourneert: (batches_df, batch_pallets_bridge_df)
    """
    if batches.empty:
        cols = ["User","DocGroup","GroupKey","BatchStart","BatchEnd","Aantal Pallets",
                "Sec per pallet (mm:ss)","Sec per pallet (s)","BatchId","Activity"]
        return pd.DataFrame(columns=cols), pd.DataFrame(columns=["BatchId","Dragernr.","Omschrijving","ProductGroup","Activity"])

    events  = events.sort_values(["User","EventRowId"]).reset_index(drop=True)
    batches = batches.sort_values(["User","BatchStart"]).reset_index(drop=True)

    # Metadata per pallet (omschrijving) voor de bridge
    pal_meta = (
        scans_pick[[COL_DRAGER, COL_OMS]].copy()
        .assign(**{COL_DRAGER: lambda d: d[COL_DRAGER].astype(str),
                   COL_OMS:    lambda d: d[COL_OMS].astype(str).str.strip()})
    )
    pal_meta = pal_meta.loc[pal_meta[COL_OMS] != ""]
    pal_meta = pal_meta.drop_duplicates(subset=[COL_DRAGER], keep="last").rename(columns={COL_OMS: "Omschrijving"})

    def compute_avg_gap(user, start, end, dg, gk):
        """Bereken gemiddelde tijdsgap tussen events binnen een batch."""
        ue = events[(events["User"]==user)&(events["EventTime"]>=start)&(events["EventTime"]<=end)&
                    (events["DocGroup"]==dg)&(events["GroupKey"]==gk)].copy().sort_values(["EventRowId"])
        if len(ue) >= 2:
            gaps = ue["EventTime"].diff().dt.total_seconds().dropna()
            if not gaps.empty: return float(gaps.mean())
        return np.nan

    batches["AvgEventGapSec"] = [compute_avg_gap(r["User"],r["BatchStart"],r["LastEvent"],r["DocGroup"],r["GroupKey"])
                                  for _, r in batches.iterrows()]

    # Bereken DocGroup-gemiddelden (uit batches met N≥2) voor fallbacks
    mask_ok = (batches["Aantal Pallets"]>=2) & batches["AvgEventGapSec"].notna() & (batches["AvgEventGapSec"]>0)
    docgroup_avg_map = batches.loc[mask_ok].groupby("DocGroup")["AvgEventGapSec"].mean().to_dict()
    global_median_per_pallet = float(batches.loc[mask_ok,"AvgEventGapSec"].median()) if mask_ok.any() else DOCGROUP_SINGLE_FALLBACK_SEC

    # Bepaal gap naar volgende batch (per user) voor idle-chaining
    batches["NextBatchStart"] = batches.groupby("User")["BatchStart"].shift(-1)
    batches["GapToNextSec"]   = (batches["NextBatchStart"] - batches["LastEvent"]).dt.total_seconds()

    def basis_end(row):
        """Basisregel voor BatchEnd (als chaining niet geldt)."""
        n, last_ev, dg, avg_gap = int(row["Aantal Pallets"]), row["LastEvent"], row["DocGroup"], row["AvgEventGapSec"]
        if n >= 2 and (not pd.isna(avg_gap)) and (avg_gap > 0):
            # Batch-scan correctie: als avg gap heel kort is (<20s) en niet Verzending,
            # dan werd alles eerst gescand en daarna pas verplaatst.
            # → Compenseer met N × DocGroup-gemiddelde als proxy voor werkelijke verplaatsingstijd
            if avg_gap < BATCH_SCAN_CORRECTION_SEC and dg != "Verzending":
                doc_avg = docgroup_avg_map.get(dg, np.nan)
                if (not pd.isna(doc_avg)) and (doc_avg > 0):
                    return last_ev + pd.to_timedelta(n * doc_avg, unit="s")
                elif global_median_per_pallet and global_median_per_pallet > 0:
                    return last_ev + pd.to_timedelta(n * global_median_per_pallet, unit="s")
            # Standaard: 1 × eigen AvgEventGap (compenseert de laatste, ontbrekende gap)
            return last_ev + pd.to_timedelta(avg_gap, unit="s")
        # N=1 fallback: gebruik DocGroup-gemiddelde, globale mediaan, of 120s
        doc_avg = docgroup_avg_map.get(dg, np.nan)
        if (not pd.isna(doc_avg)) and (doc_avg > 0): add_sec = doc_avg
        elif global_median_per_pallet and global_median_per_pallet > 0: add_sec = global_median_per_pallet
        else: add_sec = DOCGROUP_SINGLE_FALLBACK_SEC
        return last_ev + pd.to_timedelta(add_sec, unit="s")

    def pick_end(row):
        """Bepaal BatchEnd: idle-chaining als mogelijk, anders basisregel."""
        next_ts, last_ev, gap_next = row["NextBatchStart"], row["LastEvent"], row["GapToNextSec"]
        chain_ok = (pd.notna(next_ts) and pd.notna(last_ev) and pd.notna(gap_next)
                    and (gap_next >= 0) and (gap_next <= CHAIN_MAX_IDLE_SEC) and (next_ts >= last_ev))
        if chain_ok: return next_ts
        return basis_end(row)

    batches["BatchEnd"]               = batches.apply(pick_end, axis=1)
    batches["Duration_s"]             = (batches["BatchEnd"]-batches["BatchStart"]).dt.total_seconds().clip(lower=0)
    batches["Sec per pallet (s)"]     = (batches["Duration_s"]/batches["Aantal Pallets"].replace(0,np.nan)).fillna(0.0)
    batches["Sec per pallet (mm:ss)"] = [mmss(x) for x in batches["Sec per pallet (s)"]]
    batches["BatchId"]                = [make_batch_id(r["User"],r["DocGroup"],r["GroupKey"],r["BatchStart"]) for _,r in batches.iterrows()]
    batches["Activity"]               = batches.apply(lambda r: map_activity_from_docgroup(r["DocGroup"],r["GroupKey"]), axis=1)

    # --- Bridge-tabel: BatchId ↔ Dragernr. ---
    # Per batch, vind alle unieke pallets (Dragernr.) die erin voorkomen
    bridge_rows = []
    for _, r in batches.iterrows():
        user, dg, gk = r["User"], r["DocGroup"], r["GroupKey"]
        start_ts, end_ts, batch_id = r["BatchStart"], r["LastEvent"], r["BatchId"]
        sub = events[(events["User"]==user)&(events["DocGroup"]==dg)&(events["GroupKey"]==gk)&
                     (events["EventTime"]>=start_ts)&(events["EventTime"]<=end_ts)][[COL_DRAGER]].dropna()
        for dr in sorted(sub[COL_DRAGER].astype(str).unique()):
            bridge_rows.append({"BatchId": batch_id, COL_DRAGER: dr})

    batch_pallets = pd.DataFrame(bridge_rows).drop_duplicates().reset_index(drop=True)
    batch_pallets[COL_DRAGER] = batch_pallets[COL_DRAGER].astype(str)

    # Verrijk bridge met omschrijving uit brondata
    batch_pallets = batch_pallets.merge(pal_meta, on=COL_DRAGER, how="left")
    if "Omschrijving" not in batch_pallets.columns: batch_pallets["Omschrijving"] = ""
    batch_pallets["Omschrijving"] = batch_pallets["Omschrijving"].fillna("")
    batch_pallets["ProductGroup"] = ""  # Wordt later gevuld via ProductGroup mapping

    # Activity ook op bridge (voor Power BI filtering via batch_pallets)
    if WRITE_ACTIVITY_TO_BATCH_PALLETS:
        activity_map_df = batches[["BatchId","Activity"]].drop_duplicates()
        batch_pallets = batch_pallets.merge(activity_map_df, on="BatchId", how="left")

    out_batches = batches[["User","DocGroup","GroupKey","BatchStart","BatchEnd",
        "Aantal Pallets","Sec per pallet (mm:ss)","Sec per pallet (s)","BatchId","Activity","Duration_s"
    ]].sort_values(["User","BatchStart"]).reset_index(drop=True)
    out_bridge = batch_pallets[["BatchId",COL_DRAGER,"Omschrijving","ProductGroup"]+(["Activity"] if WRITE_ACTIVITY_TO_BATCH_PALLETS else [])].copy()
    return out_batches, out_bridge

# ===================================================================================
# ACTIVITY SEGMENTS — Lineaire tijdlijn per user
#
# Elke user krijgt een aaneengesloten tijdlijn van:
#   - Work-segmenten: exact de batches (één op één)
#   - Idle-segmenten: gaps >30min tussen batches (apart, Activity = "Idle/Pauze")
#
# Dit is voorbereid voor toekomstig gebruik — momenteel niet actief in Power BI
# maar kan gebruikt worden voor productiviteitsanalyse.
# ===================================================================================

def build_activity_segments(batches_out: pd.DataFrame,
                            idle_gap_threshold_sec: int = IDLE_GAP_THRESHOLD_SEC,
                            idle_label: str = ACTIVITY_IDLE_LABEL) -> pd.DataFrame:
    """Bouw lineaire tijdlijn per user uit batches + idle-gaps."""
    if batches_out.empty:
        return pd.DataFrame(columns=["User","SegmentStart","SegmentEnd","Duration_s","SegmentType","Activity",
                "BatchId","DocGroup","GroupKey","PrevBatchId","NextBatchId"])
    b = batches_out.copy()
    for c in ["BatchStart","BatchEnd"]:
        if not np.issubdtype(b[c].dtype, np.datetime64): b[c] = pd.to_datetime(b[c], errors="coerce")
    b = b.sort_values(["User","BatchStart"]).reset_index(drop=True)
    b["NextBatchStart"] = b.groupby("User")["BatchStart"].shift(-1)
    b["NextBatchId"]    = b.groupby("User")["BatchId"].shift(-1)

    segments = []

    # Work-segmenten: één op één met batches
    for _, r in b.iterrows():
        dur = float(pd.to_timedelta(r["BatchEnd"]-r["BatchStart"]).total_seconds()) if (pd.notna(r["BatchEnd"]) and pd.notna(r["BatchStart"])) else 0.0
        segments.append({"User":r["User"],"SegmentStart":r["BatchStart"],"SegmentEnd":r["BatchEnd"],
            "Duration_s":max(dur,0.0),"SegmentType":"Work",
            "Activity":r.get("Activity",map_activity_from_docgroup(r.get("DocGroup",""),r.get("GroupKey",""))),
            "BatchId":r["BatchId"],"DocGroup":r.get("DocGroup",""),"GroupKey":r.get("GroupKey",""),
            "PrevBatchId":None,"NextBatchId":None})

    # Idle-segmenten: gaps > threshold tussen batches (per user)
    gap_thr = float(idle_gap_threshold_sec)
    b["GapToNextSec"] = (b["NextBatchStart"]-b["BatchEnd"]).dt.total_seconds()
    idle_mask = b["NextBatchStart"].notna() & b["GapToNextSec"].notna() & (b["GapToNextSec"] > gap_thr)
    for _, r in b.loc[idle_mask].iterrows():
        if pd.isna(r["BatchEnd"]) or pd.isna(r["NextBatchStart"]): continue
        if r["NextBatchStart"] <= r["BatchEnd"]: continue
        dur = float(pd.to_timedelta(r["NextBatchStart"]-r["BatchEnd"]).total_seconds())
        if dur <= 0: continue
        segments.append({"User":r["User"],"SegmentStart":r["BatchEnd"],"SegmentEnd":r["NextBatchStart"],
            "Duration_s":dur,"SegmentType":"Idle","Activity":idle_label,"BatchId":None,
            "DocGroup":"","GroupKey":"","PrevBatchId":r["BatchId"],"NextBatchId":r["NextBatchId"]})

    seg_df = pd.DataFrame(segments)
    seg_df["Duration_s"] = seg_df["Duration_s"].astype(float).clip(lower=0.0)
    return seg_df.sort_values(["User","SegmentStart","SegmentEnd"]).reset_index(drop=True)

# ===================================================================================
# MAIN — Pipeline orchestratie
# ===================================================================================

def main():
    # SILVER_DIR.mkdir(parents=True, exist_ok=True)
    # GOLD_DIR.mkdir(parents=True, exist_ok=True)
    print(f"Start warehouse-model … {PIPELINE_VERSION}")
    print(f"📥 Bronze: {INPUT_XLSX}")

    # =====================================================================
    # STAP 1: Storage Days (alle scans, incl. EURO-CAPS)
    # =====================================================================
    scans_storage = storage_load_scans_raw()
    print(f"🔢 Storage loader: scans={len(scans_storage)} | dragers={scans_storage[COL_DRAGER].nunique()}")

    # Detecteer events (micropairs + singletons)
    ev_all = storage_build_micro_events(scans_storage)

    # Dedup: als een micropair en singleton op hetzelfde moment/locatie bestaan,
    # behoud de micropair (heeft meer informatie: FromLoc ≠ ToLoc)
    ev_all["SourcePri"] = ev_all["Source"].map({"micro":0,"single":1}).fillna(9).astype(int)
    ev_all = (ev_all.sort_values(["EventTime","EventOrd","SourcePri"])
              .drop_duplicates(subset=[COL_DRAGER,"EventTime","EventOrd","FromLoc","ToLoc"], keep="first")
              .drop(columns=["SourcePri"]))

    # Classificeer events als PutAway of Pick
    events_class = storage_classify_events(ev_all)
    n_ev = len(events_class)
    n_put = int(events_class["IsPutAway"].sum()) if "IsPutAway" in events_class else 0
    n_pick = int(events_class["IsPick"].sum()) if "IsPick" in events_class else 0
    n_singletons = int((events_class.get("Source","") == "single").sum())
    n_micro      = int((events_class.get("Source","") == "micro").sum())
    print(f"🧩 Storage events: total={n_ev} | micro={n_micro} | singletons={n_singletons} | put-away={n_put} | pick={n_pick}")

    # Optioneel: exporteer events naar silver voor QA
    if SAVE_SILVER_EVENTS:
        with pd.ExcelWriter(SILVER_EVENTS_XLSX, engine="openpyxl", mode="w") as xw:
            cols = [COL_DRAGER,"EventTime","EventOrd","FromLoc","ToLoc",
                "DocSoortFrom","DocSoortTo","DocGroupFrom","DocGroupTo","DocMagFrom","DocMagTo",
                "RefDocFrom","RefDocTo","DagboekFrom","DagboekTo","IsPutAway","IsPick","Source","Omschrijving"]
            keep = [c for c in cols if c in events_class.columns]
            events_class[keep].sort_values(["EventTime","EventOrd",COL_DRAGER]).to_excel(xw, index=False, sheet_name=SILVER_EVENTS_SHEET)
        print(f"✅ Silver (events) → {SILVER_EVENTS_XLSX}")

    # Samenvat tot één rij per pallet via state machine
    storage_days = storage_summarize(scans_storage, events_class)
    if not storage_days.empty:
        storage_days[COL_DRAGER] = storage_days[COL_DRAGER].astype(str)
    print(f"✅ StorageDays rows: {len(storage_days)}")

    # =====================================================================
    # STAP 2: Pick / Batches (ZONDER EURO-CAPS users)
    # =====================================================================
    scans_pick = pd.read_excel(INPUT_XLSX, engine="openpyxl")
    scans_pick.rename(columns={COL_USER: "User"}, inplace=True)
    scans_pick["Timestamp"] = parse_timestamp(scans_pick)
    scans_pick["RowId"] = np.arange(len(scans_pick), dtype=int)

    # EURO-CAPS uitsluiting: alleen voor de batch pipeline!
    # Storage pipeline hierboven bevat wél alle users.
    if EXCLUDE_USER_PREFIXES_PICK:
        prefixes = tuple(p.upper() for p in EXCLUDE_USER_PREFIXES_PICK)
        m = scans_pick["User"].astype(str).str.strip().str.upper().str.startswith(prefixes)
        removed = int(m.sum())
        scans_pick = scans_pick.loc[~m].copy()
        if removed > 0: print(f"🚫 Pick: excluded rows for users {EXCLUDE_USER_PREFIXES_PICK}: removed {removed}")

    scans_pick["DocGroup"] = scans_pick[COL_DOCMAG].map(map_doc_group_pick)
    # DagboekN nodig voor GroupKey (onderscheid Klaarzetten vs Interne Relocatie)
    scans_pick["DagboekN"] = scans_pick[COL_DAGBOEK].fillna("").astype(str).str.strip().str.upper()
    scans_pick["GroupKey"] = scans_pick.apply(group_key_for_pick, axis=1)

    events_pick = pick_build_events(scans_pick)
    batches_core = pick_build_batches(events_pick)
    batches_out, batch_pallets = pick_finalize_batches(batches_core, events_pick, scans_pick)

    # Verrijk bridge met ProductGroup en Omschrijving vanuit storage_days
    if not storage_days.empty and not batch_pallets.empty:
        sd_pg = storage_days[[COL_DRAGER,"ProductGroup","Omschrijving"]].copy().astype({COL_DRAGER:str})
        batch_pallets[COL_DRAGER] = batch_pallets[COL_DRAGER].astype(str)
        batch_pallets = (batch_pallets.drop(columns=["ProductGroup"], errors="ignore")
                         .merge(sd_pg[[COL_DRAGER,"ProductGroup"]], on=COL_DRAGER, how="left"))
        batch_pallets = (batch_pallets
                         .merge(sd_pg[[COL_DRAGER,"Omschrijving"]].rename(columns={"Omschrijving":"Omschrijving_sd"}),
                                on=COL_DRAGER, how="left"))
        batch_pallets["Omschrijving"] = np.where(
            batch_pallets["Omschrijving"].astype(str).str.strip().eq(""),
            batch_pallets["Omschrijving_sd"].fillna(""),
            batch_pallets["Omschrijving"].astype(str)).astype(str)
        batch_pallets = batch_pallets.drop(columns=["Omschrijving_sd"], errors="ignore")
        for c in ["ProductGroup","Omschrijving"]: batch_pallets[c] = batch_pallets[c].fillna("")
    print(f"✅ Batches rows: {len(batches_out)}")
    print(f"🔗 Bridge rows (pre-aanvulling): {len(batch_pallets)}")

    # =====================================================================
    # STAP 3: ProductGroup verrijking (vanuit externe mapping)
    # =====================================================================
    if PRODUCT_GROUPS_XLSX.exists():
        pg = pd.read_excel(PRODUCT_GROUPS_XLSX, engine="openpyxl")
        pg["Omschrijving"] = pg["Omschrijving"].fillna("").astype(str).str.strip()
        pg = pg[pg["Omschrijving"] != ""].drop_duplicates(subset=["Omschrijving"], keep="last")
        pg_map = pg[["Omschrijving", "ProductGroup", "Bedrijf"]].copy()

        # Verrijk storage_days
        if not storage_days.empty:
            storage_days["Omschrijving"] = storage_days["Omschrijving"].fillna("").astype(str).str.strip()
            storage_days = storage_days.drop(columns=["ProductGroup","Bedrijf"], errors="ignore")
            storage_days = storage_days.merge(pg_map, on="Omschrijving", how="left")
            storage_days["ProductGroup"] = storage_days["ProductGroup"].fillna("")
            storage_days["Bedrijf"] = storage_days["Bedrijf"].fillna("")

        # Verrijk batch_pallets
        if not batch_pallets.empty:
            batch_pallets["Omschrijving"] = batch_pallets["Omschrijving"].fillna("").astype(str).str.strip()
            batch_pallets = batch_pallets.drop(columns=["ProductGroup","Bedrijf"], errors="ignore")
            batch_pallets = batch_pallets.merge(pg_map, on="Omschrijving", how="left")
            batch_pallets["ProductGroup"] = batch_pallets["ProductGroup"].fillna("")
            batch_pallets["Bedrijf"] = batch_pallets["Bedrijf"].fillna("")

        matched_sd = int((storage_days["ProductGroup"] != "").sum()) if not storage_days.empty else 0
        matched_bp = int((batch_pallets["ProductGroup"] != "").sum()) if not batch_pallets.empty else 0
        print(f"📦 ProductGroup: mapping={len(pg_map)} | matched storage_days={matched_sd}/{len(storage_days)} | matched batch_pallets={matched_bp}/{len(batch_pallets)}")
    else:
        print(f"⚠️ ProductGroup bestand niet gevonden: {PRODUCT_GROUPS_XLSX} — ProductGroup blijft leeg")

    # =====================================================================
    # STAP 4: Aanvullen ontbrekende pallets in batch_pallets
    #
    # Probleem: pallets die alleen EURO-CAPS scans hebben, bestaan in
    # storage_days maar NIET in batch_pallets (want EURO-CAPS is uitgesloten
    # uit de batch pipeline). Door bidirectioneel filteren in Power BI
    # vallen deze pallets dan weg uit storage_days visuals.
    #
    # Oplossing: voeg deze pallets toe aan batch_pallets met lege BatchId
    # en Activity. Ze koppelen dan niet aan batches (geen picktijd-vervuiling)
    # maar bestaan wél in de bridge (geen wegfiltering in Power BI).
    # =====================================================================
    if not storage_days.empty and not batch_pallets.empty:
        existing_dragers = set(batch_pallets[COL_DRAGER].astype(str).unique())
        sd_cols = [COL_DRAGER, "Omschrijving"]
        if "ProductGroup" in storage_days.columns: sd_cols.append("ProductGroup")
        if "Bedrijf" in storage_days.columns: sd_cols.append("Bedrijf")
        all_sd = storage_days[sd_cols].copy()
        all_sd[COL_DRAGER] = all_sd[COL_DRAGER].astype(str)
        missing = all_sd[~all_sd[COL_DRAGER].isin(existing_dragers)].copy()
        if not missing.empty:
            missing["BatchId"] = ""
            if WRITE_ACTIVITY_TO_BATCH_PALLETS:
                missing["Activity"] = ""
            for c in batch_pallets.columns:
                if c not in missing.columns:
                    missing[c] = ""
            batch_pallets = pd.concat(
                [batch_pallets, missing[batch_pallets.columns]],
                ignore_index=True
            )
            print(f"➕ batch_pallets aangevuld: {len(missing)} pallets zonder batchactiviteit toegevoegd")
    print(f"🔗 Bridge rows (finaal): {len(batch_pallets)}")

    # =====================================================================
    # STAP 5: Activity Segments (Work + Idle tijdlijn per user)
    # =====================================================================
    activity_segments = build_activity_segments(batches_out,
                                                idle_gap_threshold_sec=IDLE_GAP_THRESHOLD_SEC,
                                                idle_label=ACTIVITY_IDLE_LABEL)
    print(f"🧱 Activity segments: total={len(activity_segments)} "
          f"| work={int((activity_segments['SegmentType']=='Work').sum())} "
          f"| idle={int((activity_segments['SegmentType']=='Idle').sum())}")

    # =====================================================================
    # STAP 6: QA — Massabalans controle
    #
    # Check 1: Som van Work-segmenten (s) moet gelijk zijn aan som van batch Duration_s
    # Check 2: Som van Idle-segmenten (s) moet gelijk zijn aan verwachte idle
    #          (som van gaps > threshold per user)
    # =====================================================================
    try:
        sum_work_segments = float(activity_segments.loc[activity_segments["SegmentType"]=="Work","Duration_s"].sum())
        sum_batches_dur   = float(batches_out["Duration_s"].sum())
        sum_idle_segments = float(activity_segments.loc[activity_segments["SegmentType"]=="Idle","Duration_s"].sum())
        tmp_b = batches_out.sort_values(["User","BatchStart"]).copy()
        tmp_b["NextBatchStart"] = tmp_b.groupby("User")["BatchStart"].shift(-1)
        tmp_b["GapToNextSec"]   = (tmp_b["NextBatchStart"] - tmp_b["BatchEnd"]).dt.total_seconds()
        sum_expected_idle       = float(tmp_b.loc[tmp_b["GapToNextSec"] > IDLE_GAP_THRESHOLD_SEC, "GapToNextSec"].sum())
        print(f"🧪 QA — ΣWork(seg)={sum_work_segments:.1f}s vs ΣDuration(batches)={sum_batches_dur:.1f}s "
              f"| ΣIdle(seg)={sum_idle_segments:.1f}s vs expected={sum_expected_idle:.1f}s")
        qa_rows = [
            {"Metric":"SumWorkSegments_s","Value":sum_work_segments},
            {"Metric":"SumBatchesDuration_s","Value":sum_batches_dur},
            {"Metric":"SumIdleSegments_s","Value":sum_idle_segments},
            {"Metric":"SumExpectedIdle_s","Value":sum_expected_idle},
            {"Metric":"IdleGapThreshold_sec","Value":float(IDLE_GAP_THRESHOLD_SEC)},
        ]
        qa_df = pd.DataFrame(qa_rows)
    except Exception as e:
        print(f"⚠️ QA kon niet worden berekend: {e}")
        qa_df = pd.DataFrame(columns=["Metric","Value"])

    # =====================================================================
    # STAP 7: Kosten Parameters (voor Power BI measures)
    #
    # Deze tabel wordt in Power BI gelezen via LOOKUPVALUE() en drijft alle
    # kostenmeasures. Pas de Waarde-kolom aan als er betere data binnenkomt.
    # De Python pipeline hoeft dan opnieuw gedraaid te worden en Power BI
    # refreshed — measures passen automatisch mee.
    #
    # AANNAMES / PLACEHOLDERS:
    # - Uurtarief: €22 all-in (schatting op basis van €3000 bruto + vakantiegeld + werkgeverslasten)
    # - Betaalde uren per dag: 40 (placeholder — vervang met echte klokdata van klant)
    # - EPT/Reachtruck: afschrijving per maand / 30.44 = dagkosten
    # - Onderhoud+energie: benchmark van 35% bovenop afschrijving
    # - Stellingen: €2400/maand × 12 = jaarkosten (bevestiging nodig: per maand of per jaar?)
    # =====================================================================
    kosten_params = pd.DataFrame([
        {"Parameter": "Uurtarief_AllIn",            "Waarde": 22,     "Toelichting": "Bruto + vakantiegeld + werkgeverslasten, schatting"},
        {"Parameter": "Betaalde_Uren_Per_Dag",      "Waarde": 40,    "Toelichting": "Placeholder - vervang met echte klokdata"},
        {"Parameter": "EPT_Dagkosten",              "Waarde": 11.51,  "Toelichting": "350/30.44 - afschrijving per EPT per dag"},
        {"Parameter": "Reachtruck_Dagkosten",       "Waarde": 13.14,  "Toelichting": "400/30.44 - afschrijving per reachtruck per dag"},
        {"Parameter": "Aantal_EPT",                 "Waarde": 4,      "Toelichting": "Placeholder - uitvragen bij klant"},
        {"Parameter": "Aantal_Reachtruck",          "Waarde": 1,      "Toelichting": "Placeholder - uitvragen bij klant"},
        {"Parameter": "Onderhoud_Energie_Factor",   "Waarde": 1.35,   "Toelichting": "Benchmark: 35% bovenop afschrijving"},
        {"Parameter": "Stellingen_Jaarkosten",      "Waarde": 12*2400,"Toelichting": "Totale afschrijving stellingen per jaar"},
    ])

    # =====================================================================
    # STAP 8: Export naar Excel
    # =====================================================================
    # GOLD_DIR.mkdir(parents=True, exist_ok=True)
    # OUTPUT_LOC = SHAREPOINT_LOC / OUTPUT_XLSX
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl", mode="w") as xw:
        storage_days.to_excel(xw, index=False, sheet_name=SHEET_STORAGE_DAYS)
        batches_out.to_excel(xw, index=False, sheet_name=SHEET_BATCHES)
        batch_pallets.to_excel(xw, index=False, sheet_name=SHEET_BATCH_PALLETS)
        activity_segments.to_excel(xw, index=False, sheet_name=SHEET_ACTIVITY_SEGMENTS)
        if WRITE_ACTIVITY_SEGMENTS_QA_SHEET:
            qa_df.to_excel(xw, index=False, sheet_name=SHEET_ACTIVITY_QA)
        kosten_params.to_excel(xw, index=False, sheet_name=SHEET_KOSTEN_PARAMS)

    print(f"📦 Model → {OUTPUT_XLSX}")
    print("🎉 Klaar.")

if __name__ == "__main__":
    main()