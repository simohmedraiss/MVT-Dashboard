#!/usr/bin/env python3
"""
Le Mouvement — Script de transformation Excel → data.json
Usage : python transform.py [chemin_vers_excel]
Par défaut : DATA_Situations_Mouvement.xlsx dans le même dossier
"""

import sys
import json
from pathlib import Path
import pandas as pd

# ── CONFIG ───────────────────────────────────────────────────────────────────
EXCEL_FILE  = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "DATA_Situations_Mouvement.xlsx"
OUTPUT_FILE = Path(__file__).parent / "data.json"

SHEET_P1 = "DATA_Consolidée"
SHEET_P2 = "Etape_Categorization"

PATHWAY_MAP = {
    "CAT A&B -> Idéation":         "AB",
    "CAT C -> Pré-Idéation":       "C",
    "CAT D -> Pré-Idéation":       "D",
    "Quick wins -> Move Up":       "QW",
    "Entrepreneurial Certificate": "F",
}

PATHWAY_LABEL = {
    "AB": "CAT A&B -> Demo Day",
    "C":  "CAT C -> Pré-idéation",
    "D":  "CAT D -> Pré-idéation",
    "QW": "Quick Wins -> Move Up",
    "F":  "CAT F -> Entrepreneurial Certificate",
}

PATHWAY_ORG = {
    "AB": "Le HUB",
    "C":  "LaunchX",
    "D":  "LaunchX",
    "QW": "MindX",
    "F":  "Le HUB",
}

SBU_KEEP = ["Mining","Manufacturing","Corporate","UM6P","InnovX",
            "Fondation Phosboucraa","Nutricrops","OTED","Rock","SPS"]
SITES    = ["Khouribga","Phosboucraa","Jorf Lasfar","Safi","Gantour","Casablanca"]

# ── HELPERS ───────────────────────────────────────────────────────────────────
def clean(val):
    if pd.isna(val): return None
    s = str(val).strip()
    return s if s and s not in ("00:00:00","nan") else None

def sbu_norm(val):
    s = clean(val)
    if not s: return "Autre"
    for k in SBU_KEEP:
        if k.lower() in s.lower(): return k
    return "Autre"

def site_norm(val):
    s = clean(val)
    if not s: return "Autre"
    for k in SITES:
        if k.lower() in s.lower(): return k
    return "Autre"

def to_stage(step):
    if not step: return "Framing"
    s = step.lower()
    if "soumission" in s:                         return "Framing"
    if "categor" in s or "qualification" in s:    return "Ideation"
    if "ideation" in s:                           return "Ideation"
    if "incubation" in s or "hacking" in s:       return "Incubation"
    if "acceler" in s:                            return "Acceleration"
    if "impact" in s or "development" in s:       return "Impact"
    return "Framing"

# ── PAGE 1 ────────────────────────────────────────────────────────────────────
def load_p1(path):
    df = pd.read_excel(path, sheet_name=SHEET_P1, header=1)
    df.columns = [c.strip() for c in df.columns]
    df = df[df["Team ID"].notna()].copy()
    records = []
    for _, r in df.iterrows():
        lead = f"{clean(r.get('Prénom du lead fiabilisé')) or ''} {clean(r.get('Nom du lead fiabilisé')) or ''}".strip() or None
        records.append({
            "name":    clean(r.get("Project name")),
            "lead":    lead,
            "desc":    clean(r.get("Project description")),
            "status":  clean(r.get("Current status")),
            "stage":   to_stage(clean(r.get("Current step name"))),
            "sbu":     sbu_norm(r.get("SBU/Filiales de rattachement")),
            "type":    clean(r.get("Type de Situation")),
            "horizon": clean(r.get("Horizon Stratégique")),
            "theme":   clean(r.get("Thématique principale")),
        })
    return records

# ── PAGE 2 ────────────────────────────────────────────────────────────────────
def load_p2(path):
    df = pd.read_excel(path, sheet_name=SHEET_P2, header=2)
    df.columns = [c.strip() for c in df.columns]
    df = df[df["Team ID"].notna()].copy()

    situations = []
    for _, r in df.iterrows():
        raw_pw = clean(r.get("Parcours d'accompagnement adapté .1"))
        pk     = PATHWAY_MAP.get(raw_pw)
        if not pk: continue

        lead = f"{clean(r.get('Prénom du lead fiabilisé')) or ''} {clean(r.get('Nom du lead fiabilisé')) or ''}".strip() or None
        situations.append({
            "id":          int(r["Team ID"]) if pd.notna(r["Team ID"]) else None,
            "name":        clean(r.get("Project name")),
            "lead":        lead,
            "sbu":         sbu_norm(r.get("SBU/Filiales de rattachement")),
            "site":        site_norm(r.get("Site de ratt")),
            "pathway":     raw_pw,
            "pathway_key": pk,
            "organisme":   PATHWAY_ORG[pk],
            "demarre":     str(r.get("Validation DD","")).strip().lower() == "oui",
            "step":        clean(r.get("Current step name")),
        })
    return situations

def build_p2(situations):
    pks = ["AB","C","D","QW","F"]

    p2_stats = {}
    for pk in pks:
        g = [s for s in situations if s["pathway_key"]==pk]
        p2_stats[pk] = {"total": len(g), "demarre": sum(1 for s in g if s["demarre"])}

    sbu_section = {}
    for pk in pks:
        g = [s for s in situations if s["pathway_key"]==pk]
        counts = {}
        for s in g:
            k = s["sbu"] or "Autre"; counts[k] = counts.get(k,0)+1
        counts["Total"] = p2_stats[pk]["total"]
        sbu_section[PATHWAY_LABEL[pk]] = counts

    site_section = {}
    for pk in pks:
        g = [s for s in situations if s["pathway_key"]==pk]
        counts = {}
        for s in g:
            k = s["site"] or "Autre"; counts[k] = counts.get(k,0)+1
        counts["Total"] = p2_stats[pk]["total"]
        site_section[PATHWAY_LABEL[pk]] = counts

    org_counts = {"LaunchX":0,"Le HUB":0,"MindX":0}
    for pk in pks:
        org_counts[PATHWAY_ORG[pk]] = org_counts.get(PATHWAY_ORG[pk],0) + p2_stats[pk]["total"]

    return {
        "situations":   situations,
        "sbu_section":  sbu_section,
        "site_section": site_section,
        "total":        len(situations),
        "p2_stats":     p2_stats,
        "org_counts":   org_counts,
    }

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    if not EXCEL_FILE.exists():
        print(f"❌  Fichier introuvable : {EXCEL_FILE}"); import sys; sys.exit(1)

    print(f"📖  Lecture de {EXCEL_FILE.name}...")
    p1 = load_p1(EXCEL_FILE)
    p2 = build_p2(load_p2(EXCEL_FILE))

    data = {
        "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
        "p1": {"records": p1},
        "p2": p2,
    }

    OUTPUT_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"✅  data.json généré → {OUTPUT_FILE}")
    print(f"   P1 : {len(p1)} situations")
    print(f"   P2 : {p2['total']} avec parcours")
    for pk,s in p2["p2_stats"].items():
        print(f"      {pk}: {s['total']} total · {s['demarre']} démarrés")

if __name__ == "__main__":
    main()
