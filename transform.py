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

# ── CONFIG ──────────────────────────────────────────────────────────────────
EXCEL_FILE = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "DATA_Situations_Mouvement.xlsx"
OUTPUT_FILE = Path(__file__).parent / "data.json"

SHEET_P1 = "DATA_Consolidée"
SHEET_P2 = "Etape_Categorization"

# Mapping parcours label Excel → clé interne dashboard
PATHWAY_MAP = {
    "CAT A&B -> Idéation":            "AB",
    "CAT C -> Pré-Idéation":          "C",
    "CAT D -> Pré-Idéation":          "D",
    "Quick wins -> Move Up":          "QW",
    "Entrepreneurial Certificate":    "F",
}

# Label lisible pour sbu_section / site_section (format attendu par le JS)
PATHWAY_LABEL = {
    "AB": "CAT A&B -> Demo Day",
    "C":  "CAT C -> Pré-idéation",
    "D":  "CAT D -> Pré-idéation",
    "QW": "Quick Wins -> Move Up",
    "F":  "Entrepreneurial Certificate",
}

PATHWAY_ORG = {
    "AB": "Le HUB",
    "C":  "LaunchX",
    "D":  "LaunchX",
    "QW": "MindX",
    "F":  "Le HUB",
}

SBU_KEEP = ["Mining", "Manufacturing", "Corporate", "UM6P", "InnovX",
            "Fondation Phosboucraa", "Nutricrops", "OTED", "Rock", "SPS"]

SITES = ["Khouribga", "Phosboucraa", "Jorf Lasfar", "Safi", "Gantour", "Casablanca"]

# ── HELPERS ──────────────────────────────────────────────────────────────────
def clean(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    return s if s and s not in ("00:00:00", "nan") else None

def sbu_normalize(val):
    s = clean(val)
    if not s:
        return "Autre"
    for kept in SBU_KEEP:
        if kept.lower() in s.lower():
            return kept
    return "Autre"

def site_normalize(val):
    s = clean(val)
    if not s:
        return "Autre"
    for site in SITES:
        if site.lower() in s.lower():
            return site
    return "Autre"

# ── LECTURE P1 ────────────────────────────────────────────────────────────────
def load_p1(path):
    df = pd.read_excel(path, sheet_name=SHEET_P1, header=1)
    df.columns = [c.strip() for c in df.columns]
    df = df[df["Team ID"].notna()].copy()

    records = []
    for _, r in df.iterrows():
        status = clean(r.get("Current status"))
        step   = clean(r.get("Current step name"))
        sbu    = sbu_normalize(r.get("SBU/Filiales de rattachement"))
        typ    = clean(r.get("Type de Situation"))
        horiz  = clean(r.get("Horizon Stratégique"))
        theme  = clean(r.get("Thématique principale"))
        name   = clean(r.get("Project name"))
        desc   = clean(r.get("Project description"))
        lead_p = clean(r.get("Prénom du lead fiabilisé"))
        lead_n = clean(r.get("Nom du lead fiabilisé"))
        lead   = f"{lead_p or ''} {lead_n or ''}".strip() or None
        stage  = _to_stage(step)

        records.append({
            "name":    name,
            "lead":    lead,
            "desc":    desc,
            "status":  status,
            "stage":   stage,
            "sbu":     sbu,
            "type":    typ,
            "horizon": horiz,
            "theme":   theme,
        })
    return records

def _to_stage(step):
    if not step:
        return "Framing"
    s = step.lower()
    if "soumission" in s:
        return "Framing"
    if "categor" in s or "qualification" in s:
        return "Ideation"
    if "ideation" in s:
        return "Ideation"
    if "incubation" in s or "pre-hack" in s or "hacking" in s:
        return "Incubation"
    if "acceleration" in s or "acceler" in s:
        return "Acceleration"
    if "impact" in s or "development" in s:
        return "Impact"
    return "Framing"

# ── LECTURE P2 ────────────────────────────────────────────────────────────────
def load_p2(path):
    df = pd.read_excel(path, sheet_name=SHEET_P2, header=2)
    df.columns = [c.strip() for c in df.columns]
    df = df[df["Team ID"].notna()].copy()

    parcours_label_col = "Parcours d'accompagnement adapté .1"
    validation_col     = "Validation DD"
    sbu_col            = "SBU/Filiales de rattachement"
    site_col           = "Site de ratt"

    situations = []
    for _, r in df.iterrows():
        raw_pathway = clean(r.get(parcours_label_col))
        pathway_key = PATHWAY_MAP.get(raw_pathway)
        if not pathway_key:
            continue  # Ignorer les situations sans pathway

        demarre    = str(r.get(validation_col, "")).strip().lower() == "oui"
        sbu        = sbu_normalize(r.get(sbu_col))
        site       = site_normalize(r.get(site_col))
        organisme  = PATHWAY_ORG.get(pathway_key, "")
        lead_p     = clean(r.get("Prénom du lead fiabilisé"))
        lead_n     = clean(r.get("Nom du lead fiabilisé"))
        lead       = f"{lead_p or ''} {lead_n or ''}".strip() or None
        step       = clean(r.get("Current step name"))

        situations.append({
            "id":          int(r["Team ID"]) if pd.notna(r["Team ID"]) else None,
            "name":        clean(r.get("Project name")),
            "lead":        lead,
            "sbu":         sbu,
            "site":        site,
            "pathway":     raw_pathway,
            "pathway_key": pathway_key,
            "organisme":   organisme,
            "demarre":     demarre,
            "step":        step,
        })
    return situations

def build_p2_output(situations):
    """Construit le bloc P2 dans le format exact attendu par le dashboard JS."""
    pathway_keys = ["AB", "C", "D", "QW", "F"]

    # p2_stats
    p2_stats = {}
    for pk in pathway_keys:
        group   = [s for s in situations if s["pathway_key"] == pk]
        demarre = [s for s in group if s["demarre"]]
        p2_stats[pk] = {"total": len(group), "demarre": len(demarre)}

    # sbu_section : {"CAT A&B -> Demo Day": {Mining:N, ..., Total:N}, ...}
    sbu_section = {}
    for pk in pathway_keys:
        label = PATHWAY_LABEL[pk]
        group = [s for s in situations if s["pathway_key"] == pk]
        counts = {}
        for s in group:
            k = s["sbu"] or "Autre"
            counts[k] = counts.get(k, 0) + 1
        counts["Total"] = p2_stats[pk]["total"]
        sbu_section[label] = counts

    # site_section
    site_section = {}
    for pk in pathway_keys:
        label = PATHWAY_LABEL[pk]
        group = [s for s in situations if s["pathway_key"] == pk]
        counts = {}
        for s in group:
            k = s["site"] or "Autre"
            counts[k] = counts.get(k, 0) + 1
        counts["Total"] = p2_stats[pk]["total"]
        site_section[label] = counts

    # org_counts (pour le donut des orgs)
    org_counts = {"LaunchX": 0, "Le HUB": 0, "MindX": 0}
    for pk in pathway_keys:
        org = PATHWAY_ORG[pk]
        org_counts[org] = org_counts.get(org, 0) + p2_stats[pk]["total"]

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
        print(f"❌  Fichier introuvable : {EXCEL_FILE}")
        sys.exit(1)

    print(f"📖  Lecture de {EXCEL_FILE.name}...")

    p1_records  = load_p1(EXCEL_FILE)
    p2_situs    = load_p2(EXCEL_FILE)
    p2_output   = build_p2_output(p2_situs)

    print(f"   ✓ {len(p1_records)} situations P1 chargées")
    print(f"   ✓ {p2_output['total']} situations P2 avec parcours assigné")

    data = {
        "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
        "p1": {
            "records": p1_records,
        },
        "p2": p2_output,
    }

    OUTPUT_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"✅  data.json généré → {OUTPUT_FILE}")
    print(f"   P1 : {len(p1_records)} situations")
    print(f"   P2 : {p2_output['total']} avec parcours · Détail par pathway :")
    for pk, s in p2_output["p2_stats"].items():
        print(f"      {pk}: {s['total']} total, {s['demarre']} démarrés")

if __name__ == "__main__":
    main()
