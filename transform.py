#!/usr/bin/env python3
"""
Le Mouvement — Script de transformation Excel → data.json
Usage : python transform.py [chemin_vers_excel]
Par défaut : DATA_Situations_Mouvement.xlsx dans le même dossier

Règles métier :
- P1 : toutes les situations (actives + discarded)
- P2 : actives + étape soumission/catégorisation/idéation + parcours assigné
- Démarré = Période de mentorat non vide et != "Non initié"
- Demo Day = Validation DD = "Oui" dans Etape_Categorization
- Hacking Committee = Version HC non vide dans RAW_Incubation
- Impact Gate = étape contient "Impact"
- Date affichée = date du fichier Excel
"""

import sys
import json
import os
from pathlib import Path
from datetime import datetime
import pandas as pd

EXCEL_FILE  = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "DATA_Situations_Mouvement.xlsx"
OUTPUT_FILE = Path(__file__).parent / "data.json"

SHEET_P1  = "DATA_Consolidée"
SHEET_P2  = "Etape_Categorization"
SHEET_INC = "RAW_Incubation"

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
    "AB": "Le HUB", "C": "LaunchX", "D": "LaunchX",
    "QW": "MindX",  "F": "Le HUB",
}

SBU_KEEP = ["Mining", "Manufacturing", "Corporate", "UM6P", "InnovX",
            "Fondation Phosboucraa", "Nutricrops", "OTED", "Rock", "SPS"]
SITES    = ["Khouribga", "Phosboucraa", "Jorf Lasfar", "Safi", "Gantour", "Casablanca"]

# Étapes du périmètre P2
PERIM_STEPS = ["*Soumission*", "Soumission", "Categorization", "Ideation"]

# ── Helpers ────────────────────────────────────────────────────────────────────

def clean(val):
    if pd.isna(val): return None
    s = str(val).strip()
    return s if s and s not in ("00:00:00", "nan") else None

def sbu_normalize(val):
    s = clean(val)
    if not s: return "Autre"
    for k in SBU_KEEP:
        if k.lower() in s.lower(): return k
    return "Autre"

def site_normalize(val):
    s = clean(val)
    if not s: return "Autre"
    for site in SITES:
        if site.lower() in s.lower(): return site
    return "Autre"

def to_stage(step):
    """Mappe l'étape vers un stage simplifié."""
    if not step: return "Framing"
    s = step.lower()
    if "soumission" in s:                          return "Framing"
    if "categor" in s:                             return "Ideation"
    if "ideation" in s:                            return "Ideation"
    if "pre-hack" in s or "pre_hack" in s:         return "Incubation"
    if "hacking" in s:                             return "Incubation"
    if "incubation" in s:                          return "Incubation"
    if "acceleration" in s or "acceler" in s:      return "Acceleration"
    if "impact" in s:                              return "Impact"
    return "Framing"

def is_demarre(val):
    """Démarré = Période de mentorat non vide et != 'Non initié'."""
    if pd.isna(val): return False
    s = str(val).strip()
    return s != "" and s.lower() not in ("non initié", "00:00:00", "nan")

def truncate_desc(desc, max_words=30):
    """Tronque proprement à max_words sans couper en milieu de phrase."""
    if not desc: return ""
    words = desc.split()
    if len(words) <= max_words: return desc
    truncated = " ".join(words[:max_words])
    for punct in [". ", ", ", "; ", ": "]:
        last = truncated.rfind(punct)
        if last > len(truncated) * 0.5:
            return truncated[:last + 1].strip()
    return truncated.strip() + "…"

# ── Maps de référence ──────────────────────────────────────────────────────────

def get_demo_ids(df2):
    """Demo Day = Validation DD = 'Oui' dans Etape_Categorization."""
    mask = df2['Validation DD'].astype(str).str.lower() == 'oui'
    return set(df2[mask]['Team ID'].dropna().astype(int))

def get_hc_ids(path):
    """Hacking Committee = Version HC non vide dans RAW_Incubation."""
    df = pd.read_excel(path, sheet_name=SHEET_INC, header=0)
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    mask = (df['Version HC'].notna() &
            (df['Version HC'].astype(str).str.strip() != '') &
            (df['Version HC'].astype(str).str.strip() != 'nan'))
    return set(df[mask]['Team ID'].dropna().astype(int))

# ── P1 ─────────────────────────────────────────────────────────────────────────

def load_p1(df1, demo_ids, hc_ids):
    records = []
    for _, r in df1[df1['Team ID'].notna()].iterrows():
        tid   = int(r['Team ID'])
        lead  = f"{clean(r.get('Prénom du lead fiabilisé')) or ''} {clean(r.get('Nom du lead fiabilisé')) or ''}".strip() or None
        step  = clean(r.get('Current step name'))
        goal  = clean(r.get('Objectif Sratégique')) or clean(r.get('Objectif stratégique'))
        records.append({
            "name":        clean(r.get('Project name')),
            "lead":        lead,
            "desc":        truncate_desc(clean(r.get('Project description')) or ''),
            "status":      clean(r.get('Current status')),
            "stage":       to_stage(step),
            "sbu":         sbu_normalize(r.get('SBU/Filiales de rattachement')),
            "site":        site_normalize(r.get('Site de rattachement')),
            "type":        clean(r.get('Type de Situation')),
            "horizon":     clean(r.get('Horizon Stratégique')),
            "theme":       clean(r.get('Thématique principale')),
            "goal":        goal,
            "demo_day":    tid in demo_ids,
            "hacking":     tid in hc_ids,
            "impact_gate": bool(step and 'impact' in step.lower()),
        })
    return records

# ── P2 ─────────────────────────────────────────────────────────────────────────

def load_p2(df1, df2):
    """
    P2 = situations actives + périmètre soum/catégo/idéation + parcours assigné.
    236 situations dans les données actuelles.
    """
    # IDs du périmètre : actives + étape soum/catégo/idéation
    actives   = df1[df1['Current status'] == 'active']
    df1_perim = actives[actives['Current step name'].isin(PERIM_STEPS)]
    ids_perim = set(df1_perim['Team ID'].dropna().astype(int))

    pk_col     = "Parcours d'accompagnement adapté .1"
    period_col = "Période de mentorat"

    situations = []
    for _, r in df2[df2[pk_col].isin(PATHWAY_MAP.keys())].iterrows():
        tid = int(r['Team ID']) if pd.notna(r['Team ID']) else None
        if tid not in ids_perim: continue  # filtre périmètre
        raw_pk = clean(r.get(pk_col))
        pk     = PATHWAY_MAP.get(raw_pk)
        if not pk: continue
        lead = f"{clean(r.get('Prénom du lead fiabilisé')) or ''} {clean(r.get('Nom du lead fiabilisé')) or ''}".strip() or None
        situations.append({
            "id":          tid,
            "name":        clean(r.get('Project name')),
            "lead":        lead,
            "sbu":         sbu_normalize(r.get('SBU/Filiales de rattachement')),
            "site":        site_normalize(r.get('Site de ratt')),
            "pathway":     raw_pk,
            "pathway_key": pk,
            "organisme":   PATHWAY_ORG[pk],
            "demarre":     is_demarre(r.get(period_col)),
            "periode":     clean(r.get(period_col)) or "Non initié",
            "step":        clean(r.get('Current step name')),
        })
    return situations, ids_perim, len(df1_perim)

# ── Agrégations P2 ─────────────────────────────────────────────────────────────

def build_p2_output(situations, df2, ids_perim, total_perim):
    pathway_keys = ["AB", "C", "D", "QW", "F"]
    p2_stats, sbu_section, site_section = {}, {}, {}

    for pk in pathway_keys:
        g   = [s for s in situations if s['pathway_key'] == pk]
        dem = [s for s in g if s['demarre']]
        p2_stats[pk] = {"total": len(g), "demarre": len(dem)}
        lbl = PATHWAY_LABEL[pk]
        sc, sic = {}, {}
        for s in g:
            sc[s['sbu']]   = sc.get(s['sbu'], 0) + 1
            sic[s['site']] = sic.get(s['site'], 0) + 1
        sc["Total"] = len(g); sic["Total"] = len(g)
        sbu_section[lbl]  = sc
        site_section[lbl] = sic

    org_counts = {"LaunchX": 0, "Le HUB": 0, "MindX": 0}
    for pk in pathway_keys:
        org_counts[PATHWAY_ORG[pk]] += p2_stats[pk]['total']

    # Périmètre
    pk_col        = "Parcours d'accompagnement adapté .1"
    df2_with_pk   = df2[df2[pk_col].isin(PATHWAY_MAP.keys())]
    ids_with_pk   = set(df2_with_pk['Team ID'].dropna().astype(int))
    avec_parcours = len(ids_perim & ids_with_pk)
    sans_parcours = total_perim - avec_parcours
    pct_cat       = round(avec_parcours / total_perim * 100, 1) if total_perim > 0 else 0
    total_demarre = sum(p2_stats[pk]['demarre'] for pk in pathway_keys)
    pct_demarre   = round(total_demarre / avec_parcours * 100, 1) if avec_parcours > 0 else 0

    return {
        "situations":   situations,
        "sbu_section":  sbu_section,
        "site_section": site_section,
        "total":        len(situations),
        "p2_stats":     p2_stats,
        "org_counts":   org_counts,
        "perimetre": {
            "total":          total_perim,
            "avec_parcours":  avec_parcours,
            "sans_parcours":  sans_parcours,
            "pct_cat":        pct_cat,
            "total_demarre":  total_demarre,
            "pct_demarre":    pct_demarre,
        },
    }

# ── Main ────────────────────────────────────────────────────────────────────────

def main():
    if not EXCEL_FILE.exists():
        print(f"❌  Fichier introuvable : {EXCEL_FILE}")
        sys.exit(1)

    print(f"📖  Lecture de {EXCEL_FILE.name}...")

    file_date = datetime.fromtimestamp(os.path.getmtime(EXCEL_FILE)).strftime("%d/%m/%Y")

    df1 = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_P1, header=1)
    df1.columns = [c.strip() for c in df1.columns]
    df2 = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_P2, header=2)
    df2.columns = [c.strip() for c in df2.columns]

    demo_ids = get_demo_ids(df2)
    hc_ids   = get_hc_ids(EXCEL_FILE)
    print(f"   Demo Day IDs        : {len(demo_ids)}")
    print(f"   Hacking Comm. IDs   : {len(hc_ids)}")

    p1_records              = load_p1(df1, demo_ids, hc_ids)
    situations, ids_perim, total_perim = load_p2(df1, df2)
    p2_output               = build_p2_output(situations, df2, ids_perim, total_perim)

    data = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "file_date":    file_date,
        "p1": {"records": p1_records},
        "p2": p2_output,
    }

    OUTPUT_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    p = p2_output['perimetre']
    print(f"✅  data.json généré → {OUTPUT_FILE}")
    print(f"   Date fichier        : {file_date}")
    print(f"   P1                  : {len(p1_records)} situations")
    print(f"   Périmètre P2        : {p['total']} (soum/catégo/idéa actives)")
    print(f"   Avec parcours       : {p['avec_parcours']} ({p['pct_cat']}%)")
    print(f"   Sans parcours       : {p['sans_parcours']}")
    print(f"   Démarrés            : {p['total_demarre']} / {p['avec_parcours']} ({p['pct_demarre']}%)")
    print(f"   Détail pathways :")
    for pk, s in p2_output['p2_stats'].items():
        print(f"      {pk} : {s['total']} situatons · {s['demarre']} démarrés")

if __name__ == "__main__":
    main()
