"""
Étape 1 - Nettoyage du classeur pluviométrique 2026 de la DRDR de Diourbel.

Ce script lit les feuilles mensuelles (Juin -> Octobre), extrait chaque bloc
« Semaine du ... », nettoie les valeurs, vérifie la cohérence des chiffres et
écrit des tables propres (SQLite + CSV) ainsi qu'un rapport d'anomalies.

Utilisation :
    python nettoyage.py chemin/vers/le_fichier.xlsx
"""
import re
import db
import sys
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

# --------------------------------------------------------------------------
# Paramètres
# --------------------------------------------------------------------------
MOIS = ["Juin", "Juillet", "Août", "Septembre", "Octobre"]
TOLERANCE_MM = 0.15  # écart d'arrondi accepté dans les tests (en mm)

# Les mois pré-remplis (ex. Octobre) contiennent un modèle de semaines
# futures qui n'ont pas encore eu lieu. On les ignore en s'arrêtant à la date
# de fin annoncée par le fichier lui-même, lue dans son nom
# ("... du 25 Sept au 1er oct 2026 ...") plutôt qu'écrite en dur ici : chaque
# nouveau bulletin hebdomadaire de la DRDR porte sa propre date de fin.
NOMS_MOIS = {"janv": 1, "jan": 1, "fev": 2, "fevr": 2, "mars": 3, "avr": 4,
            "mai": 5, "juin": 6, "juil": 7, "aout": 8, "sept": 9, "oct": 10,
            "nov": 11, "dec": 12}


def date_fin_depuis_nom(chemin):
    """Extrait 'au 1er oct 2026' / 'au 08 octobre 2026' du nom du fichier.
    Renvoie None si le nom ne suit pas ce format (à l'appelant de réagir)."""
    sans_accent = unicodedata.normalize("NFD", Path(chemin).stem)
    sans_accent = "".join(c for c in sans_accent
                          if unicodedata.category(c) != "Mn").lower()
    m = re.search(r"au[_ ]?(\d{1,2})(?:er)?[_ ]([a-z]+)[_ ](\d{4})",
                 sans_accent)
    if not m:
        return None
    jour, mot, an = m.groups()
    mois = NOMS_MOIS.get(mot[:5].rstrip(".")) or NOMS_MOIS.get(mot[:4])
    if mois is None:
        return None
    return date(int(an), mois, int(jour))

# Référentiel des postes : nom normalisé -> (nom officiel, département)
REFERENTIEL = {
    "bambey": ("Bambey", "Bambey"), "babagarage": ("Baba Garage", "Bambey"),
    "lambaye": ("Lambaye", "Bambey"), "ngoye": ("Ngoye", "Bambey"),
    "cnra": ("CNRA", "Bambey"), "ndondol": ("Ndondol", "Bambey"),
    "keursambakane": ("Keur Samba Kane", "Bambey"),
    "refane": ("Réfane", "Bambey"), "gawane": ("Gawane", "Bambey"),
    "diourbel": ("Diourbel", "Diourbel"), "ndindy": ("Ndindy", "Diourbel"),
    "ndoulo": ("Ndoulo", "Diourbel"),
    "ndankhsene": ("Ndankh Sène", "Diourbel"),
    "tockygare": ("Tocky Gare", "Diourbel"),
    "mbacke": ("Mbacké", "Mbacké"), "ndame": ("Ndame", "Mbacké"),
    "kael": ("Kael", "Mbacké"), "taif": ("Taif", "Mbacké"),
    "sadio": ("Sadio", "Mbacké"),
}

# Positions des colonnes dans chaque bloc hebdomadaire (numéros Excel)
COL_POSTE, COL_CUM_ANT_MM, COL_CUM_ANT_NJ = 1, 2, 3
COLS_JOURS = range(4, 11)          # D..J : les 7 jours de la semaine
COL_SEM_MM, COL_SEM_NJ = 11, 12    # K, L
COL_2026_MM, COL_2026_NJ = 13, 14  # M, N
COL_2025_MM, COL_2025_NJ = 15, 16  # O, P
COL_NORMALE, COL_ECART_NORMALE = 19, 20  # S, T

anomalies = []
DERNIERE_DATE = None  # fixée dans main(), avant tout appel à lire_feuille


def signaler(feuille, ligne, poste, type_, detail):
    """Note un problème rencontré, pour le rapport final."""
    anomalies.append(dict(feuille=feuille, ligne=ligne, poste=poste,
                          type=type_, detail=detail))


# --------------------------------------------------------------------------
# Fonctions de conversion
# --------------------------------------------------------------------------
def normaliser(nom):
    """'NdankhSène' / 'NDANKH SENE' -> 'ndankhsene' (clé de comparaison)."""
    sans_accent = unicodedata.normalize("NFD", str(nom))
    sans_accent = "".join(c for c in sans_accent
                          if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", sans_accent.lower())


def vers_date(v):
    """Date Excel (datetime) ou numéro de série (46178) -> date."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, (int, float)):
        return (datetime(1899, 12, 30) + timedelta(days=float(v))).date()
    return None


def vers_nombre(v):
    """Renvoie (valeur, statut). 'tr' = trace, 'np' = non parvenu."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return None, "vide"
    if isinstance(v, (int, float)):
        return float(v), "mesure"
    texte = str(v).strip().lower()
    if texte == "tr":
        return 0.0, "trace"
    if texte == "np":
        return None, "non_parvenu"
    return None, f"inconnu:{texte}"


# --------------------------------------------------------------------------
# Lecture des feuilles
# --------------------------------------------------------------------------
def lire_feuille(ws, hebdo, jours):
    entetes = [r for r in range(1, ws.max_row + 1)
               if isinstance(ws.cell(r, 4).value, str)
               and "emaine du" in ws.cell(r, 4).value]
    for i, debut in enumerate(entetes):
        fin_bloc = entetes[i + 1] - 1 if i + 1 < len(entetes) else ws.max_row
        dates = [vers_date(ws.cell(debut + 1, c).value) for c in COLS_JOURS]
        if None in dates:
            signaler(ws.title, debut + 1, "", "dates_illisibles",
                     "Impossible de lire les 7 dates de la semaine")
            continue
        if dates[0] > DERNIERE_DATE:
            continue  # semaine future (modèle vide)

        for r in range(debut + 2, fin_bloc + 1):
            nom = ws.cell(r, COL_POSTE).value
            if nom is None or not str(nom).strip():
                continue
            if "moyenne" in str(nom).lower():
                break  # ligne de moyennes : recalculée plus tard, on l'ignore
            cle = normaliser(nom)
            if cle not in REFERENTIEL:
                signaler(ws.title, r, nom, "poste_inconnu",
                         "Nom absent du référentiel des 19 postes")
                continue
            poste, dept = REFERENTIEL[cle]
            if str(nom).strip() != poste:
                signaler(ws.title, r, nom, "nom_harmonise", f"-> {poste}")

            def val(col):
                v, statut = vers_nombre(ws.cell(r, col).value)
                if statut.startswith("inconnu"):
                    signaler(ws.title, r, poste, "valeur_inconnue",
                             f"Colonne {col} : {ws.cell(r, col).value!r}")
                return v, statut

            # Pluie journalière (uniquement les cases renseignées)
            somme_jours, jours_np = 0.0, False
            for c, d in zip(COLS_JOURS, dates):
                v, statut = val(c)
                if statut == "vide":
                    continue
                jours.append(dict(poste=poste, date=d, mm=v, statut=statut))
                if statut == "trace":
                    signaler(ws.title, r, poste, "trace",
                             f"{d} : trace (comptée 0 mm)")
                if statut == "non_parvenu":
                    jours_np = True
                    signaler(ws.title, r, poste, "non_parvenu", f"{d} : np")
                somme_jours += v or 0.0

            mm_sem, st_sem = val(COL_SEM_MM)
            hebdo.append(dict(
                poste=poste, departement=dept,
                semaine_debut=dates[0], semaine_fin=dates[-1],
                cumul_anterieur_mm=val(COL_CUM_ANT_MM)[0],
                mm_semaine=mm_sem, statut_semaine=st_sem,
                somme_jours_mm=None if jours_np else round(somme_jours, 2),
                cumul_2026_mm=val(COL_2026_MM)[0],
                cumul_2026_nj=val(COL_2026_NJ)[0],
                cumul_2025_mm=val(COL_2025_MM)[0],
                normale_mm=val(COL_NORMALE)[0],
                ecart_normale_mm=val(COL_ECART_NORMALE)[0],
                feuille=ws.title, ligne=r,
            ))


# --------------------------------------------------------------------------
# Tests de cohérence
# --------------------------------------------------------------------------
def tests(h):
    resultats = []

    def enregistrer(nom, mauvais):
        resultats.append((nom, len(mauvais) == 0, len(mauvais)))
        for _, ligne in mauvais.iterrows():
            signaler(ligne["feuille"], ligne["ligne"], ligne["poste"],
                     f"test_echoue: {nom}",
                     f"semaine du {ligne['semaine_debut']}")

    # 1. Chaque semaine doit contenir les 19 postes
    par_semaine = h.groupby("semaine_debut")["poste"].nunique()
    incomplet = par_semaine[par_semaine != 19]
    resultats.append(("19 postes par semaine", incomplet.empty, len(incomplet)))
    for d, n in incomplet.items():
        signaler("", "", "", "test_echoue: 19 postes", f"{d} : {n} postes")

    # 2. Somme des jours = pluie de la semaine
    ok = h["somme_jours_mm"].notna() & h["mm_semaine"].notna()
    ecart = (h["somme_jours_mm"] - h["mm_semaine"]).abs()
    enregistrer("somme des jours = cumul semaine",
                h[ok & (ecart > TOLERANCE_MM)])

    # 3. Cumul antérieur + semaine = cumul 2026
    calc = h["cumul_anterieur_mm"] + h["mm_semaine"].fillna(0)
    ok = h["cumul_2026_mm"].notna() & h["cumul_anterieur_mm"].notna()
    enregistrer("cumul antérieur + semaine = cumul 2026",
                h[ok & ((calc - h["cumul_2026_mm"]).abs() > TOLERANCE_MM)])

    # 4. Continuité : cumul 2026 de la semaine n-1 = cumul antérieur de la n
    h = h.sort_values(["poste", "semaine_debut"])
    prec = h.groupby("poste")["cumul_2026_mm"].shift(1)
    ok = prec.notna() & h["cumul_anterieur_mm"].notna()
    enregistrer("continuité d'une semaine à l'autre",
                h[ok & ((prec - h["cumul_anterieur_mm"]).abs() > TOLERANCE_MM)])

    # 5. Le cumul ne peut pas diminuer
    diminue = h["cumul_2026_mm"] < prec - TOLERANCE_MM
    enregistrer("cumul jamais décroissant", h[diminue])

    # 6. Valeurs plausibles (pas de pluie négative)
    enregistrer("pluie non négative", h[h["mm_semaine"] < 0])
    return resultats


# --------------------------------------------------------------------------
# Programme principal
# --------------------------------------------------------------------------
def main(chemin_xlsx, dossier_sortie, date_fin=None):
    global DERNIERE_DATE
    if date_fin is not None:
        DERNIERE_DATE = date_fin
    else:
        DERNIERE_DATE = date_fin_depuis_nom(chemin_xlsx)
        if DERNIERE_DATE is None:
            raise SystemExit(
                "Impossible de déterminer la date de fin des données à "
                "partir du nom du fichier (attendu : '...du J MOIS au J "
                "MOIS AAAA...', ex. 'du 25 Sept au 1er oct 2026'). "
                "Renomme le fichier selon ce modèle, ou relance avec : "
                "python nettoyage.py <fichier.xlsx> AAAA-MM-JJ")
    print(f"Date de fin des données prise en compte : {DERNIERE_DATE} "
         "(les semaines postérieures, pré-remplies, sont ignorées)")

    wb = load_workbook(chemin_xlsx, data_only=True)
    hebdo, jours = [], []
    for nom_feuille in MOIS:
        lire_feuille(wb[nom_feuille], hebdo, jours)

    h = pd.DataFrame(hebdo)
    j = pd.DataFrame(jours)
    resultats = tests(h)

    postes = pd.DataFrame(
        [(p, d) for p, d in REFERENTIEL.values()],
        columns=["poste", "departement"])
    rapport = pd.DataFrame(anomalies)

    sortie = Path(dossier_sortie)
    sortie.mkdir(parents=True, exist_ok=True)
    moteur = db.obtenir_moteur(dossier_sortie)
    with moteur.begin() as con:
        postes.to_sql("postes", con, if_exists="replace", index=False)
        h.to_sql("observations_hebdo", con, if_exists="replace", index=False)
        j.to_sql("mesures_journalieres", con, if_exists="replace", index=False)
        rapport.to_sql("anomalies", con, if_exists="replace", index=False)
    h.to_csv(sortie / "observations_hebdo.csv", index=False)
    rapport.to_csv(sortie / "rapport_anomalies.csv", index=False)

    print(f"Semaines lues : {h['semaine_debut'].nunique()} "
          f"({h['semaine_debut'].min()} -> {h['semaine_fin'].max()})")
    print(f"Lignes poste x semaine : {len(h)} | jours renseignés : {len(j)}")
    print("\nTests de cohérence :")
    for nom, ok, n in resultats:
        print(f"  [{'OK' if ok else 'ECHEC'}] {nom}" + ("" if ok else f" ({n})"))
    if not rapport.empty:
        print("\nAnomalies par type :")
        print(rapport["type"].value_counts().to_string())


if __name__ == "__main__":
    fichier = sys.argv[1]
    date_manuelle = (datetime.strptime(sys.argv[2], "%Y-%m-%d").date()
                     if len(sys.argv) > 2 else None)
    main(fichier, Path(__file__).parent / "donnees_propres", date_manuelle)
