"""
Mise à jour hebdomadaire des données - à lancer à chaque nouveau bulletin.

Ce script fait, dans l'ordre, tout ce que nettoyage.py, nettoyage_reference.py
et indicateurs.py font séparément, avec en plus :
  - une sauvegarde du bulletin précédent dans historique/, pour ne jamais
    perdre une ancienne version ;
  - un résumé de ce qui a changé (semaines ajoutées, valeurs corrigées).

Utilisation (le plus courant) :
    python maj_donnees.py "Pluviométrie_du_02_au_08_oct_2026_DRDR_Diourbel.xlsx"

Si le nom du fichier ne suit pas le modèle habituel de la DRDR :
    python maj_donnees.py fichier.xlsx --date-fin 2026-10-08
"""
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

import db
import indicateurs
import nettoyage
import nettoyage_reference

DOSSIER = Path(__file__).parent / "donnees_propres"
HISTORIQUE = Path(__file__).parent / "historique"
TOL = 0.1  # mm : en dessous, un écart n'est pas signalé comme un changement

# Toutes les tables de données, pour l'archivage et les tests de
# disponibilité - qu'elles soient dans le fichier SQLite local ou dans une
# base en ligne (Supabase), la logique est la même.
TABLES = ["postes", "observations_hebdo", "mesures_journalieres",
         "anomalies", "annuel", "normales", "semaines",
         "indicateurs_poste_semaine", "indicateurs_agreges_semaine",
         "rang_2026"]


def archiver_version_precedente(moteur):
    """Exporte chaque table de la base actuelle en CSV dans
    historique/AAAA-MM-JJ_HHhMM/ avant de les remplacer, pour toujours
    pouvoir revenir en arrière - que les données soient dans le fichier
    SQLite local ou dans une base en ligne."""
    if not db.donnees_disponibles(moteur):
        return None
    dest = HISTORIQUE / datetime.now().strftime("%Y-%m-%d_%Hh%M")
    dest.mkdir(parents=True, exist_ok=True)
    for table in TABLES:
        try:
            pd.read_sql(f"select * from {table}", moteur).to_csv(
                dest / f"{table}.csv", index=False)
        except Exception:
            pass  # table pas encore créée (tout premier chargement)
    print(f"Ancienne version sauvegardée dans : {dest}")
    return dest


def etat_semaines(moteur):
    """Dernière semaine et nombre de semaines connues avant la mise à jour."""
    if not db.donnees_disponibles(moteur):
        return None
    return pd.read_sql("select distinct semaine_debut, semaine_fin, poste, "
                       "cumul_2026_mm from observations_hebdo", moteur)


def executer(chemin_xlsx, date_manuelle=None):
    """Fait toute la mise à jour et renvoie un résumé structuré, sans rien
    imprimer : c'est ce que utilise l'onglet « Ajouter des données » de
    l'application web. Ne lève jamais d'exception : les erreurs sont
    renvoyées dans le résumé pour être affichées proprement."""
    resume = dict(ok=False, erreur=None, semaines_ajoutees=[],
                  semaines_disparues=[], valeurs_modifiees=[],
                  archive=None)
    try:
        moteur = db.obtenir_moteur()
        avant = etat_semaines(moteur)
        resume["archive"] = archiver_version_precedente(moteur)
        nettoyage.main(chemin_xlsx, DOSSIER, date_manuelle)
        nettoyage_reference.main(chemin_xlsx, DOSSIER)
        indicateurs.main(chemin_xlsx)
    except SystemExit as e:
        resume["erreur"] = str(e)
        return resume
    except Exception as e:  # fichier illisible, feuille manquante, etc.
        resume["erreur"] = f"{type(e).__name__} : {e}"
        return resume

    apres = etat_semaines(db.obtenir_moteur())
    if avant is None:
        resume["semaines_ajoutees"] = sorted(apres.semaine_debut.unique())
    else:
        da, dp = set(avant.semaine_debut), set(apres.semaine_debut)
        resume["semaines_ajoutees"] = sorted(dp - da)
        resume["semaines_disparues"] = sorted(da - dp)
        communes = da & dp
        if communes:
            m = avant.merge(apres, on=["semaine_debut", "poste"],
                            suffixes=("_avant", "_apres"))
            m = m[m.semaine_debut.isin(communes)]
            chg = m[(m.cumul_2026_mm_avant - m.cumul_2026_mm_apres).abs()
                    > TOL]
            resume["valeurs_modifiees"] = chg[
                ["poste", "semaine_debut", "cumul_2026_mm_avant",
                 "cumul_2026_mm_apres"]].to_dict("records")
    resume["ok"] = True
    return resume


def comparer(avant, apres):
    """Affiche les semaines nouvelles et signale toute valeur qui aurait
    changé pour une semaine déjà connue (une correction en amont, par
    exemple)."""
    if avant is None:
        print(f"Première exécution : {apres.semaine_debut.nunique()} "
             "semaine(s) chargée(s).")
        return
    dates_avant = set(avant.semaine_debut)
    dates_apres = set(apres.semaine_debut)
    nouvelles = sorted(dates_apres - dates_avant)
    if nouvelles:
        print(f"Semaine(s) ajoutée(s) : {', '.join(map(str, nouvelles))}")
    else:
        print("Aucune semaine nouvelle par rapport à la version précédente.")
    disparues = sorted(dates_avant - dates_apres)
    if disparues:
        print(f"⚠ Semaine(s) qui n'apparaissent plus dans le nouveau "
             f"fichier : {', '.join(map(str, disparues))} — vérifie que ce "
             "n'est pas un oubli dans le fichier déposé.")

    communes = dates_avant & dates_apres
    if communes:
        m = avant.merge(apres, on=["semaine_debut", "poste"],
                        suffixes=("_avant", "_apres"))
        m = m[m.semaine_debut.isin(communes)]
        ecart = (m.cumul_2026_mm_avant - m.cumul_2026_mm_apres).abs()
        change = m[ecart > TOL]
        if change.empty:
            print("Aucune valeur des semaines déjà connues n'a changé.")
        else:
            print(f"⚠ {len(change)} valeur(s) modifiée(s) pour des semaines "
                 "déjà chargées (la DRDR a corrigé une donnée passée) :")
            for _, r in change.iterrows():
                print(f"    {r.poste}, semaine du {r.semaine_debut} : "
                     f"{r.cumul_2026_mm_avant:.1f} -> "
                     f"{r.cumul_2026_mm_apres:.1f} mm")


def main():
    if len(sys.argv) < 2:
        raise SystemExit("Utilisation : python maj_donnees.py fichier.xlsx "
                         "[--date-fin AAAA-MM-JJ]")
    chemin_xlsx = sys.argv[1]
    date_manuelle = None
    if "--date-fin" in sys.argv:
        date_manuelle = datetime.strptime(
            sys.argv[sys.argv.index("--date-fin") + 1], "%Y-%m-%d").date()

    moteur = db.obtenir_moteur()
    avant = etat_semaines(moteur)
    archiver_version_precedente(moteur)

    print("\n== 1/3 : nettoyage du bulletin hebdomadaire ==")
    nettoyage.main(chemin_xlsx, DOSSIER, date_manuelle)

    print("\n== 2/3 : tableau annuel, normales, calendrier ==")
    nettoyage_reference.main(chemin_xlsx, DOSSIER)

    print("\n== 3/3 : calcul des indicateurs ==")
    indicateurs.main(chemin_xlsx)

    print("\n== Résumé ==")
    apres = etat_semaines(db.obtenir_moteur())
    comparer(avant, apres)
    print("\nTerminé. Relance le tableau de bord (streamlit run app.py) "
         "pour voir les nouvelles données.")


if __name__ == "__main__":
    main()
