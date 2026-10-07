"""
Connexion à la base de données.

Par défaut, les données vivent dans un fichier SQLite local
(donnees_propres/pluvio.db), exactement comme avant - aucune configuration
n'est nécessaire pour travailler en local, sur ton PC.

Si une adresse de connexion est fournie - variable d'environnement
DATABASE_URL, ou secret Streamlit du même nom -, l'application se
connecte à cette base à la place : typiquement une base PostgreSQL
gratuite (Supabase). C'est ce qui permet aux données déposées via la
page « Ajouter des données » d'un site hébergé gratuitement (Streamlit
Community Cloud) de survivre aux redémarrages du serveur, qui sinon
remet son disque local à zéro à chaque réveil.

Rien d'autre dans le projet ne change de comportement : tous les
scripts continuent d'appeler obtenir_moteur() exactement comme ils
appelaient sqlite3.connect() avant.
"""
import os
from pathlib import Path

from dotenv import load_dotenv
from sqlalchemy import create_engine, inspect

load_dotenv(override=True)

DOSSIER = Path(__file__).parent / "donnees_propres"


def _adresse_configuree():
    """Lit DATABASE_URL dans l'environnement, puis dans les secrets
    Streamlit si l'application tourne sous Streamlit. Renvoie None si
    aucune des deux n'est définie (cas normal en local)."""
    url = os.environ.get("DATABASE_URL")
    if url:
        return url
    try:
        import streamlit as st
        return st.secrets["DATABASE_URL"]
    except Exception:
        return None
def obtenir_moteur(dossier=None):
    """Renvoie un moteur SQLAlchemy : la base PostgreSQL configurée si
    DATABASE_URL existe, sinon le fichier SQLite local (créé si besoin)."""
    url = _adresse_configuree()
    if url:
        return create_engine(
            url,
            pool_pre_ping=True,
            connect_args={"prepare_threshold": None},
        )
    dossier = Path(dossier) if dossier else DOSSIER
    dossier.mkdir(parents=True, exist_ok=True)
    return create_engine(f"sqlite:///{dossier / 'pluvio.db'}")





def est_local(moteur):
    """True si les données sont dans le fichier SQLite local (pas dans
    une base en ligne comme Supabase)."""
    return moteur.dialect.name == "sqlite"


def donnees_disponibles(moteur):
    """True si au moins une table de données existe déjà (donc qu'un
    premier bulletin a déjà été chargé)."""
    try:
        return "semaines" in inspect(moteur).get_table_names()
    except Exception:
        return False
