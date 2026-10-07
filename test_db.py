import os

from dotenv import load_dotenv
from sqlalchemy import create_engine, text


# Charger les variables du fichier .env
load_dotenv(override=True)

# Récupérer DATABASE_URL
database_url = os.getenv("DATABASE_URL")

# Vérifier que DATABASE_URL existe
if not database_url:
    print("DATABASE_URL introuvable dans le fichier .env")
    raise SystemExit


# Afficher uniquement le serveur utilisé
# Le mot de passe n'est pas affiché
try:
    host = database_url.split("@")[1].split(":")[0]
    print("Serveur utilisé :", host)
except Exception:
    print("DATABASE_URL semble mal configurée")
    raise SystemExit


# Créer la connexion SQLAlchemy
engine = create_engine(database_url)


# Tester la connexion PostgreSQL
try:
    with engine.connect() as connection:

        # Requête simple pour vérifier PostgreSQL
        result = connection.execute(text("SELECT version();"))

        print("Connexion à Supabase réussie !")
        print("Version PostgreSQL :")
        print(result.fetchone()[0])

except Exception as e:
    print("Erreur de connexion :")
    print(e)