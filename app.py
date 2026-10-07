"""
Tableau de bord pluviométrique 2026 - DRDR de Diourbel.

Premier lancement (depuis ce dossier) :
    pip install -r requirements.txt
    streamlit run app.py

Ensuite, chaque semaine, tout se fait dans le navigateur, sans terminal :
page « Ajouter des données » du menu de gauche, dépôt du fichier Excel de
la DRDR, puis clic sur « Mettre à jour le tableau de bord ».
"""
import io
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import db
import maj_donnees
import nettoyage


MOIS = ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août",
        "sept.", "oct.", "nov.", "déc."]

JOURS = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi",
         "dimanche"]

DEFICIT, EXCES, GRIS, CLAIR = "#b5651d", "#1f5f8b", "#6b7280", "#9db3c2"


st.set_page_config(
    page_title="Pluviométrie 2026 - Diourbel",
    page_icon="🌧️",
    layout="wide"
)


# --------------------------------------------------------- thème clair/sombre
if "theme_sombre" not in st.session_state:
    st.session_state.theme_sombre = True  # sombre par défaut, comme avant


if st.sidebar.button(
    "🌗 Changer de thème (clair / sombre)",
    width="stretch"
):
    st.session_state.theme_sombre = not st.session_state.theme_sombre


if st.session_state.theme_sombre:
    FOND, FOND2, TEXTE, BORDURE = "#0e1117", "#161a23", "#e6e6e6", "#30363d"
    pio.templates.default = "plotly_dark"
else:
    FOND, FOND2, TEXTE, BORDURE = "#ffffff", "#f7f7f9", "#111111", "#d8d8dc"
    pio.templates.default = "plotly_white"


st.markdown(f"""
<style>
.stApp {{ background-color: {FOND}; color: {TEXTE}; }}
section[data-testid="stSidebar"] {{ background-color: {FOND2};
    border-right: 1px solid {BORDURE}; }}
[data-testid="stMetricValue"], [data-testid="stMetricLabel"],
h1, h2, h3, p, label, span {{ color: {TEXTE}; }}
[data-testid="stDataFrame"] {{ background-color: {FOND2}; }}
</style>
""", unsafe_allow_html=True)


st.sidebar.title("Pluviométrie 2026")


page = st.sidebar.radio(
    "Page",
    [
        "Ajouter des données",
        "Suivi de la saison",
        "Écarts à la normale par poste",
        "Cumul mensuel",
        "Historique 14 dernières années",
        "Bulletin de la semaine"
    ]
)


# --------------------------------------------------- page d'ajout de données
# Cette page ne dépend d'aucune donnée déjà chargée : c'est volontaire, pour
# qu'elle fonctionne même au tout premier lancement, avant qu'une base existe.
if page == "Ajouter des données":

    st.title("Ajouter le bulletin de la semaine")

    st.write(
        "Dépose ici le fichier Excel envoyé par la DRDR (le même modèle "
        "que d'habitude, pas besoin de le convertir). La mise à jour se "
        "fait automatiquement : nettoyage, vérifications et calcul des "
        "indicateurs."
    )

    depot = st.file_uploader(
        "Fichier Excel de la DRDR (.xlsx)",
        type=["xlsx"]
    )

    if depot is not None:

        chemin = Path(__file__).parent / depot.name

        chemin.write_bytes(
            depot.getbuffer()
        )

        date_detectee = nettoyage.date_fin_depuis_nom(chemin)

        date_manuelle = None

        if date_detectee:

            st.caption(
                f"Date de fin de bulletin détectée dans le nom du "
                f"fichier : {date_detectee}."
            )

        else:

            st.warning(
                "Le nom du fichier ne suit pas le modèle habituel de la "
                "DRDR (« ... du J MOIS au J MOIS AAAA ... »), donc la date "
                "de fin ne peut pas être devinée automatiquement. "
                "Indique-la ci-dessous."
            )

            date_manuelle = st.date_input(
                "Dernier jour de données réelles dans ce fichier"
            )

        if st.button(
            "Mettre à jour le tableau de bord",
            type="primary"
        ):

            with st.spinner(
                "Nettoyage, vérifications et calculs en cours..."
            ):

                resume = maj_donnees.executer(
                    chemin,
                    date_manuelle
                )

            if not resume["ok"]:

                st.error(
                    f"La mise à jour a échoué : {resume['erreur']}"
                )

            else:

                st.success(
                    "Mise à jour terminée, toutes les vérifications "
                    "de cohérence sont passées."
                )

                if resume["semaines_ajoutees"]:

                    st.write(
                        "**Semaine(s) ajoutée(s) :** "
                        + ", ".join(
                            str(d)
                            for d in resume["semaines_ajoutees"]
                        )
                    )

                else:

                    st.info(
                        "Aucune semaine nouvelle : ce fichier ne "
                        "contenait rien après la dernière date connue."
                    )

                if resume["semaines_disparues"]:

                    st.warning(
                        "⚠ Semaine(s) qui n'apparaissent plus dans "
                        "ce fichier : "
                        + ", ".join(
                            str(d)
                            for d in resume["semaines_disparues"]
                        )
                    )

                if resume["valeurs_modifiees"]:

                    st.warning(
                        f"⚠ {len(resume['valeurs_modifiees'])} "
                        "valeur(s) d'une semaine déjà connue ont été "
                        "corrigées par ce nouveau fichier :"
                    )

                    vm = pd.DataFrame(
                        resume["valeurs_modifiees"]
                    ).rename(
                        columns={
                            "poste": "Poste",
                            "semaine_debut": "Semaine du",
                            "cumul_2026_mm_avant": "Avant (mm)",
                            "cumul_2026_mm_apres": "Après (mm)"
                        }
                    )

                    st.dataframe(
                        vm.round(1),
                        hide_index=True,
                        width="stretch"
                    )

                st.cache_data.clear()

                st.info(
                    "Va sur « Suivi de la saison » dans le menu à gauche "
                    "pour voir les données à jour."
                )

    st.stop()


MOTEUR = db.obtenir_moteur()

DB_PRESENTE = db.donnees_disponibles(
    MOTEUR
)


if not DB_PRESENTE:

    st.warning(
        "Aucune donnée n'a encore été chargée. Va sur « Ajouter des "
        "données » dans le menu à gauche et dépose le premier "
        "fichier Excel de la DRDR."
    )

    st.stop()


@st.cache_data
def charger():

    moteur = db.obtenir_moteur()

    noms = [
        "indicateurs_poste_semaine",
        "indicateurs_agreges_semaine",
        "observations_hebdo",
        "mesures_journalieres",
        "annuel",
        "normales",
        "semaines",
        "rang_2026"
    ]

    t = {
        n: pd.read_sql(
            f"select * from {n}",
            moteur
        )
        for n in noms
    }

    return t


T = charger()


ip, ia, oh = (
    T["indicateurs_poste_semaine"],
    T["indicateurs_agreges_semaine"],
    T["observations_hebdo"]
)


mj, an, nm, sm, rg = (
    T["mesures_journalieres"],
    T["annuel"],
    T["normales"],
    T["semaines"].sort_values("num_semaine"),
    T["rang_2026"]
)


POIDS = (
    ip.groupby("departement")
    .poste
    .nunique()
    .to_dict()
)


ORDRE_DEPT = [
    "Bambey",
    "Diourbel",
    "Mbacké"
]


DEPT_DU_POSTE = (
    ip.drop_duplicates("poste")
    .set_index("poste")
    .departement
)


def grouper_par_departement(
    df,
    col_poste="poste",
    col_dept="departement",
    tri_secondaire=None,
    ascendant=True
):
    """Trie un tableau par département (Bambey, Diourbel, Mbacké), puis par
    la colonne donnée à l'intérieur de chaque département. Utilisé partout
    où une liste de postes doit rester regroupée par département plutôt que
    mélangée."""

    df = df.copy()

    df[col_dept] = pd.Categorical(
        df[col_dept],
        categories=ORDRE_DEPT,
        ordered=True
    )

    colonnes = [
        col_dept,
        tri_secondaire if tri_secondaire else col_poste
    ]

    return df.sort_values(
        colonnes,
        ascending=[True, ascendant]
    )


def mettre_en_forme_tableau(ws, df):
    """Mise en forme d'un tableau exporté en Excel : en-tête en gras sur fond
    coloré, colonnes ajustées à leur contenu, nombres alignés à droite,
    texte à gauche, bordures légères, première ligne figée et filtre
    automatique - pour un fichier propre, prêt à être ouvert et lu."""

    entete_gras = Font(
        bold=True,
        color="FFFFFF"
    )

    fond_entete = PatternFill(
        "solid",
        fgColor="1F5F8B"
    )

    fine = Side(
        style="thin",
        color="D9D9D9"
    )

    bordure = Border(
        left=fine,
        right=fine,
        top=fine,
        bottom=fine
    )

    colonnes_texte = {
        "Poste",
        "Département",
        "Début de semaine",
        "Fin de semaine",
        "Données incomplètes"
    }

    for c, col in enumerate(
        df.columns,
        start=1
    ):

        cell = ws.cell(
            1,
            c
        )

        cell.font = entete_gras

        cell.fill = fond_entete

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        largeur = max(
            len(str(col)),
            df[col].astype(str).map(len).max()
        ) + 2

        ws.column_dimensions[
            get_column_letter(c)
        ].width = min(
            largeur,
            32
        )

        align = (
            "left"
            if col in colonnes_texte
            else "right"
        )

        for r in range(
            2,
            len(df) + 2
        ):

            ws.cell(
                r,
                c
            ).alignment = Alignment(
                horizontal=align
            )

    for r in range(
        1,
        len(df) + 2
    ):

        for c in range(
            1,
            len(df.columns) + 1
        ):

            ws.cell(
                r,
                c
            ).border = bordure

    ws.freeze_panes = "A2"

    ws.auto_filter.ref = ws.dimensions


def d_fr(s, jour=False):

    d = pd.Timestamp(s)

    j = (
        "1er"
        if d.day == 1
        else str(d.day)
    )

    return (
        (JOURS[d.weekday()] + " " if jour else "")
        + f"{j} {MOIS[d.month - 1]} {d.year}"
    )


def periode(n):

    r = sm[
        sm.num_semaine == n
    ].iloc[0]

    return (
        f"{d_fr(r.semaine_debut, True)} → "
        f"{d_fr(r.semaine_fin, True)}"
    )


def serie(ent):

    q = (
        ia[
            ia.niveau == "Région"
        ]
        if ent == "Région"
        else ia[
            (ia.niveau == "Département")
            & (ia.nom == ent)
        ]
    )

    q = (
        q.sort_values("num_semaine")
        .reset_index(drop=True)
    )

    q["fin"] = pd.to_datetime(
        q.semaine_fin
    )

    return q


def courbe_normale(ent):

    p = nm.pivot(
        index="date",
        columns="departement",
        values="cumul_normale_mm"
    )

    s = (
        p[ent]
        if ent != "Région"
        else sum(
            p[d] * w
            for d, w in POIDS.items()
        ) / sum(POIDS.values())
    )

    s.index = pd.to_datetime(
        s.index
    )

    return s


def pct_txt(v):

    return (
        None
        if pd.isna(v)
        else f"{v:+.1f} %"
    )


# ------------------------------------------------------------- barre latérale
ent = st.sidebar.selectbox(
    "Territoire",
    [
        "Région",
        "Bambey",
        "Diourbel",
        "Mbacké"
    ]
)


dernier = int(
    sm.num_semaine.max()
)


sem = st.sidebar.selectbox(
    "Semaine (vendredi → jeudi)",
    sm.num_semaine.tolist()[::-1],
    format_func=periode
)


st.sidebar.caption(
    "Normale : moyenne 1991-2020 au jeudi de fin de semaine, reprise telle "
    "quelle de la feuille « Normales cumulées » du fichier de la DRDR "
    "(une valeur par jour et par département ; la région est une moyenne "
    "pondérée par le nombre de postes de chaque département). "
    "np = non parvenu (mesure absente). tr = trace, comptée 0 mm. "
    "⚠ = poste avec au moins un jour np depuis le début de la saison."
)


fin_sem = pd.Timestamp(
    sm[
        sm.num_semaine == sem
    ].semaine_fin.iloc[0]
)


# ---------------------------------------------------------------- page 1
if page == "Suivi de la saison":

    st.title(
        f"{ent} : la saison au {d_fr(fin_sem, True)}"
    )

    q = serie(ent)

    c = q[
        q.num_semaine == sem
    ].iloc[0]

    up = q[
        q.num_semaine <= sem
    ]

    a, b, c3, d = st.columns(4)

    a.metric(
        "Cumul 2026",
        f"{c.cumul_2026_mm:.1f} mm"
    )

    b.metric(
        "Normale à cette date",
        f"{c.normale_mm:.1f} mm",
        help=(
            "Moyenne 1991-2020 au même jour (jeudi de fin de "
            "semaine), depuis la feuille « Normales cumulées » du "
            "fichier de la DRDR. "
            f"{ent if ent != 'Région' else 'Région : moyenne des 3 départements, pondérée par leur nombre de postes'}."
        )
    )

    c3.metric(
        "Écart à la normale",
        f"{c.ecart_normale_mm:+.1f} mm",
        pct_txt(c.ecart_normale_pct),
        help=(
            "Cumul 2026 moins la normale à cette date (colonne "
            "précédente). Négatif = moins de pluie que d'habitude "
            "à ce point de la saison."
        )
    )

    d.metric(
        "Jours de pluie cumulés",
        f"{c.cumul_2026_nj:.1f}",
        f"{c.intensite_mm_par_jour:.1f} mm par jour de pluie",
        delta_color="off"
    )

    if c.nb_postes_incomplets > 0:

        st.warning(
            f"{int(c.nb_postes_incomplets)} poste(s) sur "
            f"{int(c.nb_postes)} ont au moins un jour « np » depuis le "
            "début de la saison : leur cumul peut être sous-estimé."
        )

    if sem == dernier:

        r = rg[
            (rg.nom == ent)
            & (rg.indicateur == "cumul_mm")
        ].iloc[0]

        st.info(
            f"Classement : {int(r.rang_du_plus_bas)}e année la plus sèche "
            f"sur {int(r.n_annees)} (2012-2025 plus 2026). Attention, "
            "2026 est un cumul au 1er octobre, les autres années sont des "
            "totaux de saison."
        )

    cn = courbe_normale(ent)

    f1 = go.Figure()

    f1.add_scatter(
        x=cn.index,
        y=cn.values,
        name="Normale 1991-2020",
        line=dict(
            color=GRIS,
            dash="dash"
        )
    )

    f1.add_scatter(
        x=up.fin,
        y=up.cumul_2025_mm,
        name="2025",
        line=dict(
            color=CLAIR
        )
    )

    f1.add_scatter(
        x=up.fin,
        y=up.cumul_2026_mm,
        name="2026",
        mode="lines+markers",
        line=dict(
            color=DEFICIT,
            width=3
        )
    )

    f1.update_layout(
        title="Cumul depuis le début de la saison (mm)",
        hovermode="x unified",
        legend=dict(
            orientation="h"
        ),
        margin=dict(
            t=50,
            b=10
        )
    )

    st.plotly_chart(
        f1,
        width="stretch"
    )

    pluie = np.diff(
        np.r_[
            0,
            up.cumul_2026_mm.values
        ]
    )

    norm_sem = (
        cn.reindex(up.fin).values
        - cn.reindex(
            up.fin - pd.Timedelta(days=7)
        ).values
    )

    lab = up.fin.dt.strftime(
        "%d/%m"
    )

    f2 = go.Figure()

    f2.add_bar(
        x=lab,
        y=pluie,
        name="2026",
        marker_color=DEFICIT
    )

    f2.add_bar(
        x=lab,
        y=norm_sem,
        name="Normale",
        marker_color=CLAIR
    )

    f2.update_layout(
        title="Pluie de chaque semaine (mm), date = jeudi de fin",
        barmode="group",
        legend=dict(
            orientation="h"
        ),
        margin=dict(
            t=50,
            b=10
        )
    )

    st.plotly_chart(
        f2,
        width="stretch"
    )


# ---------------------------------------------------------------- page 2
elif page == "Écarts à la normale par poste":

    st.title(
        f"Écarts à la normale par poste au {d_fr(fin_sem, True)}"
    )

    d = ip[
        ip.num_semaine == sem
    ]

    if ent != "Région":

        d = d[
            d.departement == ent
        ]

    d = d.assign(
        poste_aff=d.poste
        + np.where(
            d.donnees_incompletes,
            " ⚠",
            ""
        ),
        sens=np.where(
            d.ecart_normale_mm < 0,
            "Sous la normale",
            "Au-dessus de la normale"
        )
    )

    # Postes regroupés par département (Bambey, puis Diourbel, puis Mbacké),
    # et à l'intérieur de chaque département, du plus sec au plus humide.
    d = grouper_par_departement(
        d,
        tri_secondaire="ecart_normale_mm"
    )

    nb_dept = d.departement.nunique()

    f = px.bar(
        d,
        x="ecart_normale_mm",
        y="poste_aff",
        orientation="h",
        color="sens",
        text=d.ecart_normale_mm.round(1),
        facet_row="departement" if nb_dept > 1 else None,
        category_orders={
            "departement": ORDRE_DEPT,
            "poste_aff": d.poste_aff.tolist()
        },
        color_discrete_map={
            "Sous la normale": DEFICIT,
            "Au-dessus de la normale": EXCES
        },
        labels={
            "ecart_normale_mm": "Écart à la normale (mm)",
            "poste_aff": "",
            "sens": ""
        }
    )

    if nb_dept > 1:

        f.update_yaxes(
            matches=None
        )

        for a in f.layout.annotations:

            a.text = a.text.split("=")[-1]

    f.update_layout(
        height=40 * nb_dept + 30 * len(d),
        legend=dict(
            orientation="h"
        ),
        margin=dict(
            t=10,
            b=10
        )
    )

    st.plotly_chart(
        f,
        width="stretch"
    )

    t = d[
        [
            "poste_aff",
            "departement",
            "cumul_2026_mm",
            "normale_mm",
            "ecart_normale_mm",
            "ecart_normale_pct",
            "cumul_2025_mm",
            "ecart_2025_mm",
            "cumul_2026_nj",
            "intensite_mm_par_jour"
        ]
    ]

    t.columns = [
        "Poste",
        "Département",
        "Cumul 2026 (mm)",
        "Normale (mm)",
        "Écart normale (mm)",
        "Écart normale (%)",
        "Cumul 2025 (mm)",
        "Écart 2025 (mm)",
        "Jours de pluie",
        "mm par jour de pluie"
    ]

    st.dataframe(
        t.round(1),
        hide_index=True,
        width="stretch"
    )

    st.caption(
        "Colonne « Normale » : moyenne 1991-2020 au jeudi de fin de "
        "semaine, par département (feuille « Normales cumulées » du "
        "fichier de la DRDR). Tous les postes d'un même département "
        "partagent la même normale, le fichier ne la donne pas poste "
        "par poste."
    )

    if sem == dernier:

        r = rg[
            (rg.niveau == "Poste")
            & (rg.indicateur == "cumul_mm")
        ]

        if ent != "Région":

            r = r[
                r.nom.isin(d.poste)
            ]

        with st.expander(
            "Position de chaque poste par rapport à 2012-2025"
        ):

            r = r.sort_values("z")[
                [
                    "nom",
                    "valeur_2026",
                    "moyenne_2012_2025",
                    "z",
                    "rang_du_plus_bas",
                    "n_annees"
                ]
            ]

            r.columns = [
                "Poste",
                "2026 (mm)",
                "Moyenne 2012-2025 (mm)",
                "Écart réduit z",
                "Rang (1 = plus sec)",
                "Années"
            ]

            st.dataframe(
                r.round(1),
                hide_index=True,
                width="stretch"
            )


# ---------------------------------------------------------------- page 3
elif page == "Cumul mensuel":

    st.title(
        f"Cumul mensuel : {ent}"
    )

    st.caption(
        "Reprend la feuille « Cumul mensuel » du fichier de la DRDR : "
        "un groupe de barres par poste, une couleur par mois."
    )

    dep_poste = (
        ip.drop_duplicates("poste")
        .set_index("poste")
        .departement
    )

    j = mj.copy()

    j["departement"] = j.poste.map(
        dep_poste
    )

    if ent != "Région":

        j = j[
            j.departement == ent
        ]

    j["mois_num"] = pd.to_datetime(
        j.date
    ).dt.month

    ordre_mois = sorted(
        j.mois_num.unique()
    )

    lib_mois = [
        MOIS[m - 1]
        for m in ordre_mois
    ]

    mm = (
        j.groupby(
            ["poste", "mois_num"]
        ).mm.sum()
        .unstack(fill_value=0)
        .reindex(
            columns=ordre_mois,
            fill_value=0
        )
    )

    long = (
        mm.reset_index()
        .melt(
            id_vars="poste",
            var_name="mois_num",
            value_name="mm"
        )
    )

    long["mois"] = long.mois_num.map(
        lambda m: MOIS[m - 1]
    )

    long["departement"] = long.poste.map(
        DEPT_DU_POSTE
    )

    # Comme dans le fichier de la DRDR : chaque poste reste une barre à part
    # (jamais sommé ni empilé avec un autre poste), avec une couleur par mois
    # - au lieu des 19 couleurs par poste de l'ancien graphique, on n'a plus
    # que 5 couleurs (les mois), beaucoup plus lisible.
    ordre_postes = grouper_par_departement(
        long.drop_duplicates("poste")
    ).poste.tolist()

    nb_dept = long.departement.nunique()

    f = px.bar(
        long,
        x="poste",
        y="mm",
        color="mois",
        barmode="group",
        facet_row="departement" if nb_dept > 1 else None,
        category_orders={
            "mois": lib_mois,
            "poste": ordre_postes,
            "departement": ORDRE_DEPT
        },
        labels={
            "mm": "Pluie du mois (mm)",
            "poste": "",
            "mois": ""
        }
    )

    if nb_dept > 1:

        f.update_xaxes(
            matches=None
        )

        for a in f.layout.annotations:

            a.text = a.text.split("=")[-1]

    f.update_layout(
        legend=dict(
            orientation="h",
            title=None
        ),
        margin=dict(
            t=10,
            b=10
        ),
        height=(
            320
            if nb_dept == 1
            else 230 * nb_dept
        )
    )

    st.plotly_chart(
        f,
        width="stretch"
    )

    st.caption(
        f"Le dernier mois affiché ({lib_mois[-1]}) peut être "
        "incomplet : il ne compte que les semaines déjà parues dans "
        "le bulletin. Le tableau ci-dessous garde le détail par poste."
    )

    tbl = mm.copy()

    tbl.columns = [
        MOIS[m - 1]
        for m in tbl.columns
    ]

    st.dataframe(
        tbl.round(1)
        .reset_index()
        .rename(
            columns={"poste": "Poste"}
        ),
        hide_index=True,
        width="stretch"
    )

    nb_np = (
        j[
            j.statut == "non_parvenu"
        ]
        .groupby("poste")
        .size()
    )

    if not nb_np.empty:

        st.warning(
            "Cumul(s) mensuel(s) possiblement sous-estimé(s) (jour(s) "
            "« np » dans la saison) pour : "
            + ", ".join(nb_np.index)
        )


# ---------------------------------------------------------------- page 4
elif page == "Historique 14 dernières années":

    st.title(
        f"Historique 14 dernières années (2012-2026) : {ent}"
    )

    st.caption(
        "Reprend la feuille « Pluvio 14 dernières années » du fichier "
        "de la DRDR, complétée par la saison 2026 en cours."
    )

    choix = {
        "Cumul (mm)": "mm",
        "Jours de pluie": "jours_pluie",
        "Intensité (mm par jour de pluie)": "intensite"
    }

    lib = st.radio(
        "Indicateur",
        list(choix),
        horizontal=True
    )

    col = choix[lib]

    a = an.assign(
        intensite=an.mm / an.jours_pluie
    )

    f26 = (
        ip[
            ip.num_semaine == dernier
        ]
        .assign(
            mm=lambda x: x.cumul_2026_mm,
            jours_pluie=lambda x: x.cumul_2026_nj
        )
    )

    f26["intensite"] = (
        f26.mm
        / f26.jours_pluie.replace(
            0,
            np.nan
        )
    )

    if ent != "Région":

        a = a[
            a.departement == ent
        ]

        f26 = f26[
            f26.departement == ent
        ]

    fin_saison = d_fr(
        ip[
            ip.num_semaine == dernier
        ].semaine_fin.iloc[0]
    )

    wide = a.pivot(
        index="poste",
        columns="annee",
        values=col
    )

    wide["2026*"] = (
        f26.set_index("poste")[col]
    )

    # Postes regroupés par département (plutôt que triés uniquement par
    # valeur), du plus sec au plus humide à l'intérieur de chaque groupe.
    wide["departement"] = (
        wide.index.map(
            DEPT_DU_POSTE
        )
    )

    wide = grouper_par_departement(
        wide.reset_index(),
        col_poste="poste",
        tri_secondaire="2026*"
    ).set_index("poste")

    tailles_dept = (
        wide.groupby(
            "departement",
            sort=False,
            observed=True
        ).size()
    )

    wide = wide.drop(
        columns="departement"
    )

    hm = px.imshow(
        wide.values,
        x=[
            str(c)
            for c in wide.columns
        ],
        y=wide.index,
        aspect="auto",
        color_continuous_scale="YlGnBu",
        text_auto=(
            ".0f"
            if col != "intensite"
            else ".1f"
        ),
        labels=dict(
            color=lib
        )
    )

    if ent == "Région":

        # Lignes pointillées séparant les 3 départements, avec leur nom
        pos = 0

        for dept, n in tailles_dept.items():

            pos += n

            if pos < len(wide):

                hm.add_hline(
                    y=pos - 0.5,
                    line_dash="dot",
                    line_color="white",
                    line_width=1
                )

            hm.add_annotation(
                x=1.02,
                xref="paper",
                y=pos - n / 2,
                text=dept,
                showarrow=False,
                xanchor="left",
                font=dict(
                    size=11
                )
            )

    hm.update_layout(
        height=140 + 26 * len(wide),
        margin=dict(
            t=10,
            b=10,
            r=90
        )
    )

    st.plotly_chart(
        hm,
        width="stretch"
    )

    y = (
        a.dropna(
            subset=["mm"]
        )
        .groupby("annee")
        .agg(
            mm=("mm", "mean"),
            jours_pluie=("jours_pluie", "mean")
        )
    )

    y["intensite"] = (
        y.mm
        / y.jours_pluie
    )

    v = {
        "mm": f26.mm.mean(),
        "jours_pluie": f26.jours_pluie.mean()
    }

    v["intensite"] = (
        v["mm"]
        / v["jours_pluie"]
    )

    moy = y[col].mean()

    f = go.Figure(
        go.Bar(
            x=[
                str(i)
                for i in y.index
            ] + ["2026*"],
            y=list(y[col]) + [v[col]],
            marker_color=[
                CLAIR
            ] * len(y) + [
                DEFICIT
            ]
        )
    )

    f.add_hline(
        y=moy,
        line_dash="dash",
        line_color=GRIS,
        annotation_text=(
            f"Moyenne 2012-2025 : {moy:.1f}"
        )
    )

    f.update_layout(
        title=f"{lib}, moyenne des postes : {ent}",
        margin=dict(
            t=50,
            b=10
        )
    )

    st.plotly_chart(
        f,
        width="stretch"
    )

    st.caption(
        f"* 2026 : cumul au {fin_saison}. Les autres années sont des "
        "totaux de saison. Cases vides : pas de donnée pour ce poste."
    )


# ---------------------------------------------------------------- page 5
else:

    s = sm[
        sm.num_semaine == sem
    ].iloc[0]

    st.title(
        "Bulletin de la semaine"
    )

    st.subheader(
        periode(sem)
    )

    # -----------------------------------------------------------
    # Préparation des dates avant le merge.
    #
    # Après lecture depuis PostgreSQL, semaine_debut peut être de type
    # datetime dans ip et de type objet/texte dans oh.
    # On convertit explicitement les deux colonnes au même type.
    # -----------------------------------------------------------

    b = ip[
        ip.num_semaine == sem
    ].copy()

    b["semaine_debut"] = pd.to_datetime(
        b["semaine_debut"]
    )

    oh_bulletin = oh[
        [
            "poste",
            "semaine_debut",
            "mm_semaine"
        ]
    ].copy()

    oh_bulletin["semaine_debut"] = pd.to_datetime(
        oh_bulletin["semaine_debut"]
    )

    b = b.merge(
        oh_bulletin,
        on=[
            "poste",
            "semaine_debut"
        ]
    )

    b["poste_aff"] = (
        b.poste
        + np.where(
            b.donnees_incompletes,
            " ⚠",
            ""
        )
    )

    cols = {
        "poste_aff": "Poste",
        "mm_semaine": "Pluie de la semaine (mm)",
        "cumul_2026_mm": "Cumul 2026 (mm)",
        "cumul_2026_nj": "Jours de pluie cumulés",
        "normale_mm": "Normale (mm)",
        "ecart_normale_mm": "Écart (mm)",
        "cumul_2025_mm": "Cumul 2025 (mm)"
    }

    for dep in [
        "Bambey",
        "Diourbel",
        "Mbacké"
    ]:

        st.markdown(
            f"**Département de {dep}**"
        )

        st.dataframe(
            b[
                b.departement == dep
            ][
                list(cols)
            ]
            .rename(
                columns=cols
            )
            .round(1),
            hide_index=True,
            width="stretch"
        )

    st.caption(
        "Colonne « Normale » : moyenne 1991-2020 au jeudi de fin de "
        "cette semaine, depuis la feuille « Normales cumulées » du "
        "fichier de la DRDR (une valeur par département)."
    )

    # -----------------------------------------------------------
    # Correction du deuxième problème de dates.
    #
    # mj.date est lu comme datetime.date tandis que
    # s.semaine_debut / s.semaine_fin peuvent être des chaînes.
    # On convertit explicitement les trois dates en datetime pandas
    # avant d'effectuer les comparaisons >= et <=.
    # -----------------------------------------------------------

    mj_bulletin = mj.copy()

    mj_bulletin["date"] = pd.to_datetime(
        mj_bulletin["date"]
    )

    date_debut = pd.to_datetime(
        s.semaine_debut
    )

    date_fin = pd.to_datetime(
        s.semaine_fin
    )

    j = mj_bulletin[
        (mj_bulletin.date >= date_debut)
        & (mj_bulletin.date <= date_fin)
    ].copy()

    if not j.empty:

        j["txt"] = np.select(
            [
                j.statut == "non_parvenu",
                j.statut == "trace"
            ],
            [
                "np",
                "tr"
            ],
            default=j.mm.map(
                lambda x: f"{x:g}"
            )
        )

        j["jour"] = pd.to_datetime(
            j.date
        ).dt.strftime(
            "%d/%m"
        )

        ordre = [
            x.strftime("%d/%m")
            for x in pd.date_range(
                date_debut,
                date_fin
            )
        ]

        piv = (
            j.pivot_table(
                index="poste",
                columns="jour",
                values="txt",
                aggfunc="first"
            )
            .reindex(
                columns=ordre
            )
            .fillna("")
            .reset_index()
            .rename(
                columns={"poste": "Poste"}
            )
        )

        with st.expander(
            "Détail journalier (jours inscrits au bulletin)"
        ):

            st.dataframe(
                piv,
                hide_index=True,
                width="stretch"
            )

            st.caption(
                "Case vide : rien d'inscrit ce jour-là."
            )

    # ------------------------------------------------- export propre (xlsx/csv)
    # Mêmes colonnes que le tableau affiché à l'écran, mêmes intitulés en
    # français, postes regroupés par département, nombres arrondis et jours
    # au format JJ/MM/AAAA : un fichier lisible tel quel, pas un export brut
    # de la base de données.

    export = grouper_par_departement(
        b[
            [
                "poste_aff",
                "departement",
                "semaine_debut",
                "semaine_fin",
                "mm_semaine",
                "cumul_2026_mm",
                "cumul_2026_nj",
                "normale_mm",
                "ecart_normale_mm",
                "ecart_normale_pct",
                "cumul_2025_mm",
                "donnees_incompletes"
            ]
        ].copy(),
        col_poste="poste_aff"
    )

    export["semaine_debut"] = (
        pd.to_datetime(
            export.semaine_debut
        )
        .dt.strftime(
            "%d/%m/%Y"
        )
    )

    export["semaine_fin"] = (
        pd.to_datetime(
            export.semaine_fin
        )
        .dt.strftime(
            "%d/%m/%Y"
        )
    )

    export["donnees_incompletes"] = np.where(
        export.donnees_incompletes,
        "Oui",
        "Non"
    )

    num_cols = [
        "mm_semaine",
        "cumul_2026_mm",
        "cumul_2026_nj",
        "normale_mm",
        "ecart_normale_mm",
        "ecart_normale_pct",
        "cumul_2025_mm"
    ]

    export[num_cols] = (
        export[num_cols]
        .round(1)
    )

    export.columns = [
        "Poste",
        "Département",
        "Début de semaine",
        "Fin de semaine",
        "Pluie de la semaine (mm)",
        "Cumul 2026 (mm)",
        "Jours de pluie cumulés",
        "Normale (mm)",
        "Écart à la normale (mm)",
        "Écart à la normale (%)",
        "Cumul 2025 (mm)",
        "Données incomplètes"
    ]

    buf = io.BytesIO()

    with pd.ExcelWriter(
        buf,
        engine="openpyxl"
    ) as w:

        export.to_excel(
            w,
            index=False,
            sheet_name="Bulletin"
        )

        mettre_en_forme_tableau(
            w.sheets["Bulletin"],
            export
        )

    x, y = st.columns(2)

    x.download_button(
        "Télécharger en Excel",
        buf.getvalue(),
        f"bulletin_{s.semaine_fin}.xlsx",
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    )

    y.download_button(
        "Télécharger en CSV",
        export.to_csv(
            index=False,
            sep=";"
        ),
        f"bulletin_{s.semaine_fin}.csv",
        "text/csv"
    )

    st.caption(
        "Le CSV utilise le point-virgule (;) comme séparateur, pour "
        "s'ouvrir directement en colonnes dans Excel en français."
    )