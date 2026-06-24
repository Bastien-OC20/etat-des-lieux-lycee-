#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Application Streamlit - Etat des lieux Lycee
Lancer avec : streamlit run app.py
"""

from datetime import date

import streamlit as st

from constants import (
    ANNEE_SCOLAIRE,
    LOGO_PRINCIPALE,
    LOGO_SECONDAIRE,
    ORDRE_PRIO,
)
from data import charger_xlsx
from exports import generer_docx, generer_xlsx
from views import (
    render_tab_dashboard,
    render_tab_edition,
    render_tab_liste,
    render_tab_salle,
)

# ── Config page ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Etat des lieux - Lycee",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS personnalise ───────────────────────────────────────────────────────────
st.markdown(
    """
<style>
    .main-title {
        font-size: 2rem; font-weight: 800; color: #1A3C6E;
        border-bottom: 3px solid #1A3C6E; padding-bottom: 0.3rem;
        margin-bottom: 1rem;
    }
    .kpi-card {
        background: white; border-radius: 12px; padding: 1rem 1.2rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-left: 5px solid; text-align: center;
    }
    .kpi-num { font-size: 2.2rem; font-weight: 800; line-height: 1; }
    .kpi-label { font-size: 0.875rem; color: #555; margin-top: 4px; }
    .salle-header {
        background: #1A3C6E; color: white; padding: 0.4rem 0.8rem;
        border-radius: 6px; font-weight: 700; margin: 0.8rem 0 0.3rem;
    }
    .badge {
        display: inline-block; padding: 2px 10px; border-radius: 20px;
        font-size: 0.75rem; font-weight: 700;
    }
    div[data-testid="stSidebarContent"] { background: #FFFF; }
    div[data-testid="stSidebarContent"] h1,
    div[data-testid="stSidebarContent"] h2,
    div[data-testid="stSidebarContent"] h3,
    div[data-testid="stSidebarContent"] p{
        color: #000000 !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    col_logo1, col_logo2 = st.columns(2)
    with col_logo1:
        st.image(LOGO_PRINCIPALE, use_container_width=True)
    with col_logo2:
        st.image(LOGO_SECONDAIRE, use_container_width=True)
    st.markdown("## 🏫 Etat des lieux Lycee")
    st.divider()

    fichier = st.file_uploader(
        "📂 Charger le fichier Excel",
        type=["xlsx"],
        help=f"Fichier Etat_des_lieux_{ANNEE_SCOLAIRE.replace('/', '_')}.xlsx",
    )

    if fichier:
        st.success(f"✅ {fichier.name}")
        st.divider()

        base_absente = "df_base" not in st.session_state
        fichier_change = st.session_state.get("fichier_nom") != fichier.name
        if base_absente or fichier_change:
            try:
                st.session_state.df_base = charger_xlsx(fichier.read())
                st.session_state.fichier_nom = fichier.name
                st.session_state.traites = {
                    i: False for i in st.session_state.df_base.index
                }
            except ValueError as e:
                st.error(f"❌ Erreur lors du chargement : {e}")
                st.stop()

        df_base = st.session_state.df_base.copy()
        df_base["Traite"] = df_base.index.map(st.session_state.traites)

        st.markdown("### 🔍 Filtres")

        salles_dispo = sorted(df_base["Salle"].unique())
        salles_sel = st.multiselect(
            "Salle(s)", salles_dispo,
            placeholder="Toutes les salles",
        )

        prios_sel = st.multiselect(
            "Priorite(s)", ORDRE_PRIO,
            default=ORDRE_PRIO,
            help="Filtrer par niveau de priorite",
        )

        elements_dispo = sorted(df_base["Element"].unique())
        elements_sel = st.multiselect(
            "Element(s)", elements_dispo,
            placeholder="Tous les elements",
        )

        afficher_traites = st.toggle("Afficher les traitees", value=True)

        st.divider()
        st.markdown("### 📥 Exports")
        tri_par = st.radio(
            "Classer par",
            options=["Salle", "Element", "Priorite"],
            format_func=lambda x: {
                "Salle": "🏫 Classe / Salle",
                "Element": "🔧 Type d'intervention",
                "Priorite": "🚨 Urgence",
            }[x],
            horizontal=False,
        )
    else:
        salles_sel = []
        prios_sel = ORDRE_PRIO
        elements_sel = []
        afficher_traites = True
        tri_par = "Salle"

# ── Page d'accueil (sans fichier) ─────────────────────────────────────────────
if not fichier:
    st.markdown(
        '<div class="main-title">🏫 Etat des lieux - Lycee</div>',
        unsafe_allow_html=True,
    )
    st.info(
        "👈 **Commencez par charger votre fichier Excel** dans la barre laterale.",
        icon="📂",
    )
    st.markdown("""
    #### Fonctionnalites disponibles
    - 📊 **Tableau de bord** avec statistiques globales
    - 🔍 **Filtres** par salle, priorite, element
    - ✅ **Marquer les interventions** comme traitees
    - 📄 **Export Word** (.docx) du rapport complet
    - 📊 **Export Excel** (.xlsx) du tableau filtre
    """)
    st.stop()

# ── Application des filtres ────────────────────────────────────────────────────
df = st.session_state.df_base.copy()
df["Traite"] = df.index.map(st.session_state.traites)
if salles_sel:
    df = df[df["Salle"].isin(salles_sel)]
if prios_sel:
    df = df[df["Priorite"].isin(prios_sel)]
if elements_sel:
    df = df[df["Element"].isin(elements_sel)]
if not afficher_traites:
    df = df[~df["Traite"]]

# ── Onglets ────────────────────────────────────────────────────────────────────
tab_dash, tab_liste, tab_salle, tab_edition = st.tabs([
    "📊 Tableau de bord",
    "📋 Liste des interventions",
    "🏠 Vue par salle",
    "✅ Gestion des traitements",
])

with tab_dash:
    df_tout = st.session_state.df_base.copy()
    df_tout["Traite"] = df_tout.index.map(st.session_state.traites)
    render_tab_dashboard(df_tout)

with tab_liste:
    render_tab_liste(df)

with tab_salle:
    render_tab_salle(df, salles_sel)

with tab_edition:
    render_tab_edition()

# ── Boutons d'export dans la sidebar ──────────────────────────────────────────
with st.sidebar:
    df_export = st.session_state.df_base.copy()
    df_export["Traite"] = df_export.index.map(st.session_state.traites)

    try:
        docx_bytes = generer_docx(df_export, tri_par=tri_par)
        st.download_button(
            label="📄 Telecharger Word (.docx)",
            data=docx_bytes,
            file_name=(
                f"interventions_lycee_{date.today().strftime('%Y%m%d')}.docx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.document"
            ),
            use_container_width=True,
        )
    except ImportError:
        st.warning("python-docx requis : pip install python-docx")

    xlsx_bytes = generer_xlsx(df_export, tri_par=tri_par)
    st.download_button(
        label="📊 Telecharger Excel (.xlsx)",
        data=xlsx_bytes,
        file_name=(
            f"interventions_lycee_{date.today().strftime('%Y%m%d')}.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument"
            ".spreadsheetml.sheet"
        ),
        use_container_width=True,
    )

    st.divider()
