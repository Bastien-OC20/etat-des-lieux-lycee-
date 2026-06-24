#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Rendu des quatre onglets de l'application."""

import html as _html

import pandas as pd
import plotly.express as px
import streamlit as st

from constants import (
    COULEURS_PRIO,
    ORDRE_PRIO,
    ORDRE_PRIO_MAP,
)


def _card_salle_html(row: pd.Series) -> str:
    """Construit la carte HTML d'une intervention pour la vue par salle."""
    prio = row["Priorite"]
    color = COULEURS_PRIO.get(prio, "#888")
    element = _html.escape(str(row["Element"]))
    valeur = _html.escape(str(row["Valeur"]))
    prio_safe = _html.escape(str(prio))
    trait_badge = (
        '<span style="background:#1E8449;color:#FFF;'
        'padding:1px 8px;border-radius:3px;'
        'font-size:0.75rem;font-weight:700;'
        'margin-left:6px"> ✅ Traite</span>'
        if row["Traite"] else ""
    )
    return (
        f'<div style="background:#F9FAFB;border-left:4px solid {color};'
        f'padding:6px 12px;border-radius:4px;margin:4px 0;font-size:0.9rem">'
        f'<span style="background:{color};color:#FFF;padding:1px 8px;'
        f'border-radius:3px;font-size:0.75rem;font-weight:700">'
        f'{row["Emoji"]} {prio_safe}</span>'
        f' <b style="color:#1A3C6E">{element}</b>'
        f' <span style="color:#555">&mdash; {valeur}</span>'
        f'{trait_badge}</div>'
    )


def _card_edition_html(row: pd.Series) -> str:
    """Construit la carte HTML d'une intervention pour la vue edition."""
    prio = row["Priorite"]
    color = COULEURS_PRIO.get(prio, "#888")
    salle = _html.escape(str(row["Salle"]))
    element = _html.escape(str(row["Element"]))
    valeur = _html.escape(str(row["Valeur"]))
    prio_safe = _html.escape(str(prio))
    return (
        f'<div style="background:#FFFFFF;border-left:4px solid {color};'
        f'padding:5px 10px;border-radius:4px;margin:2px 0;font-size:0.88rem">'
        f'<span style="background:{color};color:#FFF;padding:1px 8px;'
        f'border-radius:3px;font-size:0.75rem;font-weight:700">'
        f'{row["Emoji"]} {prio_safe}</span>'
        f' <b style="color:#000000">{salle}</b>'
        f' &gt; <b style="color:#000000">{element}</b>'
        f' <span style="color:#000000">&mdash; {valeur}</span>'
        f'</div>'
    )


def render_tab_dashboard(df_tout: pd.DataFrame) -> None:
    """Affiche le tableau de bord avec KPIs et graphiques."""
    st.markdown(
        '<div class="main-title">📊 Tableau de bord</div>',
        unsafe_allow_html=True,
    )

    total = len(df_tout)
    critiques = len(df_tout[df_tout["Priorite"] == "CRITIQUE"])
    remplace = len(df_tout[df_tout["Priorite"] == "REMPLACEMENT"])
    travaux = len(df_tout[df_tout["Priorite"] == "TRAVAUX"])
    traites = len(df_tout[df_tout["Traite"]])
    nb_salles = df_tout["Salle"].nunique()

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    kpis = [
        (c1, total,     "Total",            "#1A3C6E"),
        (c2, critiques, "Critiques 🔴",     "#CC0000"),
        (c3, remplace,  "Remplacements 🟠", "#A04000"),
        (c4, travaux,   "Travaux 🟡",       "#7D6608"),
        (c5, traites,   "Traitees ✅",      "#1E8449"),
        (c6, nb_salles, "Salles",           "#1A3C6E"),
    ]
    for col, val, label, color in kpis:
        with col:
            st.markdown(
                f'<div class="kpi-card" style="border-color:{color}">'
                f'<div class="kpi-num" style="color:{color}">{val}</div>'
                f'<div class="kpi-label">{label}</div></div>',
                unsafe_allow_html=True,
            )

    st.markdown("")
    col_g1, col_g2 = st.columns(2)

    with col_g1:
        st.markdown("#### Repartition par priorite")
        prio_counts = (
            df_tout["Priorite"].value_counts()
            .reindex(ORDRE_PRIO, fill_value=0)
        )
        fig = px.pie(
            values=prio_counts.values,
            names=prio_counts.index,
            color=prio_counts.index,
            color_discrete_map={
                "CRITIQUE":     "#FF4444",
                "REMPLACEMENT": "#E67E22",
                "TRAVAUX":      "#F0A500",
                "A VERIFIER":   "#888888",
            },
            hole=0.4,
        )
        fig.update_layout(margin=dict(t=10, b=10), height=300)
        st.plotly_chart(fig, use_container_width=True)

    with col_g2:
        st.markdown("#### Top 10 salles avec le plus d'interventions")
        top_salles = (
            df_tout[~df_tout["Traite"]]
            .groupby("Salle")
            .size()
            .sort_values(ascending=False)
            .head(10)
            .reset_index(name="nb")
        )
        fig2 = px.bar(
            top_salles, x="nb", y="Salle", orientation="h",
            color="nb",
            color_continuous_scale=["#FFFDE7", "#E67E22", "#CC0000"],
            labels={"nb": "Nb interventions", "Salle": ""},
        )
        fig2.update_layout(
            margin=dict(t=10, b=10), height=300,
            coloraxis_showscale=False,
            yaxis=dict(autorange="reversed"),
        )
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### Elements les plus souvent en cause")
    top_elem = (
        df_tout.groupby("Element").size()
        .sort_values(ascending=False).head(12)
        .reset_index(name="nb")
    )
    fig3 = px.bar(
        top_elem, x="Element", y="nb",
        color="nb",
        color_continuous_scale=["#FFFDE7", "#E67E22", "#CC0000"],
        labels={"nb": "Nb interventions", "Element": ""},
    )
    fig3.update_layout(
        margin=dict(t=10, b=10), height=280, coloraxis_showscale=False
    )
    st.plotly_chart(fig3, use_container_width=True)


def render_tab_liste(df: pd.DataFrame) -> None:
    """Affiche la liste paginee des interventions avec tri interactif."""
    st.markdown(
        f'<div class="main-title">📋 Interventions ({len(df)})</div>',
        unsafe_allow_html=True,
    )

    if df.empty:
        st.success("✅ Aucune intervention selon les filtres appliques.")
        return

    col_tri, col_ordre = st.columns([2, 1])
    with col_tri:
        tri = st.selectbox(
            "Trier par", ["Priorite", "Salle", "Element", "Traite"]
        )
    with col_ordre:
        ordre_asc = st.toggle("Croissant", value=True)

    df_affiche = df.copy()
    if tri == "Priorite":
        df_affiche["_ordre"] = df_affiche["Priorite"].map(ORDRE_PRIO_MAP)
        df_affiche = (
            df_affiche.sort_values("_ordre", ascending=ordre_asc)
            .drop(columns="_ordre")
        )
    else:
        df_affiche = df_affiche.sort_values(tri, ascending=ordre_asc)

    cols = ["Emoji", "Salle", "Element", "Valeur", "Priorite", "Traite"]
    df_display = df_affiche[cols].copy()
    df_display.columns = [
        "", "Salle", "Element", "Valeur", "Priorite", "Traite"
    ]
    df_display["Traite"] = df_display["Traite"].map(
        {True: "✅ Oui", False: "⏳ Non"}
    )

    def coloriser_ligne(row):
        return ["background-color: #111111; color: #FFFFFF"] * len(row)

    styled = df_display.style.apply(coloriser_ligne, axis=1)
    st.dataframe(styled, use_container_width=True, height=500, hide_index=True)


def render_tab_salle(df: pd.DataFrame, salles_sel: list) -> None:
    """Affiche les interventions groupees par salle sous forme de cartes."""
    st.markdown(
        '<div class="main-title">🏠 Vue par salle</div>',
        unsafe_allow_html=True,
    )

    if df.empty:
        st.success("✅ Aucune intervention pour les filtres selectionnes.")
        return

    salles_filtrees = df["Salle"].unique() if not salles_sel else salles_sel

    for salle in sorted(salles_filtrees):
        df_s = df[df["Salle"] == salle]
        if df_s.empty:
            continue

        nb = len(df_s)
        nb_crit = len(df_s[df_s["Priorite"] == "CRITIQUE"])
        nb_trait = len(df_s[df_s["Traite"]])

        with st.expander(
            f"**{salle}** — {nb} intervention(s)"
            + (f"  |  🔴 {nb_crit} critique(s)" if nb_crit else "")
            + (f"  |  ✅ {nb_trait} traitee(s)" if nb_trait else ""),
            expanded=(nb_crit > 0),
        ):
            df_sorted = df_s.sort_values(
                "Priorite", key=lambda x: x.map(ORDRE_PRIO_MAP)
            )
            # Batch : une seule injection HTML par salle au lieu de N appels
            cards = df_sorted.apply(_card_salle_html, axis=1)
            st.markdown("\n".join(cards), unsafe_allow_html=True)


def render_tab_edition() -> None:
    """Permet de cocher les interventions traitees et de suivre la progression."""
    st.markdown(
        '<div class="main-title">✅ Gestion des traitements</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Cochez les interventions realisees. "
        "L'etat est conserve pendant la session."
    )

    df_edit = st.session_state.df_base.copy()
    df_edit["Traite"] = df_edit.index.map(st.session_state.traites)

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        salle_edit = st.selectbox(
            "Filtrer par salle",
            ["Toutes"] + sorted(df_edit["Salle"].unique().tolist()),
        )
    with col_f2:
        prio_edit = st.selectbox(
            "Filtrer par priorite", ["Toutes"] + ORDRE_PRIO
        )

    df_filtered_edit = df_edit.copy()
    if salle_edit != "Toutes":
        df_filtered_edit = df_filtered_edit[
            df_filtered_edit["Salle"] == salle_edit
        ]
    if prio_edit != "Toutes":
        df_filtered_edit = df_filtered_edit[
            df_filtered_edit["Priorite"] == prio_edit
        ]

    c_tout, c_aucun = st.columns([1, 1])
    with c_tout:
        if st.button("✅ Tout marquer comme traite (vue actuelle)"):
            for idx in df_filtered_edit.index:
                st.session_state.traites[idx] = True
            st.rerun()
    with c_aucun:
        if st.button("↩️ Tout reinitialiser (vue actuelle)"):
            for idx in df_filtered_edit.index:
                st.session_state.traites[idx] = False
            st.rerun()

    st.divider()

    df_filtered_edit = df_filtered_edit.sort_values(
        "Priorite",
        key=lambda x: x.map(ORDRE_PRIO_MAP),
    )

    # Les checkboxes sont des widgets interactifs : boucle obligatoire
    for _, row in df_filtered_edit.iterrows():
        idx = row.name
        col_check, col_info = st.columns([1, 11])
        with col_check:
            checked = st.checkbox(
                "",
                value=st.session_state.traites.get(idx, False),
                key=f"chk_{idx}",
            )
            st.session_state.traites[idx] = checked
        with col_info:
            st.markdown(_card_edition_html(row), unsafe_allow_html=True)

    total_base = len(st.session_state.df_base)
    nb_traites = sum(st.session_state.traites.values())
    pct = int(nb_traites / total_base * 100) if total_base else 0
    st.divider()
    st.markdown(
        f"**Progression globale : {nb_traites}/{total_base} traitees ({pct}%)**"
    )
    st.progress(pct / 100)
