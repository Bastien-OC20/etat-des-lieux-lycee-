#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Chargement et analyse du fichier Excel etat des lieux."""

import io
from typing import Optional

import pandas as pd
import streamlit as st

from constants import VALEURS_KO, VALEURS_OK, normaliser


def determiner_priorite(valeur: str) -> Optional[tuple[str, str]]:
    """Retourne (priorite, emoji) si la valeur indique un probleme, None sinon."""
    norm = normaliser(valeur)
    if norm in VALEURS_OK:
        return None
    matched = VALEURS_KO.get(norm)
    if matched is None:
        for motcle, info in VALEURS_KO.items():
            if motcle in norm:
                matched = info
                break
    if matched:
        return matched  # (priorite, emoji)
    return ("A VERIFIER", "⚪")


@st.cache_data
def charger_xlsx(contenu: bytes) -> pd.DataFrame:
    """Charge le xlsx et retourne un DataFrame plat des interventions.

    Raises:
        ValueError: si le fichier est invalide, vide ou manque de colonnes.
    """
    try:
        xl = pd.ExcelFile(io.BytesIO(contenu))
    except Exception as e:
        raise ValueError(f"Impossible de lire le fichier Excel : {e}") from e

    if not xl.sheet_names:
        raise ValueError("Le fichier Excel ne contient aucune feuille.")

    try:
        ws = xl.parse(xl.sheet_names[0], header=1)
    except Exception as e:
        raise ValueError(f"Erreur lors de la lecture de la feuille : {e}") from e

    ws.columns = [str(c).strip() for c in ws.columns]

    if len(ws.columns) < 2:
        raise ValueError(
            "La feuille Excel ne contient pas assez de colonnes "
            "(colonne Salle + au moins un element attendu)."
        )

    col_salle = ws.columns[0]

    # Feuille avec structure valide mais sans donnees
    if ws.empty:
        return pd.DataFrame(
            columns=["Salle", "Element", "Valeur", "Priorite", "Emoji", "Traite"]
        )

    # Filtrage vectorise : exclure les lignes sans salle valide
    # isna() couvre les vrais NaN float, isin() couvre les strings residuelles
    mask_valide = (
        ~ws[col_salle].isna() &
        ~ws[col_salle].astype(str).str.strip().isin(["nan", "None", ""])
    )
    ws_clean = ws[mask_valide].copy()
    # Normalise les entiers lus comme float par pandas (ex: 102.0 → "102")
    ws_clean[col_salle] = ws_clean[col_salle].apply(
        lambda v: str(int(v)) if isinstance(v, float) and v.is_integer()
        else str(v).strip()
    )

    cols_elements = [
        c for c in ws.columns[1:]
        if not str(c).startswith("Unnamed") and str(c) != "nan"
    ]

    if not cols_elements:
        raise ValueError("Aucune colonne d'element valide trouvee dans le fichier.")

    # Toutes les salles filtrées (ex : colonne salle entièrement vide)
    if ws_clean.empty:
        return pd.DataFrame(
            columns=["Salle", "Element", "Valeur", "Priorite", "Emoji", "Traite"]
        )

    # Pivot long : une ligne par (salle, element)
    df_long = ws_clean[[col_salle] + cols_elements].melt(
        id_vars=[col_salle], var_name="Element", value_name="Valeur_raw"
    )
    df_long["Valeur"] = df_long["Valeur_raw"].apply(
        lambda v: str(v).strip() if pd.notna(v) else ""
    )
    df_long = df_long.drop(columns="Valeur_raw").rename(columns={col_salle: "Salle"})

    # Determination des priorites via apply (evite iterrows)
    resultats = df_long["Valeur"].apply(determiner_priorite)
    mask_ko = resultats.notna()
    df_ko = df_long[mask_ko].copy()
    df_ko["Priorite"] = resultats[mask_ko].apply(lambda x: x[0])
    df_ko["Emoji"] = resultats[mask_ko].apply(lambda x: x[1])
    df_ko["Traite"] = False

    return df_ko[
        ["Salle", "Element", "Valeur", "Priorite", "Emoji", "Traite"]
    ].reset_index(drop=True)
