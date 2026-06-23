#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Chargement et analyse du fichier Excel etat des lieux."""

import io
import unicodedata

import pandas as pd
import streamlit as st

from constants import VALEURS_KO, VALEURS_OK


def normaliser(s: str) -> str:
    return (
        unicodedata.normalize("NFD", str(s))
        .encode("ascii", "ignore")
        .decode("ascii")
        .lower()
        .strip()
    )


def determiner_priorite(valeur: str):
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
    """Charge le xlsx et retourne un DataFrame plat des interventions."""
    xl = pd.ExcelFile(io.BytesIO(contenu))
    ws = xl.parse(xl.sheet_names[0], header=1)
    ws.columns = [str(c).strip() for c in ws.columns]
    col_salle = ws.columns[0]

    rows = []
    for _, row in ws.iterrows():
        salle = str(row[col_salle]).strip()
        if salle in ("nan", "None", ""):
            continue
        for col in ws.columns[1:]:
            if str(col).startswith("Unnamed") or str(col) == "nan":
                continue
            valeur = str(row[col]).strip() if pd.notna(row[col]) else ""
            res = determiner_priorite(valeur)
            if res:
                prio, emoji = res
                rows.append({
                    "Salle":    salle,
                    "Element":  col,
                    "Valeur":   valeur,
                    "Priorite": prio,
                    "Emoji":    emoji,
                    "Traite":   False,
                })
    return pd.DataFrame(rows)
