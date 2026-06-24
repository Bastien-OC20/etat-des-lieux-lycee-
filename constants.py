#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Constantes metier et utilitaires partages par toute l'application."""

import unicodedata

# ── Configuration generale ─────────────────────────────────────────────────────
ANNEE_SCOLAIRE = "2025/2026"
LOGO_PRINCIPALE = "logo.png"
LOGO_SECONDAIRE = "LOL.jpg"

# ── Valeurs metier ─────────────────────────────────────────────────────────────
VALEURS_KO = {
    "manquant":          ("CRITIQUE",     "🔴"),
    "manquante":         ("CRITIQUE",     "🔴"),
    "a changer":         ("REMPLACEMENT", "🟠"),
    "travaux a prevoir": ("TRAVAUX",      "🟡"),
    "dalles":            ("TRAVAUX",      "🟡"),
    "changer piles":     ("REMPLACEMENT", "🟠"),
    "deteriore":         ("CRITIQUE",     "🔴"),
    "deterioree":        ("CRITIQUE",     "🔴"),
}
VALEURS_OK = {"bon", "tres bon", "presente", "present", "x", "", "none", "nan"}

# ── Apparence par priorite ─────────────────────────────────────────────────────
COULEURS_PRIO = {
    "CRITIQUE":     "#C0392B",  # 5,41:1 sur blanc  (RGAA ✅)
    "REMPLACEMENT": "#A04000",  # 6,55:1 sur blanc  (RGAA ✅)
    "TRAVAUX":      "#7D6608",  # 5,59:1 sur blanc  (RGAA ✅)
    "A VERIFIER":   "#5D5D5D",  # 6,86:1 sur blanc  (RGAA ✅)
}
FONDS_PRIO = {
    "CRITIQUE":     "#FFEBEB",
    "REMPLACEMENT": "#FFF3E0",
    "TRAVAUX":      "#FFFDE7",
    "A VERIFIER":   "#F5F5F5",
}
EMOJIS_PRIO = {
    "CRITIQUE":     "🔴",
    "REMPLACEMENT": "🟠",
    "TRAVAUX":      "🟡",
    "A VERIFIER":   "⚪",
}
LABELS_PRIO = {
    "CRITIQUE":     "Manquant ou deteriore",
    "REMPLACEMENT": "A remplacer",
    "TRAVAUX":      "Travaux a prevoir",
    "A VERIFIER":   "Valeur a controler",
}
ORDRE_PRIO = ["CRITIQUE", "REMPLACEMENT", "TRAVAUX", "A VERIFIER"]
ORDRE_PRIO_MAP = {
    "CRITIQUE": 0, "REMPLACEMENT": 1, "TRAVAUX": 2, "A VERIFIER": 3
}


# ── Utilitaires partages ───────────────────────────────────────────────────────
def hex2rgb(h: str) -> tuple[int, int, int]:
    """Convertit une couleur hex (#RRGGBB ou RRGGBB) en tuple (R, G, B)."""
    h = h.lstrip("#")
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def normaliser(s: str) -> str:
    """Normalise une chaine : supprime les accents, met en minuscules et strip."""
    return (
        unicodedata.normalize("NFD", str(s))
        .encode("ascii", "ignore")
        .decode("ascii")
        .lower()
        .strip()
    )
