#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Constantes metier partagees par toute l'application."""

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
ORDRE_PRIO = ["CRITIQUE", "REMPLACEMENT", "TRAVAUX", "A VERIFIER"]
ORDRE_PRIO_MAP = {
    "CRITIQUE": 0, "REMPLACEMENT": 1, "TRAVAUX": 2, "A VERIFIER": 3
}
