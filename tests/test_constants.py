#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tests pour constants.py : normaliser, hex2rgb, et coherence des constantes."""

import pytest

from constants import (
    COULEURS_PRIO,
    EMOJIS_PRIO,
    FONDS_PRIO,
    LABELS_PRIO,
    ORDRE_PRIO,
    ORDRE_PRIO_MAP,
    VALEURS_KO,
    VALEURS_OK,
    hex2rgb,
    normaliser,
)


# ── normaliser ─────────────────────────────────────────────────────────────────

class TestNormaliser:
    def test_supprime_accents(self):
        assert normaliser("Détérioré") == "deteriore"

    def test_minuscules(self):
        assert normaliser("MANQUANT") == "manquant"

    def test_strip(self):
        assert normaliser("  bon  ") == "bon"

    def test_accents_multiples(self):
        assert normaliser("À changer") == "a changer"

    def test_chaine_vide(self):
        assert normaliser("") == ""

    def test_none_converti(self):
        assert normaliser("None") == "none"

    def test_valeur_numerique(self):
        assert normaliser("42") == "42"


# ── hex2rgb ────────────────────────────────────────────────────────────────────

class TestHex2rgb:
    def test_avec_diese(self):
        assert hex2rgb("#C0392B") == (0xC0, 0x39, 0x2B)

    def test_sans_diese(self):
        assert hex2rgb("C0392B") == (0xC0, 0x39, 0x2B)

    def test_blanc(self):
        assert hex2rgb("#FFFFFF") == (255, 255, 255)

    def test_noir(self):
        assert hex2rgb("#000000") == (0, 0, 0)

    def test_bleu_lycee(self):
        assert hex2rgb("#1A3C6E") == (0x1A, 0x3C, 0x6E)

    def test_retourne_tuple_3_int(self):
        r, g, b = hex2rgb("#AABBCC")
        assert isinstance(r, int)
        assert isinstance(g, int)
        assert isinstance(b, int)


# ── Coherence des constantes ───────────────────────────────────────────────────

class TestCoherenceConstantes:
    def test_toutes_prios_dans_ordre(self):
        for prio in ["CRITIQUE", "REMPLACEMENT", "TRAVAUX", "A VERIFIER"]:
            assert prio in ORDRE_PRIO

    def test_ordre_prio_map_aligne_ordre_prio(self):
        for i, prio in enumerate(ORDRE_PRIO):
            assert ORDRE_PRIO_MAP[prio] == i

    def test_couleurs_couvrent_toutes_prios(self):
        for prio in ORDRE_PRIO:
            assert prio in COULEURS_PRIO
            assert prio in FONDS_PRIO
            assert prio in EMOJIS_PRIO
            assert prio in LABELS_PRIO

    def test_couleurs_prio_format_hex(self):
        for prio, couleur in COULEURS_PRIO.items():
            assert couleur.startswith("#"), f"{prio}: {couleur} doit commencer par #"
            assert len(couleur) == 7, f"{prio}: {couleur} doit faire 7 caracteres"

    def test_valeurs_ko_format(self):
        for motcle, valeur in VALEURS_KO.items():
            prio, emoji = valeur
            assert prio in ORDRE_PRIO, f"Priorite inconnue '{prio}' pour '{motcle}'"
            assert len(emoji) > 0

    def test_valeurs_ok_est_un_set(self):
        assert isinstance(VALEURS_OK, set)
        assert "" in VALEURS_OK
        assert "nan" in VALEURS_OK
