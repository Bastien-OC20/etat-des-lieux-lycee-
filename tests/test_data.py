#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tests pour data.py : determiner_priorite et charger_xlsx."""

import io
import sys
from unittest.mock import MagicMock

import openpyxl
import pandas as pd
import pytest

# Mock streamlit avant l'import de data pour eviter le contexte Streamlit
_mock_st = MagicMock()
_mock_st.cache_data = lambda func: func
sys.modules.setdefault("streamlit", _mock_st)

from data import charger_xlsx, determiner_priorite  # noqa: E402


# ── Helpers ────────────────────────────────────────────────────────────────────

def _make_excel(data: dict) -> bytes:
    """Cree un fichier Excel en memoire au format attendu par charger_xlsx.

    charger_xlsx utilise header=1, donc :
    - Ligne Excel 1 : titre ignore
    - Ligne Excel 2 : en-tetes de colonnes
    - Lignes Excel 3+ : donnees
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Etat des lieux"])      # ligne titre (ignoree)
    ws.append(list(data.keys()))       # en-tetes
    n = max((len(v) for v in data.values()), default=0)
    for i in range(n):
        row = [data[k][i] if i < len(data[k]) else None for k in data]
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── determiner_priorite ────────────────────────────────────────────────────────

class TestDeterminerPriorite:
    def test_valeur_ok_retourne_none(self):
        assert determiner_priorite("bon") is None
        assert determiner_priorite("tres bon") is None
        assert determiner_priorite("x") is None
        assert determiner_priorite("") is None

    def test_manquant_est_critique(self):
        prio, emoji = determiner_priorite("manquant")
        assert prio == "CRITIQUE"
        assert emoji == "🔴"

    def test_deteriore_est_critique(self):
        prio, _ = determiner_priorite("deteriore")
        assert prio == "CRITIQUE"

    def test_a_changer_est_remplacement(self):
        prio, emoji = determiner_priorite("a changer")
        assert prio == "REMPLACEMENT"
        assert emoji == "🟠"

    def test_travaux_a_prevoir_est_travaux(self):
        prio, emoji = determiner_priorite("travaux a prevoir")
        assert prio == "TRAVAUX"
        assert emoji == "🟡"

    def test_valeur_inconnue_est_a_verifier(self):
        prio, _ = determiner_priorite("valeur_xyz_inconnue")
        assert prio == "A VERIFIER"

    def test_insensible_a_la_casse(self):
        prio, _ = determiner_priorite("MANQUANT")
        assert prio == "CRITIQUE"

    def test_insensible_aux_accents(self):
        prio, _ = determiner_priorite("détérioré")
        assert prio == "CRITIQUE"

    def test_correspondance_partielle(self):
        prio, _ = determiner_priorite("element manquant depuis 2024")
        assert prio == "CRITIQUE"

    def test_nan_est_ok(self):
        assert determiner_priorite("nan") is None
        assert determiner_priorite("None") is None


# ── charger_xlsx ───────────────────────────────────────────────────────────────

class TestChargerXlsx:
    def test_chargement_valide(self):
        excel = _make_excel({
            "Salle": ["101", "102"],
            "Chaise": ["bon", "manquant"],
            "Bureau": ["a changer", "tres bon"],
        })
        df = charger_xlsx(excel)
        assert isinstance(df, pd.DataFrame)
        assert len(df) == 2
        assert set(df.columns) == {
            "Salle", "Element", "Valeur", "Priorite", "Emoji", "Traite"
        }

    def test_colonnes_ok_ne_genere_pas_de_lignes(self):
        excel = _make_excel({
            "Salle": ["101"],
            "Chaise": ["bon"],
            "Bureau": ["tres bon"],
        })
        df = charger_xlsx(excel)
        assert df.empty

    def test_salle_nulle_ignoree(self):
        excel = _make_excel({
            "Salle": [None, "102"],
            "Chaise": ["manquant", "manquant"],
        })
        df = charger_xlsx(excel)
        assert len(df) == 1
        assert df.iloc[0]["Salle"] == "102"

    def test_df_vide_retourne_colonnes_correctes(self):
        excel = _make_excel({"Salle": [], "Chaise": []})
        df = charger_xlsx(excel)
        assert df.empty
        assert list(df.columns) == [
            "Salle", "Element", "Valeur", "Priorite", "Emoji", "Traite"
        ]

    def test_fichier_corrompu_leve_valueerror(self):
        with pytest.raises(ValueError, match="Impossible de lire"):
            charger_xlsx(b"ceci n'est pas un fichier Excel")

    def test_caracteres_speciaux_dans_salle(self):
        excel = _make_excel({
            "Salle": ["Salle <B>101</B>"],
            "Fenetre": ["manquant"],
        })
        df = charger_xlsx(excel)
        assert len(df) == 1
        assert "<B>" in df.iloc[0]["Salle"]

    def test_caracteres_speciaux_dans_valeur(self):
        excel = _make_excel({
            "Salle": ["101"],
            "Note": ["<script>alert('xss')</script>"],
        })
        df = charger_xlsx(excel)
        # La valeur inconnue est marquee A VERIFIER, pas injectee
        assert len(df) == 1
        assert df.iloc[0]["Priorite"] == "A VERIFIER"

    def test_priorites_correctes(self):
        excel = _make_excel({
            "Salle": ["101", "101", "101", "101"],
            "Col":   ["manquant", "a changer", "travaux a prevoir", "valeur_xyz"],
        })
        df = charger_xlsx(excel)
        prios = set(df["Priorite"].tolist())
        assert "CRITIQUE" in prios
        assert "REMPLACEMENT" in prios
        assert "TRAVAUX" in prios
        assert "A VERIFIER" in prios

    def test_colonne_traite_initialisee_a_false(self):
        excel = _make_excel({
            "Salle": ["101"],
            "Chaise": ["manquant"],
        })
        df = charger_xlsx(excel)
        assert not df["Traite"].any()
        assert df["Traite"].sum() == 0

    def test_fichier_une_seule_colonne_leve_valueerror(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Titre"])
        ws.append(["Salle"])
        ws.append(["101"])
        buf = io.BytesIO()
        wb.save(buf)
        with pytest.raises(ValueError):
            charger_xlsx(buf.getvalue())
