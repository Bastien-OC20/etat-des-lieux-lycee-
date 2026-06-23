#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generation des exports DOCX et XLSX."""

import io
from datetime import date

import pandas as pd

from constants import COULEURS_PRIO, FONDS_PRIO


def generer_docx(df: pd.DataFrame) -> bytes:
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    def hex2rgb(h):
        h = h.lstrip("#")
        return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

    def set_cell_bg(cell, hex_color):
        hex_color = hex_color.lstrip("#")
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), hex_color)
        tcPr.append(shd)

    def set_para_bg(para, hex_color):
        hex_color = hex_color.lstrip("#")
        pPr = para._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), hex_color)
        pPr.append(shd)

    def add_border_bottom(para, color="1A3C6E"):
        pPr = para._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        b = OxmlElement("w:bottom")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "1")
        b.set(qn("w:color"), color)
        pBdr.append(b)
        pPr.append(pBdr)

    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    # Titre
    p = doc.add_paragraph()
    r = p.add_run("Etat des lieux - Interventions a entreprendre")
    r.bold = True
    r.font.size = Pt(18)
    r.font.color.rgb = RGBColor(0x1A, 0x3C, 0x6E)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_border_bottom(p)

    p2 = doc.add_paragraph()
    date_str = date.today().strftime('%d/%m/%Y')
    p2.add_run(f"Lycee - Annee 2025/2026  |  Genere le {date_str}")
    p2.runs[0].font.size = Pt(10)
    p2.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # Bilan global
    total_i = len(df)
    non_traites = (
        len(df[~df["Traite"]]) if "Traite" in df.columns else total_i
    )
    p_stat = doc.add_paragraph()
    r_stat = p_stat.add_run(
        f"Total interventions : {total_i}  |  Non traitees : {non_traites}"
    )
    r_stat.bold = True
    r_stat.font.size = Pt(11)
    r_stat.font.color.rgb = RGBColor(0x1A, 0x3C, 0x6E)
    p_stat.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    for salle in df["Salle"].unique():
        df_s = df[df["Salle"] == salle]
        p_hdr = doc.add_paragraph()
        r_hdr = p_hdr.add_run(
            f"  Salle : {salle}  ({len(df_s)} intervention(s))"
        )
        r_hdr.bold = True
        r_hdr.font.size = Pt(12)
        r_hdr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_para_bg(p_hdr, "1A3C6E")

        table = doc.add_table(rows=1, cols=4)
        table.style = "Table Grid"
        headers = ["Element", "Valeur originale", "Priorite", "Statut"]
        for cell, label in zip(table.rows[0].cells, headers):
            cell.text = label
            set_cell_bg(cell, "D5E8F0")
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].runs[0].font.size = Pt(9)

        for _, row in df_s.sort_values("Priorite").iterrows():
            cells = table.add_row().cells
            cells[0].text = str(row["Element"])
            cells[1].text = str(row["Valeur"])
            cells[2].text = str(row["Priorite"])
            cells[3].text = "✓ Traite" if row.get("Traite") else "En attente"

            bg = FONDS_PRIO.get(row["Priorite"], "#F5F5F5").lstrip("#")
            tc_hex = COULEURS_PRIO.get(row["Priorite"], "#888888").lstrip("#")
            set_cell_bg(cells[2], bg)
            r_prio = cells[2].paragraphs[0].runs[0]
            r_prio.bold = True
            r_prio.font.size = Pt(9)
            r_prio.font.color.rgb = RGBColor(*hex2rgb(tc_hex))
            for c in cells:
                c.paragraphs[0].runs[0].font.size = Pt(9)

        for row in table.rows:
            row.cells[0].width = Cm(5)
            row.cells[1].width = Cm(5)
            row.cells[2].width = Cm(3.5)
            row.cells[3].width = Cm(2.5)

        doc.add_paragraph()

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def generer_xlsx(df: pd.DataFrame) -> bytes:
    from openpyxl.styles import PatternFill, Font, Alignment

    FILLS = {
        "CRITIQUE":     PatternFill("solid", fgColor="FFEBEB"),
        "REMPLACEMENT": PatternFill("solid", fgColor="FFF3E0"),
        "TRAVAUX":      PatternFill("solid", fgColor="FFFDE7"),
        "A VERIFIER":   PatternFill("solid", fgColor="F5F5F5"),
    }
    FONTS = {
        "CRITIQUE":     Font(bold=True, color="C0392B"),
        "REMPLACEMENT": Font(bold=True, color="A04000"),
        "TRAVAUX":      Font(bold=True, color="7D6608"),
        "A VERIFIER":   Font(bold=True, color="5D5D5D"),
    }

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        export_df = df[
            ["Salle", "Element", "Valeur", "Priorite", "Traite"]
        ].copy()
        export_df["Traite"] = export_df["Traite"].map(
            {True: "Oui", False: "Non"}
        )
        export_df.to_excel(writer, index=False, sheet_name="Interventions")

        ws = writer.sheets["Interventions"]
        header_fill = PatternFill("solid", fgColor="1A3C6E")
        header_font = Font(bold=True, color="FFFFFF")
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        for row in ws.iter_rows(min_row=2):
            prio = row[3].value
            if prio in FILLS:
                row[3].fill = FILLS[prio]
                row[3].font = FONTS[prio]

        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 30
        ws.column_dimensions["C"].width = 25
        ws.column_dimensions["D"].width = 18
        ws.column_dimensions["E"].width = 10

    return buf.getvalue()
