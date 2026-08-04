#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Générateur du classeur MG Hôte - Gestion de conciergerie."""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.pagebreak import Break

# ---------------------------------------------------------------------------
# PALETTE MG HÔTE
# ---------------------------------------------------------------------------
BLACK = "1A1A1A"
GOLD = "C9A227"
GOLD_LIGHT = "F6E9C4"
GOLD_INPUT = "FFF6DF"   # fond des cellules modifiables (saisie utilisateur)
WHITE = "FFFFFF"
GREY_LIGHT = "F2F2F2"
GREY_BORDER = "BFBFBF"
RED_ALERT = "F8CBAD"
RED_TEXT = "C00000"
GREEN_OK = "C6EFCE"
GREEN_TEXT = "006100"

FONT_NAME = "Arial"

thin = Side(style="thin", color=GREY_BORDER)
BORDER_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
wb.remove(wb.active)

# ---------------------------------------------------------------------------
# STYLE HELPERS
# ---------------------------------------------------------------------------

def banner(ws, row, text, col_span=12, height=26):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=col_span)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT_NAME, size=14, bold=True, color=GOLD)
    c.fill = PatternFill("solid", fgColor=BLACK)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = height
    for col in range(1, col_span + 1):
        ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor=BLACK)


def subtitle(ws, row, text, col_span=12):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=col_span)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT_NAME, size=10, italic=True, color=BLACK)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)


def section(ws, row, text, col_span=12, height=20):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=col_span)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT_NAME, size=11, bold=True, color=BLACK)
    c.fill = PatternFill("solid", fgColor=GOLD)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = height


def header_row(ws, row, headers, start_col=1):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col + i, value=h)
        c.font = Font(name=FONT_NAME, size=10, bold=True, color=GOLD)
        c.fill = PatternFill("solid", fgColor=BLACK)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER_ALL
    ws.row_dimensions[row].height = 30


def input_cell(cell, number_format=None, align="left"):
    cell.fill = PatternFill("solid", fgColor=GOLD_INPUT)
    cell.font = Font(name=FONT_NAME, size=10, color=BLACK)
    cell.border = BORDER_ALL
    cell.protection = Protection(locked=False)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    if number_format:
        cell.number_format = number_format
    return cell


def formula_cell(cell, number_format=None, align="left", bold=False):
    cell.font = Font(name=FONT_NAME, size=10, color=BLACK, bold=bold)
    cell.fill = PatternFill("solid", fgColor=WHITE)
    cell.border = BORDER_ALL
    cell.alignment = Alignment(horizontal=align, vertical="center")
    if number_format:
        cell.number_format = number_format
    return cell


def label_cell(cell, bold=True, size=10):
    cell.font = Font(name=FONT_NAME, size=size, bold=bold, color=BLACK)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    return cell


def kpi_tile(ws, row, col, title, formula, number_format="0", title_color=GOLD):
    """Crée une tuile KPI 2 lignes x 3 colonnes (titre + valeur)."""
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 2)
    t = ws.cell(row=row, column=col, value=title)
    t.font = Font(name=FONT_NAME, size=9, bold=True, color=WHITE)
    t.fill = PatternFill("solid", fgColor=BLACK)
    t.alignment = Alignment(horizontal="center", vertical="center")
    for cc in range(col, col + 3):
        ws.cell(row=row, column=cc).fill = PatternFill("solid", fgColor=BLACK)
    ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 2, end_column=col + 2)
    v = ws.cell(row=row + 1, column=col, value=formula)
    v.font = Font(name=FONT_NAME, size=18, bold=True, color=GOLD)
    v.fill = PatternFill("solid", fgColor=GREY_LIGHT)
    v.alignment = Alignment(horizontal="center", vertical="center")
    v.number_format = number_format
    v.border = BORDER_ALL
    for cc in range(col, col + 3):
        for rr in (row + 1, row + 2):
            ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=GREY_LIGHT)
            ws.cell(row=rr, column=cc).border = BORDER_ALL


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def add_table(ws, name, ref, style="TableStyleLight1", stripes=False):
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(
        name=style, showFirstColumn=False, showLastColumn=False,
        showRowStripes=stripes, showColumnStripes=False
    )
    ws.add_table(t)
    return t


def add_dropdown(ws, cell_range, source, allow_blank=True):
    dv = DataValidation(type="list", formula1=source, allow_blank=allow_blank, showDropDown=False)
    ws.add_data_validation(dv)
    dv.add(cell_range)
    return dv


def protect_sheet(ws, password="MGHote2026", insert_rows=False):
    ws.protection.sheet = True
    ws.protection.password = password
    ws.protection.formatCells = False
    ws.protection.formatColumns = True
    ws.protection.formatRows = True
    ws.protection.insertRows = insert_rows
    ws.protection.deleteRows = False
    ws.protection.sort = True
    ws.protection.autoFilter = True

# ---------------------------------------------------------------------------
# 1) PARAMÈTRES (listes de référence + seuils) — créé en premier car référencé
#    par les listes déroulantes des autres onglets
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Paramètres")
set_col_widths(ws, [26, 30, 4, 4, 4, 4, 4, 4])
banner(ws, 1, "PARAMÈTRES — Listes de référence & seuils (ne pas supprimer/déplacer les colonnes)", col_span=8)
subtitle(ws, 2, "Cet onglet alimente les listes déroulantes de tout le classeur. Vous pouvez modifier les seuils en bas de page.", col_span=8)

section(ws, 4, "Listes déroulantes", col_span=8)
lists = {
    "A": ("Plateformes", ["Airbnb", "Booking", "Direct"]),
    "B": ("Statuts réservation", ["Confirmée", "En cours", "Terminée", "Annulée"]),
    "C": ("Types de bien", ["Villa", "Appartement", "Riad", "Autre"]),
    "D": ("Statuts ménage", ["Fait", "En attente"]),
    "E": ("Priorités maintenance", ["Basse", "Moyenne", "Haute", "Urgente"]),
    "F": ("Statuts maintenance", ["Signalé", "En cours", "Résolu"]),
    "G": ("Paiement soldé", ["Oui", "Non"]),
}
for col_letter, (title, values) in lists.items():
    col = openpyxl.utils.column_index_from_string(col_letter)
    hc = ws.cell(row=6, column=col, value=title)
    hc.font = Font(name=FONT_NAME, size=10, bold=True, color=BLACK)
    hc.fill = PatternFill("solid", fgColor=GOLD_LIGHT)
    hc.border = BORDER_ALL
    hc.alignment = Alignment(horizontal="center", wrap_text=True)
    for i, v in enumerate(values):
        cell = ws.cell(row=7 + i, column=col, value=v)
        cell.font = Font(name=FONT_NAME, size=10, color=BLACK)
        cell.border = BORDER_ALL

section(ws, 14, "Seuils d'alerte & paramètres généraux", col_span=8)
label_cell(ws.cell(row=16, column=2, value="Seuil rentabilité faible (marge nette)"))
input_cell(ws.cell(row=16, column=4, value=0.20), number_format="0%")
label_cell(ws.cell(row=17, column=2, value="Seuil logement vide (jours sans réservation)"))
input_cell(ws.cell(row=17, column=4, value=14), number_format="0")
label_cell(ws.cell(row=18, column=2, value="Commission MG Hôte par défaut"))
input_cell(ws.cell(row=18, column=4, value=0.20), number_format="0%")
label_cell(ws.cell(row=19, column=2, value="Mot de passe protection des onglets"))
c = ws.cell(row=19, column=4, value="MGHote2026")
c.font = Font(name=FONT_NAME, size=10, italic=True, color=BLACK)
c.border = BORDER_ALL

# Defined names (référencés par toutes les formules du classeur)
wb.defined_names["Seuil_Rentabilite"] = DefinedName("Seuil_Rentabilite", attr_text="Paramètres!$D$16")
wb.defined_names["Seuil_Vide"] = DefinedName("Seuil_Vide", attr_text="Paramètres!$D$17")
wb.defined_names["Commission_Defaut"] = DefinedName("Commission_Defaut", attr_text="Paramètres!$D$18")

ws.sheet_view.showGridLines = False
ws_parametres = ws
print("Paramètres OK")

# ---------------------------------------------------------------------------
# 2) BIENS
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Biens")
set_col_widths(ws, [10, 22, 13, 22, 14, 10, 10, 18, 16, 12, 10, 26])
banner(ws, 1, "BIENS — Propriétés gérées par MG Hôte", col_span=12)
subtitle(ws, 2, "Une ligne par bien. Pour ajouter un bien, insérez une ligne dans le tableau (Bien 2, Bien 3…). Cellules dorées = à saisir.", col_span=12)

biens_headers = ["ID Bien", "Nom / Désignation", "Type", "Adresse", "Ville",
                  "Capacité (pers.)", "Nb chambres", "Propriétaire",
                  "Plateformes actives", "Commission MG Hôte", "Statut", "Notes"]
header_row(ws, 4, biens_headers)

# ligne 1 = exemple réel (Bien 1) ; lignes suivantes = pré-formatées pour la croissance (Bien 2, Bien 3…)
N_BIENS = 15
first_row, last_row = 5, 5 + N_BIENS - 1
for i in range(N_BIENS):
    row = 5 + i
    is_example = (i == 0)
    input_cell(ws.cell(row=row, column=1, value=f"Bien {i+1}"))
    input_cell(ws.cell(row=row, column=2, value="À compléter" if is_example else ""))
    input_cell(ws.cell(row=row, column=3, value="Villa" if is_example else ""))
    input_cell(ws.cell(row=row, column=4, value=""))
    input_cell(ws.cell(row=row, column=5, value=""))
    input_cell(ws.cell(row=row, column=6, value=4 if is_example else None), number_format="0", align="center")
    input_cell(ws.cell(row=row, column=7, value=2 if is_example else None), number_format="0", align="center")
    input_cell(ws.cell(row=row, column=8, value="À compléter" if is_example else ""))
    input_cell(ws.cell(row=row, column=9, value="Airbnb, Booking" if is_example else ""))
    input_cell(ws.cell(row=row, column=10, value="=Commission_Defaut"), number_format="0%", align="center")
    input_cell(ws.cell(row=row, column=11, value="Actif" if is_example else "Inactif"), align="center")
    input_cell(ws.cell(row=row, column=12, value=""))
r = last_row

add_table(ws, "T_Biens", f"A4:L{r}")
add_dropdown(ws, f"C{first_row}:C{r}", "=Paramètres!$C$7:$C$10")
add_dropdown(ws, f"K{first_row}:K{r}", '"Actif,Inactif"')
ws.freeze_panes = "A5"
ws.sheet_view.showGridLines = False

# mise en forme conditionnelle: statut Inactif en gris
ws.conditional_formatting.add(
    f"K{first_row}:K{r}",
    FormulaRule(formula=[f'K{first_row}="Inactif"'], fill=PatternFill("solid", fgColor=GREY_LIGHT))
)
ws_biens = ws
print("Biens OK")

# ---------------------------------------------------------------------------
# 3) PROPRIÉTAIRES
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Propriétaires")
set_col_widths(ws, [10, 24, 16, 24, 26, 22, 20, 18, 14, 26])
banner(ws, 1, "PROPRIÉTAIRES — Fiches propriétaires", col_span=10)
subtitle(ws, 2, "Une ligne par propriétaire. Ces informations alimentent les fiches et factures imprimables.", col_span=10)

prop_headers = ["ID Propriétaire", "Nom complet", "Téléphone", "Email", "Adresse",
                "Bien(s) géré(s)", "RIB / IBAN", "Mode paiement préféré",
                "Début contrat", "Notes"]
header_row(ws, 4, prop_headers)

N_PROP = 15
for i in range(N_PROP):
    row = 5 + i
    is_example = (i == 0)
    input_cell(ws.cell(row=row, column=1, value=f"P{i+1}"))
    input_cell(ws.cell(row=row, column=2, value="À compléter" if is_example else ""))
    input_cell(ws.cell(row=row, column=3, value=""))
    input_cell(ws.cell(row=row, column=4, value=""))
    input_cell(ws.cell(row=row, column=5, value=""))
    input_cell(ws.cell(row=row, column=6, value="Bien 1" if is_example else ""))
    input_cell(ws.cell(row=row, column=7, value=""))
    input_cell(ws.cell(row=row, column=8, value="Virement" if is_example else ""), align="center")
    input_cell(ws.cell(row=row, column=9, value=None), number_format="dd/mm/yyyy", align="center")
    input_cell(ws.cell(row=row, column=10, value=""))
r = 4 + N_PROP

add_table(ws, "T_Proprietaires", f"A4:J{r}")
ws.freeze_panes = "A5"
ws.sheet_view.showGridLines = False
ws_proprietaires = ws
print("Propriétaires OK")

# ---------------------------------------------------------------------------
# 4) RÉSERVATIONS
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Réservations")
set_col_widths(ws, [6, 10, 12, 20, 20, 13, 13, 8, 9, 12, 10, 12, 14, 12, 12, 10, 10, 10, 10, 24])
banner(ws, 1, "RÉSERVATIONS — Directes, Airbnb & Booking", col_span=20)
subtitle(ws, 2, "Ajoutez une réservation en bas du tableau (utilisez l'onglet 'Formulaire Réservation' comme aide de saisie). Cellules dorées = à saisir.", col_span=20)

res_headers = ["N°", "Bien", "Plateforme", "Client", "Contact client",
               "Arrivée", "Départ", "Nuits", "Voyageurs", "Prix total (TTC)",
               "Commission plateforme %", "Montant net", "Statut",
               "Acompte reçu", "Solde dû", "Soldé", "Ménage fait",
               "Retard paiement", "Mois", "Notes"]
header_row(ws, 4, res_headers)

N_RES = 150
for i in range(N_RES):
    r = 5 + i
    is_example = (i == 0)
    formula_cell(ws.cell(row=r, column=1, value=f'=IF(D{r}="","",ROW()-4)'), number_format="0", align="center")
    input_cell(ws.cell(row=r, column=2, value="Bien 1" if is_example else ""), align="center")
    input_cell(ws.cell(row=r, column=3, value="Airbnb" if is_example else ""), align="center")
    input_cell(ws.cell(row=r, column=4, value="Ex: Jean Dupont" if is_example else ""))
    input_cell(ws.cell(row=r, column=5, value="jean.dupont@email.com" if is_example else ""))
    input_cell(ws.cell(row=r, column=6, value="2026-07-10" if is_example else None), number_format="dd/mm/yyyy", align="center")
    input_cell(ws.cell(row=r, column=7, value="2026-07-15" if is_example else None), number_format="dd/mm/yyyy", align="center")
    formula_cell(ws.cell(row=r, column=8, value=f'=IF(AND(G{r}<>"",F{r}<>""),G{r}-F{r},"")'), number_format="0", align="center")
    input_cell(ws.cell(row=r, column=9, value=2 if is_example else None), number_format="0", align="center")
    input_cell(ws.cell(row=r, column=10, value=5000 if is_example else None), number_format="#,##0 MAD", align="right")
    input_cell(ws.cell(row=r, column=11, value=0.03 if is_example else None), number_format="0%", align="center")
    formula_cell(ws.cell(row=r, column=12, value=f'=IF(J{r}="","",J{r}*(1-K{r}))'), number_format="#,##0 MAD", align="right")
    input_cell(ws.cell(row=r, column=13, value="Confirmée" if is_example else ""), align="center")
    input_cell(ws.cell(row=r, column=14, value=5000 if is_example else None), number_format="#,##0 MAD", align="right")
    formula_cell(ws.cell(row=r, column=15, value=f'=IF(J{r}="","",IF(M{r}="Annulée",0,J{r}-N{r}))'), number_format="#,##0 MAD", align="right")
    formula_cell(ws.cell(row=r, column=16, value=f'=IF(J{r}="","",IF(M{r}="Annulée","-",IF(O{r}<=0,"Oui","Non")))'), align="center")
    input_cell(ws.cell(row=r, column=17, value="Oui" if is_example else ""), align="center")
    formula_cell(ws.cell(row=r, column=18, value=f'=IF(OR(J{r}="",G{r}=""),"",IF(AND(M{r}<>"Annulée",O{r}>0,G{r}<TODAY()),"⚠️ Retard",""))'), align="center")
    formula_cell(ws.cell(row=r, column=19, value=f'=IF(F{r}="","",TEXT(F{r},"yyyy-mm"))'), align="center")
    input_cell(ws.cell(row=r, column=20, value=""))
r = 4 + N_RES

add_table(ws, "T_Reservations", f"A4:T{r}")
add_dropdown(ws, f"B5:B{r}", "=Biens!$A$5:$A$300")
add_dropdown(ws, f"C5:C{r}", "=Paramètres!$A$7:$A$9")
add_dropdown(ws, f"M5:M{r}", "=Paramètres!$B$7:$B$10")
add_dropdown(ws, f"Q5:Q{r}", "=Paramètres!$G$7:$G$8")

# mise en forme conditionnelle
ws.conditional_formatting.add(f"M5:M{r+500}", FormulaRule(formula=['M5="Annulée"'], fill=PatternFill("solid", fgColor=RED_ALERT)))
ws.conditional_formatting.add(f"M5:M{r+500}", FormulaRule(formula=['M5="Confirmée"'], fill=PatternFill("solid", fgColor=GREEN_OK)))
ws.conditional_formatting.add(f"R5:R{r+500}", FormulaRule(formula=['R5="⚠️ Retard"'], font=Font(color=RED_TEXT, bold=True)))

ws.freeze_panes = "A5"
ws.sheet_view.showGridLines = False
ws.auto_filter.ref = f"A4:T{r}"
ws_reservations = ws
print("Réservations OK")

# ---------------------------------------------------------------------------
# 5) FORMULAIRE RÉSERVATION (saisie guidée, sans macro)
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Formulaire Réservation")
set_col_widths(ws, [30, 30, 4, 30, 30])
banner(ws, 1, "➕ NOUVELLE RÉSERVATION — Formulaire de saisie", col_span=5)
subtitle(ws, 2, "Remplissez les cellules dorées ci-dessous, puis reportez la ligne dans l'onglet 'Réservations' (copier/coller la ligne complétée en bas du tableau).", col_span=5)

form_fields = [
    ("Bien", "B4", "list_bien"),
    ("Plateforme", "B5", "list_plateforme"),
    ("Nom du client", "B6", None),
    ("Contact client (tél / email)", "B7", None),
    ("Date d'arrivée", "B8", "date"),
    ("Date de départ", "B9", "date"),
    ("Nombre de voyageurs", "B10", "num"),
    ("Prix total (TTC)", "B11", "money"),
    ("Commission plateforme (%)", "B12", "pct"),
    ("Statut de la réservation", "B13", "list_statut"),
    ("Acompte reçu", "B14", "money"),
    ("Notes", "B15", None),
]
row = 4
for label, addr, kind in form_fields:
    lc = ws.cell(row=row, column=1, value=label)
    label_cell(lc)
    lc.fill = PatternFill("solid", fgColor=GOLD_LIGHT)
    lc.border = BORDER_ALL
    lc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    cell = ws.cell(row=row, column=2)
    if kind == "date":
        input_cell(cell, number_format="dd/mm/yyyy", align="center")
    elif kind == "num":
        input_cell(cell, number_format="0", align="center")
    elif kind == "money":
        input_cell(cell, number_format="#,##0 MAD", align="right")
    elif kind == "pct":
        input_cell(cell, number_format="0%", align="center")
    else:
        input_cell(cell)
    row += 1
    ws.row_dimensions[row - 1].height = 20

add_dropdown(ws, "B4", "=Biens!$A$5:$A$300")
add_dropdown(ws, "B5", "=Paramètres!$A$7:$A$9")
add_dropdown(ws, "B13", "=Paramètres!$B$7:$B$10")

section(ws, 17, "Calculs automatiques (aperçu avant report)", col_span=5)
calc_fields = [
    ("Nuits", "=IF(AND(B9<>\"\",B8<>\"\"),B9-B8,\"\")", "0"),
    ("Montant net (après commission plateforme)", "=IF(B11<>\"\",B11*(1-B12),\"\")", "#,##0 MAD"),
    ("Solde dû", "=IF(AND(B11<>\"\",B14<>\"\"),B11-B14,\"\")", "#,##0 MAD"),
]
row = 18
for label, formula, fmt in calc_fields:
    lc = ws.cell(row=row, column=1, value=label)
    label_cell(lc, bold=False)
    lc.border = BORDER_ALL
    fc = ws.cell(row=row, column=2, value=formula)
    formula_cell(fc, number_format=fmt, align="center", bold=True)
    row += 1

note = ws.cell(row=row + 1, column=1, value=(
    "Astuce : ce formulaire sert d'aide de saisie visuelle. La méthode la plus rapide au quotidien reste "
    "d'ajouter directement une nouvelle ligne dans le tableau de l'onglet 'Réservations' (les listes déroulantes "
    "et les formules se recopient automatiquement)."
))
ws.merge_cells(start_row=row + 1, start_column=1, end_row=row + 3, end_column=5)
note.font = Font(name=FONT_NAME, size=9, italic=True, color=BLACK)
note.alignment = Alignment(wrap_text=True, vertical="top")

ws.sheet_view.showGridLines = False
ws_formulaire = ws
print("Formulaire Réservation OK")

# ---------------------------------------------------------------------------
# 6) MÉNAGE
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Ménage")
set_col_widths(ws, [13, 10, 14, 13, 34, 16, 14])
banner(ws, 1, "MÉNAGE — Suivi par réservation", col_span=7)
subtitle(ws, 2, "Une ligne par passage de ménage. Le statut 'En attente' avec une date passée déclenche une alerte automatique.", col_span=7)

menage_headers = ["Date", "Bien", "Réservation liée (N°)", "Statut",
                   "Compte-rendu", "Responsable", "Alerte"]
header_row(ws, 4, menage_headers)

N_MEN = 150
for i in range(N_MEN):
    r = 5 + i
    is_example = (i == 0)
    input_cell(ws.cell(row=r, column=1, value="2026-07-15" if is_example else None), number_format="dd/mm/yyyy", align="center")
    input_cell(ws.cell(row=r, column=2, value="Bien 1" if is_example else ""), align="center")
    input_cell(ws.cell(row=r, column=3, value=1 if is_example else None), number_format="0", align="center")
    input_cell(ws.cell(row=r, column=4, value="Fait" if is_example else ""), align="center")
    input_cell(ws.cell(row=r, column=5, value="RAS, logement en bon état" if is_example else ""))
    input_cell(ws.cell(row=r, column=6, value="Ex: Fatima" if is_example else ""))
    formula_cell(ws.cell(row=r, column=7, value=f'=IF(A{r}="","",IF(AND(D{r}="En attente",A{r}<TODAY()),"⚠️ En retard",""))'), align="center")
r = 4 + N_MEN

add_table(ws, "T_Menage", f"A4:G{r}")
add_dropdown(ws, f"B5:B{r}", "=Biens!$A$5:$A$300")
add_dropdown(ws, f"D5:D{r}", "=Paramètres!$D$7:$D$8")
ws.conditional_formatting.add(f"G5:G{r}", FormulaRule(formula=[f'G5="⚠️ En retard"'], fill=PatternFill("solid", fgColor=RED_ALERT), font=Font(color=RED_TEXT, bold=True)))
ws.conditional_formatting.add(f"D5:D{r}", FormulaRule(formula=['D5="Fait"'], fill=PatternFill("solid", fgColor=GREEN_OK)))
ws.conditional_formatting.add(f"D5:D{r}", FormulaRule(formula=['D5="En attente"'], fill=PatternFill("solid", fgColor=RED_ALERT)))
ws.freeze_panes = "A5"
ws.sheet_view.showGridLines = False
ws.auto_filter.ref = f"A4:G{r}"
ws_menage = ws
print("Ménage OK")

# ---------------------------------------------------------------------------
# 7) MAINTENANCE
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Maintenance")
set_col_widths(ws, [14, 10, 34, 11, 12, 13, 14, 26, 10])
banner(ws, 1, "MAINTENANCE — Suivi des incidents", col_span=9)
subtitle(ws, 2, "Une ligne par incident signalé. Priorisez selon l'urgence pour le confort du voyageur.", col_span=9)

maint_headers = ["Date signalement", "Bien", "Description incident", "Priorité",
                  "Statut", "Coût", "Date résolution", "Notes", "Mois"]
header_row(ws, 4, maint_headers)

N_MAINT = 60
for i in range(N_MAINT):
    r = 5 + i
    is_example = (i == 0)
    input_cell(ws.cell(row=r, column=1, value="2026-07-05" if is_example else None), number_format="dd/mm/yyyy", align="center")
    input_cell(ws.cell(row=r, column=2, value="Bien 1" if is_example else ""), align="center")
    input_cell(ws.cell(row=r, column=3, value="Ex: Climatiseur en panne" if is_example else ""))
    input_cell(ws.cell(row=r, column=4, value="Haute" if is_example else ""), align="center")
    input_cell(ws.cell(row=r, column=5, value="Résolu" if is_example else ""), align="center")
    input_cell(ws.cell(row=r, column=6, value=350 if is_example else None), number_format="#,##0 MAD", align="right")
    input_cell(ws.cell(row=r, column=7, value="2026-07-06" if is_example else None), number_format="dd/mm/yyyy", align="center")
    input_cell(ws.cell(row=r, column=8, value=""))
    formula_cell(ws.cell(row=r, column=9, value=f'=IF(A{r}="","",TEXT(A{r},"yyyy-mm"))'), align="center")
r = 4 + N_MAINT

add_table(ws, "T_Maintenance", f"A4:I{r}")
add_dropdown(ws, f"B5:B{r}", "=Biens!$A$5:$A$300")
add_dropdown(ws, f"D5:D{r}", "=Paramètres!$E$7:$E$10")
add_dropdown(ws, f"E5:E{r}", "=Paramètres!$F$7:$F$9")

ws.conditional_formatting.add(f"D5:D{r}", FormulaRule(formula=['D5="Urgente"'], fill=PatternFill("solid", fgColor=RED_ALERT), font=Font(color=RED_TEXT, bold=True)))
ws.conditional_formatting.add(f"D5:D{r}", FormulaRule(formula=['D5="Haute"'], fill=PatternFill("solid", fgColor="FCE4D6")))
ws.conditional_formatting.add(f"E5:E{r}", FormulaRule(formula=['E5="Résolu"'], fill=PatternFill("solid", fgColor=GREEN_OK)))
ws.freeze_panes = "A5"
ws.sheet_view.showGridLines = False
ws.auto_filter.ref = f"A4:I{r}"
ws_maintenance = ws
print("Maintenance OK")

# ---------------------------------------------------------------------------
# 8) FINANCES (mini comptabilité mensuelle par bien)
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Finances")
set_col_widths(ws, [10, 10, 13, 11, 13, 12, 9, 11, 13, 12, 15, 13, 12, 14])
banner(ws, 1, "FINANCES — Mini comptabilité mensuelle par bien", col_span=14)
subtitle(ws, 2, "Une ligne par bien et par mois. Le revenu locatif se calcule automatiquement depuis l'onglet Réservations ; saisissez les charges du mois en jaune.", col_span=14)

fin_headers = ["Mois", "Bien", "Revenu locatif", "Ménage", "Consommables",
               "Électricité", "Eau", "Publicité", "Maintenance", "Total charges",
               "Commission MG Hôte %", "Commission MG Hôte", "Résultat net propriétaire",
               "Marge nette", "Alerte rentabilité"]
# 15 headers but col widths defined 14 -> fix widths list length
set_col_widths(ws, [10, 10, 13, 11, 13, 12, 9, 11, 13, 12, 15, 15, 16, 11, 15])
header_row(ws, 4, fin_headers)

N_FIN = 60
for i in range(N_FIN):
    r = 5 + i
    is_example = (i == 0)
    input_cell(ws.cell(row=r, column=1, value="2026-07" if is_example else None), align="center")
    input_cell(ws.cell(row=r, column=2, value="Bien 1" if is_example else ""), align="center")
    formula_cell(ws.cell(row=r, column=3, value=(
        f'=IF(OR(A{r}="",B{r}=""),"",SUMIFS(T_Reservations[Prix total (TTC)],T_Reservations[Bien],B{r},T_Reservations[Mois],A{r},T_Reservations[Statut],"<>Annulée"))'
    )), number_format="#,##0 MAD", align="right")
    input_cell(ws.cell(row=r, column=4, value=400 if is_example else None), number_format="#,##0 MAD", align="right")
    input_cell(ws.cell(row=r, column=5, value=150 if is_example else None), number_format="#,##0 MAD", align="right")
    input_cell(ws.cell(row=r, column=6, value=300 if is_example else None), number_format="#,##0 MAD", align="right")
    input_cell(ws.cell(row=r, column=7, value=100 if is_example else None), number_format="#,##0 MAD", align="right")
    input_cell(ws.cell(row=r, column=8, value=200 if is_example else None), number_format="#,##0 MAD", align="right")
    formula_cell(ws.cell(row=r, column=9, value=(
        f'=IF(OR(A{r}="",B{r}=""),"",SUMIFS(T_Maintenance[Coût],T_Maintenance[Bien],B{r},T_Maintenance[Mois],A{r}))'
    )), number_format="#,##0 MAD", align="right")
    formula_cell(ws.cell(row=r, column=10, value=f'=IF(A{r}="","",SUM(D{r}:I{r}))'), number_format="#,##0 MAD", align="right", bold=True)
    input_cell(ws.cell(row=r, column=11, value="=Commission_Defaut" if is_example else None), number_format="0%", align="center")
    formula_cell(ws.cell(row=r, column=12, value=f'=IF(OR(A{r}="",K{r}=""),"",C{r}*K{r})'), number_format="#,##0 MAD", align="right")
    formula_cell(ws.cell(row=r, column=13, value=f'=IF(A{r}="","",C{r}-J{r}-L{r})'), number_format="#,##0 MAD", align="right", bold=True)
    formula_cell(ws.cell(row=r, column=14, value=f'=IF(A{r}="","",IFERROR(M{r}/C{r},0))'), number_format="0%", align="center")
    formula_cell(ws.cell(row=r, column=15, value=f'=IF(A{r}="","",IF(N{r}<Seuil_Rentabilite,"⚠️ Faible","OK"))'), align="center")
r = 4 + N_FIN

add_table(ws, "T_Finances", f"A4:O{r}")
add_dropdown(ws, f"B5:B{r}", "=Biens!$A$5:$A$300")

ws.conditional_formatting.add(f"O5:O{r}", FormulaRule(formula=['O5="⚠️ Faible"'], fill=PatternFill("solid", fgColor=RED_ALERT), font=Font(color=RED_TEXT, bold=True)))
ws.conditional_formatting.add(f"O5:O{r}", FormulaRule(formula=['O5="OK"'], fill=PatternFill("solid", fgColor=GREEN_OK)))
ws.freeze_panes = "A5"
ws.sheet_view.showGridLines = False
ws.auto_filter.ref = f"A4:O{r}"
ws_finances = ws
print("Finances OK")

# ---------------------------------------------------------------------------
# 9) ALERTES
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Alertes")
set_col_widths(ws, [4, 14, 14, 14, 4, 14, 14, 14, 4, 14, 14, 14, 4, 14, 14, 14])
banner(ws, 1, "ALERTES — Ce qui mérite votre attention aujourd'hui", col_span=16)
subtitle(ws, 2, "Ces compteurs se mettent à jour automatiquement. Filtrez les colonnes correspondantes dans chaque onglet pour voir le détail.", col_span=16)

kpi_tile(ws, 5, 2, "PAIEMENTS EN RETARD", "=COUNTIF(T_Reservations[Retard paiement],\"⚠️ Retard\")", "0")
kpi_tile(ws, 5, 6, "MÉNAGES OUBLIÉS", '=COUNTIF(T_Menage[Alerte],"⚠️ En retard")', "0")
kpi_tile(ws, 5, 10, "RENTABILITÉ FAIBLE (MOIS EN COURS)", '=COUNTIFS(T_Finances[Mois],TEXT(TODAY(),"yyyy-mm"),T_Finances[Alerte rentabilité],"⚠️ Faible")', "0")
kpi_tile(ws, 5, 14, "LOGEMENTS VIDES > SEUIL", "=COUNTIF($E$14:$E$23,\"⚠️ Vide trop longtemps\")", "0")

# conditional colour on tile values (row 6-7 merged value cell anchor at row6)
for col in (2, 6, 10, 14):
    anchor = f"{get_column_letter(col)}6"
    ws.conditional_formatting.add(anchor, CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor=RED_ALERT), font=Font(name=FONT_NAME, size=18, bold=True, color=RED_TEXT)))
    ws.conditional_formatting.add(anchor, CellIsRule(operator="equal", formula=["0"], fill=PatternFill("solid", fgColor=GREEN_OK), font=Font(name=FONT_NAME, size=18, bold=True, color=GREEN_TEXT)))

section(ws, 11, "Détail — Logements vides depuis trop longtemps (vs seuil défini dans Paramètres)", col_span=16)
header_row(ws, 13, ["Bien", "Dernier départ", "Jours sans réservation", "Alerte"], start_col=2)
for i in range(10):
    row = 14 + i
    bien_ref = f"Biens!A{5+i}"
    formula_cell(ws.cell(row=row, column=2, value=f'=IF({bien_ref}="","",{bien_ref})'), align="center")
    dep = ws.cell(row=row, column=3, value=(
        f'=IF(B{row}="","",IF(COUNTIFS(T_Reservations[Bien],B{row},T_Reservations[Statut],"<>Annulée")=0,"",'
        f'_xlfn.MAXIFS(T_Reservations[Départ],T_Reservations[Bien],B{row},T_Reservations[Statut],"<>Annulée")))'
    ))
    formula_cell(dep, number_format="dd/mm/yyyy", align="center")
    jours = ws.cell(row=row, column=4, value=f'=IF(OR(B{row}="",C{row}=""),"",TODAY()-C{row})')
    formula_cell(jours, number_format="0", align="center")
    alerte = ws.cell(row=row, column=5, value=f'=IF(D{row}="","",IF(D{row}>Seuil_Vide,"⚠️ Vide trop longtemps",""))')
    formula_cell(alerte, align="center")

ws.conditional_formatting.add("E14:E23", FormulaRule(formula=['E14="⚠️ Vide trop longtemps"'], fill=PatternFill("solid", fgColor=RED_ALERT), font=Font(color=RED_TEXT, bold=True)))
ws.sheet_view.showGridLines = False
ws_alertes = ws
print("Alertes OK")

# ---------------------------------------------------------------------------
# 10) TABLEAU DE BORD
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Tableau de bord")
set_col_widths(ws, [4, 14, 14, 14, 4, 14, 14, 14, 4, 14, 14, 14, 4, 14, 14, 14])
banner(ws, 1, "MG HÔTE — TABLEAU DE BORD", col_span=16)
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=16)
dcell = ws.cell(row=2, column=1, value='="Données au "&TEXT(TODAY(),"dd/mm/yyyy")')
dcell.font = Font(name=FONT_NAME, size=10, italic=True, color=BLACK)
dcell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

section(ws, 4, "Aujourd'hui", col_span=16)
kpi_tile(ws, 5, 2, "ARRIVÉES AUJOURD'HUI", '=COUNTIFS(T_Reservations[Arrivée],TODAY(),T_Reservations[Statut],"<>Annulée")', "0")
kpi_tile(ws, 5, 6, "TAUX D'OCCUPATION (MOIS)", "=IFERROR(J6/(COUNTIF(T_Biens[Statut],\"Actif\")*DAY(EOMONTH(TODAY(),0))),0)", "0%")
kpi_tile(ws, 5, 10, "NUITS RÉSERVÉES (MOIS)", '=SUMIFS(T_Reservations[Nuits],T_Reservations[Mois],TEXT(TODAY(),"yyyy-mm"),T_Reservations[Statut],"<>Annulée")', "0")

section(ws, 9, "Bloc Financier (mois en cours)", col_span=16)
fin_rows = [
    ("Chiffre d'affaires brut du mois", '=SUMIFS(T_Reservations[Prix total (TTC)],T_Reservations[Mois],TEXT(TODAY(),"yyyy-mm"),T_Reservations[Statut],"<>Annulée")', "#,##0 MAD"),
    ("Charges du mois", '=SUMIFS(T_Finances[Total charges],T_Finances[Mois],TEXT(TODAY(),"yyyy-mm"))', "#,##0 MAD"),
    ("Résultat net du mois", '=SUMIFS(T_Finances[Résultat net propriétaire],T_Finances[Mois],TEXT(TODAY(),"yyyy-mm"))', "#,##0 MAD"),
    ("Impayés en cours (tous biens)", '=SUMIFS(T_Reservations[Solde dû],T_Reservations[Soldé],"Non")', "#,##0 MAD"),
    ("Résultat net cumulé (année)", '=SUMIFS(T_Finances[Résultat net propriétaire],T_Finances[Mois],TEXT(TODAY(),"yyyy")&"*")', "#,##0 MAD"),
]
row = 11
for label, formula, fmt in fin_rows:
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
    lc = ws.cell(row=row, column=2, value=label)
    label_cell(lc, bold=False)
    lc.border = BORDER_ALL
    lc.fill = PatternFill("solid", fgColor=WHITE)
    ws.merge_cells(start_row=row, start_column=7, end_row=row, end_column=9)
    fc = ws.cell(row=row, column=7, value=formula)
    formula_cell(fc, number_format=fmt, align="right", bold=True)
    row += 1

section(ws, 17, "Bloc Opérationnel — Alertes", col_span=16)
kpi_tile(ws, 19, 2, "PAIEMENTS EN RETARD", "='Alertes'!B6", "0")
kpi_tile(ws, 19, 6, "MÉNAGES OUBLIÉS", "='Alertes'!F6", "0")
kpi_tile(ws, 19, 10, "RENTABILITÉ FAIBLE", "='Alertes'!J6", "0")
kpi_tile(ws, 19, 14, "LOGEMENTS VIDES", "='Alertes'!N6", "0")
for col in (2, 6, 10, 14):
    anchor = f"{get_column_letter(col)}20"
    ws.conditional_formatting.add(anchor, CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor=RED_ALERT), font=Font(name=FONT_NAME, size=18, bold=True, color=RED_TEXT)))
    ws.conditional_formatting.add(anchor, CellIsRule(operator="equal", formula=["0"], fill=PatternFill("solid", fgColor=GREEN_OK), font=Font(name=FONT_NAME, size=18, bold=True, color=GREEN_TEXT)))

section(ws, 23, "Bloc Performance — Évolution mensuelle (CA & taux d'occupation)", col_span=16)
ws.merge_cells(start_row=25, start_column=2, end_row=40, end_column=16)  # zone réservée au graphique

# --- zone de données du graphique (12 derniers mois) ---
CHART_ROW0 = 45
section(ws, CHART_ROW0 - 1, "Données du graphique — ne pas supprimer (12 derniers mois)", col_span=16)
header_row(ws, CHART_ROW0, ["Mois", "Chiffre d'affaires", "Taux d'occupation"], start_col=2)
for i in range(12):
    months_back = 11 - i
    row = CHART_ROW0 + 1 + i
    mcell = ws.cell(row=row, column=2, value=f"=EOMONTH(TODAY(),-{months_back}-1)+1")
    formula_cell(mcell, number_format="mmm-yy", align="center")
    ca = ws.cell(row=row, column=3, value=(
        f'=SUMIFS(T_Reservations[Prix total (TTC)],T_Reservations[Mois],TEXT(B{row},"yyyy-mm"),T_Reservations[Statut],"<>Annulée")'
    ))
    formula_cell(ca, number_format="#,##0 MAD", align="right")
    tx = ws.cell(row=row, column=4, value=(
        f'=IFERROR(SUMIFS(T_Reservations[Nuits],T_Reservations[Mois],TEXT(B{row},"yyyy-mm"),T_Reservations[Statut],"<>Annulée")'
        f'/(COUNTIF(T_Biens[Statut],"Actif")*DAY(EOMONTH(B{row},0))),0)'
    ))
    formula_cell(tx, number_format="0%", align="center")

note = ws.cell(row=CHART_ROW0 + 14, column=2, value=(
    "Hypothèse : le taux d'occupation mensuel affiché ici rattache chaque nuit réservée au mois de la date d'arrivée "
    "(approximation simple, sans répartition au prorata pour les séjours à cheval sur deux mois)."
))
ws.merge_cells(start_row=CHART_ROW0 + 14, start_column=2, end_row=CHART_ROW0 + 15, end_column=16)
note.font = Font(name=FONT_NAME, size=9, italic=True, color=BLACK)
note.alignment = Alignment(wrap_text=True, vertical="top")

chart1 = LineChart()
chart1.title = "Évolution du chiffre d'affaires"
chart1.style = 2
chart1.y_axis.title = "CA (MAD)"
chart1.x_axis.title = "Mois"
chart1.height = 7.5
chart1.width = 15
data = Reference(ws, min_col=3, min_row=CHART_ROW0, max_row=CHART_ROW0 + 12)
cats = Reference(ws, min_col=2, min_row=CHART_ROW0 + 1, max_row=CHART_ROW0 + 12)
chart1.add_data(data, titles_from_data=True)
chart1.set_categories(cats)
chart1.series[0].graphicalProperties.line.solidFill = GOLD
chart1.series[0].graphicalProperties.line.width = 25000
ws.add_chart(chart1, "B25")

chart2 = LineChart()
chart2.title = "Évolution du taux d'occupation"
chart2.style = 2
chart2.y_axis.title = "Taux d'occupation"
chart2.y_axis.numFmt = "0%"
chart2.x_axis.title = "Mois"
chart2.height = 7.5
chart2.width = 15
data2 = Reference(ws, min_col=4, min_row=CHART_ROW0, max_row=CHART_ROW0 + 12)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats)
chart2.series[0].graphicalProperties.line.solidFill = BLACK
chart2.series[0].graphicalProperties.line.width = 25000
ws.add_chart(chart2, "J25")

ws.sheet_view.showGridLines = False
ws_dashboard = ws
print("Tableau de bord OK")

# ---------------------------------------------------------------------------
# 11) FACTURE PROPRIÉTAIRE (imprimable / export PDF)
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Facture Propriétaire")
set_col_widths(ws, [3, 26, 22, 4, 22, 22, 3])
banner(ws, 1, "MG HÔTE — FACTURE PROPRIÉTAIRE", col_span=7)
subtitle(ws, 2, "Sélectionnez le bien et le mois : la facture se génère automatiquement depuis l'onglet Finances.", col_span=7)

section(ws, 4, "Sélection", col_span=7)
label_cell(ws.cell(row=6, column=2, value="Bien"))
input_cell(ws.cell(row=6, column=3, value="Bien 1"), align="center")
label_cell(ws.cell(row=6, column=5, value="Mois (aaaa-mm)"))
input_cell(ws.cell(row=6, column=6, value="2026-07"), align="center")
add_dropdown(ws, "C6", "=Biens!$A$5:$A$300")

label_cell(ws.cell(row=7, column=2, value="N° de facture"))
formula_cell(ws.cell(row=7, column=3, value='="FACT-"&C6&"-"&F6'), bold=True, align="center")
label_cell(ws.cell(row=7, column=5, value="Date d'émission"))
formula_cell(ws.cell(row=7, column=6, value="=TODAY()"), number_format="dd/mm/yyyy", align="center")

section(ws, 9, "Propriétaire", col_span=7)
label_cell(ws.cell(row=11, column=2, value="Nom"))
formula_cell(ws.cell(row=11, column=3, value='=IFERROR(INDEX(Biens!$H:$H,MATCH($C$6,Biens!$A:$A,0)),"")'))
label_cell(ws.cell(row=11, column=5, value="Mode de paiement"))
formula_cell(ws.cell(row=11, column=6, value='=IFERROR(INDEX(Propriétaires!$H:$H,MATCH($C$11,Propriétaires!$B:$B,0)),"Virement")'))
label_cell(ws.cell(row=12, column=2, value="RIB / IBAN"))
formula_cell(ws.cell(row=12, column=3, value='=IFERROR(INDEX(Propriétaires!$G:$G,MATCH($C$11,Propriétaires!$B:$B,0)),"")'))

section(ws, 14, "Détail financier du mois", col_span=7)
header_row(ws, 16, ["Poste", "Montant"], start_col=2)
inv_rows = [
    ("Revenu locatif brut", '=SUMIFS(T_Reservations[Prix total (TTC)],T_Reservations[Bien],$C$6,T_Reservations[Mois],$F$6,T_Reservations[Statut],"<>Annulée")'),
    ("Ménage", '=-SUMIFS(T_Finances[Ménage],T_Finances[Bien],$C$6,T_Finances[Mois],$F$6)'),
    ("Consommables", '=-SUMIFS(T_Finances[Consommables],T_Finances[Bien],$C$6,T_Finances[Mois],$F$6)'),
    ("Électricité", '=-SUMIFS(T_Finances[Électricité],T_Finances[Bien],$C$6,T_Finances[Mois],$F$6)'),
    ("Eau", '=-SUMIFS(T_Finances[Eau],T_Finances[Bien],$C$6,T_Finances[Mois],$F$6)'),
    ("Publicité", '=-SUMIFS(T_Finances[Publicité],T_Finances[Bien],$C$6,T_Finances[Mois],$F$6)'),
    ("Maintenance", '=-SUMIFS(T_Finances[Maintenance],T_Finances[Bien],$C$6,T_Finances[Mois],$F$6)'),
    ("Commission MG Hôte", '=-SUMIFS(T_Finances[Commission MG Hôte],T_Finances[Bien],$C$6,T_Finances[Mois],$F$6)'),
]
row = 17
for label, formula in inv_rows:
    lc = ws.cell(row=row, column=2, value=label)
    label_cell(lc, bold=False)
    lc.border = BORDER_ALL
    fc = ws.cell(row=row, column=3, value=formula)
    formula_cell(fc, number_format="#,##0 MAD;(#,##0) MAD", align="right")
    row += 1

lc = ws.cell(row=row, column=2, value="NET À VERSER AU PROPRIÉTAIRE")
label_cell(lc, bold=True)
lc.fill = PatternFill("solid", fgColor=GOLD)
lc.border = BORDER_ALL
fc = ws.cell(row=row, column=3, value=f"=SUM(C17:C{row-1})")
formula_cell(fc, number_format="#,##0 MAD", align="right", bold=True)
fc.fill = PatternFill("solid", fgColor=GOLD_LIGHT)

foot_row = row + 3
foot = ws.cell(row=foot_row, column=2, value="Merci de votre confiance. — MG Hôte, Conciergerie Airbnb & Booking")
ws.merge_cells(start_row=foot_row, start_column=2, end_row=foot_row, end_column=6)
foot.font = Font(name=FONT_NAME, size=9, italic=True, color=BLACK)
ws.cell(row=foot_row + 3, column=2, value="Signature MG Hôte").font = Font(name=FONT_NAME, size=9, color=BLACK)
ws.cell(row=foot_row + 3, column=5, value="Signature Propriétaire").font = Font(name=FONT_NAME, size=9, color=BLACK)

ws.sheet_view.showGridLines = False
ws.print_area = f"A1:G{foot_row + 6}"
ws.page_setup.orientation = "portrait"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws_facture = ws
print("Facture Propriétaire OK")

# ---------------------------------------------------------------------------
# 12) RAPPORT MENSUEL (imprimable / export PDF)
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Rapport Mensuel")
set_col_widths(ws, [3, 30, 18, 3, 30, 18, 3])
banner(ws, 1, "MG HÔTE — RAPPORT MENSUEL DE SYNTHÈSE", col_span=7)
subtitle(ws, 2, "Sélectionnez le mois : la synthèse se génère automatiquement, tous biens confondus.", col_span=7)

label_cell(ws.cell(row=4, column=2, value="Mois (aaaa-mm)"))
input_cell(ws.cell(row=4, column=3, value="2026-07"), align="center")
label_cell(ws.cell(row=4, column=5, value="Date d'édition"))
formula_cell(ws.cell(row=4, column=6, value="=TODAY()"), number_format="dd/mm/yyyy", align="center")

section(ws, 6, "Synthèse financière", col_span=7)
fin_rows = [
    ("Chiffre d'affaires brut total", '=SUMIFS(T_Reservations[Prix total (TTC)],T_Reservations[Mois],$C$4,T_Reservations[Statut],"<>Annulée")', "#,##0 MAD"),
    ("Charges totales", '=SUMIFS(T_Finances[Total charges],T_Finances[Mois],$C$4)', "#,##0 MAD"),
    ("Commission MG Hôte totale", '=SUMIFS(T_Finances[Commission MG Hôte],T_Finances[Mois],$C$4)', "#,##0 MAD"),
    ("Résultat net global (propriétaires)", '=SUMIFS(T_Finances[Résultat net propriétaire],T_Finances[Mois],$C$4)', "#,##0 MAD"),
    ("Marge nette moyenne", '=IFERROR(C11/C8,0)', "0%"),
]
row = 8
for label, formula, fmt in fin_rows:
    lc = ws.cell(row=row, column=2, value=label)
    label_cell(lc, bold=False)
    lc.border = BORDER_ALL
    fc = ws.cell(row=row, column=3, value=formula)
    formula_cell(fc, number_format=fmt, align="right", bold=(row == 11))
    if row == 11:
        fc.fill = PatternFill("solid", fgColor=GOLD_LIGHT)
    row += 1

section(ws, 14, "Activité réservations", col_span=7)
act_rows = [
    ("Réservations confirmées / en cours / terminées", '=COUNTIFS(T_Reservations[Mois],$C$4,T_Reservations[Statut],"<>Annulée")', "0"),
    ("Annulations", '=COUNTIFS(T_Reservations[Mois],$C$4,T_Reservations[Statut],"Annulée")', "0"),
    ("Nuits réservées", '=SUMIFS(T_Reservations[Nuits],T_Reservations[Mois],$C$4,T_Reservations[Statut],"<>Annulée")', "0"),
    ("Taux d'occupation moyen", '=IFERROR(C18/(COUNTIF(T_Biens[Statut],"Actif")*DAY(EOMONTH(DATEVALUE($C$4&"-01"),0))),0)', "0%"),
]
row = 16
for label, formula, fmt in act_rows:
    lc = ws.cell(row=row, column=2, value=label)
    label_cell(lc, bold=False)
    lc.border = BORDER_ALL
    fc = ws.cell(row=row, column=3, value=formula)
    formula_cell(fc, number_format=fmt, align="right")
    row += 1

section(ws, 21, "Opérationnel", col_span=7)
op_rows = [
    ("Ménages effectués", '=COUNTIFS(T_Menage[Date],">="&DATEVALUE($C$4&"-01"),T_Menage[Date],"<="&EOMONTH(DATEVALUE($C$4&"-01"),0),T_Menage[Statut],"Fait")', "0"),
    ("Ménages en attente", '=COUNTIFS(T_Menage[Date],">="&DATEVALUE($C$4&"-01"),T_Menage[Date],"<="&EOMONTH(DATEVALUE($C$4&"-01"),0),T_Menage[Statut],"En attente")', "0"),
    ("Incidents maintenance signalés", '=COUNTIFS(T_Maintenance[Mois],$C$4)', "0"),
    ("Incidents maintenance résolus", '=COUNTIFS(T_Maintenance[Mois],$C$4,T_Maintenance[Statut],"Résolu")', "0"),
]
row = 23
for label, formula, fmt in op_rows:
    lc = ws.cell(row=row, column=2, value=label)
    label_cell(lc, bold=False)
    lc.border = BORDER_ALL
    fc = ws.cell(row=row, column=3, value=formula)
    formula_cell(fc, number_format=fmt, align="right")
    row += 1

foot = ws.cell(row=row + 2, column=2, value="Rapport généré automatiquement — MG Hôte, Conciergerie Airbnb & Booking")
ws.merge_cells(start_row=row + 2, start_column=2, end_row=row + 2, end_column=6)
foot.font = Font(name=FONT_NAME, size=9, italic=True, color=BLACK)

ws.sheet_view.showGridLines = False
ws.print_area = f"A1:G{row + 4}"
ws.page_setup.orientation = "portrait"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws_rapport = ws
print("Rapport Mensuel OK")

# ---------------------------------------------------------------------------
# 13) ACCUEIL (page de garde + mode d'emploi)
# ---------------------------------------------------------------------------
ws = wb.create_sheet("Accueil")
set_col_widths(ws, [3, 40, 40, 3])
ws.sheet_view.showGridLines = False

ws.merge_cells("A1:D5")
logo = ws.cell(row=1, column=1, value="MG HÔTE")
logo.font = Font(name=FONT_NAME, size=36, bold=True, color=GOLD)
logo.fill = PatternFill("solid", fgColor=BLACK)
logo.alignment = Alignment(horizontal="center", vertical="center")
for rr in range(1, 6):
    for cc in range(1, 5):
        ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=BLACK)
ws.merge_cells("A6:D6")
tagline = ws.cell(row=6, column=1, value="Conciergerie Airbnb & Booking — Villas, Appartements, Riads")
tagline.font = Font(name=FONT_NAME, size=11, italic=True, color=BLACK)
tagline.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[6].height = 22
note_logo = ws.cell(row=7, column=1, value=(
    "Note : insérez votre logo officiel ici (Insertion > Images) à la place de ce bandeau texte — "
    "les couleurs noir/or du classeur sont déjà calées sur votre charte."
))
ws.merge_cells("A7:D7")
note_logo.font = Font(name=FONT_NAME, size=8, italic=True, color="808080")
note_logo.alignment = Alignment(horizontal="center", wrap_text=True)

section(ws, 9, "Sommaire du classeur", col_span=4)
sommaire = [
    ("Tableau de bord", "Vue d'ensemble : arrivées du jour, occupation, finances, alertes, évolution mensuelle."),
    ("Réservations", "Journal de toutes les réservations (Airbnb, Booking, directes) et annulations."),
    ("Formulaire Réservation", "Aide de saisie visuelle pour ajouter une nouvelle réservation."),
    ("Biens", "Liste des biens gérés (Bien 1, Bien 2…)."),
    ("Propriétaires", "Fiches des propriétaires (contact, RIB, mode de paiement)."),
    ("Finances", "Mini comptabilité mensuelle par bien : revenus, charges, résultat net."),
    ("Ménage", "Suivi du ménage par réservation : fait / en attente, compte-rendu."),
    ("Maintenance", "Suivi des incidents et de leur priorité."),
    ("Alertes", "Paiements en retard, ménages oubliés, rentabilité faible, logements vides."),
    ("Facture Propriétaire", "Facture imprimable/PDF générée automatiquement par bien et par mois."),
    ("Rapport Mensuel", "Synthèse mensuelle imprimable/PDF, tous biens confondus."),
    ("Paramètres", "Listes déroulantes et seuils d'alerte — à ne modifier qu'avec précaution."),
]
row = 11
for tab, desc in sommaire:
    lc = ws.cell(row=row, column=2, value=tab)
    label_cell(lc)
    lc.border = BORDER_ALL
    lc.fill = PatternFill("solid", fgColor=GOLD_LIGHT)
    dc = ws.cell(row=row, column=3, value=desc)
    dc.font = Font(name=FONT_NAME, size=9, color=BLACK)
    dc.alignment = Alignment(wrap_text=True, vertical="center")
    dc.border = BORDER_ALL
    ws.row_dimensions[row].height = 28
    row += 1

section(ws, row + 1, "Légende des couleurs", col_span=4)
row += 3
c1 = ws.cell(row=row, column=2, value="Cellule dorée")
c1.fill = PatternFill("solid", fgColor=GOLD_INPUT)
c1.border = BORDER_ALL
ws.cell(row=row, column=3, value="= zone modifiable (saisie libre).").font = Font(name=FONT_NAME, size=9)
row += 1
c2 = ws.cell(row=row, column=2, value="Cellule blanche")
c2.fill = PatternFill("solid", fgColor=WHITE)
c2.border = BORDER_ALL
ws.cell(row=row, column=3, value="= calcul automatique (formule protégée).").font = Font(name=FONT_NAME, size=9)
row += 2

section(ws, row, "Sécurité & sauvegarde", col_span=4)
row += 2
security_notes = [
    "Les onglets sont protégés : seules les cellules dorées sont modifiables, les formules sont verrouillées "
    "(mot de passe : voir onglet Paramètres).",
    "Structure du classeur protégée pour éviter la suppression ou le déplacement accidentel d'un onglet.",
    "Activez l'enregistrement automatique dans Excel : enregistrez le fichier sur OneDrive/SharePoint puis "
    "activez « Enregistrement automatique » en haut de la fenêtre Excel.",
    "Pour retrouver un historique des versions : Fichier > Informations > Historique des versions (nécessite "
    "OneDrive/SharePoint).",
    "Pour un suivi collaboratif des modifications : Révision > Partager le classeur (ou co-édition via OneDrive "
    "si vous travaillez à plusieurs).",
]
for n in security_notes:
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
    c = ws.cell(row=row, column=2, value="• " + n)
    c.font = Font(name=FONT_NAME, size=9, color=BLACK)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[row].height = 28
    row += 1

row += 1
section(ws, row, "Export PDF", col_span=4)
row += 2
pdf_notes = [
    "Facture Propriétaire et Rapport Mensuel sont préformatés pour l'impression (zone d'impression + mise en page "
    "1 page définies).",
    "Pour exporter en PDF : ouvrez l'onglet souhaité, puis Fichier > Exporter > Créer un document PDF/XPS "
    "(ou Ctrl+P > Imprimante « Microsoft Print to PDF »).",
]
for n in pdf_notes:
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
    c = ws.cell(row=row, column=2, value="• " + n)
    c.font = Font(name=FONT_NAME, size=9, color=BLACK)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[row].height = 28
    row += 1

ws_accueil = ws
print("Accueil OK")

# ---------------------------------------------------------------------------
# FINALISATION : ordre des onglets, couleurs, protection, propriétés
# ---------------------------------------------------------------------------
sheet_order = [
    "Accueil", "Tableau de bord", "Réservations", "Formulaire Réservation",
    "Biens", "Propriétaires", "Finances", "Ménage", "Maintenance", "Alertes",
    "Facture Propriétaire", "Rapport Mensuel", "Paramètres",
]
wb._sheets = [wb[name] for name in sheet_order]

for name in sheet_order:
    wb[name].sheet_properties.tabColor = GOLD if name != "Paramètres" else GREY_BORDER

# Protection : onglets à croissance (ajout de lignes autorisé)
for name in ["Réservations", "Ménage", "Maintenance", "Finances", "Biens", "Propriétaires"]:
    protect_sheet(wb[name], insert_rows=True)
# Autres onglets : protégés, sans besoin d'insertion de ligne
for name in ["Tableau de bord", "Formulaire Réservation", "Alertes",
             "Facture Propriétaire", "Rapport Mensuel", "Paramètres", "Accueil"]:
    protect_sheet(wb[name], insert_rows=False)

wb.security = openpyxl.workbook.protection.WorkbookProtection(
    workbookPassword="MGHote2026", lockStructure=True
)

wb.properties.title = "MG Hôte — Gestion de Conciergerie"
wb.properties.creator = "MG Hôte"
wb.properties.subject = "Conciergerie Airbnb & Booking — Villas, Appartements, Riads"
wb.properties.description = "Classeur de gestion MG Hôte : réservations, finances, ménage, maintenance, alertes, factures."

wb.active = sheet_order.index("Tableau de bord")

OUT_PATH = "MG_Hote_Gestion_Conciergerie.xlsx"
wb.save(OUT_PATH)
print(f"Classeur enregistré : {OUT_PATH}")
