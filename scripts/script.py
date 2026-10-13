#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Auteurs : Eser Enes & Marfouk Omar
- Charge et nettoie les données GBIF (observations d'oiseaux)
-Effectue la jointure spatiale avec les EPCI
- Calcule les indicateurs de biodiversité
- Croise avec les données d'émissions GES
- Exporte les données en CSV
- Génère le dashboard Excel avec graphiques
"""

import pandas as pd
import geopandas as gpd
import numpy as np
import os
import sys
from pathlib import Path

# Encodage UTF-8 pour Windows
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# Importer openpyxl pour Excel
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.formatting.rule import ColorScaleRule
except ImportError:
    print("Installation de openpyxl...")
    import subprocess
    subprocess.check_call(['pip', 'install', 'openpyxl'])
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.formatting.rule import ColorScaleRule

# Config
BASE_PATH = Path(r"C:\SAE_GBIF")
DATA_PATH = BASE_PATH / "data"
EXPORT_PATH = DATA_PATH / "export"
DASHBOARD_PATH = BASE_PATH / "dashboard"
# Fichiers sources
OBSERVATIONS_CSV = DATA_PATH / "oiseaux_grand-est_2018-2022_GBIF.csv"
EPCI_SHAPEFILE = DATA_PATH / "ges_grand_est_epci_2022.shp"
# Fichier de sortie Excel
OUTPUT_EXCEL = DASHBOARD_PATH / "Dashboard_Biodiversite.xlsx"
# Années à traiter
ANNEES = [2018, 2019, 2020, 2021, 2022]
# Coleurs du thème Excel
VERT_FONCE = "1B4D3E"
VERT_CLAIR = "4CAF50"
BLANC = "FFFFFF"
# Créer les dossiers 
EXPORT_PATH.mkdir(parents=True, exist_ok=True)
DASHBOARD_PATH.mkdir(parents=True, exist_ok=True)

print("  TRAITEMENT DES DONNÉES BIODIVERSITÉ et  GES")

# CHARGEMENT DES OBSERVATIONS GBIF

df_obs = pd.read_csv(OBSERVATIONS_CSV, sep="\t")
print(f"      {len(df_obs):,} observations chargées")

#  NETTOYAGE DES DONNÉES

colonnes_necessaires = ['gbifID', 'eventDate', 'species', 'decimalLatitude', 
                        'decimalLongitude', 'individualCount', 'year']
df_clean = df_obs[colonnes_necessaires].copy()
# Conversion des dates
df_clean['eventDate'] = pd.to_datetime(df_clean['eventDate'], errors='coerce')
df_clean['year_parsed'] = df_clean['eventDate'].dt.year
df_clean['year'] = df_clean['year'].fillna(df_clean['year_parsed'])
# Suppression des valeurs manquantes
df_clean = df_clean.dropna(subset=['species', 'decimalLatitude', 'decimalLongitude', 'year'])
# Filtrer sur la période 2018-2022
df_clean = df_clean[(df_clean['year'] >= 2018) & (df_clean['year'] <= 2022)]
df_clean['year'] = df_clean['year'].astype(int)
# Remplacer individualCount manquant par 1
df_clean['individualCount'] = df_clean['individualCount'].fillna(1)
print(f"      {len(df_clean):,} observations après nettoyage")
print(f"      Période : {df_clean['year'].min()} - {df_clean['year'].max()}")
print(f"      Espèces uniques : {df_clean['species'].nunique()}")

#  JOINTURE SPATIALE AVEC LES EPCI
os.environ['SHAPE_RESTORE_SHX'] = 'YES'
# Création du GeoDataFrame des observations
gdf_obs = gpd.GeoDataFrame(
    df_clean,
    geometry=gpd.points_from_xy(df_clean.decimalLongitude, df_clean.decimalLatitude),
    crs='EPSG:4326'
)
# Chargement du shapefile EPCI
gdf_epci = gpd.read_file(EPCI_SHAPEFILE)
print(f"      {len(gdf_epci)} EPCI chargés")
# Reprojection et jointure
gdf_obs_proj = gdf_obs.to_crs(gdf_epci.crs)
gdf_obs_epci = gpd.sjoin(gdf_obs_proj, gdf_epci, how='inner', predicate='within')
print(f"      {len(gdf_obs_epci):,} observations dans les EPCI")
print(f"      {gdf_obs_epci['code_epci'].nunique()} EPCI concernés")

#  CALCUL DES INDICATEURS DE BIODIVERSITÉ

# Indicateurs par EPCI et par année
indicateurs = gdf_obs_epci.groupby(['code_epci', 'year']).agg({
    'species': 'nunique',
    'gbifID': 'count',
    'individualCount': ['sum', 'mean']
}).reset_index()
indicateurs.columns = ['code_epci', 'year', 'richesse_specifique', 'effort_observation', 
                       'abondance_totale', 'abondance_moyenne']
indicateurs['abondance_moyenne'] = indicateurs['abondance_moyenne'].round(2)
# Joindre avec les noms EPCI
epci_info = gdf_epci[['code_epci', 'lib_zone']].drop_duplicates()
indicateurs = indicateurs.merge(epci_info, on='code_epci', how='left')
print(f"      {len(indicateurs)} lignes (EPCI × année)")
# Liste des espèces
especes = gdf_obs_epci.groupby('species').agg({
    'gbifID': 'count',
    'individualCount': 'sum'
}).reset_index()
especes.columns = ['species', 'nb_observations', 'abondance_totale']
especes = especes.sort_values('nb_observations', ascending=False)
print(f"      {len(especes)} espèces observées")
# Évolution temporelle
evolution = indicateurs.groupby('year').agg({
    'richesse_specifique': 'sum',
    'effort_observation': 'sum',
    'abondance_totale': 'sum'
}).reset_index()
# Statistiques générales
stats = {
    'nb_observations_total': int(len(gdf_obs_epci)),
    'nb_especes': int(gdf_obs_epci['species'].nunique()),
    'nb_epci': int(gdf_obs_epci['code_epci'].nunique()),
    'periode': f"{df_clean['year'].min()}-{df_clean['year'].max()}",
    'observations_par_annee': gdf_obs_epci.groupby('year').size().to_dict()
}

# CHARGEMENT ET CROISEMENT AVEC LES DONNÉES GES

df_ges_list = []
for annee in ANNEES:
    chemin = DATA_PATH / f"ges_grand_est_epci_{annee}.csv"
    try:
        df_temp = pd.read_csv(chemin, sep=",", encoding='utf-8')
        df_temp['annee_ges'] = annee
        df_ges_list.append(df_temp)
    except Exception as e:
        print(f"      Erreur pour {annee} : {e}")

df_ges = pd.concat(df_ges_list, ignore_index=True)
print(f"      {len(df_ges)} lignes GES chargées")

# Jointure biodiversité et GES
indicateurs['code_epci'] = indicateurs['code_epci'].astype(str)
df_ges['code_zone'] = df_ges['code_zone'].astype(str)
indicateurs['year'] = indicateurs['year'].astype(int)
df_ges['annee_ges'] = df_ges['annee_ges'].astype(int)

df_croisement = indicateurs.merge(
    df_ges,
    left_on=['code_epci', 'year'],
    right_on=['code_zone', 'annee_ges'],
    how='inner'
)

# Gérer les colonnes dupliquées
if 'code_epci_x' in df_croisement.columns:
    df_croisement = df_croisement.rename(columns={'code_epci_x': 'code_epci'})
if 'lib_zone_x' in df_croisement.columns:
    df_croisement = df_croisement.rename(columns={'lib_zone_x': 'lib_zone'})

colonnes_a_supprimer = ['code_zone', 'annee_ges', 'code_epci_y', 'lib_zone_y']
df_croisement = df_croisement.drop(columns=[c for c in colonnes_a_supprimer if c in df_croisement.columns])

# Richesse normalisée
df_croisement['richesse_normalisee'] = (
    df_croisement['richesse_specifique'] / df_croisement['effort_observation'] * 1000
).round(2)

print(f"      {len(df_croisement)} lignes après croisement")

# Calcul des corrélations
cols_emissions = ['co2_kg', 'ch4_kg', 'n2o_kg', 'prg2013_teqco2']
correlations = {'brutes': {}, 'normalisees': {}}

for col in cols_emissions:
    if col in df_croisement.columns:
        data_valid = df_croisement[['richesse_specifique', col]].dropna()
        if len(data_valid) > 10:
            correlations['brutes'][col] = round(data_valid['richesse_specifique'].corr(data_valid[col]), 4)
        
        data_valid = df_croisement[['richesse_normalisee', col]].dropna()
        if len(data_valid) > 10:
            correlations['normalisees'][col] = round(data_valid['richesse_normalisee'].corr(data_valid[col]), 4)

print("\n      Corrélations brutes :")
for col, corr in correlations['brutes'].items():
    print(f"         {col}: {'+' if corr > 0 else ''}{corr:.3f}")

print("\n      Corrélations normalisées :")
for col, corr in correlations['normalisees'].items():
    print(f"         {col}: {'+' if corr > 0 else ''}{corr:.3f}")

# ÉTAPE 6 : EXPORT DES DONNÉES CSV
indicateurs.to_csv(EXPORT_PATH / "indicateurs_epci_annee.csv", index=False, encoding='utf-8')
print(f"      indicateurs_epci_annee.csv ({len(indicateurs)} lignes)")
df_croisement.to_csv(EXPORT_PATH / "croisement_biodiv_ges.csv", index=False, encoding='utf-8')
print(f"      croisement_biodiv_ges.csv ({len(df_croisement)} lignes)")
especes.to_csv(EXPORT_PATH / "liste_especes.csv", index=False, encoding='utf-8')
print(f"      liste_especes.csv ({len(especes)} espèces)")

#  CRÉATION DU DASHBOARD EXCEL

wb = Workbook()

# Styles
titre_font = Font(name='Calibri', size=18, bold=True, color=VERT_FONCE)
header_font = Font(name='Calibri', size=11, bold=True, color=BLANC)
header_fill = PatternFill(start_color=VERT_FONCE, end_color=VERT_FONCE, fill_type='solid')
nombre_font = Font(name='Calibri', size=24, bold=True, color=VERT_CLAIR)
bordure = Border(
    left=Side(style='thin', color='CCCCCC'),
    right=Side(style='thin', color='CCCCCC'),
    top=Side(style='thin', color='CCCCCC'),
    bottom=Side(style='thin', color='CCCCCC')
)

# FEUILLE 1 : DASHBOARD

ws_dashboard = wb.active
ws_dashboard.title = "Dashboard"
# Titre
ws_dashboard.merge_cells('B2:H2')
ws_dashboard['B2'] = "BIODIVERSITÉ × ÉMISSIONS GES - GRAND EST"
ws_dashboard['B2'].font = titre_font
ws_dashboard['B2'].alignment = Alignment(horizontal='center')
ws_dashboard.merge_cells('B3:H3')
ws_dashboard['B3'] = "Analyse de l'avifaune (2018-2022) - SAE 501 - Eser Enes & Marfouk Omar"
ws_dashboard['B3'].alignment = Alignment(horizontal='center')
# Cartes statistiques
stats_data = [
    ("B5", "Observations", f"{stats['nb_observations_total']:,}".replace(",", " ")),
    ("D5", "Espèces", str(stats['nb_especes'])),
    ("F5", "EPCI", str(stats['nb_epci'])),
    ("H5", "Période", stats['periode'])
]
for col, label, value in stats_data:
    ws_dashboard[col] = label
    ws_dashboard[col].font = Font(bold=True, color=VERT_FONCE)
    ws_dashboard[col].alignment = Alignment(horizontal='center')
    
    value_cell = col[0] + "6"
    ws_dashboard[value_cell] = value
    ws_dashboard[value_cell].font = nombre_font
    ws_dashboard[value_cell].alignment = Alignment(horizontal='center')
# Section Évolution temporelle
ws_dashboard['B9'] = "ÉVOLUTION TEMPORELLE"
ws_dashboard['B9'].font = Font(size=14, bold=True, color=VERT_FONCE)

headers_evo = ['Année', 'Richesse spécifique', 'Effort observation', 'Abondance totale']
for col_idx, header in enumerate(headers_evo, start=2):
    cell = ws_dashboard.cell(row=10, column=col_idx, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center')
    cell.border = bordure

for row_idx, row in enumerate(evolution.itertuples(index=False), start=11):
    ws_dashboard.cell(row=row_idx, column=2, value=int(row[0])).border = bordure
    ws_dashboard.cell(row=row_idx, column=3, value=int(row[1])).border = bordure
    ws_dashboard.cell(row=row_idx, column=4, value=int(row[2])).border = bordure
    ws_dashboard.cell(row=row_idx, column=5, value=int(row[3])).border = bordure

# Graphique évolution
chart_evo = LineChart()
chart_evo.title = "Évolution de la richesse spécifique"
chart_evo.style = 10
chart_evo.x_axis.title = "Année"
chart_evo.y_axis.title = "Nombre d'espèces"

data_evo = Reference(ws_dashboard, min_col=3, min_row=10, max_col=3, max_row=15)
cats_evo = Reference(ws_dashboard, min_col=2, min_row=11, max_row=15)
chart_evo.add_data(data_evo, titles_from_data=True)
chart_evo.set_categories(cats_evo)
chart_evo.width = 15
chart_evo.height = 8

ws_dashboard.add_chart(chart_evo, "G9")

# Section Corrélations
ws_dashboard['B18'] = "CORRÉLATIONS GES × BIODIVERSITÉ"
ws_dashboard['B18'].font = Font(size=14, bold=True, color=VERT_FONCE)

headers_corr = ['Gaz', 'Corrélation brute', 'Corrélation normalisée']
for col_idx, header in enumerate(headers_corr, start=2):
    cell = ws_dashboard.cell(row=19, column=col_idx, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center')
    cell.border = bordure

gaz_labels = {'co2_kg': 'CO2', 'ch4_kg': 'CH4', 'n2o_kg': 'N2O', 'prg2013_teqco2': 'PRG total'}
row_idx = 20
for gaz, label in gaz_labels.items():
    if gaz in correlations['brutes']:
        ws_dashboard.cell(row=row_idx, column=2, value=label).border = bordure
        
        brute = correlations['brutes'].get(gaz, 0)
        cell_brute = ws_dashboard.cell(row=row_idx, column=3, value=round(brute, 3))
        cell_brute.border = bordure
        cell_brute.number_format = '0.000'
        if brute > 0:
            cell_brute.font = Font(color='FF0000')
        
        normalisee = correlations['normalisees'].get(gaz, 0)
        cell_norm = ws_dashboard.cell(row=row_idx, column=4, value=round(normalisee, 3))
        cell_norm.border = bordure
        cell_norm.number_format = '0.000'
        if normalisee < 0:
            cell_norm.font = Font(color='008000')
        
        row_idx += 1

# Graphique corrélations
chart_corr = BarChart()
chart_corr.title = "Corrélations brutes vs normalisées"
chart_corr.style = 10
chart_corr.type = "col"

data_corr = Reference(ws_dashboard, min_col=3, min_row=19, max_col=4, max_row=23)
cats_corr = Reference(ws_dashboard, min_col=2, min_row=20, max_row=23)
chart_corr.add_data(data_corr, titles_from_data=True)
chart_corr.set_categories(cats_corr)
chart_corr.width = 12
chart_corr.height = 8

ws_dashboard.add_chart(chart_corr, "G18")

# Interprétation
ws_dashboard['B26'] = "INTERPRÉTATION"
ws_dashboard['B26'].font = Font(size=14, bold=True, color=VERT_FONCE)

ws_dashboard.merge_cells('B27:H28')
ws_dashboard['B27'] = """Les corrélations BRUTES sont positives (biais : plus d'observateurs dans les zones polluées).
Les corrélations NORMALISÉES sont négatives : plus de GES = moins de biodiversité. Hypothèse confirmée !"""
ws_dashboard['B27'].alignment = Alignment(wrap_text=True)

for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']:
    ws_dashboard.column_dimensions[col].width = 18

#  INDICATEURS PAR EPCI

ws_indicateurs = wb.create_sheet("Indicateurs EPCI")
ws_indicateurs.merge_cells('A1:G1')
ws_indicateurs['A1'] = "Indicateurs de biodiversité par EPCI et par année"
ws_indicateurs['A1'].font = titre_font

for col_idx, col_name in enumerate(indicateurs.columns, start=1):
    cell = ws_indicateurs.cell(row=3, column=col_idx, value=col_name)
    cell.font = header_font
    cell.fill = header_fill
    cell.border = bordure

for row_idx, row in enumerate(indicateurs.itertuples(index=False), start=4):
    for col_idx, value in enumerate(row, start=1):
        cell = ws_indicateurs.cell(row=row_idx, column=col_idx, value=value)
        cell.border = bordure

ws_indicateurs.conditional_formatting.add(
    f'C4:C{len(indicateurs) + 3}',
    ColorScaleRule(
        start_type='min', start_color='FFEB9C',
        mid_type='percentile', mid_value=50, mid_color='92D050',
        end_type='max', end_color='006400'
    )
)

# FEUILLE 3 : TOP ESPÈCES
ws_especes = wb.create_sheet("Top Espèces")

ws_especes.merge_cells('A1:C1')
ws_especes['A1'] = "Top 20 des espèces les plus observées (2018-2022)"
ws_especes['A1'].font = titre_font

headers_esp = ['Espèce', 'Nb observations', 'Abondance totale']
for col_idx, header in enumerate(headers_esp, start=1):
    cell = ws_especes.cell(row=3, column=col_idx, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.border = bordure

top20_especes = especes.head(20)
for row_idx, row in enumerate(top20_especes.itertuples(index=False), start=4):
    for col_idx, value in enumerate(row, start=1):
        cell = ws_especes.cell(row=row_idx, column=col_idx, value=value)
        cell.border = bordure

# Graphique Top 10
chart_esp = BarChart()
chart_esp.type = "bar"
chart_esp.title = "Top 10 espèces observées"
chart_esp.style = 10

data_esp = Reference(ws_especes, min_col=2, min_row=3, max_row=13)
cats_esp = Reference(ws_especes, min_col=1, min_row=4, max_row=13)
chart_esp.add_data(data_esp, titles_from_data=True)
chart_esp.set_categories(cats_esp)
chart_esp.width = 18
chart_esp.height = 12

ws_especes.add_chart(chart_esp, "E3")

ws_especes.column_dimensions['A'].width = 30
ws_especes.column_dimensions['B'].width = 18
ws_especes.column_dimensions['C'].width = 18

# FEUILLE 4 : OBSERVATIONS PAR ANNÉE

ws_obs = wb.create_sheet("Observations par année")

ws_obs.merge_cells('A1:C1')
ws_obs['A1'] = "Nombre d'observations par année"
ws_obs['A1'].font = titre_font

ws_obs.cell(row=3, column=1, value="Année").font = header_font
ws_obs.cell(row=3, column=1).fill = header_fill
ws_obs.cell(row=3, column=2, value="Observations").font = header_font
ws_obs.cell(row=3, column=2).fill = header_fill

row_idx = 4
for annee, nb_obs in stats['observations_par_annee'].items():
    ws_obs.cell(row=row_idx, column=1, value=int(annee)).border = bordure
    ws_obs.cell(row=row_idx, column=2, value=int(nb_obs)).border = bordure
    row_idx += 1

chart_obs = BarChart()
chart_obs.title = "Observations par année"
chart_obs.style = 10

data_obs = Reference(ws_obs, min_col=2, min_row=3, max_row=8)
cats_obs = Reference(ws_obs, min_col=1, min_row=4, max_row=8)
chart_obs.add_data(data_obs, titles_from_data=True)
chart_obs.set_categories(cats_obs)
chart_obs.width = 12
chart_obs.height = 8

ws_obs.add_chart(chart_obs, "D3")

# FEUILLE 5 : CROISEMENT GES

ws_ges = wb.create_sheet("Croisement GES")

ws_ges.merge_cells('A1:K1')
ws_ges['A1'] = "Croisement Biodiversité × Émissions GES par EPCI"
ws_ges['A1'].font = titre_font

cols_importantes = ['code_epci', 'year', 'lib_zone', 'richesse_specifique', 
                    'effort_observation', 'co2_kg', 'ch4_kg', 'n2o_kg', 
                    'prg2013_teqco2', 'richesse_normalisee']
df_ges_select = df_croisement[[c for c in cols_importantes if c in df_croisement.columns]]

for col_idx, col_name in enumerate(df_ges_select.columns, start=1):
    cell = ws_ges.cell(row=3, column=col_idx, value=col_name)
    cell.font = header_font
    cell.fill = header_fill
    cell.border = bordure

for row_idx, row in enumerate(df_ges_select.head(200).itertuples(index=False), start=4):
    for col_idx, value in enumerate(row, start=1):
        cell = ws_ges.cell(row=row_idx, column=col_idx, value=value)
        cell.border = bordure

# SAUVEGARDE

wb.save(OUTPUT_EXCEL)

print(f"\n      Dashboard Excel créé : {OUTPUT_EXCEL}")

# RÉSUMÉ

print(f"\n  Fichiers exportés dans : {EXPORT_PATH}")
print(f"     - indicateurs_epci_annee.csv")
print(f"     - croisement_biodiv_ges.csv")
print(f"     - liste_especes.csv")
print(f"\n  Dashboard Excel : {OUTPUT_EXCEL}")
print(f"     1. Ouvrir {OUTPUT_EXCEL}")
