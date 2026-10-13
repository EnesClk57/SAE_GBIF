# SAE_GBIF — Biodiversité et émissions de GES dans le Grand Est

Projet réalisé par **Eser Enes** et **Marfouk Omar** dans le cadre de la SAE 501. Il explore les observations d'oiseaux du Grand Est recensées par GBIF entre 2018 et 2022 et les met en regard des émissions de gaz à effet de serre (GES) à l'échelle des établissements publics de coopération intercommunale (EPCI).

Le traitement est effectué avec Python. Les résultats comprennent des indicateurs de biodiversité, des fichiers CSV et un classeur Excel généré par le script. Le dépôt contient également un tableau de bord Excel avec macros.

## Objectifs

- Nettoyer et analyser les observations d'oiseaux sur la période 2018–2022.
- Associer géographiquement les observations aux territoires EPCI.
- Calculer des indicateurs par EPCI et par année.
- Croiser ces indicateurs avec les données annuelles d'émissions de GES.
- Produire des fichiers de résultats et des visualisations dans Excel.

Les corrélations présentées servent à explorer les données : elles ne permettent pas, à elles seules, de conclure à un lien de causalité entre émissions de GES et biodiversité.

## Indicateurs calculés

| Indicateur | Description |
| --- | --- |
| Richesse spécifique | Nombre d'espèces différentes observées |
| Effort d'observation | Nombre d'observations enregistrées |
| Abondance totale | Somme des individus observés |
| Abondance moyenne | Nombre moyen d'individus par observation |
| Richesse normalisée | Richesse spécifique rapportée à l'effort d'observation, pour 1 000 observations |

## Organisation du dépôt

```text
.
├── README.md
├── dashboard/
│   ├── Dashboard_Biodiversite.xlsm
│   └── module_vba.bas
├── data/
│   ├── export/                         # Résultats CSV déjà exportés
│   ├── oiseaux_grand-est_2018-2022_GBIF.csv
│   ├── ges_grand_est_epci_2018.csv
│   ├── ...
│   └── ges_grand_est_epci_2022.shp    # Limites géographiques EPCI
├── notebooks/
│   └── enes_omar_traitement_donnees.ipynb
├── scripts/
│   └── script.py
└── Guide.md
```

Le jeu d'observations GBIF, volumineux (environ 741 Mo), est stocké avec **Git LFS**. Les limites EPCI sont fournies sous forme de shapefile et utilisent plusieurs fichiers associés (`.shp`, `.dbf`, `.shx`, `.prj` et, selon le jeu, `.cpg`).

## Prérequis

- Python 3
- Les bibliothèques Python `pandas`, `geopandas`, `numpy` et `openpyxl`
- Microsoft Excel 2019 ou une version ultérieure pour ouvrir les tableaux de bord
- Git LFS pour récupérer le jeu de données GBIF lors du clonage

## Installation et exécution

### 1. Cloner le dépôt et récupérer les données

Le script Python et le notebook utilisent actuellement le chemin fixe `C:\SAE_GBIF` pour trouver les fichiers. Pour les exécuter sans modifier leurs chemins, clonez le dépôt à cet emplacement :

```powershell
git lfs install
git clone https://github.com/EnesClk57/SAE_GBIF.git C:\SAE_GBIF
cd C:\SAE_GBIF
git lfs pull
```

Si le dépôt est installé dans un autre dossier, adaptez le chemin `BASE_PATH` dans `scripts/script.py`. Les cellules du notebook qui lisent des fichiers utilisent également des chemins `C:\SAE_GBIF` ; adaptez-les au besoin.

### 2. Installer les dépendances

Dans PowerShell, depuis le dossier du projet :

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install pandas geopandas numpy openpyxl
```

### 3. Lancer le traitement

```powershell
python scripts\script.py
```

Le script nettoie les données, effectue une jointure spatiale avec les EPCI, calcule les indicateurs, réalise le croisement avec les fichiers GES, puis exporte les résultats.

## Résultats générés

Le script écrit les fichiers suivants dans `data/export/` :

- `indicateurs_epci_annee.csv` — indicateurs de biodiversité par EPCI et par année ;
- `croisement_biodiv_ges.csv` — indicateurs de biodiversité croisés avec les émissions de GES ;
- `liste_especes.csv` — observations et abondance totale par espèce.

Il génère aussi `dashboard/Dashboard_Biodiversite.xlsx`, avec des feuilles consacrées aux indicateurs par EPCI, aux principales espèces observées, aux observations par année et au croisement avec les GES.

## Tableaux de bord Excel

- `dashboard/Dashboard_Biodiversite.xlsx` est le classeur généré par `scripts/script.py`. Il est recréé lors de chaque exécution du script.
- `dashboard/Dashboard_Biodiversite.xlsm` est un classeur distinct, déjà présent dans le dépôt. Il contient des macros : ouvrez-le avec Excel et activez les macros si vous souhaitez utiliser ses fonctionnalités interactives.

## Notebook

Le notebook [`notebooks/enes_omar_traitement_donnees.ipynb`](notebooks/enes_omar_traitement_donnees.ipynb) présente des étapes d'exploration et de traitement des données. Ouvrez-le dans Jupyter ou Visual Studio Code après avoir installé les dépendances. Comme le script, certaines cellules s'attendent à trouver les données sous `C:\SAE_GBIF`.

## Documentation complémentaire

Consultez [`Guide.md`](Guide.md) pour retrouver les informations sur le tableau de bord Excel et ses indicateurs.
