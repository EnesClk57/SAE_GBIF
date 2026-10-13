**Auteurs** : Eser Enes & Marfouk Omar  
---
**Ouvrir directement le dashboard Excel**
 Ouvrir le fichier `dashboard/Dashboard_Biodiversite.xlsm`

notebooks/
    └── enes_omar_traitement_donnees.ipynb
dashboard/
  ├── Dashboard_Biodiversite.xlsm    
  └── module_vba.bas 
scripts/
│   └── script.py 

##  Logiciels utilisés

| Logiciel | Version | Utilisation |
|----------|---------|-------------|
| Python | 3.x | Traitement des données |
| Microsoft Excel | 2019+ | Dashboard interactif (VBA) |
| Visual Studio Code | - | Développement |

### Bibliothèques Python requises

```
pandas
geopandas
openpyxl
```

---

##  Étapes pour lancer la visualisation

### Option 1 : Ouvrir directement le dashboard Excel 

1. Ouvrir le fichier `dashboard/Dashboard_Biodiversite.xlsm`
2. **Activer les macros** quand Excel le demande
3. Utiliser les boutons interactifs sur la feuille "Dashboard"

## Fonctionnalités du Dashboard

### Feuilles disponibles

| Feuille | Contenu |
|---------|---------|
| Dashboard | Vue principale avec boutons |
| Indicateurs EPCI | Richesse spécifique, abondance par EPCI et année |
| Top Espèces | Classement des espèces les plus observées |
| Observations par année | Évolution temporelle 2018-2022 |
| Croisement GES | Corrélation biodiversité × émissions |
| Analyse TCD | Tableau croisé dynamique |

---

## Indicateurs calculés

| Indicateur | Description |
|------------|-------------|
| **Richesse spécifique** | Nombre d'espèces différentes observées |
| **Abondance totale** | Nombre total d'individus observés |
| **Effort d'observation** | Nombre d'observations réalisées |
| **Abondance moyenne** | Nombre moyen d'individus par observation |

---
##  Remarques

- **Excel avec macros** : Le fichier `.xlsm` nécessite d'activer les macros pour fonctionner

---
