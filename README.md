# EOD Zootopia : Impact du Climat sur la Gestion Hydraulique

Optimisation économique du dispatching électrique avec analyse de trois scénarios climatiques (sèche, normale, humide) sur le système électrique zootopien.

## 📁 Structure
```
eod-zootopia/
├── README.md
├── .gitignore
├── code/
│   └── EOD_Zootopia.jl
├── data/
│   └── Donnees_etude_de_cas_ETE305.xlsx
├── notebook/
│   └── Notebook_EOD_Results.ipynb
├── results/
│   └── .gitkeep
└── requirements/
    ├── requirements.jl
    ├── requirements.txt
    └── environment.yml
```

## 🚀 Utilisation

### Installation
```bash
git clone https://github.com/TonUser/eod-zootopia.git
cd eod-zootopia

# Julia
julia -e "using Pkg; Pkg.activate(\".\"); include(\"requirements/requirements.jl\")"

# Python (optionnel)
pip install -r requirements/requirements.txt
```

### Exécution
```bash
cd code
julia EOD_Zootopia.jl
# Génère results/results_dry.csv, results_normal.csv, results_wet.csv
```

### Analyse
```bash
cd ../notebook
jupyter notebook Notebook_EOD_Results.ipynb
```

## 📊 Méthodologie

- **Type** : MILP (Mixed-Integer Linear Program)
- **Horizon** : 8760 heures (1 année)
- **Fenêtre roulante** : 9 jours (7 optimisés + 2 raccordement)
- **Solveur** : HiGHS (primal-dual révisé)
- **Scénarios** : DRY (-40%), NORMAL, WET (+40%)

## 🔑 Modèle

✅ Min/max stocks hydrauliques saisonniers  
✅ Écrêtage hydraulique été (contrainte agricole)  
✅ Dynamique STEP (rendement 75%)  
✅ 23 générateurs (nucléaire, charbon, gaz, hydro, ENR, bioénergies)

---

**Février 2025** | Projet EOD - Sujet 1 : Modélisation de l'Hydraulique
```