# requirements.jl - Dépendances Julia
# Exécute avec : julia -e "include(\"requirements/requirements.jl\")"

using Pkg

# Ajoute les dépendances nécessaires
Pkg.add([
    "JuMP",        # Modélisation optimisation
    "HiGHS",       # Solveur MILP
    "XLSX",        # Lecture fichiers Excel
    "DataFrames",  # Manipulation données
    "CSV",         # Lecture/écriture CSV
    "Plots",       # Visualisation
    "Statistics",  # Statistiques
    "Dates",       # Gestion dates
])

println("✅ Dépendances Julia installées !")