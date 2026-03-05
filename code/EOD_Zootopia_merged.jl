# ==============================================================================
# PROJET EOD - PARC ÉLECTRIQUE ZOOTOPIA (VERSION FUSIONNÉE)
#
# Optimisation économique du dispatching électrique avec :
# - Fenêtre roulante 9 jours (7+2)
# - Gestion hydraulique saisonnière avec contraintes de stock et d'écrêtage
# - Gestion STEP avec rendement 75%
# - 3 scénarios météo (dry/normal/wet)
# - [IMPROVED] Arguments en ligne de commande pour scénarios
# - [IMPROVED] Fonctions de conversion robustes pour lecture Excel
# - [IMPROVED] Tranches d'eau (socle/surplus) pour meilleure gestion hydraulique
# - [IMPROVED] Fonction helper add_uc_constraints!
# - [IMPROVED] Export enrichi (infeasible_flag, S_socle, S_surplus, Phy_lac)
# - [IMPROVED] Stats détaillées par scénario
#
# DIFFÉRENCES VS ORIGINAL (régressions corrigées) :
#   ✅ COST_UNSUPPLIED restauré à 1 milliard €/MWh (was 10 000 dans improved)
#   ✅ STEP_CAPACITY restauré à 1 TWh (was 0.5 TWh dans improved)
#   ✅ Phy_fdl forcé (== min(apport, cap), was optionnel dans improved)
#   ✅ WATER_PRICE_MONTHLY restauré (supprimé dans improved)
#   ✅ HYDRO_LAC_ECRÊTAGE_MONTHLY restauré (supprimé dans improved)
#   ✅ Pénalité slack stock restaurée à 500M €/MWh (was 2 000 dans improved)
#   ✅ Pcharge_STEP dans dynamique lac restauré (supprimé dans improved)
#   ✅ recale_hydro_values restaurée (manquante dans improved → crash)
#   ✅ ENR must-take (== wind[t]) restauré pour cohérence physique
#
# UTILISATION :
#   julia EOD_Zootopia_merged.jl                    # tous scénarios
#   julia EOD_Zootopia_merged.jl --scenario dry     # scénario sec seulement
#   julia EOD_Zootopia_merged.jl --scenario normal  # scénario normal seulement
#   julia EOD_Zootopia_merged.jl --scenario wet     # scénario humide seulement
#   julia EOD_Zootopia_merged.jl --data chemin/fichier.xlsx
# ==============================================================================

using JuMP
using HiGHS
using XLSX
using Dates
using Statistics
using Random
using Plots
using DataFrames
using CSV

# ==============================================================================
# SECTION 0 : PARSING D'ARGUMENTS [IMPROVED]
# ==============================================================================

function parse_args(args)
    """Parse les arguments de la ligne de commande"""
    params = Dict(
        "scenario" => "all",
        "data_file" => ".//data//Donnees_etude_de_cas_ETE305.xlsx"
    )

    i = 1
    while i <= length(args)
        if i < length(args)
            key = args[i]
            val = args[i+1]
            if key == "--scenario"
                if val in ["dry", "normal", "wet", "all"]
                    params["scenario"] = val
                    i += 2
                else
                    println("⚠️  Scénario invalide: $val (utiliser dry/normal/wet/all)")
                    i += 2
                end
            elseif key == "--data"
                params["data_file"] = val
                i += 2
            else
                i += 1
            end
        else
            i += 1
        end
    end

    return params
end

# ==============================================================================
# SECTION 1 : STRUCTURES DE DONNÉES
# ==============================================================================

struct UniteCentrale
    id::String
    centrale::String
    type::String
    Pmax::Float64
    Pmin::Float64
    dmin::Int
    prix_marche::Float64
    annee_lancement::Int
end

struct StockHydro
    date::String
    lev_low::Float64
    lev_high::Float64
end

struct ApportMensuel
    mois::String
    apport::Int
end

struct ConsommationHoraire
    date::String
    heure::String
    fil_de_leau::Int
    lacs::Int
    step::Int
    conso_charge::Int
    eolien_cons::Int
    solaire_cons::Int
end

# ==============================================================================
# SECTION 2 : FONCTIONS DE CONVERSION ROBUSTES [IMPROVED]
# ==============================================================================

"""Convertir une valeur en Float64, gérer les cas problématiques"""
function _val_f64(val)
    if ismissing(val) || val === nothing
        return 0.0
    elseif val isa Bool
        return Float64(val)
    elseif val isa Number
        return Float64(val)
    else
        str_val = replace(string(val), "," => ".")
        return parse(Float64, strip(str_val))
    end
end

"""Convertir un vecteur en Vector{Float64}"""
function to_f64(x)
    return [_val_f64(v) for v in vec(x)]
end

"""Convertir une valeur en Int, gérer les cas problématiques"""
function _val_int(val)
    if ismissing(val) || val === nothing
        return 0
    end
    if val isa Bool
        return Int(val)
    end
    if val isa Number
        return Int(round(Float64(val)))
    end
    return parse(Int, strip(string(val)))
end

"""Convertir un vecteur en Vector{Int}"""
function to_int(x)
    return [_val_int(v) for v in vec(x)]
end

"""Extraire un scalaire Float64 d'une cellule XLSX"""
function scalar_f64(x)
    val = isa(x, AbstractArray) ? x[1] : x
    if ismissing(val) || val === nothing
        return 0.0
    end
    if val isa Bool
        return Float64(val)
    end
    if val isa Number
        return Float64(val)
    end
    str_val = replace(string(val), "," => ".")
    return parse(Float64, strip(str_val))
end

# ==============================================================================
# SECTION 3 : DONNÉES DES CENTRALES
# ==============================================================================

function initialize_unites()
    """Initialiser la liste des unités de production"""
    unites = UniteCentrale[]

    centrales_data = [
        # Nom, Type, Capacité totale (MW), Nb unités, Pmax/unité (MW), Pmin/unité (MW), Dmin (h), Prix marché (€/MWh), Année
        ("Iconuc",       "Nucléaire",          1800, 2, 900,  300, 24, 12,  1977),
        ("Tabarnuc",     "Nucléaire",          1800, 2, 900,  300, 24, 12,  1981),
        ("NucPlusUltra", "Nucléaire",          2600, 2, 1300, 330, 24, 12,  1987),
        ("Gazby",        "CCG gaz",             390, 1, 390,  135,  4, 40,  1994),
        ("Pégaz",        "CCG gaz",             788, 2, 394,  137,  4, 40,  2000),
        ("Samagaz",      "CCG gaz",             430, 1, 430,  151,  4, 40,  2005),
        ("Omaïgaz",      "CCG gaz",             860, 2, 430,  151,  4, 40,  2009),
        ("Gastafiore",   "CCG gaz",             430, 1, 430,  151,  4, 40,  2015),
        ("Igaznodon",    "TAC gaz",             170, 2,  85,   34,  1, 70,  2004),
        ("Cogénération", "cogénération gaz",    882, 1, 882,    0,  0, 70,  2000),
        ("Coron",        "charbon",            1200, 2, 600,  210,  8, 36,  1997),
        ("Mockingjay",   "charbon",             600, 1, 600,  210,  8, 36,  1990),
        ("Lantier",      "charbon",             600, 1, 600,  210,  8, 36,  1984),
        ("Déchets",      "déchets",              60, 1,  60,   60,  0,  0,  2000),
        ("Biomasse",     "petite biomasse",     365, 1, 365,    0,  0,  0,  2000),
        ("Tacotac",      "fioul",                65, 1,  65,   20,  1, 100, 1990),
        ("TicEtTac",     "fioul",                95, 1,  95,   30,  1, 100, 2005),
        ("HydroFDE",     "hydraulique fil de l'eau", 1000, 1, 1000, 0, 0, 0, 1980),
        ("HydroLac",     "hydraulique lac",    6000, 1, 6000,   0,  0,  0,  1980),
        ("Polochon",     "STEP",               1200, 1, 1200,   0,  0,  0,  2000),
        ("Eolien",       "éolien",             5900, 1, 5900,   0,  0,  0,  2017),
        ("Zéphyr",       "éolien",              500, 1,  500,   0,  0,  0,  2018),
        ("Solaire",      "solaire",            3000, 1, 3000,   0,  0,  0,  2019)
    ]

    for (c_name, c_type, cap_installee, nb_unites, Pmax_unite, Pmin_unite, dmin_val, prix_marche_val, annee_lancement) in centrales_data
        for u in 1:nb_unites
            unit_id = "$(c_name)_$u"
            push!(unites, UniteCentrale(
                unit_id, c_name, c_type, Float64(Pmax_unite), Float64(Pmin_unite),
                dmin_val, Float64(prix_marche_val), annee_lancement
            ))
        end
    end

    return unites
end

# ==============================================================================
# SECTION 4 : LECTURE DES DONNÉES EXCEL [IMPROVED - fonctions robustes]
# ==============================================================================

function load_excel_data(data_file::String)
    """Charger les données depuis le fichier Excel avec conversions robustes"""

    println("📂 Chargement des données depuis: $data_file")

    # Stock hydraulique
    sheet_stock_hydro = "Stock hydro"
    lev_low  = XLSX.readdata(data_file, sheet_stock_hydro, "B4:B368")
    lev_high = XLSX.readdata(data_file, sheet_stock_hydro, "C4:C368")
    dates    = XLSX.readdata(data_file, sheet_stock_hydro, "A4:A368")

    # Conversions robustes [IMPROVED]
    lev_low_clean  = to_f64(lev_low)
    lev_high_clean = to_f64(lev_high)

    stocks_hydro = StockHydro[]
    for i in eachindex(dates)
        date = string(dates[i])
        low  = lev_low_clean[i]
        high = lev_high_clean[i]
        if !ismissing(low) && !ismissing(high) && low != 0.0 && high != 0.0
            push!(stocks_hydro, StockHydro(date, low, high))
        end
    end

    # Apports mensuels et consommations horaires
    sheet_details = "Détails historique hydro"
    mois    = XLSX.readdata(data_file, sheet_details, "A2:A13")
    apports = XLSX.readdata(data_file, sheet_details, "B2:B13")

    apports_mensuels = ApportMensuel[]
    for i in eachindex(mois)
        push!(apports_mensuels, ApportMensuel(string(mois[i]), _val_int(apports[i])))
    end

    # Consommations horaires avec conversions robustes [IMPROVED]
    dates_cons       = XLSX.readdata(data_file, sheet_details, "K2:K8761")
    heures_cons      = XLSX.readdata(data_file, sheet_details, "L2:L8761")
    fil_de_leau_cons = to_int(XLSX.readdata(data_file, sheet_details, "M2:M8761"))
    lacs_cons        = to_int(XLSX.readdata(data_file, sheet_details, "N2:N8761"))
    step_cons        = to_int(XLSX.readdata(data_file, sheet_details, "O2:O8761"))
    conso_charge     = to_int(XLSX.readdata(data_file, sheet_details, "Q2:Q8761"))
    eolien_cons      = to_int(XLSX.readdata(data_file, sheet_details, "R2:R8761"))
    solaire_cons     = to_int(XLSX.readdata(data_file, sheet_details, "S2:S8761"))

    consommations_horaires = ConsommationHoraire[]
    for i in eachindex(dates_cons)
        push!(consommations_horaires, ConsommationHoraire(
            string(dates_cons[i]), string(heures_cons[i]),
            fil_de_leau_cons[i], lacs_cons[i], step_cons[i],
            conso_charge[i], eolien_cons[i], solaire_cons[i]
        ))
    end

    println("✅ Données chargées: $(length(stocks_hydro)) stocks hydro, $(length(apports_mensuels)) mois d'apports, $(length(consommations_horaires)) heures")

    return stocks_hydro, apports_mensuels, consommations_horaires
end

# ==============================================================================
# SECTION 5 : PARAMÈTRES GLOBAUX
# ==============================================================================

# Constantes de simulation
const HOURS_PER_DAY  = 24
const WINDOW_SIZE    = 9 * HOURS_PER_DAY   # 216 heures
const RESULTS_SIZE   = 7 * HOURS_PER_DAY   # 168 heures
const YEAR_HOURS     = 8760
const HOURS          = 1:YEAR_HOURS
const CONSO_ANNUELLE_TWH = 85.0
const HYDRO_SHIFT_DAYS   = 181
const HYDRO_SHIFT_HOURS  = HYDRO_SHIFT_DAYS * HOURS_PER_DAY

# Heures par mois
const HOURS_PER_MONTH = Dict(
    1 => (744, "Janvier"),  2 => (672, "Février"),  3 => (744, "Mars"),
    4 => (720, "Avril"),    5 => (744, "Mai"),       6 => (720, "Juin"),
    7 => (744, "Juillet"),  8 => (744, "Août"),      9 => (720, "Septembre"),
    10 => (744, "Octobre"), 11 => (720, "Novembre"), 12 => (744, "Décembre")
)

# Hydraulique
const HYDRO_CAPACITY_TWh       = 1.0
const HYDRO_CAPACITY_MWh       = HYDRO_CAPACITY_TWh * 1_000_000
const HYDRO_FDL_CAPACITY_MW    = 1000
const HYDRO_LAC_CAPACITY_MW    = 6000

# STEP [ORIGINAL - 1 TWh ; improved avait réduit à 0.5 TWh]
const STEP_CAPACITY_MWh = 1_000_000
const STEP_POWER_MW     = 1200
const STEP_EFFICIENCY   = 0.75

# Scénarios météo (biais sur ENR)
const METEO_SCENARIOS = Dict(
    "dry"    => (wind_bias=-0.05, solar_bias=+0.10),
    "normal" => (wind_bias=0.00,  solar_bias=0.00),
    "wet"    => (wind_bias=+0.05, solar_bias=-0.08)
)

const SEASONAL_STOCK_TARGETS = Dict(
    "target_1"  => (dry=0.50, normal=0.50, wet=0.50),  # h500  : hiver, forcer réserves
    "target_2"  => (dry=0.35, normal=0.35, wet=0.35),  # h1500 : fin hiver, tenir le stock
    "target_3"  => (dry=0.35, normal=0.35, wet=0.35),  # h2500 : début printemps
    "target_4"  => (dry=0.50, normal=0.50, wet=0.50),  # h3500 : fonte des neiges
    "target_45" => (dry=0.60, normal=0.60, wet=0.60),  # h4000 : pic été (WET bridé)
    "target_5"  => (dry=0.65, normal=0.65, wet=0.65),  # h4500 : pic été (WET bridé)
    "target_55" => (dry=0.70, normal=0.70, wet=0.70),  # h5000 : pic été (WET bridé)
    "target_6"  => (dry=0.80, normal=0.80, wet=0.80),  # h5500 : fin été (WET bridé)
    "target_7"  => (dry=0.80, normal=0.80, wet=0.80),  # h6500 : automne
    "target_8"  => (dry=0.80, normal=0.80, wet=0.80),  # h7500 : automne avancé
    "target_9"  => (dry=0.50, normal=0.50, wet=0.50),  # h8500 : préparer hiver
)

const SEASONAL_WINDOWS = Dict(
    "target_1"  => (start_hour=416,  end_hour=584),
    "target_2"  => (start_hour=1416, end_hour=1584),
    "target_3"  => (start_hour=2416, end_hour=2584),
    "target_4"  => (start_hour=3416, end_hour=3584),
    "target_45" => (start_hour=3916, end_hour=4084),
    "target_5"  => (start_hour=4416, end_hour=4584),
    "target_55" => (start_hour=4916, end_hour=5084),
    "target_6"  => (start_hour=5416, end_hour=5584),
    "target_7"  => (start_hour=6416, end_hour=6584),
    "target_8"  => (start_hour=7416, end_hour=7584),
    "target_9"  => (start_hour=8416, end_hour=8584),
)

const HYDRO_LAC_ECRÊTAGE_MONTHLY = Dict(
    1 => 1.00, 2 => 1.00, 3 => 1.00, 4 => 1.00,
    5 => 1.00, 6 => 1.00, 7 => 0.70,
    8 => 0.60, 9 => 0.60, 10 => 0.60,
    11 => 0.70, 12 => 1.00
)

# Coûts [ORIGINAL - COST_UNSUPPLIED critique : improved avait 10 000 au lieu de 1 milliard]
const COST_PUMP_STEP        = 1.0            # €/MWh
const COST_SPILL            = 8000.0         # €/MWh (écrêtage)
const COST_HYDRO_OPPORTUNITY = 5.0           # €/MWh
# Max variation de production entre deux heures consécutives
# Le stock final ne peut pas varier de plus de X% de la capacité par rapport au stock initial
# const HYDRO_STOCK_RAMP_PCT = 0.15  # 15% de HYDRO_CAPACITY_MWh max de variation par fenêtre
const FUNNEL_HORIZON = 672  # 21 jours avant la fenêtre, l'entonnoir s'active


const WATER_PRICE_MONTHLY = Dict(
    1  => 22.0,   # Janvier
    2  => 12.0,   # Février
    3  => 12.0,   # Mars
    4  => 3.0,    # Avril    - fonte des neiges → turbine à fond
    5  => 3.0,    # Mai
    6  => 15.0,   # Juin
    7  => 200.0,  # Juillet  - eau ultra-précieuse → même le fioul (100€) est préféré
    8  => 200.0,  # Août     - pic de préservation
    9  => 200.0,  # Septembre
    10 => 200.0,  # Octobre  - commence à redescendre
    11 => 200.0,  # Novembre - transition
    12 => 22.0    # Décembre - retour hiver
)

# const WATER_PRICE_MONTHLY = Dict(
#     1  => 22.0,  # Janvier   - entre Nuc (12) et Charbon (36) → préserve modérément
#     2  => 12.0,  # Février   - = Nucléaire → neutre, stock peut baisser
#     3  => 12.0,  # Mars      - idem, fin d'hiver
#     4  => 3.0,   # Avril     - < Nucléaire → turbine agressivement (fonte)
#     5  => 3.0,   # Mai       - idem
#     6  => 15.0,  # Juin      - entre Nuc et Charbon → turbine librement
#     7  => 50.0,  # Juillet   - entre CCG (40) et TAC (70) → préfère CCG à l'hydro
#     8  => 50.0,  # Août      - idem
#     9  => 75.0,  # Septembre - entre TAC (70) et Fioul (100) → préserve fortement
#     10 => 75.0,  # Octobre   - idem
#     11 => 42.0,  # Novembre  - entre Charbon (36) et CCG (40) → transition
#     12 => 22.0   # Décembre  - = Janvier
# )

function water_price_at_hour(hour::Int)::Float64
    # Heure centrale de chaque mois (cumul des heures)
    month_centers = [
        (372,  1),   # milieu janvier
        (1020, 2),   # milieu février
        (1380, 3),   # milieu mars (672+744/2)
        (2232, 4),
        (2952, 5),
        (3696, 6),
        (4428, 7),
        (5172, 8),
        (5916, 9),
        (6660, 10),
        (7404, 11),
        (8148, 12),
    ]

    # Avant le premier centre → prix janvier
    if hour <= month_centers[1][1]
        return WATER_PRICE_MONTHLY[1]
    end
    # Après le dernier centre → prix décembre
    if hour >= month_centers[end][1]
        return WATER_PRICE_MONTHLY[12]
    end
    # Interpolation linéaire entre les deux centres encadrants
    for i in 1:(length(month_centers) - 1)
        h0, m0 = month_centers[i]
        h1, m1 = month_centers[i+1]
        if h0 <= hour <= h1
            alpha = (hour - h0) / (h1 - h0)
            p0 = WATER_PRICE_MONTHLY[m0]
            p1 = WATER_PRICE_MONTHLY[m1]
            return p0 + alpha * (p1 - p0)
        end
    end
    return WATER_PRICE_MONTHLY[12]
end

const COST_UNSUPPLIED   = 1_000_000_000.0   # inchangé
const COST_SLACK_STOCK  = 5_000_000_000.0   # 5× plus cher que Puns
const COST_SLACK_SEASONAL = 2_000_000_000.0

# ==============================================================================
# SECTION 6 : FONCTIONS UTILITAIRES
# ==============================================================================

function month_from_hour(hour::Int)
    cumsum = 0
    for m in 1:12
        cumsum += HOURS_PER_MONTH[m][1]
        if hour <= cumsum
            return m
        end
    end
    return 12
end

function get_seasonal_stock_target(scenario::String, season::String)
    """Retourner la cible de stock minimum pour une saison et un scénario"""
    targets = SEASONAL_STOCK_TARGETS[season]
    if scenario == "dry"
        return targets.dry
    elseif scenario == "normal"
        return targets.normal
    elseif scenario == "wet"
        return targets.wet
    else
        error("Scénario inconnu: $scenario")
    end
end

function is_in_seasonal_window(hour_global::Int, season::String)
    """Vérifier si une heure est dans la fenêtre d'une saison"""
    window = SEASONAL_WINDOWS[season]
    return window.start_hour <= hour_global <= window.end_hour
end

function recale_hydro_stocks(stocks::Vector{StockHydro}, shift_days::Int)
    """Décaler les stocks hydrauliques de shift_days jours"""
    n = length(stocks)
    return vcat(stocks[shift_days+1:end], stocks[1:shift_days])
end

function recale_hydro_values(v::Vector{Float64}, shift::Int)::Vector{Float64}
    """Décaler un vecteur de shift heures [ORIGINAL - manquait dans improved → crash]"""
    return vcat(v[shift+1:end], v[1:shift])
end

# ==============================================================================
# SECTION 7 : GÉNÉRATION DES DONNÉES ANNUELLES
# ==============================================================================

function generate_complete_year_data(scenario::String)::NTuple{8, Vector{Float64}}

    # ====== HYDRAULIQUE, CONSO ET ENR (depuis Excel) ======
    inflows_fdl = Float64[cons.fil_de_leau for cons in consommations_horaires[1:YEAR_HOURS]]
    load        = Float64[cons.conso_charge for cons in consommations_horaires[1:YEAR_HOURS]]
    wind        = Float64[cons.eolien_cons  for cons in consommations_horaires[1:YEAR_HOURS]]
    solar       = Float64[cons.solaire_cons for cons in consommations_horaires[1:YEAR_HOURS]]

    inflows_fdl = recale_hydro_values(inflows_fdl, HYDRO_SHIFT_HOURS)
    load        = recale_hydro_values(load,        HYDRO_SHIFT_HOURS)
    wind        = recale_hydro_values(wind,        HYDRO_SHIFT_HOURS)
    solar       = recale_hydro_values(solar,       HYDRO_SHIFT_HOURS)

    # ====== Apports lacs mensuels distribués ======
    apports_mensuels_gwh = [
        (1, 228.687), (2, 402.087), (3, 463.140), (4, 472.486),
        (5, 556.264), (6, 471.685), (7, 452.414), (8, 344.348),
        (9, 261.450), (10, 251.472), (11, 191.610), (12, 209.405)
    ]

    inflows_lac_opt = Float64[]
    for day in 1:365
        cumsum_days = 0
        month = 1
        for m in 1:12
            days_in_month = HOURS_PER_MONTH[m][1] ÷ 24
            cumsum_days += days_in_month
            if day <= cumsum_days
                month = m
                break
            end
        end

        apport_gwh       = apports_mensuels_gwh[month][2]
        hours_in_month   = HOURS_PER_MONTH[month][1]
        apport_mw_moyen  = (apport_gwh * 1000) / hours_in_month

        for h_day in 1:24
            variation = 0.85 + 0.30 * rand()
            push!(inflows_lac_opt, apport_mw_moyen * variation)
        end
    end

    inflows_lac = inflows_lac_opt[1:YEAR_HOURS]

    # Appliquer les facteurs scénario
    if scenario == "dry"
        inflows_lac .*= 0.6
        inflows_fdl .*= 0.8
    elseif scenario == "wet"
        inflows_lac .*= 1.4
        inflows_fdl .*= 1.2
    end

    # ====== BIOÉNERGIES ======
    dechets  = fill(60.0, YEAR_HOURS)
    biomasse = fill(365.0, YEAR_HOURS)
    fatale   = dechets .+ biomasse

    return load, wind, solar, inflows_fdl, inflows_lac, dechets, biomasse, fatale
end

# ==============================================================================
# SECTION 8 : CRÉATION DU MODÈLE D'OPTIMISATION
# ==============================================================================

function create_eod_model(
    window_num::Int, month::Int, scenario::String,
    load::Vector, wind::Vector, solar::Vector,
    inflows_fdl::Vector, inflows_lac::Vector,
    fatale::Vector, dechets::Vector, biomasse::Vector,
    stock_hydro_init::Float64, stock_step_init::Float64,
    tmax::Int, window_start::Int,
    unites::Vector{UniteCentrale},
    stocks_hydro::Vector{StockHydro},
    uc_init::Dict{String, Vector{Float64}}      # ← AJOUT
)::Tuple{Model, Vector, Vector, Vector, Vector}

    model = Model(HiGHS.Optimizer)
    set_optimizer_attribute(model, "log_to_console", false)
    set_optimizer_attribute(model, "time_limit", 120.0)      # réduit de 300 à 120s
    set_optimizer_attribute(model, "mip_rel_gap", 0.02)      # stop à 2% de l'optimal
    set_optimizer_attribute(model, "mip_abs_gap", 1e6)       # ou 1M€ d'écart absolu

    # Pré-calculer les minima/maxima de stock pour chaque heure de la fenêtre
    stock_limits = []
    for t in 1:tmax
        date_index = (window_start + t - 2) % length(stocks_hydro) + 1
        stock_data = stocks_hydro[date_index]
        stock_min  = max(0.0, (stock_data.lev_low  / 100) * HYDRO_CAPACITY_MWh)
        stock_max  = min(HYDRO_CAPACITY_MWh, (stock_data.lev_high / 100) * HYDRO_CAPACITY_MWh)
        push!(stock_limits, (stock_min, stock_max))
    end

    nuclear_units = [u for (u, unit) in enumerate(unites) if occursin("Nucléaire", unit.type)]
    @variable(model, P_nuc[1:tmax, nuclear_units] >= 0)
    @variable(model, UC_nuc[1:tmax, nuclear_units], Bin)

    charbon_units = [u for (u, unit) in enumerate(unites) if unit.type == "charbon"]
    @variable(model, P_charbon[1:tmax, charbon_units] >= 0)
    @variable(model, UC_charbon[1:tmax, charbon_units], Bin)

    ccg_units = [u for (u, unit) in enumerate(unites) if unit.type == "CCG gaz"]
    @variable(model, P_ccg[1:tmax, ccg_units] >= 0)
    @variable(model, UC_ccg[1:tmax, ccg_units], Bin)

    tac_units = [u for (u, unit) in enumerate(unites) if unit.type == "TAC gaz"]
    @variable(model, P_tac[1:tmax, tac_units] >= 0)
    @variable(model, UC_tac[1:tmax, tac_units], Bin)

    cogen_units = [u for (u, unit) in enumerate(unites) if unit.type == "cogénération gaz"]
    @variable(model, P_cogen[1:tmax, cogen_units] >= 0)
    @variable(model, UC_cogen[1:tmax, cogen_units], Bin)

    fioul_units = [u for (u, unit) in enumerate(unites) if unit.type == "fioul"]
    @variable(model, P_fioul[1:tmax, fioul_units] >= 0)
    @variable(model, UC_fioul[1:tmax, fioul_units], Bin)

    eolien_units = [u for (u, unit) in enumerate(unites) if unit.type == "éolien"]
    @variable(model, P_eolien[1:tmax, eolien_units] >= 0)

    solaire_units = [u for (u, unit) in enumerate(unites) if unit.type == "solaire"]
    @variable(model, P_solaire[1:tmax, solaire_units] >= 0)

    dechets_units = [u for (u, unit) in enumerate(unites) if unit.type == "déchets"]
    @variable(model, P_dechets[1:tmax, dechets_units] >= 0)
    @variable(model, UC_dechets[1:tmax, dechets_units], Bin)

    biomasse_units = [u for (u, unit) in enumerate(unites) if unit.type == "petite biomasse"]
    @variable(model, P_biomasse[1:tmax, biomasse_units] >= 0)
    @variable(model, UC_biomasse[1:tmax, biomasse_units], Bin)

    # Hydraulique
    @variable(model, Phy_fdl[1:tmax] >= 0)
    @variable(model, Phy_lac[1:tmax] >= 0)
    @variable(model, stock_hydro[1:tmax] >= 0)

    # STEP
    @variable(model, Pcharge_STEP[1:tmax] >= 0)
    @variable(model, Pdecharge_STEP[1:tmax] >= 0)
    @variable(model, stock_STEP[1:tmax] >= 0)

    # Bilan
    @variable(model, Puns[1:tmax] >= 0)
    @variable(model, Pspill[1:tmax] >= 0)

    # Slacks (faisabilité)
    @variable(model, slack_stock_min[1:tmax] >= 0)
    @variable(model, slack_stock_max[1:tmax] >= 0)
    @variable(model, slack_seasonal[1:tmax] >= 0)

    # [IMPROVED] Tranches d'eau (socle à conserver / surplus utilisable librement)
    @variable(model, S_socle[1:tmax] >= 0)
    @variable(model, S_surplus[1:tmax] >= 0)

    # ====== VARIABLES UP/DO POUR DMIN ======
    n_nuc   = length(nuclear_units)
    n_char  = length(charbon_units)
    n_ccg   = length(ccg_units)
    n_tac   = length(tac_units)
    n_fioul = length(fioul_units)

    @variable(model, UP_nuc[1:tmax,    1:n_nuc],   Bin)
    @variable(model, DO_nuc[1:tmax,    1:n_nuc],   Bin)
    @variable(model, UP_charbon[1:tmax, 1:n_char],  Bin)
    @variable(model, DO_charbon[1:tmax, 1:n_char],  Bin)
    @variable(model, UP_ccg[1:tmax,    1:n_ccg],   Bin)
    @variable(model, DO_ccg[1:tmax,    1:n_ccg],   Bin)
    @variable(model, UP_tac[1:tmax,    1:n_tac],   Bin)
    @variable(model, DO_tac[1:tmax,    1:n_tac],   Bin)
    @variable(model, UP_fioul[1:tmax,  1:n_fioul], Bin)
    @variable(model, DO_fioul[1:tmax,  1:n_fioul], Bin)

    # ====== CONTRAINTES DURÉE MINIMALE avec UP/DO ======
    function add_dmin!(model, UC, UP, DO, global_indices, unites, tmax, uc_init_vec)
        for (i_local, u_global) in enumerate(global_indices)
            dmin = unites[u_global].dmin
            for t in 1:tmax
                uc_curr = UC[t, u_global]
                uc_prev = t == 1 ? uc_init_vec[i_local] : UC[t-1, u_global]
                up_t    = UP[t, i_local]
                do_t    = DO[t, i_local]
                @constraint(model, up_t - do_t == uc_curr - uc_prev)
                @constraint(model, up_t + do_t <= 1)
            end
            if dmin > 0
                for t in 1:tmax
                    tend = min(t + dmin - 1, tmax)
                    @constraint(model,
                        sum(UC[tau, u_global] for tau in t:tend) >= dmin * UP[t, i_local]
                    )
                end
            end
        end
    end

    add_dmin!(model, UC_nuc,     UP_nuc,     DO_nuc,     nuclear_units, unites, tmax, uc_init["nuc"])
    add_dmin!(model, UC_charbon, UP_charbon, DO_charbon, charbon_units, unites, tmax, uc_init["charbon"])
    add_dmin!(model, UC_ccg,     UP_ccg,     DO_ccg,     ccg_units,     unites, tmax, uc_init["ccg"])
    add_dmin!(model, UC_tac,     UP_tac,     DO_tac,     tac_units,     unites, tmax, uc_init["tac"])
    add_dmin!(model, UC_fioul,   UP_fioul,   DO_fioul,   fioul_units,   unites, tmax, uc_init["fioul"])

    # ====== CONTRAINTES THERMIQUES ======
    for u in nuclear_units
        unit = unites[u]
        @constraint(model, [t=1:tmax], P_nuc[t, u] <= unit.Pmax * UC_nuc[t, u])
        @constraint(model, [t=1:tmax], P_nuc[t, u] >= unit.Pmin * UC_nuc[t, u])
    end

    for u in charbon_units
        unit = unites[u]
        pmax = unit.Pmax
        @constraint(model, [t=1:tmax], P_charbon[t, u] <= pmax * UC_charbon[t, u])
        @constraint(model, [t=1:tmax], P_charbon[t, u] >= unit.Pmin * UC_charbon[t, u])
    end

    for u in ccg_units
        unit = unites[u]
        pmax = unit.Pmax
        @constraint(model, [t=1:tmax], P_ccg[t, u] <= pmax * UC_ccg[t, u])
        @constraint(model, [t=1:tmax], P_ccg[t, u] >= unit.Pmin * UC_ccg[t, u])
    end

    for u in tac_units
        unit = unites[u]
        pmax = unit.Pmax
        @constraint(model, [t=1:tmax], P_tac[t, u] <= pmax * UC_tac[t, u])
        @constraint(model, [t=1:tmax], P_tac[t, u] >= unit.Pmin * UC_tac[t, u])
    end

    for u in cogen_units
        unit = unites[u]
        pmax = unit.Pmax
        @constraint(model, [t=1:tmax], P_cogen[t, u] <= pmax * UC_cogen[t, u])
        @constraint(model, [t=1:tmax], P_cogen[t, u] >= unit.Pmin * UC_cogen[t, u])
    end

    for u in fioul_units
        unit = unites[u]
        @constraint(model, [t=1:tmax], P_fioul[t, u] <= unit.Pmax * UC_fioul[t, u])
        @constraint(model, [t=1:tmax], P_fioul[t, u] >= unit.Pmin * UC_fioul[t, u])
    end

    # ====== CONTRAINTES ENR (MUST-TAKE) [ORIGINAL]  ======
    # == : l'ENR est obligatoirement injecté (must-take physique)
    # Pspill peut absorber les surplus via le bilan
    @constraint(model, [t=1:tmax], sum(P_eolien[t, :]) == wind[t])
    @constraint(model, [t=1:tmax], sum(P_solaire[t, :]) == solar[t])

    # ====== CONTRAINTES BIOÉNERGIES (MUST-RUN) ======
    for u in dechets_units
        unit = unites[u]
        @constraint(model, [t=1:tmax], UC_dechets[t, u] == 1)
        @constraint(model, [t=1:tmax], P_dechets[t, u] == unit.Pmax)
    end

    for u in biomasse_units
        unit = unites[u]
        pmax = unit.Pmax
        @constraint(model, [t=1:tmax], UC_biomasse[t, u] == 1)
        @constraint(model, [t=1:tmax], P_biomasse[t, u] == pmax)
    end

    # ====== CONTRAINTES HYDRAULIQUE FDL [ORIGINAL - forcé] ======
    @constraint(model, [t=1:tmax], Phy_fdl[t] == min(inflows_fdl[t], HYDRO_FDL_CAPACITY_MW))

    # ====== CONTRAINTES HYDRAULIQUE LACS ======
        for t in 1:tmax
            stock_min, stock_max = stock_limits[t]

            @constraint(model, stock_hydro[t] >= stock_min - slack_stock_min[t])
            @constraint(model, stock_hydro[t] <= stock_max + slack_stock_max[t])
        end

    # ====== CONTRAINTE DE RAMPE HYDRAULIQUE LAC ======

    # max_variation = HYDRO_STOCK_RAMP_PCT * HYDRO_CAPACITY_MWh

    # @constraint(model, stock_hydro[tmax] >= stock_hydro_init - max_variation)
    # @constraint(model, stock_hydro[tmax] <= stock_hydro_init + max_variation)

    # [IMPROVED] Tranches d'eau : stock total = socle + surplus
    @constraint(model, [t=1:tmax], stock_hydro[t] == S_socle[t] + S_surplus[t])

    # Contraintes saisonnières sur stock [décommentables]
    for t in 1:tmax
        hour_global = window_start + t - 1
        for (season, _) in SEASONAL_WINDOWS
            if is_in_seasonal_window(hour_global, season)
                target_pct = get_seasonal_stock_target(scenario, season)
                target_mwh = target_pct * HYDRO_CAPACITY_MWh
                @constraint(model, stock_hydro[t] + slack_seasonal[t] >= target_mwh)
            end
        end
    end

    # ====== CONTRAINTES SAISONNIÈRES AVEC APPROCHE PROGRESSIVE ======
    

    for t in 1:tmax
        hour_global = window_start + t - 1

        for (season, window) in SEASONAL_WINDOWS
            target_pct = get_seasonal_stock_target(scenario, season)
            target_mwh = target_pct * HYDRO_CAPACITY_MWh

            # Dans la fenêtre : contrainte pleine
            if window.start_hour <= hour_global <= window.end_hour
                @constraint(model, stock_hydro[t] + slack_seasonal[t] >= target_mwh)

            # Avant la fenêtre : entonnoir progressif
            elseif hour_global < window.start_hour
                hours_to_window = window.start_hour - hour_global
                if hours_to_window <= FUNNEL_HORIZON
                    # Progression linéaire : 0% de la target à FUNNEL_HORIZON heures,
                    # 100% à l'entrée de la fenêtre
                    progress = 1.0 - hours_to_window / FUNNEL_HORIZON
                    partial_target = target_mwh * progress
                    @constraint(model, stock_hydro[t] + slack_seasonal[t] >= partial_target)
                end
            end
        end
    end

    # Écrêtage saisonnier
    for t in 1:tmax
        hour_global = window_start + t - 1
        month_t     = month_from_hour(hour_global)
        ecrêtage    = HYDRO_LAC_ECRÊTAGE_MONTHLY[month_t]
        @constraint(model, Phy_lac[t] <= HYDRO_LAC_CAPACITY_MW * ecrêtage)
    end

    # Dynamique lac : la STEP ne touche pas au lac
    @constraint(model, [t=1:tmax],
        stock_hydro[t] ==
        (t == 1 ? stock_hydro_init : stock_hydro[t-1]) +
        inflows_lac[t] - Phy_lac[t] - Pspill[t]
    )

    # ====== CONTRAINTES STEP ======
    @constraint(model, [t=1:tmax], Pcharge_STEP[t]   <= STEP_POWER_MW)
    @constraint(model, [t=1:tmax], Pdecharge_STEP[t] <= STEP_POWER_MW)

    @constraint(model, [t=1:tmax],
        stock_STEP[t] ==
        (t == 1 ? stock_step_init : stock_STEP[t-1]) +
        Pcharge_STEP[t] * STEP_EFFICIENCY - Pdecharge_STEP[t]
    )
    @constraint(model, [t=1:tmax], stock_STEP[t] <= STEP_CAPACITY_MWh)

    # # Fermeture hebdomadaire STEP
    # hours_per_week = 168
    # num_weeks = floor(Int, tmax / hours_per_week)
    # for week in 2:num_weeks
    #     hour_start = (week - 1) * 168 + 1
    #     hour_end   = week * 168
    #     if hour_end <= tmax
    #         @constraint(model, stock_STEP[hour_start] == stock_STEP[hour_end])
    #     end
    # end
    # @constraint(model, stock_STEP[1] >= stock_step_init * 0.8)
    # @constraint(model, stock_STEP[1] <= stock_step_init * 1.2)

    # Condition initiale exacte
    @constraint(model, stock_STEP[1] == stock_step_init)

    # Contrainte de clôture souple : le stock final ne doit pas être
    # trop inférieur au stock initial (éviter de vider le STEP sans recharger)
    @constraint(model, stock_STEP[tmax] >= stock_step_init * 0.5)

    # ====== CONTRAINTE BILAN ======
    @constraint(model, [t=1:tmax],
        sum(P_nuc[t, :]) + sum(P_charbon[t, :]) + sum(P_ccg[t, :]) +
        sum(P_tac[t, :]) + sum(P_cogen[t, :])  + sum(P_fioul[t, :]) +
        sum(P_eolien[t, :]) + sum(P_solaire[t, :]) +
        sum(P_dechets[t, :]) + sum(P_biomasse[t, :]) +
        Phy_fdl[t] + Phy_lac[t] + Pdecharge_STEP[t] + Pspill[t] ==
        load[t] + Pcharge_STEP[t] + Puns[t]
    )

    # ====== FONCTION OBJECTIF ======
    cost_thermal = sum(
        unites[u].prix_marche * P_nuc[t, u] for t in 1:tmax, u in nuclear_units
    ) + sum(
        unites[u].prix_marche * P_charbon[t, u] for t in 1:tmax, u in charbon_units
    ) + sum(
        unites[u].prix_marche * P_ccg[t, u] for t in 1:tmax, u in ccg_units
    ) + sum(
        unites[u].prix_marche * P_tac[t, u] for t in 1:tmax, u in tac_units
    ) + sum(
        unites[u].prix_marche * P_cogen[t, u] for t in 1:tmax, u in cogen_units
    ) + sum(
        unites[u].prix_marche * P_fioul[t, u] for t in 1:tmax, u in fioul_units
    )

    cost_hydro = sum(
        water_price_at_hour(window_start + t - 1) * (Phy_fdl[t] + Phy_lac[t])
        for t in 1:tmax
    )

    cost_pump      = sum(COST_PUMP_STEP    * Pcharge_STEP[t] for t in 1:tmax)
    cost_spill     = sum(COST_SPILL        * Pspill[t]       for t in 1:tmax)
    cost_unserved  = sum(COST_UNSUPPLIED   * Puns[t]         for t in 1:tmax)

    # Pénalités slack [ORIGINAL - improved avait réduit à 2 000]
    cost_slack = sum(
        COST_SLACK_STOCK    * (slack_stock_min[t] + slack_stock_max[t]) +
        COST_SLACK_SEASONAL * slack_seasonal[t]
        for t in 1:tmax
    )

    @objective(model, Min,
        cost_thermal + cost_hydro + cost_pump + cost_spill + cost_unserved + cost_slack
    )

    return model, stock_limits, slack_stock_min, slack_stock_max, slack_seasonal
end

# ==============================================================================
# SECTION 9 : RÉSOLUTION FENÊTRE ROULANTE
# ==============================================================================

function solve_year_rolling(
    load::Vector, wind::Vector, solar::Vector,
    inflows_fdl::Vector, inflows_lac::Vector,
    dechets::Vector, biomasse::Vector, fatale::Vector,
    scenario::String, stock_init_pct::Float64,
    unites::Vector{UniteCentrale}, stocks_hydro::Vector{StockHydro}
)::Vector{Dict}

    year_hours = length(load)
    all_results = []

    # Stock initial calé sur les limites du mois 1 [ORIGINAL]
    stock_data_h1   = stocks_hydro[1]
    stock_min_h1    = max(0.0, (stock_data_h1.lev_low  / 100) * HYDRO_CAPACITY_MWh)
    stock_max_h1    = min(HYDRO_CAPACITY_MWh, (stock_data_h1.lev_high / 100) * HYDRO_CAPACITY_MWh)
    stock_hydro_raw = HYDRO_CAPACITY_MWh * stock_init_pct
    stock_hydro_current = clamp(stock_hydro_raw, stock_min_h1, stock_max_h1)
    stock_step_current  = 500_000.0  # 0.5 TWh de départ
    # État UC inter-fenêtres pour Dmin
    uc_prev = Dict(
        "nuc"     => zeros(Float64, sum(1 for unit in unites if occursin("Nucléaire", unit.type))),
        "charbon" => zeros(Float64, sum(1 for unit in unites if unit.type == "charbon")),
        "ccg"     => zeros(Float64, sum(1 for unit in unites if unit.type == "CCG gaz")),
        "tac"     => zeros(Float64, sum(1 for unit in unites if unit.type == "TAC gaz")),
        "fioul"   => zeros(Float64, sum(1 for unit in unites if unit.type == "fioul")),
    )

    println("  Stock hydro init (avant clamp) : $(round(stock_hydro_raw/1e6, digits=2)) TWh")
    println("  Limites janvier : [$(round(stock_min_h1/1e6, digits=2)), $(round(stock_max_h1/1e6, digits=2))] TWh")
    println("  Stock hydro init (après clamp) : $(round(stock_hydro_current/1e6, digits=2)) TWh")

    hour       = 1
    window_num = 1

    println("\nRésolution fenêtre roulante : $scenario")
    println("Heures totales : $year_hours\n")

    while hour <= year_hours
        window_start = hour
        window_end   = min(window_start + WINDOW_SIZE - 1, year_hours)
        results_end  = min(window_start + RESULTS_SIZE - 1, year_hours)
        hours_to_keep = results_end - window_start + 1

        month      = month_from_hour(window_start)
        month_name = HOURS_PER_MONTH[month][2]

        println("  ▶ Fenêtre $window_num ($month_name) - Heures $window_start-$window_end")
        println("     Stock hydro : $(round(stock_hydro_current/1e6, digits=3)) TWh | STEP : $(round(stock_step_current/1e6, digits=3)) TWh")

        window_range = window_start:window_end
        load_w = load[window_range]
        wind_w = wind[window_range]
        solar_w = solar[window_range]
        fdl_w  = inflows_fdl[window_range]
        lac_w  = inflows_lac[window_range]
        fat_w  = fatale[window_range]
        dec_w  = dechets[window_range]
        bio_w  = biomasse[window_range]

        hours_in_window = length(window_range)

        model, stock_limits, slack_stock_min, slack_stock_max, slack_seasonal = create_eod_model(
            window_num, month, scenario,
            load_w, wind_w, solar_w, fdl_w, lac_w, fat_w, dec_w, bio_w,
            stock_hydro_current, stock_step_current, hours_in_window, window_start,
            unites, stocks_hydro,
            uc_prev                                                           # ← AJOUT
        )

        stock_hydro_before = stock_hydro_current

        optimize!(model)
        status = termination_status(model)

        # Extraction résultats (même si non optimal)
        P_nuc        = value.(model[:P_nuc])
        P_charbon    = value.(model[:P_charbon])
        P_ccg        = value.(model[:P_ccg])
        P_tac        = value.(model[:P_tac])
        P_cogen      = value.(model[:P_cogen])
        P_fioul      = value.(model[:P_fioul])
        P_eolien     = value.(model[:P_eolien])
        P_solaire    = value.(model[:P_solaire])
        P_dechets    = value.(model[:P_dechets])
        P_biomasse   = value.(model[:P_biomasse])
        Phy_fdl_vals = value.(model[:Phy_fdl])
        Phy_lac_vals = value.(model[:Phy_lac])
        stock_hydro_vals = value.(model[:stock_hydro])
        stock_STEP   = value.(model[:stock_STEP])
        Puns         = value.(model[:Puns])
        Pspill       = value.(model[:Pspill])
        UC_nuc       = value.(model[:UC_nuc])
        UC_charbon   = value.(model[:UC_charbon])
        UC_ccg       = value.(model[:UC_ccg])
        UC_tac       = value.(model[:UC_tac])
        UC_cogen     = value.(model[:UC_cogen])
        UC_fioul     = value.(model[:UC_fioul])
        Pcharge_STEP   = value.(model[:Pcharge_STEP])
        Pdecharge_STEP = value.(model[:Pdecharge_STEP])
        S_socle_vals   = value.(model[:S_socle])
        S_surplus_vals = value.(model[:S_surplus])

        # Diagnostic si non optimal [ORIGINAL]
        if status != OPTIMAL
            println("     ⚠️  Statut : $status")
            println("\n     🔍 DIAGNOSTIC :")
            println("       - Stock hydro init: $(round(stock_hydro_before/1e6, digits=3)) TWh")
            println("       - Stock STEP init : $(round(stock_step_current/1e6, digits=3)) TWh")
            println("       - Charge moyenne  : $(round(mean(load_w)/1000, digits=2)) GW")
            println("       - Éolien moyen    : $(round(mean(wind_w)/1000, digits=2)) GW")
            println("       - Apports hydro   : $(round((mean(fdl_w) + mean(lac_w))/1000, digits=2)) GW")
            if stock_hydro_before < 0.12 * HYDRO_CAPACITY_MWh
                println("       ❌ STOCK HYDRO TRÈS BAS ($(round(stock_hydro_before/1e6, digits=3)) TWh < 0.12 TWh)")
            end
            if mean(load_w) > 11500
                println("       ❌ CHARGE TRÈS ÉLEVÉE ($(round(mean(load_w)/1000, digits=2)) GW)")
            end
        end

        # Mise à jour stock pour fenêtre suivante
        if status != OPTIMAL
            stock_hydro_current = stock_hydro_before
            stock_step_current = 500_000.0
        else
            stock_hydro_current = stock_hydro_vals[hours_to_keep]
            stock_step_current  = stock_STEP[hours_to_keep]

            if !isfinite(stock_hydro_current) || stock_hydro_current < 0
                stock_hydro_current = HYDRO_CAPACITY_MWh * 0.5
            end
            # if stock_step_current < 500_000.0
            #     println("     ⚠️  Stock STEP bas → reset à 0.5 TWh")
            #     stock_step_current = 500_000.0
            # end
        end

        # Mise à jour UC pour la fenêtre suivante
        if status == OPTIMAL
            nuc_idx     = [u for (u, unit) in enumerate(unites) if occursin("Nucléaire", unit.type)]
            charbon_idx = [u for (u, unit) in enumerate(unites) if unit.type == "charbon"]
            ccg_idx     = [u for (u, unit) in enumerate(unites) if unit.type == "CCG gaz"]
            tac_idx     = [u for (u, unit) in enumerate(unites) if unit.type == "TAC gaz"]
            fioul_idx   = [u for (u, unit) in enumerate(unites) if unit.type == "fioul"]

            uc_prev["nuc"]     = [round(value(model[:UC_nuc][hours_to_keep,     u])) for u in nuc_idx]
            uc_prev["charbon"] = [round(value(model[:UC_charbon][hours_to_keep,  u])) for u in charbon_idx]
            uc_prev["ccg"]     = [round(value(model[:UC_ccg][hours_to_keep,      u])) for u in ccg_idx]
            uc_prev["tac"]     = [round(value(model[:UC_tac][hours_to_keep,      u])) for u in tac_idx]
            uc_prev["fioul"]   = [round(value(model[:UC_fioul][hours_to_keep,    u])) for u in fioul_idx]
        end

        # Stocker résultats (toujours, même si infeasible)
        for t in 1:hours_to_keep
            push!(all_results, Dict(
                "hour_global"    => window_start + t - 1,
                "month"          => month,
                "load"           => load[window_start + t - 1],
                "Phy_fdl"        => Phy_fdl_vals[t],
                "Phy_lac"        => Phy_lac_vals[t],
                "P_nuc"          => sum(P_nuc[t, :]),
                "P_charbon"      => sum(P_charbon[t, :]),
                "P_ccg"          => sum(P_ccg[t, :]),
                "P_tac"          => sum(P_tac[t, :]),
                "P_cogen"        => sum(P_cogen[t, :]),
                "P_fioul"        => sum(P_fioul[t, :]),
                "P_eolien"       => sum(P_eolien[t, :]),
                "P_solaire"      => sum(P_solaire[t, :]),
                "P_dechets"      => sum(P_dechets[t, :]),
                "P_biomasse"     => sum(P_biomasse[t, :]),
                "stock_hydro"    => stock_hydro_vals[t],
                "stock_STEP"     => stock_STEP[t],
                "Puns"           => Puns[t],
                "Pspill"         => Pspill[t],
                "UC_nuc"         => sum(UC_nuc[t, :]),
                "UC_charbon"     => sum(UC_charbon[t, :]),
                "UC_ccg"         => sum(UC_ccg[t, :]),
                "UC_tac"         => sum(UC_tac[t, :]),
                "UC_cogen"       => sum(UC_cogen[t, :]),
                "UC_fioul"       => sum(UC_fioul[t, :]),
                "Pcharge_STEP"   => Pcharge_STEP[t],
                "Pdecharge_STEP" => Pdecharge_STEP[t],
                "stock_hydro_min" => stock_limits[t][1],
                "stock_hydro_max" => stock_limits[t][2],
                "slack_seasonal"  => value(slack_seasonal[t]),
                "slack_stock_min" => value(slack_stock_min[t]),
                "slack_stock_max" => value(slack_stock_max[t]),
                "S_socle"         => S_socle_vals[t],    # [IMPROVED]
                "S_surplus"       => S_surplus_vals[t],  # [IMPROVED]
                "inflows_lac"     => inflows_lac[window_start + t - 1],
                "inflows_fdl"     => inflows_fdl[window_start + t - 1],
                "infeasible_flag" => (status != OPTIMAL ? 1 : 0),
                "water_price"  => water_price_at_hour(window_start + t - 1),
                "cost_nuc"     => sum(unites[u].prix_marche * value(P_nuc[t, u]) for (u, unit) in enumerate(unites) if occursin("Nucléaire", unit.type)),
                "cost_charbon" => sum(unites[u].prix_marche * value(P_charbon[t, u]) for (u, unit) in enumerate(unites) if unit.type == "charbon"),
                "cost_ccg"     => sum(unites[u].prix_marche * value(P_ccg[t, u]) for (u, unit) in enumerate(unites) if unit.type == "CCG gaz"),
                "cost_tac"     => sum(unites[u].prix_marche * value(P_tac[t, u]) for (u, unit) in enumerate(unites) if unit.type == "TAC gaz"),
                "cost_fioul"   => sum(unites[u].prix_marche * value(P_fioul[t, u]) for (u, unit) in enumerate(unites) if unit.type == "fioul"),
            ))
        end

        hour = window_start + RESULTS_SIZE
        window_num += 1
    end

    return all_results
end

# ==============================================================================
# SECTION 10 : EXÉCUTION PRINCIPALE
# ==============================================================================

println("\n" * "="^80)
println("EOD ZOOTOPIA - OPTIMISATION DISPATCHING ÉLECTRIQUE (VERSION FUSIONNÉE)")
println("="^80 * "\n")

# [IMPROVED] Parser les arguments CLI
params         = parse_args(ARGS)
scenario_cli   = params["scenario"]
data_file      = params["data_file"]

println("📋 Paramètres:")
println("   Scénario(s) : $scenario_cli")
println("   Fichier     : $data_file")

# Initialiser
unites = initialize_unites()
stocks_hydro, apports_mensuels, consommations_horaires = load_excel_data(data_file)
stocks_hydro = recale_hydro_stocks(stocks_hydro, HYDRO_SHIFT_DAYS)

# Paramètres scénarios (stock initial %)
scenario_params = Dict(
    "dry"    => 0.60,
    "normal" => 0.70,
    "wet"    => 0.80
)

# [IMPROVED] Sélection des scénarios à résoudre
scenarios_to_solve = if scenario_cli in ["dry", "normal", "wet"]
    [scenario_cli]
else
    ["dry", "normal", "wet"]
end

results_all = Dict()

for scenario in scenarios_to_solve
    println("\n" * "-"^80)
    println("SCÉNARIO : $scenario")
    println("-"^80)

    load, wind, solar, fdl, lac, dechets, biomasse, fatale =
        generate_complete_year_data(scenario)

    results = solve_year_rolling(
        load, wind, solar, fdl, lac, dechets, biomasse, fatale,
        scenario, scenario_params[scenario],
        unites, stocks_hydro
    )

    results_all[scenario] = results

    # [IMPROVED] Stats détaillées
    println("\n📊 Résumé annuel $scenario :")
    println("  Prod. nucléaire  : $(round(sum(r["P_nuc"]     for r in results)/1e6, digits=2)) TWh")
    println("  Prod. charbon    : $(round(sum(r["P_charbon"] for r in results)/1e6, digits=2)) TWh")
    println("  Prod. gaz (CCG)  : $(round(sum(r["P_ccg"]    for r in results)/1e6, digits=2)) TWh")
    println("  Prod. gaz (TAC)  : $(round(sum(r["P_tac"]    for r in results)/1e6, digits=2)) TWh")
    println("  Prod. cogén.     : $(round(sum(r["P_cogen"]  for r in results)/1e6, digits=2)) TWh")
    println("  Prod. fioul      : $(round(sum(r["P_fioul"]  for r in results)/1e6, digits=2)) TWh")
    println("  Prod. éolienne   : $(round(sum(r["P_eolien"] for r in results)/1e6, digits=2)) TWh")
    println("  Prod. solaire    : $(round(sum(r["P_solaire"] for r in results)/1e6, digits=2)) TWh")
    println("  Prod. hydro FDE  : $(round(sum(r["Phy_fdl"]  for r in results)/1e6, digits=2)) TWh")
    println("  Prod. hydro lacs : $(round(sum(r["Phy_lac"]  for r in results)/1e6, digits=2)) TWh")
    println("  Défaillance      : $(round(sum(r["Puns"]     for r in results)/1e6, digits=6)) TWh")
    println("  Spill hydro      : $(round(sum(r["Pspill"]   for r in results)/1e6, digits=2)) TWh")
    n_inf = sum(r["infeasible_flag"] for r in results)
    if n_inf > 0
        println("  ⚠️  Heures infaisables : $n_inf")
    end
end

# Exporter résultats
mkpath("./results")

for scenario in keys(results_all)
    df = DataFrame(
        hour             = [r["hour_global"]    for r in results_all[scenario]],
        month            = [r["month"]          for r in results_all[scenario]],
        load             = [r["load"]           for r in results_all[scenario]],
        Phy_fdl          = [r["Phy_fdl"]        for r in results_all[scenario]],
        Phy_lac          = [r["Phy_lac"]        for r in results_all[scenario]],
        P_nuc            = [r["P_nuc"]          for r in results_all[scenario]],
        P_charbon        = [r["P_charbon"]      for r in results_all[scenario]],
        P_ccg            = [r["P_ccg"]          for r in results_all[scenario]],
        P_tac            = [r["P_tac"]          for r in results_all[scenario]],
        P_cogen          = [r["P_cogen"]        for r in results_all[scenario]],
        P_fioul          = [r["P_fioul"]        for r in results_all[scenario]],
        P_eolien         = [r["P_eolien"]       for r in results_all[scenario]],
        P_solaire        = [r["P_solaire"]      for r in results_all[scenario]],
        P_dechets        = [r["P_dechets"]      for r in results_all[scenario]],
        P_biomasse       = [r["P_biomasse"]     for r in results_all[scenario]],
        stock_hydro      = [r["stock_hydro"]    for r in results_all[scenario]],
        stock_STEP       = [r["stock_STEP"]     for r in results_all[scenario]],
        Puns             = [r["Puns"]           for r in results_all[scenario]],
        Pspill           = [r["Pspill"]         for r in results_all[scenario]],
        UC_nuc           = [r["UC_nuc"]         for r in results_all[scenario]],
        UC_charbon       = [r["UC_charbon"]     for r in results_all[scenario]],
        UC_ccg           = [r["UC_ccg"]         for r in results_all[scenario]],
        UC_tac           = [r["UC_tac"]         for r in results_all[scenario]],
        UC_cogen         = [r["UC_cogen"]       for r in results_all[scenario]],
        UC_fioul         = [r["UC_fioul"]       for r in results_all[scenario]],
        Pcharge_STEP     = [r["Pcharge_STEP"]   for r in results_all[scenario]],
        Pdecharge_STEP   = [r["Pdecharge_STEP"] for r in results_all[scenario]],
        stock_hydro_min  = [r["stock_hydro_min"] for r in results_all[scenario]],
        stock_hydro_max  = [r["stock_hydro_max"] for r in results_all[scenario]],
        slack_seasonal   = [r["slack_seasonal"]  for r in results_all[scenario]],
        slack_stock_min  = [r["slack_stock_min"] for r in results_all[scenario]],
        slack_stock_max  = [r["slack_stock_max"] for r in results_all[scenario]],
        S_socle          = [r["S_socle"]         for r in results_all[scenario]],    # [IMPROVED]
        S_surplus        = [r["S_surplus"]       for r in results_all[scenario]],    # [IMPROVED]
        inflows_lac      = [r["inflows_lac"]     for r in results_all[scenario]],
        inflows_fdl      = [r["inflows_fdl"]     for r in results_all[scenario]],
        infeasible_flag  = [r["infeasible_flag"] for r in results_all[scenario]],
        water_price     = [r["water_price"]     for r in results_all[scenario]],
        cost_nuc        = [r["cost_nuc"]        for r in results_all[scenario]],
        cost_charbon    = [r["cost_charbon"]    for r in results_all[scenario]],
        cost_ccg        = [r["cost_ccg"]        for r in results_all[scenario]],
        cost_tac        = [r["cost_tac"]        for r in results_all[scenario]],
        cost_fioul      = [r["cost_fioul"]      for r in results_all[scenario]]
    )

    filepath = "./results/results_$scenario.csv"
    CSV.write(filepath, df)
    println("✅ Exporté : $filepath")
end

println("\n" * "="^80)
println("✅ SIMULATION TERMINÉE")
println("="^80 * "\n")
println("📁 Résultats dans : ./results/")
println("🆕 Nouvelles colonnes : S_socle, S_surplus, infeasible_flag")
