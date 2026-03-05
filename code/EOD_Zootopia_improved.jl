# ==============================================================================
# PROJET EOD - PARC ÉLECTRIQUE ZOOTOPIA (VERSION AMÉLIORÉE)
#
# Optimisation économique du dispatching électrique avec :
# - Fenêtre roulante 9 jours (7+2)
# - Gestion hydraulique saisonnière avec contraintes de stock et d'écrêtage
# - Gestion STEP avec rendement 75%
# - 3 scénarios météo (dry/normal/wet)
# - Arguments en ligne de commande pour scénarios
# - Fonctions de conversion robustes pour lecture Excel
# - Tranches d'eau (socle/surplus) pour meilleure gestion hydraulique
# ==============================================================================
#
# UTILISATION :
#   julia EOD_Zootopia_improved.jl                    # scénario normal par défaut
#   julia EOD_Zootopia_improved.jl --scenario dry     # scénario sec
#   julia EOD_Zootopia_improved.jl --scenario wet     # scénario humide
#
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
# SECTION 0 : PARSING D'ARGUMENTS
# ==============================================================================

function parse_args(args)
    """Parse les arguments de la ligne de commande"""
    params = Dict(
        "scenario" => "normal",
        "data_file" => ".//data//Donnees_etude_de_cas_ETE305.xlsx"
    )
    
    i = 1
    while i <= length(args)
        if i < length(args)
            key = args[i]
            val = args[i+1]
            if key == "--scenario"
                if val in ["dry", "normal", "wet"]
                    params["scenario"] = val
                    i += 2
                else
                    println("⚠️  Scénario invalide: $val (utiliser dry/normal/wet)")
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
# SECTION 2 : FONCTIONS DE CONVERSION ROBUSTES (depuis optim_hydroVF.jl)
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
        # Nom, Type, Capacité totale (MW), Nombre d'unités, Pmax/unité (MW), Pmin/unité (MW), Dmin (h), Prix marché (€/MWh), Année de lancement
        ("Iconuc", "Nucléaire", 1800, 2, 900, 300, 24, 12, 1977),
        ("Tabarnuc", "Nucléaire", 1800, 2, 900, 300, 24, 12, 1981),
        ("NucPlusUltra", "Nucléaire", 2600, 2, 1300, 330, 24, 12, 1987),
        ("Gazby", "CCG gaz", 390, 1, 390, 135, 4, 40, 1994),
        ("Pégaz", "CCG gaz", 788, 2, 394, 137, 4, 40, 2000),
        ("Samagaz", "CCG gaz", 430, 1, 430, 151, 4, 40, 2005),
        ("Omaïgaz", "CCG gaz", 860, 2, 430, 151, 4, 40, 2009),
        ("Gastafiore", "CCG gaz", 430, 1, 430, 151, 4, 40, 2015),
        ("Igaznodon", "TAC gaz", 170, 2, 85, 34, 1, 70, 2004),
        ("Cogénération", "cogénération gaz", 882, 1, 882, 0, 0, 70, 2000),
        ("Coron", "charbon", 1200, 2, 600, 210, 8, 36, 1997),
        ("Mockingjay", "charbon", 600, 1, 600, 210, 8, 36, 1990),
        ("Lantier", "charbon", 600, 1, 600, 210, 8, 36, 1984),
        ("Déchets", "déchets", 60, 1, 60, 60, 0, 0, 2000),
        ("Biomasse", "petite biomasse", 365, 1, 365, 0, 0, 0, 2000),
        ("Tacotac", "fioul", 65, 1, 65, 20, 1, 100, 1990),
        ("TicEtTac", "fioul", 95, 1, 95, 30, 1, 100, 2005),
        ("HydroFDE", "hydraulique fil de l'eau", 1000, 1, 1000, 0, 0, 0, 1980),
        ("HydroLac", "hydraulique lac", 6000, 1, 6000, 0, 0, 0, 1980),
        ("Polochon", "STEP", 1200, 1, 1200, 0, 0, 0, 2000),
        ("Eolien", "éolien", 5900, 1, 5900, 0, 0, 0, 2017),
        ("Zéphyr", "éolien", 500, 1, 500, 0, 0, 0, 2018),
        ("Solaire", "solaire", 3000, 1, 3000, 0, 0, 0, 2019)
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
# SECTION 4 : LECTURE DES DONNÉES EXCEL (avec fonctions robustes)
# ==============================================================================

function load_excel_data(data_file::String)
    """Charger les données depuis le fichier Excel avec conversions robustes"""
    
    println("📂 Chargement des données depuis: $data_file")
    
    # Stock hydraulique
    sheet_stock_hydro = "Stock hydro"
    lev_low = XLSX.readdata(data_file, sheet_stock_hydro, "B4:B368")
    lev_high = XLSX.readdata(data_file, sheet_stock_hydro, "C4:C368")
    dates = XLSX.readdata(data_file, sheet_stock_hydro, "A4:A368")
    
    # Nettoyer et construire les données avec fonctions robustes
    lev_low_clean = to_f64(lev_low)
    lev_high_clean = to_f64(lev_high)
    
    stocks_hydro = StockHydro[]
    for i in eachindex(dates)
        date = string(dates[i])
        low = lev_low_clean[i]
        high = lev_high_clean[i]
        if low != 0.0 && high != 0.0  # Éviter les valeurs nulles
            push!(stocks_hydro, StockHydro(date, low, high))
        end
    end
    
    # Apports mensuels et Consommations horaires
    sheet_details = "Détails historique hydro"
    mois = XLSX.readdata(data_file, sheet_details, "A2:A13")
    apports = XLSX.readdata(data_file, sheet_details, "B2:B13")
    
    apports_mensuels = ApportMensuel[]
    for i in eachindex(mois)
        push!(apports_mensuels, ApportMensuel(string(mois[i]), _val_int(apports[i])))
    end
    
    # Consommations horaires
    dates_cons = XLSX.readdata(data_file, sheet_details, "K2:K8761")
    heures_cons = XLSX.readdata(data_file, sheet_details, "L2:L8761")
    fil_de_leau_cons = to_int(XLSX.readdata(data_file, sheet_details, "M2:M8761"))
    lacs_cons = to_int(XLSX.readdata(data_file, sheet_details, "N2:N8761"))
    step_cons = to_int(XLSX.readdata(data_file, sheet_details, "O2:O8761"))
    conso_charge = to_int(XLSX.readdata(data_file, sheet_details, "Q2:Q8761"))
    eolien_cons = to_int(XLSX.readdata(data_file, sheet_details, "R2:R8761"))
    solaire_cons = to_int(XLSX.readdata(data_file, sheet_details, "S2:S8761"))
    
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
const HOURS_PER_DAY = 24
const WINDOW_SIZE = 9 * HOURS_PER_DAY  # 216 heures
const RESULTS_SIZE = 7 * HOURS_PER_DAY  # 168 heures
const YEAR_HOURS = 8760
const HOURS = 1:YEAR_HOURS
const CONSO_ANNUELLE_TWH = 85.0
const HYDRO_SHIFT_DAYS = 181
const HYDRO_SHIFT_HOURS = HYDRO_SHIFT_DAYS * HOURS_PER_DAY

# Heures par mois
const HOURS_PER_MONTH = Dict(
    1 => (744, "Janvier"), 2 => (672, "Février"), 3 => (744, "Mars"),
    4 => (720, "Avril"), 5 => (744, "Mai"), 6 => (720, "Juin"),
    7 => (744, "Juillet"), 8 => (744, "Août"), 9 => (720, "Septembre"),
    10 => (744, "Octobre"), 11 => (720, "Novembre"), 12 => (744, "Décembre")
)

# Hydraulique
const HYDRO_CAPACITY_TWh = 1.0
const HYDRO_CAPACITY_MWh = HYDRO_CAPACITY_TWh * 1e6
const HYDRO_LAC_PMAX = 6000.0

# STEP
const STEP_PMAX = 1200.0
const STEP_RENDEMENT = 0.75
const STEP_CAPACITY_TWh = 0.5
const STEP_CAPACITY_MWh = STEP_CAPACITY_TWh * 1e6

# Pénalités
const PENALTY_UNS = 10000.0
const PENALTY_SPILL = 100.0
const PENALTY_SLACK_SEASONAL = 5000.0
const PENALTY_SLACK_STOCK = 2000.0

# ==============================================================================
# SECTION 6 : DÉCALAGE TEMPOREL HYDRAULIQUE
# ==============================================================================

function recale_hydro_stocks(stocks_hydro::Vector{StockHydro}, shift_days::Int)
    """Décaler les stocks hydrauliques de shift_days jours (1er juillet = jour 1)"""
    n = length(stocks_hydro)
    shifted = StockHydro[]
    
    for i in 1:n
        old_idx = i + shift_days
        if old_idx > n
            old_idx -= n
        end
        push!(shifted, stocks_hydro[old_idx])
    end
    
    return shifted
end

# ==============================================================================
# SECTION 7 : GÉNÉRATION DES DONNÉES ANNUELLES
# ==============================================================================

function generate_complete_year_data(scenario::String)::NTuple{8, Vector{Float64}}
    
    # ====== HYDRAULIQUE, CONSO ET ENR ======
    inflows_fdl = Float64[cons.fil_de_leau for cons in consommations_horaires[1:YEAR_HOURS]]
    load = Float64[cons.conso_charge for cons in consommations_horaires[1:YEAR_HOURS]]
    wind = Float64[cons.eolien_cons for cons in consommations_horaires[1:YEAR_HOURS]]
    solar = Float64[cons.solaire_cons for cons in consommations_horaires[1:YEAR_HOURS]]
    inflows_fdl = recale_hydro_values(inflows_fdl, HYDRO_SHIFT_HOURS)
    load = recale_hydro_values(load, HYDRO_SHIFT_HOURS)
    wind = recale_hydro_values(wind, HYDRO_SHIFT_HOURS)
    solar = recale_hydro_values(solar, HYDRO_SHIFT_HOURS)
    
    # ====== Apports mensuels distribués ======
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
        
        apport_gwh = apports_mensuels_gwh[month][2]
        hours_in_month = HOURS_PER_MONTH[month][1]
        apport_mw_moyen = (apport_gwh * 1000) / hours_in_month
        
        for h_day in 1:24
            variation = 0.85 + 0.30 * rand()
            apport_horaire = apport_mw_moyen * variation
            push!(inflows_lac_opt, apport_horaire)
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
    dechets = fill(60.0, YEAR_HOURS)
    biomasse = fill(365.0 * BIOMASSE_FACTOR, YEAR_HOURS)
    fatale = dechets .+ biomasse
    
    return load, wind, solar, inflows_fdl, inflows_lac, dechets, biomasse, fatale
end

# ==============================================================================
# SECTION 8 : RÉSOLUTION PAR FENÊTRE ROULANTE
# ==============================================================================

function solve_year_rolling(
    load::Vector{Float64},
    wind::Vector{Float64},
    solar::Vector{Float64},
    fdl::Vector{Float64},
    lac::Vector{Float64},
    dechets::Vector{Float64},
    biomasse::Vector{Float64},
    fatale::Vector{Float64},
    scenario::String,
    seasonal_target_factor::Float64,
    unites::Vector{UniteCentrale},
    stocks_hydro::Vector{StockHydro}
)
    """Résoudre l'année complète avec fenêtre roulante"""
    
    println("\n🚀 Début de la résolution pour scénario: $scenario")
    println("   Cible saisonnière: $(seasonal_target_factor * 100)%")
    
    # Séparer les unités par type
    unites_nuc = filter(u -> u.type == "Nucléaire", unites)
    unites_charbon = filter(u -> u.type == "charbon", unites)
    unites_ccg = filter(u -> occursin("CCG", u.type), unites)
    unites_tac = filter(u -> occursin("TAC", u.type), unites)
    unites_cogen = filter(u -> occursin("cogénération", u.type), unites)
    unites_fioul = filter(u -> u.type == "fioul", unites)
    unites_eolien = filter(u -> u.type == "éolien", unites)
    unites_solaire = filter(u -> u.type == "solaire", unites)
    unites_dechets = filter(u -> u.type == "déchets", unites)
    unites_biomasse = filter(u -> occursin("biomasse", u.type), unites)
    
    # Calculer les apports mensuels moyens
    monthly_inflows = Dict{Int, Float64}()
    for m in 1:12
        month_hours = HOURS_PER_MONTH[m][1]
        start_h = sum([HOURS_PER_MONTH[i][1] for i in 1:(m-1)], init=0) + 1
        end_h = start_h + month_hours - 1
        monthly_inflows[m] = mean(lac[start_h:end_h])
    end
    
    # Calculer apports horaires (produit de fil de l'eau + lacs)
    inflows_fdl = fdl
    inflows_lac = lac
    
    # États initiaux
    stock_hydro_current = HYDRO_CAPACITY_MWh * 0.6  # 60% au départ
    stock_step_current = STEP_CAPACITY_MWh * 0.6    # 60% au départ
    
    # Stockage des résultats
    all_results = []
    
    # Fenêtre roulante
    hour = 1
    window_num = 1
    total_windows = div(YEAR_HOURS - WINDOW_SIZE, RESULTS_SIZE) + 1
    
    while hour <= YEAR_HOURS - WINDOW_SIZE + 1
        window_start = hour
        window_end = min(window_start + WINDOW_SIZE - 1, YEAR_HOURS)
        Tmax = window_end - window_start + 1
        
        # Déterminer le mois
        month = 1
        cumul_hours = 0
        for m in 1:12
            cumul_hours += HOURS_PER_MONTH[m][1]
            if window_start <= cumul_hours
                month = m
                break
            end
        end
        
        println("\n" * "="^70)
        println("Fenêtre $window_num/$total_windows | Heures $window_start-$window_end | Mois $month")
        println("="^70)
        println("  Stock hydro initial: $(round(stock_hydro_current/1e6, digits=3)) TWh")
        println("  Stock STEP initial: $(round(stock_step_current/1e6, digits=3)) TWh")
        
        # Extraire les données de la fenêtre
        load_w = load[window_start:window_end]
        wind_w = wind[window_start:window_end]
        solar_w = solar[window_start:window_end]
        fdl_w = fdl[window_start:window_end]
        lac_w = lac[window_start:window_end]
        dechets_w = dechets[window_start:window_end]
        biomasse_w = biomasse[window_start:window_end]
        fatale_w = fatale[window_start:window_end]
        
        # Stocks limites pour cette fenêtre
        stock_limits = []
        for t in window_start:window_end
            idx = ((t - 1) % 365) + 1
            if idx <= length(stocks_hydro)
                stock = stocks_hydro[idx]
                low = stock.lev_low / 100.0 * HYDRO_CAPACITY_MWh
                high = stock.lev_high / 100.0 * HYDRO_CAPACITY_MWh
                push!(stock_limits, (low, high))
            else
                push!(stock_limits, (0.3 * HYDRO_CAPACITY_MWh, 0.8 * HYDRO_CAPACITY_MWh))
            end
        end
        
        # Cible saisonnière médiane
        target_seasonal = [seasonal_target_factor * (sl[1] + sl[2]) / 2 for sl in stock_limits]
        
        # Créer le modèle d'optimisation
        model = Model(HiGHS.Optimizer)
        set_optimizer_attribute(model, "time_limit", 300.0)
        set_optimizer_attribute(model, "output_flag", false)
        
        # Variables de production
        @variable(model, P_nuc[1:Tmax, 1:length(unites_nuc)] >= 0)
        @variable(model, P_charbon[1:Tmax, 1:length(unites_charbon)] >= 0)
        @variable(model, P_ccg[1:Tmax, 1:length(unites_ccg)] >= 0)
        @variable(model, P_tac[1:Tmax, 1:length(unites_tac)] >= 0)
        @variable(model, P_cogen[1:Tmax, 1:length(unites_cogen)] >= 0)
        @variable(model, P_fioul[1:Tmax, 1:length(unites_fioul)] >= 0)
        @variable(model, P_eolien[1:Tmax, 1:length(unites_eolien)] >= 0)
        @variable(model, P_solaire[1:Tmax, 1:length(unites_solaire)] >= 0)
        @variable(model, P_dechets[1:Tmax, 1:length(unites_dechets)] >= 0)
        @variable(model, P_biomasse[1:Tmax, 1:length(unites_biomasse)] >= 0)
        
        # Variables hydrauliques
        @variable(model, Phy_fdl[1:Tmax] >= 0)
        @variable(model, Phy_lac[1:Tmax] >= 0)
        @variable(model, stock_hydro_vals[1:Tmax] >= 0)
        
        # Variables STEP
        @variable(model, Pcharge_STEP[1:Tmax] >= 0)
        @variable(model, Pdecharge_STEP[1:Tmax] >= 0)
        @variable(model, stock_STEP[1:Tmax] >= 0)
        
        # Variables de défaillance
        @variable(model, Puns[1:Tmax] >= 0)
        @variable(model, Pspill[1:Tmax] >= 0)
        
        # Variables binaires UC
        @variable(model, UC_nuc[1:Tmax, 1:length(unites_nuc)], Bin)
        @variable(model, UC_charbon[1:Tmax, 1:length(unites_charbon)], Bin)
        @variable(model, UC_ccg[1:Tmax, 1:length(unites_ccg)], Bin)
        @variable(model, UC_tac[1:Tmax, 1:length(unites_tac)], Bin)
        @variable(model, UC_cogen[1:Tmax, 1:length(unites_cogen)], Bin)
        @variable(model, UC_fioul[1:Tmax, 1:length(unites_fioul)], Bin)
        
        # Variables de slack
        @variable(model, slack_seasonal[1:Tmax] >= 0)
        @variable(model, slack_stock_min[1:Tmax] >= 0)
        @variable(model, slack_stock_max[1:Tmax] >= 0)
        
        # 🆕 NOUVELLES VARIABLES : Tranches d'eau (idée de optim_hydroVF.jl)
        @variable(model, S_socle[1:Tmax] >= 0)      # Tranche socle (à conserver)
        @variable(model, S_surplus[1:Tmax] >= 0)    # Tranche surplus (utilisable librement)
        
        # Fonction objectif
        @objective(model, Min,
            # Coûts de production
            sum(unites_nuc[g].prix_marche * P_nuc[t, g] for t in 1:Tmax, g in 1:length(unites_nuc)) +
            sum(unites_charbon[g].prix_marche * P_charbon[t, g] for t in 1:Tmax, g in 1:length(unites_charbon)) +
            sum(unites_ccg[g].prix_marche * P_ccg[t, g] for t in 1:Tmax, g in 1:length(unites_ccg)) +
            sum(unites_tac[g].prix_marche * P_tac[t, g] for t in 1:Tmax, g in 1:length(unites_tac)) +
            sum(unites_cogen[g].prix_marche * P_cogen[t, g] for t in 1:Tmax, g in 1:length(unites_cogen)) +
            sum(unites_fioul[g].prix_marche * P_fioul[t, g] for t in 1:Tmax, g in 1:length(unites_fioul)) +
            # Pénalités
            sum(PENALTY_UNS * Puns[t] for t in 1:Tmax) +
            sum(PENALTY_SPILL * Pspill[t] for t in 1:Tmax) +
            sum(PENALTY_SLACK_SEASONAL * slack_seasonal[t] for t in 1:Tmax) +
            sum(PENALTY_SLACK_STOCK * (slack_stock_min[t] + slack_stock_max[t]) for t in 1:Tmax)
        )
        
        # Contraintes d'équilibre
        @constraint(model, [t=1:Tmax],
            sum(P_nuc[t, :]) + sum(P_charbon[t, :]) + sum(P_ccg[t, :]) +
            sum(P_tac[t, :]) + sum(P_cogen[t, :]) + sum(P_fioul[t, :]) +
            sum(P_eolien[t, :]) + sum(P_solaire[t, :]) +
            sum(P_dechets[t, :]) + sum(P_biomasse[t, :]) +
            Phy_fdl[t] + Phy_lac[t] + Pdecharge_STEP[t] + Puns[t]
            ==
            load_w[t] + Pcharge_STEP[t] + Pspill[t]
        )
        
        # Contraintes de capacité ENR
        @constraint(model, [t=1:Tmax], sum(P_eolien[t, :]) <= wind_w[t])
        @constraint(model, [t=1:Tmax], sum(P_solaire[t, :]) <= solar_w[t])
        
        # Contraintes fatales
        @constraint(model, [t=1:Tmax], sum(P_dechets[t, :]) == dechets_w[t])
        @constraint(model, [t=1:Tmax], sum(P_biomasse[t, :]) == biomasse_w[t])
        
        # Contraintes hydrauliques fil de l'eau
        @constraint(model, [t=1:Tmax], Phy_fdl[t] <= fdl_w[t])
        
        # Contraintes hydrauliques lacs avec tranches
        @constraint(model, [t=1:Tmax], Phy_lac[t] <= HYDRO_LAC_PMAX)
        
        # 🆕 CONTRAINTES TRANCHES D'EAU (idée de optim_hydroVF.jl)
        # Le stock total = socle + surplus
        @constraint(model, [t=1:Tmax], stock_hydro_vals[t] == S_socle[t] + S_surplus[t])
        
        # La tranche socle ne dépasse pas la cible médiane
        @constraint(model, [t=1:Tmax], S_socle[t] <= target_seasonal[t] + slack_seasonal[t])
        
        # Contraintes de stock min/max (avec slack)
        @constraint(model, [t=1:Tmax], 
            stock_hydro_vals[t] >= stock_limits[t][1] - slack_stock_min[t])
        @constraint(model, [t=1:Tmax], 
            stock_hydro_vals[t] <= stock_limits[t][2] + slack_stock_max[t])
        
        # Dynamique du stock hydraulique
        @constraint(model, stock_hydro_vals[1] == stock_hydro_current)
        @constraint(model, [t=1:Tmax-1],
            stock_hydro_vals[t+1] == stock_hydro_vals[t] - Phy_lac[t] + lac_w[t]
        )
        
        # Contraintes STEP
        @constraint(model, [t=1:Tmax], Pcharge_STEP[t] <= STEP_PMAX)
        @constraint(model, [t=1:Tmax], Pdecharge_STEP[t] <= STEP_PMAX)
        @constraint(model, [t=1:Tmax], stock_STEP[t] <= STEP_CAPACITY_MWh)
        @constraint(model, stock_STEP[1] == stock_step_current)
        
        # Dynamique STEP
        @constraint(model, [t=1:Tmax-1],
            stock_STEP[t+1] == stock_STEP[t] + STEP_RENDEMENT * Pcharge_STEP[t] - Pdecharge_STEP[t]
        )
        
        # Fonction helper pour contraintes UC
        function add_uc_constraints!(P, UC, units, Tmax)
            for (g, unit) in enumerate(units)
                # Limites de puissance
                @constraint(model, [t=1:Tmax],
                    P[t, g] >= unit.Pmin * UC[t, g])
                @constraint(model, [t=1:Tmax],
                    P[t, g] <= unit.Pmax * UC[t, g])
                
                # Durée minimale (simple)
                if unit.dmin > 0
                    @constraint(model, [t=unit.dmin:Tmax],
                        UC[t, g] >= sum(UC[i, g] - UC[i-1, g] 
                            for i in max(1, t-unit.dmin+1):t if i > 1)
                    )
                end
            end
        end
        
        # Appliquer les contraintes UC
        add_uc_constraints!(P_nuc, UC_nuc, unites_nuc, Tmax)
        add_uc_constraints!(P_charbon, UC_charbon, unites_charbon, Tmax)
        add_uc_constraints!(P_ccg, UC_ccg, unites_ccg, Tmax)
        add_uc_constraints!(P_tac, UC_tac, unites_tac, Tmax)
        add_uc_constraints!(P_cogen, UC_cogen, unites_cogen, Tmax)
        add_uc_constraints!(P_fioul, UC_fioul, unites_fioul, Tmax)
        
        # Résolution
        println("  ⏳ Résolution en cours...")
        optimize!(model)
        status = termination_status(model)
        
        println("  ✓ Statut: $status")
        
        # Extraction des résultats
        stock_hydro_before = stock_hydro_current
        hours_to_keep = min(RESULTS_SIZE, Tmax)
        
        # Gestion du statut
        if status == OPTIMAL
            stock_hydro_current = value(stock_hydro_vals[hours_to_keep])
            stock_step_current = value(stock_STEP[hours_to_keep])
            
            # Vérifications
            if !isfinite(stock_hydro_current) || stock_hydro_current < 0
                stock_hydro_current = HYDRO_CAPACITY_MWh * 0.5
            end
            
            if stock_step_current < STEP_CAPACITY_MWh * 0.3
                stock_step_current = STEP_CAPACITY_MWh * 0.3
            end
            
            println("  Stock hydro final: $(round(stock_hydro_current/1e6, digits=3)) TWh")
            println("  Stock STEP final: $(round(stock_step_current/1e6, digits=3)) TWh")
        else
            println("  ⚠️  Statut non-optimal: $status")
            stock_hydro_current = stock_hydro_before
            stock_step_current = STEP_CAPACITY_MWh * 0.5
        end
        
        # Stocker les résultats
        for t in 1:hours_to_keep
            # 🆕 DÉCOMMENTER : Extraire la production hydraulique des lacs
            prod_hydro_lac = value(Phy_lac[t])
            
            push!(all_results, Dict(
                "hour_global" => window_start + t - 1,
                "month" => month,
                "load" => load[window_start + t - 1],
                "Phy_fdl" => value(Phy_fdl[t]),
                "Phy_lac" => prod_hydro_lac,  # 🆕 Production lacs décommentée
                "P_nuc" => sum(value(P_nuc[t, :])),
                "P_charbon" => sum(value(P_charbon[t, :])),
                "P_ccg" => sum(value(P_ccg[t, :])),
                "P_tac" => sum(value(P_tac[t, :])),
                "P_cogen" => sum(value(P_cogen[t, :])),
                "P_fioul" => sum(value(P_fioul[t, :])),
                "P_eolien" => sum(value(P_eolien[t, :])),
                "P_solaire" => sum(value(P_solaire[t, :])),
                "P_dechets" => sum(value(P_dechets[t, :])),
                "P_biomasse" => sum(value(P_biomasse[t, :])),
                "stock_hydro" => value(stock_hydro_vals[t]),
                "stock_STEP" => value(stock_STEP[t]),
                "Puns" => value(Puns[t]),
                "Pspill" => value(Pspill[t]),
                "UC_nuc" => sum(value(UC_nuc[t, :])),
                "UC_charbon" => sum(value(UC_charbon[t, :])),
                "UC_ccg" => sum(value(UC_ccg[t, :])),
                "UC_tac" => sum(value(UC_tac[t, :])),
                "UC_cogen" => sum(value(UC_cogen[t, :])),
                "UC_fioul" => sum(value(UC_fioul[t, :])),
                "Pcharge_STEP" => value(Pcharge_STEP[t]),
                "Pdecharge_STEP" => value(Pdecharge_STEP[t]),
                "stock_hydro_min" => stock_limits[t][1],
                "stock_hydro_max" => stock_limits[t][2],
                "slack_seasonal" => value(slack_seasonal[t]),
                "slack_stock_min" => value(slack_stock_min[t]),
                "slack_stock_max" => value(slack_stock_max[t]),
                "S_socle" => value(S_socle[t]),        # 🆕 Tranche socle
                "S_surplus" => value(S_surplus[t]),    # 🆕 Tranche surplus
                "inflows_lac" => lac[window_start + t - 1],
                "inflows_fdl" => fdl[window_start + t - 1],
                "infeasible_flag" => (status != OPTIMAL ? 1 : 0)
            ))
        end
        
        hour = window_start + RESULTS_SIZE
        window_num += 1
    end
    
    return all_results
end

# ==============================================================================
# SECTION 9 : EXÉCUTION PRINCIPALE
# ==============================================================================

println("\n" * "="^80)
println("EOD ZOOTOPIA - OPTIMISATION DISPATCHING ÉLECTRIQUE (VERSION AMÉLIORÉE)")
println("="^80 * "\n")

# Parser les arguments
params = parse_args(ARGS)
scenario_to_run = params["scenario"]
data_file = params["data_file"]

println("📋 Paramètres:")
println("   Scénario: $scenario_to_run")
println("   Fichier données: $data_file")

# Initialiser
unites = initialize_unites()
stocks_hydro, apports_mensuels, consommations_horaires = load_excel_data(data_file)
stocks_hydro = recale_hydro_stocks(stocks_hydro, HYDRO_SHIFT_DAYS)

# Paramètres scénarios
scenario_params = Dict(
    "dry" => 0.60,
    "normal" => 0.70,
    "wet" => 0.80
)

results_all = Dict()

# Si un scénario spécifique est demandé, ne résoudre que celui-là
scenarios_to_solve = if scenario_to_run in ["dry", "normal", "wet"]
    [scenario_to_run]
else
    ["dry", "normal", "wet"]
end

# Résoudre pour chaque scénario
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
    
    # Statistiques
    println("\n📊 Résumé annuel $scenario :")
    println("  Prod. nucléaire : $(round(sum(r["P_nuc"] for r in results)/1e6, digits=2)) TWh")
    println("  Prod. charbon   : $(round(sum(r["P_charbon"] for r in results)/1e6, digits=2)) TWh")
    println("  Prod. gaz (CCG) : $(round(sum(r["P_ccg"] for r in results)/1e6, digits=2)) TWh")
    println("  Prod. éolienne  : $(round(sum(r["P_eolien"] for r in results)/1e6, digits=2)) TWh")
    println("  Prod. solaire   : $(round(sum(r["P_solaire"] for r in results)/1e6, digits=2)) TWh")
    println("  Prod. hydro FDE : $(round(sum(r["Phy_fdl"] for r in results)/1e6, digits=2)) TWh")
    println("  🆕 Prod. hydro lacs : $(round(sum(r["Phy_lac"] for r in results)/1e6, digits=2)) TWh")  # 🆕 Exporté
    println("  Défaillance     : $(round(sum(r["Puns"] for r in results)/1e6, digits=6)) TWh")
    println("  Spill hydro     : $(round(sum(r["Pspill"] for r in results)/1e6, digits=2)) TWh")
end

# Exporter résultats dans results/
mkpath("./results")

for scenario in keys(results_all)
    df = DataFrame(
        hour = [r["hour_global"] for r in results_all[scenario]],
        month = [r["month"] for r in results_all[scenario]],
        load = [r["load"] for r in results_all[scenario]],
        Phy_fdl = [r["Phy_fdl"] for r in results_all[scenario]],
        Phy_lac = [r["Phy_lac"] for r in results_all[scenario]],  # 🆕 Exporté
        P_nuc = [r["P_nuc"] for r in results_all[scenario]],
        P_charbon = [r["P_charbon"] for r in results_all[scenario]],
        P_ccg = [r["P_ccg"] for r in results_all[scenario]],
        P_tac = [r["P_tac"] for r in results_all[scenario]],
        P_cogen = [r["P_cogen"] for r in results_all[scenario]],
        P_fioul = [r["P_fioul"] for r in results_all[scenario]],
        P_eolien = [r["P_eolien"] for r in results_all[scenario]],
        P_solaire = [r["P_solaire"] for r in results_all[scenario]],
        P_dechets = [r["P_dechets"] for r in results_all[scenario]],
        P_biomasse = [r["P_biomasse"] for r in results_all[scenario]],
        stock_hydro = [r["stock_hydro"] for r in results_all[scenario]],
        stock_STEP = [r["stock_STEP"] for r in results_all[scenario]],
        Puns = [r["Puns"] for r in results_all[scenario]],
        Pspill = [r["Pspill"] for r in results_all[scenario]],
        UC_nuc = [r["UC_nuc"] for r in results_all[scenario]],
        UC_charbon = [r["UC_charbon"] for r in results_all[scenario]],
        UC_ccg = [r["UC_ccg"] for r in results_all[scenario]],
        UC_tac = [r["UC_tac"] for r in results_all[scenario]],
        UC_cogen = [r["UC_cogen"] for r in results_all[scenario]],
        UC_fioul = [r["UC_fioul"] for r in results_all[scenario]],
        Pdecharge_STEP = [r["Pdecharge_STEP"] for r in results_all[scenario]],
        Pcharge_STEP = [r["Pcharge_STEP"] for r in results_all[scenario]],
        stock_hydro_min = [r["stock_hydro_min"] for r in results_all[scenario]],
        stock_hydro_max = [r["stock_hydro_max"] for r in results_all[scenario]],
        slack_seasonal = [r["slack_seasonal"] for r in results_all[scenario]],
        slack_stock_min = [r["slack_stock_min"] for r in results_all[scenario]],
        slack_stock_max = [r["slack_stock_max"] for r in results_all[scenario]],
        S_socle = [r["S_socle"] for r in results_all[scenario]],          # 🆕 Tranche socle
        S_surplus = [r["S_surplus"] for r in results_all[scenario]],      # 🆕 Tranche surplus
        inflows_lac = [r["inflows_lac"] for r in results_all[scenario]],
        inflows_fdl = [r["inflows_fdl"] for r in results_all[scenario]],
        infeasible_flag = [r["infeasible_flag"] for r in results_all[scenario]]
    )
    
    filepath = "./results/results_$scenario.csv"
    CSV.write(filepath, df)
    println("✅ Exporté : $filepath")
end

println("\n" * "="^80)
println("✅ SIMULATION TERMINÉE AVEC SUCCÈS")
println("="^80 * "\n")
println("📁 Résultats disponibles dans: ./results/")
println("📊 Production hydraulique des lacs exportée dans la colonne 'Phy_lac'")
println("🆕 Nouvelles colonnes : S_socle, S_surplus (tranches d'eau)")
