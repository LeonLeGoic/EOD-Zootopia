##############################################################################
#  Lancement des différents scénarios directement depuis un terminal:
#
#    julia optim_hydro.jl                         # scénario de référence
#    julia optim_hydro.jl --scenario sec        # année sèche
#    julia optim_hydro.jl --scenario humide       # année humide
#    julia optim_hydro.jl --semaines 10           # modélisation de 10 semaines seulement
#    julia optim_hydro.jl --bord 5                # nombres de jours ajoutés pour contrer l'effet de bord
#    julia optim_hydro.jl --apport_fois 1.5       # multiplicateur d'apport en eau manuel
#
#  Entrées (dossier Donnees/) :
#    • Donnees_prod.xlsx    — parc électrique et données hydraulique
#    • Donnees_conso.xlsx   — consommation et productions fatales
#
#  Sorties (dossier Resultats/) :
#    • final_results_<scenario>.xlsx
##############################################################################

##############################################################################
# Packages
##############################################################################
using JuMP
using HiGHS
using  XLSX
using DataFrames

##############################################################################
# Paramètres des arguments d'entrée au lancement de l'optimisation
##############################################################################

function parse_args(args)
    p = Dict(
        "scenario"     => "reference",
        "semaines"     => 52,
        "week_start"   => 1,
        "bord"         => 3,
        "min_en_plus"  => 0.0,
        "apport_fois"  => -1.0,   
    )
    i = 1
    while i <= length(args)
        k = args[i]
        if i+1 <= length(args)
            v = args[i+1]
            if k == "--scenario";    p["scenario"]    = v;              i += 2
            elseif k == "--semaines";    p["semaines"]    = parse(Int, v);   i += 2
            elseif k == "--bord";        p["bord"]        = parse(Int, v);   i += 2
            elseif k == "--min_en_plus"; p["min_en_plus"] = parse(Float64,v);i += 2
            elseif k == "--apport_fois"; p["apport_fois"] = parse(Float64,v);i += 2
            else i += 1 end
        else i += 1 end
    end
    return p
end

params       = parse_args(ARGS)
SCENARIO     = params["scenario"]
NB_SEMAINES  = params["semaines"]
WEEK_START   = params["week_start"]
JOUR_EN_TROP = params["bord"]
MIN_EN_PLUS  = params["min_en_plus"]

# Valeur de l'apport hydraulique selon le scénario (année sèche, année humide) :
apport_fois = if params["apport_fois"] > 0
    params["apport_fois"]           # valeur explicitement passée en argument
elseif SCENARIO == "sec"
    0.5                             # année sèche : apports divisés par 2 (expliqué dans le rapport)
elseif SCENARIO == "humide"
    2.0                             # année humide : apports multipliés par 2
else
    1.0                             # scénario de référence
end

println("="^60)
println("Scénario : $SCENARIO  |  semaines : $NB_SEMAINES")
println("Jours de bord : $JOUR_EN_TROP  |  apport × $apport_fois")
println("="^60)

mkpath("Resultats")

##############################################################################
# Données d'entrée
##############################################################################

data_prod  = joinpath(@__DIR__, "Donnees", "Donnees_prod.xlsx")
data_conso = joinpath(@__DIR__, "Donnees", "Donnees_conso.xlsx")

##############################################################################
# Fonctions de conversion (pour éviter les erreurs liées à des cellules vides ou des nombres non sous le format Float64)
##############################################################################

# Pour les erreurs sur les Float64 :
function _val_f64(val)
    if ismissing(val) || val === nothing
        return 0.0
    elseif val isa Bool
        return Float64(val)
    elseif val isa Number
        return Float64(val)
    else
        return parse(Float64, strip(string(val)))
    end
end

# Pour éviter les erreurs sur les Vector{Float64} :
function to_f64(x)
    return [_val_f64(v) for v in vec(x)]
end

# Pour les erreurs sur les Int (durées minimales) :
function _val_int(val)
    ismissing(val) || val === nothing && return 0
    isa(val, Bool)   && return Int(val)
    isa(val, Number) && return Int(round(Float64(val)))
    return parse(Int, strip(string(val)))
end

# Pour éviter les erreurs sur les Vector{Int} :
function to_int(x)
    return [_val_int(v) for v in vec(x)]
end

# Pour extraitre une scalaire Float64 d'une cellule XLSX :
function scalar_f64(x)
    val = isa(x, AbstractArray) ? x[1] : x
    ismissing(val) || val === nothing && return 0.0
    isa(val, Bool) && return Float64(val)
    isa(val, Number) && return Float64(val)
    return parse(Float64, strip(string(val)))
end

##############################################################################
# Parc de production électrique 
##############################################################################

# Nucléaire :
# 3 centrales de 2 unités donc 6 unités au total  

Nnuc_raw   = 3
names_nuc_raw  = string.(vec(XLSX.readdata(data_prod, "Parc électrique", "A2:A4")))
costs_nuc_raw  = to_f64(XLSX.readdata(data_prod, "Parc électrique", "H2:H4"))
Pmin_nuc_raw   = to_f64(XLSX.readdata(data_prod, "Parc électrique", "F2:F4"))
Pmax_nuc_raw   = to_f64(XLSX.readdata(data_prod, "Parc électrique", "E2:E4"))
dmin_nuc_raw   = to_int(XLSX.readdata(data_prod, "Parc électrique", "G2:G4"))

const NB_UNITS_NUC = 2 # dédoublement des 2 unités par centrale

names_nuc = String[]
costs_nuc = Float64[]
Pmin_nuc  = Float64[]
Pmax_nuc  = Float64[]
dmin_nuc  = Int[]

for i in 1:Nnuc_raw
    for u in 1:NB_UNITS_NUC
        push!(names_nuc, names_nuc_raw[i] * "_U$u")
        push!(costs_nuc, costs_nuc_raw[i])
        push!(Pmin_nuc,  Pmin_nuc_raw[i])
        push!(Pmax_nuc,  Pmax_nuc_raw[i])
        push!(dmin_nuc,  dmin_nuc_raw[i])
    end
end

Nnuc = length(names_nuc)

# Gaz :
# 5 centrales et 7 unités

Ngaz_raw   = 5
names_gaz_raw  = string.(vec(XLSX.readdata(data_prod, "Parc électrique", "A5:A9")))
nb_units_gaz   = to_int(XLSX.readdata(data_prod, "Parc électrique", "D5:D9")) 
costs_gaz_raw  = to_f64(XLSX.readdata(data_prod, "Parc électrique", "H5:H9"))
Pmin_gaz_raw   = to_f64(XLSX.readdata(data_prod, "Parc électrique", "F5:F9"))
Pmax_gaz_raw   = to_f64(XLSX.readdata(data_prod, "Parc électrique", "E5:E9"))
dmin_gaz_raw   = to_int(XLSX.readdata(data_prod, "Parc électrique", "G5:G9"))

names_gaz = String[]
costs_gaz = Float64[]
Pmin_gaz  = Float64[]
Pmax_gaz  = Float64[]
dmin_gaz  = Int[]

for i in 1:Ngaz_raw # dédoublement des centrales multi-unités
    n_units = nb_units_gaz[i] > 0 ? nb_units_gaz[i] : 1
    for u in 1:n_units
        suffix = n_units > 1 ? "_U$u" : ""
        push!(names_gaz, names_gaz_raw[i] * suffix)
        push!(costs_gaz, costs_gaz_raw[i])
        push!(Pmin_gaz,  Pmin_gaz_raw[i])
        push!(Pmax_gaz,  Pmax_gaz_raw[i])
        push!(dmin_gaz,  dmin_gaz_raw[i])
    end
end

# et une centrale TAC avec 2 unités :
Ntac       = 2
names_tac  = "Igaznodon"
cost_tac  = scalar_f64(XLSX.readdata(data_prod, "Parc électrique", "H10"))
Pmin_tac  = scalar_f64(XLSX.readdata(data_prod, "Parc électrique", "F10"))
Pmax_tac  = scalar_f64(XLSX.readdata(data_prod, "Parc électrique", "E10"))
dmin_tac  = Int(round(scalar_f64(XLSX.readdata(data_prod, "Parc électrique", "G10"))))

for u in 1:Ntac
    push!(names_gaz, "$(names_tac)_U$u")
    push!(costs_gaz, cost_tac)
    push!(Pmin_gaz,  Pmin_tac)
    push!(Pmax_gaz,  Pmax_tac)
    push!(dmin_gaz,  dmin_tac)
end

Ngaz = length(names_gaz)

# Cogénération :

Ncogen     = 1
costs_cogen = [scalar_f64(XLSX.readdata(data_prod, "Parc électrique", "H11"))]
Pmin_cogen  = zeros(Ncogen)
Pmax_cogen  = [scalar_f64(XLSX.readdata(data_prod, "Parc électrique", "C11"))]


# Charbon :
# 3 centrales dont une de 2 unités donc 4 unités au total

Ncoal_raw      = 3
names_coal_raw = string.(vec(XLSX.readdata(data_prod, "Parc électrique", "A12:A14")))
nb_units_coal  = to_int(XLSX.readdata(data_prod, "Parc électrique", "D12:D14"))
costs_coal_raw = to_f64(XLSX.readdata(data_prod, "Parc électrique", "H12:H14"))
Pmin_coal_raw  = to_f64(XLSX.readdata(data_prod, "Parc électrique", "F12:F14"))
Pmax_coal_raw  = to_f64(XLSX.readdata(data_prod, "Parc électrique", "E12:E14"))
dmin_coal_raw  = to_int(XLSX.readdata(data_prod, "Parc électrique", "G12:G14"))

names_coal = String[]
costs_coal = Float64[]
Pmin_coal  = Float64[]
Pmax_coal  = Float64[]
dmin_coal  = Int[]

for i in 1:Ncoal_raw
    n_units = nb_units_coal[i] > 0 ? nb_units_coal[i] : 1
    for u in 1:n_units
        suffix = n_units > 1 ? "_U$u" : ""
        push!(names_coal, names_coal_raw[i] * suffix)
        push!(costs_coal, costs_coal_raw[i])
        push!(Pmin_coal,  Pmin_coal_raw[i])
        push!(Pmax_coal,  Pmax_coal_raw[i])
        push!(dmin_coal,  dmin_coal_raw[i])
    end
end

Ncoal = length(names_coal)

# Déchets et biomasse
# pris en compte ensemble (60 et 365)

Pmax_biomasse = 425.0

# Fioul :
# 2 centrales 

Nfuel      = 2
names_fuel = string.(vec(XLSX.readdata(data_prod, "Parc électrique", "A17:A18")))
costs_fuel = to_f64(XLSX.readdata(data_prod, "Parc électrique", "H17:H18"))
Pmin_fuel  = to_f64(XLSX.readdata(data_prod, "Parc électrique", "F17:F18"))
Pmax_fuel  = to_f64(XLSX.readdata(data_prod, "Parc électrique", "E17:E18"))
dmin_fuel  = to_int(XLSX.readdata(data_prod, "Parc électrique", "G17:G18"))


# Hydraulique (réservoir) :
Nhy           = 1
Pmin_hy       = zeros(Nhy)
Pmax_hy       = [scalar_f64(XLSX.readdata(data_prod, "Parc électrique", "C20"))]
costs_hy      = zeros(Nhy)
e_hy_stockmax = [scalar_f64(XLSX.readdata(data_prod, "Stock hydro", "B1")) * 1_000_000] # Valeur en TWh convertie en MWh

#  STEP :
Pmax_STEP = scalar_f64(XLSX.readdata(data_prod, "Parc électrique", "E21"))
rSTEP     = 0.75

##############################################################################
# Gestion des transitions entres semaines (effets de bord)
##############################################################################
# On transmet les 24 premières heures de la zone "débordement" de la semaine k comme conditions initiales forcées de la semaine k+1.
# 24h = 1 journée complète, ce qui dépasse la durée minimale maximale du parc (dmin du nucléaire).
# Voir le rapport section 2.2 pour l'explication du phénomène d'effet de bord. 

mutable struct EtatBord
    Pnuc   :: Matrix{Float64}   
    Pgaz   :: Matrix{Float64}   
    Pcoal  :: Matrix{Float64}   
    Pfuel  :: Matrix{Float64}   
    UCnuc  :: Matrix{Float64};  UPnuc  :: Matrix{Float64};  DOnuc  :: Matrix{Float64}
    UCgaz  :: Matrix{Float64};  UPgaz  :: Matrix{Float64};  DOgaz  :: Matrix{Float64}
    UCcoal :: Matrix{Float64};  UPcoal :: Matrix{Float64};  DOcoal :: Matrix{Float64}
    UCfuel :: Matrix{Float64};  UPfuel :: Matrix{Float64};  DOfuel :: Matrix{Float64}
    stock_hy_pct :: Float64     # pour transmettre le pourcentage de remplissage du stock hydraulique
end

bord_actuel = nothing   # rempli après la semaine 1

##############################################################################
# Initialisation du tableau de résultats
##############################################################################
col_names = vcat(
    names_nuc, names_gaz, names_coal, names_fuel,
    ["Cogén", "Biomasse", "Hydro", "STEP turbinage", "Puissance résiduelle",
     "STEP pompage", "Conso", "Conso nette", "Puissance excès", "Déficit offre","Stock eau (%)"],
    ["Semaine", "Jour", "Heure", "Jour cumulé", "Heure cumulée"]
)

final_df = DataFrame([n => Float64[] for n in col_names]...)


##############################################################################
# Lecture anticipée de toutes les données 
##############################################################################
# On lit les données en dehors de la boucle d'optimisation pour gagner du temps de calcul

println("Lecture des données annuelles...")
t_data = time()

# On lit d'un coup 8760 heures (1 an).
H_TOT = 8760

# Données de consommation et production fatale
conso_all         = to_f64(XLSX.readdata(data_conso, "TP", "C10:C$(9+H_TOT)"))
wind_all          = to_f64(XLSX.readdata(data_conso, "TP", "D10:D$(9+H_TOT)"))
solar_all         = to_f64(XLSX.readdata(data_conso, "TP", "E10:E$(9+H_TOT)"))
thermal_fatal_all = to_f64(XLSX.readdata(data_conso, "TP", "G10:G$(9+H_TOT)"))

# Données de production hydraulique fatale
hydro_fatal_all = to_f64(XLSX.readdata(data_prod, "Détails historique hydro", "M2:M$(1+H_TOT)"))

# Données du stock hydraulique
e_hy_low_pct_all  = to_f64(XLSX.readdata(data_prod, "Stock hydro", "M4:M$(3+H_TOT)"))
e_hy_high_pct_all = to_f64(XLSX.readdata(data_prod, "Stock hydro", "N4:N$(3+H_TOT)"))
apports_raw_all   = to_f64(XLSX.readdata(data_prod, "Stock hydro", "O4:O$(3+H_TOT)"))

Pres_all = wind_all .+ solar_all .+ hydro_fatal_all .+ thermal_fatal_all 

println("   → terminées en $(round(time() - t_data, digits=1))s")


##############################################################################
# Boucle hebdomadaire principale
##############################################################################
t_total_start = time()

for semaine in WEEK_START:(WEEK_START + NB_SEMAINES - 1)

    println("\n── Semaine $semaine / $(WEEK_START + NB_SEMAINES - 1) ──")
    t_sem = time()

    week = semaine - 1   # On commence à la semaine 0.

    # Fenêtre temporelle et sélection des données pré-chargées :
    Tmax             = (7 + JOUR_EN_TROP) * 24
    start_h          = 7 * 24 * week + 1
    end_h            = start_h + Tmax - 1

    if end_h > H_TOT
        println("Attention : l'horizon de simulation dépasse les données chargées (", H_TOT, "h)")
        # On tronque pour éviter une erreur
        end_h = H_TOT
        if start_h > end_h
            println("Plus de données à traiter, arrêt de la boucle.")
            break
        end
        # Corrige Tmax pour qu'il corresponde à la taille réelle des données
        Tmax = end_h - start_h + 1
    end

    h_range = start_h:end_h

    # Sélection des données pour la semaine en cours depuis les vecteurs pré-chargés
    conso         = conso_all[h_range]
    Pres          = Pres_all[h_range]
    
    e_hy_low_pct  = e_hy_low_pct_all[h_range]
    e_hy_high_pct = e_hy_high_pct_all[h_range]
    apports_raw   = apports_raw_all[h_range]

    e_hy_low  = e_hy_stockmax[1] .* (e_hy_low_pct  ./ 100) .+ e_hy_stockmax[1] .* MIN_EN_PLUS ./ 100
    e_hy_high = e_hy_stockmax[1] .* (e_hy_high_pct ./ 100)
    e_hy_cible = (e_hy_low .+ e_hy_high) ./ 2.0 # Pour la régularisation vers une cible médiane

    apports_hydro = apport_fois .* apports_raw

    # Stock initial du réservoir 
    if week == 0 # pour la première semaine, on prend une valeur moyenne 
        e_hy_init_val = (e_hy_high[1] + e_hy_low[1]) / 2.0
    else
        e_hy_init_val = bord_actuel.stock_hy_pct / 100.0 * e_hy_stockmax[1] + apports_hydro[1]
        # On borne par sécurité pour éviter des infaisabilités numériques
        e_hy_init_val = clamp(e_hy_init_val, e_hy_low[1], e_hy_high[1])
    end

    # Modèle d'optimisation :

    model = Model(HiGHS.Optimizer)
    set_silent(model)
    set_attribute(model, "mip_rel_gap", 0.01) # Tolérance de 1% pour diminuer le temps de calcul.


    # Matrices de coûts répétées sur l'horizon Tmax

    cnuc   = repeat(costs_nuc',   Tmax)
    cgaz   = repeat(costs_gaz',   Tmax)
    ccoal  = repeat(costs_coal',  Tmax)
    cfuel  = repeat(costs_fuel',  Tmax)
    ccogen = repeat(costs_cogen', Tmax)
    chy    = repeat(costs_hy',    Tmax)
    cuns   = fill(5000.0, Tmax)
    cexc   = fill(0.0,    Tmax)

    # Variables :

    @variable(model, Pnuc[1:Tmax,    1:Nnuc]  >= 0)
    @variable(model, UCnuc[1:Tmax,   1:Nnuc], Bin)
    @variable(model, UPnuc[1:Tmax,   1:Nnuc], Bin)
    @variable(model, DOnuc[1:Tmax,   1:Nnuc], Bin)

    @variable(model, Pgaz[1:Tmax,    1:Ngaz]  >= 0)
    @variable(model, UCgaz[1:Tmax,   1:Ngaz], Bin)
    @variable(model, UPgaz[1:Tmax,   1:Ngaz], Bin)
    @variable(model, DOgaz[1:Tmax,   1:Ngaz], Bin)

    @variable(model, Pcoal[1:Tmax,   1:Ncoal] >= 0)
    @variable(model, UCcoal[1:Tmax,  1:Ncoal], Bin)
    @variable(model, UPcoal[1:Tmax,  1:Ncoal], Bin)
    @variable(model, DOcoal[1:Tmax,  1:Ncoal], Bin)

    @variable(model, Pfuel[1:Tmax,   1:Nfuel] >= 0)
    @variable(model, UCfuel[1:Tmax,  1:Nfuel], Bin)
    @variable(model, UPfuel[1:Tmax,  1:Nfuel], Bin)
    @variable(model, DOfuel[1:Tmax,  1:Nfuel], Bin)

    @variable(model, Pcogen[1:Tmax,  1:Ncogen] >= 0)
    @variable(model, UCcogen[1:Tmax, 1:Ncogen], Bin)

    @variable(model, Pbio[1:Tmax] >= 0)

    @variable(model, Phy[1:Tmax,     1:Nhy]   >= 0)
    @variable(model, e_hy[1:Tmax,    1:Nhy]   >= 0)

    @variable(model, Puns[1:Tmax]            >= 0)
    @variable(model, Pexc[1:Tmax]            >= 0)
    @variable(model, Pcharge_STEP[1:Tmax]    >= 0)
    @variable(model, Pdecharge_STEP[1:Tmax]  >= 0)
    @variable(model, stock_STEP[1:Tmax]      >= 0)
    @variable(model, S_socle[1:Tmax, 1:Nhy]   >= 0) # Eau sous la cible
    @variable(model, S_surplus[1:Tmax, 1:Nhy] >= 0) # Eau au-dessus de la cible

    # Objectif :
    @objective(model, Min,
        sum(Pnuc   .* cnuc)   +
        sum(Pgaz   .* cgaz)   +
        sum(Pcoal  .* ccoal)  +
        sum(Pfuel  .* cfuel)  +
        sum(Pcogen .* ccogen) +
        sum(Phy    .* chy)    +
        sum(Puns   .* cuns)    +
        sum(Pexc   .* cexc) -
        #  On gagne 13€ pour chaque MWh d'eau en dessous de la limite cible
        sum(S_socle[t, h] * 13.0 for t=1:Tmax, h=1:Nhy)  
    )

    # Contraintes : 
    
    # Equilibre offre-demande :
    @constraint(model, balance[t=1:Tmax],
        sum(Pnuc[t,g]  for g=1:Nnuc)  +
        sum(Pgaz[t,g]  for g=1:Ngaz)  +
        sum(Pcoal[t,g] for g=1:Ncoal) +
        sum(Pfuel[t,g] for g=1:Nfuel) +
        sum(Pcogen[t,g] for g=1:Ncogen) +
        Pbio[t] +
        sum(Phy[t,h]   for h=1:Nhy)   +
        Pres[t] + Pdecharge_STEP[t] + Puns[t]
        ==
        conso[t] + Pexc[t] + Pcharge_STEP[t]
    )

    #  Contraintes UC génériques :
    # Pour chaque groupe on applique : Pmin·UC ≤ P ≤ Pmax·UC, la logique UP/DO et les durées minimales.
    # init_P / init_UC / init_UP / init_DO : matrices 24 × N à forcer
    # Pas de forçage pour la semaine 0.

    function add_uc!(P, UC, UP, DO, Pmin_v, Pmax_v, dmin_v, N,
                     init_P, init_UC, init_UP, init_DO, nh_force)
        # Bornes de puissance
        @constraint(model, [t=1:Tmax, g=1:N], P[t,g] <= Pmax_v[g] * UC[t,g])
        @constraint(model, [t=1:Tmax, g=1:N], Pmin_v[g] * UC[t,g] <= P[t,g])

        for g in 1:N
            # Première semaine : UP/DO nuls au départ
            if week == 0
                @constraint(model, UP[1,g] == 0)
                @constraint(model, DO[1,g] == 0)
            end

            d = dmin_v[g]
            if d > 1
                # Lien UC / UP / DO
                @constraint(model, [t=2:Tmax],
                    UC[t,g] - UC[t-1,g] == UP[t,g] - DO[t,g])
                # On ne peut pas allumer ET éteindre au même pas
                @constraint(model, [t=1:Tmax], UP[t,g] + DO[t,g] <= 1)
                # Durée minimale (début de fenêtre, régime transitoire)
                @constraint(model, [t=1:min(d-1, Tmax)],
                    UC[t,g] >= sum(UP[i,g] for i=1:t))
                @constraint(model, [t=1:min(d-1, Tmax)],
                    UC[t,g] <= 1 - sum(DO[i,g] for i=1:t))
                # Durée minimale (régime permanent)
                if 2*d <= Tmax
                    @constraint(model, [t=2*d:Tmax],
                        UC[t,g] >= sum(UP[i,g] for i=(t-d+1):t))
                    @constraint(model, [t=2*d:Tmax],
                        UC[t,g] <= 1 - sum(DO[i,g] for i=(t-d+1):t))
                end
            end
        end

        # Forçage de continuité depuis la semaine précédente
        if init_P !== nothing && nh_force > 0
            nf = min(nh_force, Tmax)
            @constraint(model, [t=1:nf, g=1:N], P[t,g]  == init_P[t,g])
            @constraint(model, [t=1:nf, g=1:N], UC[t,g] == init_UC[t,g])
            @constraint(model, [t=1:nf, g=1:N], UP[t,g] == init_UP[t,g])
            @constraint(model, [t=1:nf, g=1:N], DO[t,g] == init_DO[t,g])
        end
    end

    # Nombre d'heures forcées selon la technologie (durée minimale max du groupe)
    nh_nuc  = week > 0 ? 24 : 0    # 24h pour le nucléaire
    nh_gaz  = week > 0 ? 4  : 0    # 4h pour le gaz
    nh_coal = week > 0 ? 8  : 0    # 8h pour le charbon
    nh_fuel = week > 0 ? 1  : 0    # 1h pour le fioul

    global bord_actuel
    b = bord_actuel

    add_uc!(Pnuc,  UCnuc,  UPnuc,  DOnuc,
            Pmin_nuc,  Pmax_nuc,  dmin_nuc,  Nnuc,
            week > 0 ? b.Pnuc  : nothing,
            week > 0 ? b.UCnuc  : nothing,
            week > 0 ? b.UPnuc  : nothing,
            week > 0 ? b.DOnuc  : nothing, nh_nuc)

    add_uc!(Pgaz,  UCgaz,  UPgaz,  DOgaz,
            Pmin_gaz,  Pmax_gaz,  dmin_gaz,  Ngaz,
            week > 0 ? b.Pgaz  : nothing,
            week > 0 ? b.UCgaz  : nothing,
            week > 0 ? b.UPgaz  : nothing,
            week > 0 ? b.DOgaz  : nothing, nh_gaz)

    add_uc!(Pcoal, UCcoal, UPcoal, DOcoal,
            Pmin_coal, Pmax_coal, dmin_coal, Ncoal,
            week > 0 ? b.Pcoal : nothing,
            week > 0 ? b.UCcoal : nothing,
            week > 0 ? b.UPcoal : nothing,
            week > 0 ? b.DOcoal : nothing, nh_coal)

    add_uc!(Pfuel, UCfuel, UPfuel, DOfuel,
            Pmin_fuel, Pmax_fuel, dmin_fuel, Nfuel,
            week > 0 ? b.Pfuel : nothing,
            week > 0 ? b.UCfuel : nothing,
            week > 0 ? b.UPfuel : nothing,
            week > 0 ? b.DOfuel : nothing, nh_fuel)

    # Cogénération :
    @constraint(model, [t=1:Tmax, g=1:Ncogen],
        Pcogen[t,g] <= Pmax_cogen[g] * UCcogen[t,g])

    # Biomasse et déchets
    @constraint(model, [t=1:Tmax], Pbio[t] <= Pmax_biomasse)

    # Hydraulique :
    @constraint(model, [t=1:Tmax, h=1:Nhy], e_hy[t,h] <= e_hy_stockmax[h])
    @constraint(model, [t=1:Tmax, h=1:Nhy], e_hy[t,h] >= e_hy_low[t])
    @constraint(model, [t=1:Tmax, h=1:Nhy], e_hy[t,h] <= e_hy_high[t])
    @constraint(model, [t=1:Tmax, h=1:Nhy], Phy[t,h]  <= Pmax_hy[h])
    @constraint(model, [h=1:Nhy], e_hy[1,h] == e_hy_init_val)
    @constraint(model, [t=1:Tmax-1, h=1:Nhy],
        e_hy[t+1,h] == e_hy[t,h] - Phy[t,h] + apports_hydro[t])

    # STEP :
    T_sem = 7 * 24   # Pour le STEP, on réinitialise chaque semaine
    @constraint(model, [t=1:Tmax], Pcharge_STEP[t]   <= Pmax_STEP)
    @constraint(model, [t=1:Tmax], Pdecharge_STEP[t] <= Pmax_STEP)
    @constraint(model, stock_STEP[1] == 0)
    @constraint(model, Pdecharge_STEP[1] == 0)
    @constraint(model, stock_STEP[T_sem] == 0)   # stock nul en fin de semaine utile
    @constraint(model, [t=1:Tmax-1],
        stock_STEP[t+1] == stock_STEP[t] + rSTEP * Pcharge_STEP[t] - Pdecharge_STEP[t])
    @constraint(model, [t=1:Tmax], stock_STEP[t] <= T_sem * Pmax_STEP)
    #  Contraintes des tranches d'eau
    # 1. Le stock total est la somme des deux tranches
    @constraint(model, [t=1:Tmax, h=1:Nhy], e_hy[t, h] == S_socle[t, h] + S_surplus[t, h])
    # 2. La tranche "socle" ne peut pas dépasser la cible (notre médiane)
    @constraint(model, [t=1:Tmax, h=1:Nhy], S_socle[t, h] <= e_hy_cible[t])


    ## Résolution :
    optimize!(model)
    status = termination_status(model)
    if status != MOI.OPTIMAL && status != MOI.ALMOST_OPTIMAL
        @warn "Semaine $semaine : statut $status — résultats potentiellement incorrects"
    end

    # Récupération des valeurs optimales :
    nuc_gen   = value.(Pnuc)
    gaz_gen   = value.(Pgaz)
    coal_gen  = value.(Pcoal)
    fuel_gen  = value.(Pfuel)
    cogen_gen = value.(Pcogen)
    bio_gen   = value.(Pbio)
    hy_gen    = value.(Phy)
    stock_hy  = value.(e_hy)
    exc_gen   = value.(Pexc)
    uns_gen   = value.(Puns)   
    STEP_ch   = value.(Pcharge_STEP)
    STEP_dch  = value.(Pdecharge_STEP)

    # Arrondi des binaires pour éviter les valeurs 1e-15
    UCnuc_v  = round.(value.(UCnuc));  UPnuc_v  = round.(value.(UPnuc));  DOnuc_v  = round.(value.(DOnuc))
    UCgaz_v  = round.(value.(UCgaz));  UPgaz_v  = round.(value.(UPgaz));  DOgaz_v  = round.(value.(DOgaz))
    UCcoal_v = round.(value.(UCcoal)); UPcoal_v = round.(value.(UPcoal)); DOcoal_v = round.(value.(DOcoal))
    UCfuel_v = round.(value.(UCfuel)); UPfuel_v = round.(value.(UPfuel)); DOfuel_v = round.(value.(DOfuel))

    # Mise à jour de l'état de bord pour la semaine suivante : 
    # On prend les 24 premières heures de la zone débordement (heures 169..192)
    bd = 7 * 24 + 1   # premier indice de la zone de bord

    global bord_actuel = EtatBord(
        nuc_gen[bd:bd+23,  :],
        gaz_gen[bd:bd+23,  :],
        coal_gen[bd:bd+23, :],
        fuel_gen[bd:bd+23, :],
        UCnuc_v[bd:bd+23,  :], UPnuc_v[bd:bd+23,  :], DOnuc_v[bd:bd+23,  :],
        UCgaz_v[bd:bd+23,  :], UPgaz_v[bd:bd+23,  :], DOgaz_v[bd:bd+23,  :],
        UCcoal_v[bd:bd+23, :], UPcoal_v[bd:bd+23, :], DOcoal_v[bd:bd+23, :],
        UCfuel_v[bd:bd+23, :], UPfuel_v[bd:bd+23, :], DOfuel_v[bd:bd+23, :],
        stock_hy[bd, 1] / e_hy_stockmax[1] * 100.0
    )

    # Stockage des 168h utiles dans le DataFrame 
    for t in 1:(7*24)
        jour   = (t - 1) ÷ 24 + 1
        heure  = (t - 1) % 24 + 1
        jour_c  = jour  + (semaine - 1) * 7
        heure_c = t     + (semaine - 1) * 7 * 24

        row = Dict{String, Float64}()
        for g in 1:Nnuc;  row[names_nuc[g]]  = nuc_gen[t,g];  end
        for g in 1:Ngaz;  row[names_gaz[g]]  = gaz_gen[t,g];  end
        for g in 1:Ncoal; row[names_coal[g]] = coal_gen[t,g]; end
        for g in 1:Nfuel; row[names_fuel[g]] = fuel_gen[t,g]; end
        row["Cogén"]                = cogen_gen[t,1]
        row["Biomasse"]             = bio_gen[t,1]
        row["Hydro"]                = hy_gen[t,1]
        row["STEP turbinage"]       = STEP_dch[t]
        row["Puissance résiduelle"] = Pres[t]
        row["STEP pompage"]         = -STEP_ch[t]
        row["Conso"]                = conso[t]
        row["Conso nette"]          = conso[t] + STEP_ch[t] + exc_gen[t] - Pres[t]
        row["Puissance excès"]      = exc_gen[t]
        row["Déficit offre"]         = uns_gen[t]
        row["Stock eau (%)"]        = stock_hy[t,1] / e_hy_stockmax[1] * 100.0
        row["Semaine"]              = Float64(semaine)
        row["Jour"]                 = Float64(jour)
        row["Heure"]                = Float64(heure)
        row["Jour cumulé"]          = Float64(jour_c)
        row["Heure cumulée"]        = Float64(heure_c)

        push!(final_df, row)
    end

    elapsed = round(time() - t_sem, digits=1)
    println("   → terminée en $(elapsed)s")
end


##############################################################################
# Téléchargement des résultats dans un Excel
##############################################################################
fichier_sortie = "./results/final_results_$(SCENARIO).xlsx"
println("\nÉcriture de $fichier_sortie …")

XLSX.openxlsx(fichier_sortie, mode="w") do xf
    sheet = xf[1]
    XLSX.rename!(sheet, "Résultats")
    headers = names(final_df)
    for (j, h) in enumerate(headers)
        sheet[1, j] = h
    end
    for i in 1:nrow(final_df)
        for (j, col) in enumerate(headers)
            sheet[i+1, j] = final_df[i, col]
        end
    end
end

elapsed_total = round(time() - t_total_start, digits=1)
println("Terminé ! Temps total : $(elapsed_total)s  →  $fichier_sortie")
