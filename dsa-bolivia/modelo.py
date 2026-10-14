"""DSA Bolivia con bloque externo integrado.

Motor del modelo. Las mismas ecuaciones se escriben como fórmulas en el Excel
(construir_excel.py); este archivo corre escenarios deterministas, calibra
y hace la simulación estocástica (Monte Carlo).

Índice temporal: t=0 -> 2025 (dato), t=1 -> 2026 (estimación calibrada),
t=2..10 -> 2027-2035 (proyección).
"""
import numpy as np

YEARS = list(range(2025, 2036))
T = len(YEARS)


def path(d, default=None):
    """Expande un dict {año: valor} a una lista; repite el último valor hacia adelante."""
    out, last = [], default
    for y in YEARS:
        if y in d:
            last = d[y]
        out.append(last)
    return out


# ---------------------------------------------------------------------------
# Supuestos por año (unidades: tasas como fracción, USD miles de millones,
# Bs miles de millones)
# ---------------------------------------------------------------------------
SERIES = {
    # Macro
    "g_base": path({2025: -0.016, 2026: -0.030, 2027: -0.015, 2028: 0.005, 2029: 0.020, 2030: 0.025, 2031: 0.028}),
    "defl_base": path({2025: 0.20, 2026: 0.18, 2027: 0.14, 2028: 0.09, 2029: 0.06, 2030: 0.05, 2031: 0.04}),
    "pi_us": path({2025: 0.025, 2026: 0.025, 2027: 0.023}),
    "rer_base": path({2025: 0.0, 2027: -0.03, 2028: -0.02, 2029: -0.01, 2030: 0.0}),
    "er_input": path({2025: 6.96, 2026: 11.80}),
    # Precios externos (índices 2025=100)
    "p_prec": path({2025: 100, 2026: 175, 2027: 161, 2028: 153}),
    "p_basem": path({2025: 100, 2026: 115}),
    "p_oil": path({2025: 100, 2026: 112, 2027: 100, 2028: 97}),
    "p_gold": path({2025: 100, 2026: 140, 2027: 129, 2028: 122}),
    "p_agro_g": path({2025: 0.0}),
    # Volúmenes (crecimiento anual)
    "v_prec": path({2025: 0.0, 2027: 0.02}),
    "v_basem": path({2025: 0.0, 2027: 0.01}),
    "v_hydro": path({2025: 0.0, 2027: -0.12, 2031: -0.06}),
    "v_agro": path({2025: 0.0, 2027: 0.03}),
    "v_fuel": path({2025: 0.0, 2027: -0.04, 2028: 0.02}),
    # Balanza de pagos: valores 2025 (dato) y 2026 (estimación)
    "x_prec0": [2.8, 6.6] + [None] * 9,
    "x_basem0": [2.2, 2.7] + [None] * 9,
    "x_hydro0": [1.6, 1.0] + [None] * 9,
    "x_agro0": [1.9, 2.2] + [None] * 9,
    "x_other0": [1.0, 1.1] + [None] * 9,
    "m_fuel0": [2.8, 2.8] + [None] * 9,
    "m_nf0": [6.37, 8.8] + [None] * 9,
    "serv0": [-1.5, -1.4] + [None] * 9,
    "sec0": [1.05, 1.35] + [None] * 9,
    # Cuenta financiera
    "fdi": path({2025: 0.1, 2026: 0.3, 2027: 0.5, 2028: 0.7, 2029: 0.9, 2030: 1.0}),
    "disb": path({2025: 1.3, 2026: 1.6, 2027: 2.0, 2028: 1.7, 2029: 1.5, 2030: 1.3}),
    "amort": path({2025: 1.0, 2026: 0.95, 2027: 1.05, 2028: 1.10, 2029: 1.05, 2030: 1.30,
                   2031: 1.10, 2032: 1.10, 2033: 1.15, 2034: 1.15, 2035: 1.20}),
    "reprof_relief": path({2025: 0.0, 2027: 0.30, 2028: 0.32, 2029: 0.20, 2030: 0.48, 2031: 0.20, 2032: 0.0}),
    "reprof_repay": path({2025: 0.0, 2032: 0.30, 2033: 0.32, 2034: 0.20, 2035: 0.48}),
    "imf_disb": path({2025: 0.0, 2026: 0.43, 2027: 0.62, 2028: 0.62, 2029: 0.23, 2030: 0.0}),
    "priv": path({2025: -1.6, 2026: -1.5, 2027: -0.9, 2028: -0.6, 2029: -0.4, 2030: -0.3}),
    # Fiscal
    "pb_nc": None,  # calibrado abajo
    "sub_share": path({2025: 0.50, 2026: 0.15, 2027: 0.05, 2028: 0.0}),
    "i_dom": path({2025: 0.050, 2026: 0.050, 2027: 0.065, 2028: 0.075, 2029: 0.075, 2030: 0.070, 2031: 0.065}),
    "i_ext": path({2025: 0.045, 2026: 0.050, 2027: 0.052}),
    "bcb_share": path({2025: 0.55, 2026: 0.55, 2027: 0.53, 2028: 0.51, 2029: 0.49, 2030: 0.47,
                       2031: 0.45, 2032: 0.43, 2033: 0.41, 2034: 0.39, 2035: 0.37}),
    "ddom_usd": path({2025: 2.0}),
}

SCALARS = {
    "theta_h": 0.45,   # renta fiscal hidrocarburífera / valor exportado (IDH+regalías)
    "theta_m": 0.06,   # renta fiscal minera / valor exportado (regalías+IUE)
    "eps_x": 0.5,      # elasticidad export. no tradicionales al TCR
    "eps_m": 0.8,      # elasticidad import. no combustibles al TCR
    "eta_y": 1.2,      # elasticidad ingreso de importaciones
    "eps_sr": 0.5,     # elasticidad de corto plazo (cierre de brecha vía devaluación)
    "kappa": 0.15,     # costo en crecimiento por punto de devaluación de emergencia
    "pt": 0.5,         # pass-through al deflactor
    "floor_m": 2.5,    # piso duro de reservas (meses de importaciones)
    "target_m": 4.0,   # nivel cómodo de reservas (meses de importaciones)
    "lam": 0.5,        # fracción de la brecha vs nivel cómodo que se cierra vía tipo de cambio
    "mu": 0.35,        # propensión a importar del impulso fiscal (déficits gemelos)
    "pi_world": 0.02,  # inflación externa en USD
    "w_nom": 0.045,    # crecimiento nominal de demanda externa
    "prof_share": 0.06,  # remesas de utilidades / export. extractivas
    "remit_g": 0.03,   # crecimiento de remesas
    "imf_rate": 0.055,  # cargos FMI (tasa DEG + margen + recargos)
    "dom_mat": 5.0,    # vida media deuda interna (años)
    "st_share": 0.4,   # fracción de la deuda externa privada que vence en el año
    "priv_ext0": 3.0,  # deuda externa privada 2026 (USD mm de millones)
    # anclas 2025
    "y0": 382.8, "dext0": 14.164, "ddom_lc0": 206.4, "gold0": 2.20, "ngold0": 1.61,
    "ob0": -0.114, "imfout0": 0.0, "pb0": -0.078,
}

SCENARIOS = [
    # nombre, parámetros
    ("Base", {}),
    ("Réplica FMI (TCR que replica su sendero)", {"rer_adj": None}),  # calibrado
    ("Reversión de metales preciosos", {"sh_prec": -0.35, "sh_basem": -0.10, "g_shock": -0.010}),
    ("Declive acelerado del gas + petróleo caro", {"hydro_extra": -0.10, "sh_oil": 0.25, "sub_return": 0.15}),
    ("Desliz fiscal (-2 pp PIB)", {"fis_slip": -0.02}),
    ("Salida de capitales / menos IED", {"priv_shift": -0.8, "fdi_mult": 0.5}),
    ("Combinado (tormenta)", {"sh_prec": -0.25, "sh_basem": -0.10, "hydro_extra": -0.05, "fis_slip": -0.01,
                              "priv_shift": -0.4, "fdi_mult": 0.7, "g_shock": -0.010, "cont": 0.03}),
    ("Base + reperfilamiento bilateral/bonos", {"reprof": 1}),
]
PARAM_KEYS = ["sh_prec", "sh_basem", "sh_oil", "hydro_extra", "fis_slip", "g_shock", "rer_adj",
              "priv_shift", "fdi_mult", "reprof", "sub_return", "cont"]
PARAM_DEFAULT = {k: 0.0 for k in PARAM_KEYS}
PARAM_DEFAULT["fdi_mult"] = 1.0

TARGET_OB = {2026: -0.092, 2027: -0.067, 2028: -0.045, 2029: -0.032, 2030: -0.030}


def run(params=None, series=None, shocks=None, n=1):
    """Corre el modelo. shocks: dict de arrays (n, T) para Monte Carlo."""
    p = dict(PARAM_DEFAULT)
    p.update(params or {})
    S = series or SERIES
    c = SCALARS
    sh = shocks or {}
    z = lambda k: sh.get(k, np.zeros((n, T)))

    def arr():
        return np.zeros((n, T))

    o = {k: arr() for k in [
        "P_prec", "P_basem", "P_oil", "P_gold", "X_prec", "X_basem", "X_hydro", "X_agro", "X_other", "X",
        "M_fuel", "M_nf_pre", "M_nf", "M", "serv", "prof", "pubint", "imfch", "sec", "CA_pre", "CA",
        "fdi", "disb", "amort", "imf_disb", "imf_rep", "imf_out", "priv", "gold", "ngold_pre", "R_pre",
        "floor", "target", "gap", "d", "R", "rd", "rer", "g_pre", "g", "pi", "Y", "ER", "ERavg", "Yusd",
        "rev_h", "rev_m", "sub", "pb_nc", "PB", "int_ext", "int_dom", "OB", "netext", "cont", "domfin",
        "Dext", "Ddom_lc", "Ddom_usd", "Ddom", "D", "d_gdp", "Dx_bcb", "dx_gdp", "dom_amort", "GFN",
        "gfn_gdp", "priv_ext", "ED", "ed_gdp", "ed_x", "ds", "ds_x", "res_m", "gg", "ca_gdp",
        "ob_gdp", "pb_gdp",
        "c_pb", "c_int", "c_g", "c_pi", "c_er", "c_cont", "c_res"]}

    # ---- t = 0 (2025, datos) ----
    t = 0
    o["P_prec"][:, 0] = S["p_prec"][0]; o["P_basem"][:, 0] = S["p_basem"][0]
    o["P_oil"][:, 0] = S["p_oil"][0]; o["P_gold"][:, 0] = S["p_gold"][0]
    for k, s in [("X_prec", "x_prec0"), ("X_basem", "x_basem0"), ("X_hydro", "x_hydro0"), ("X_agro", "x_agro0"),
                 ("X_other", "x_other0"), ("M_fuel", "m_fuel0"), ("M_nf", "m_nf0"), ("serv", "serv0"), ("sec", "sec0")]:
        o[k][:, 0] = S[s][0]
    o["M_nf_pre"][:, 0] = o["M_nf"][:, 0]
    o["X"][:, 0] = o["X_prec"][:, 0] + o["X_basem"][:, 0] + o["X_hydro"][:, 0] + o["X_agro"][:, 0] + o["X_other"][:, 0]
    o["M"][:, 0] = o["M_fuel"][:, 0] + o["M_nf"][:, 0]
    o["prof"][:, 0] = -c["prof_share"] * (o["X_prec"][:, 0] + o["X_basem"][:, 0] + o["X_hydro"][:, 0])
    o["pubint"][:, 0] = -S["i_ext"][0] * c["dext0"]
    o["CA"][:, 0] = o["X"][:, 0] - o["M"][:, 0] + o["serv"][:, 0] + o["prof"][:, 0] + o["pubint"][:, 0] + o["sec"][:, 0]
    o["gold"][:, 0] = c["gold0"]; o["R"][:, 0] = c["gold0"] + c["ngold0"]
    o["ngold_pre"][:, 0] = c["ngold0"]
    o["imf_out"][:, 0] = c["imfout0"]
    o["Y"][:, 0] = c["y0"]; o["ER"][:, 0] = S["er_input"][0]; o["ERavg"][:, 0] = S["er_input"][0]
    o["g"][:, 0] = S["g_base"][0]; o["pi"][:, 0] = S["defl_base"][0]
    o["Yusd"][:, 0] = o["Y"][:, 0] / o["ERavg"][:, 0]
    o["Dext"][:, 0] = c["dext0"]; o["Ddom_lc"][:, 0] = c["ddom_lc0"]; o["Ddom_usd"][:, 0] = S["ddom_usd"][0]
    o["Ddom"][:, 0] = c["ddom_lc0"] + S["ddom_usd"][0] * o["ER"][:, 0]
    o["D"][:, 0] = o["Dext"][:, 0] * o["ER"][:, 0] + o["Ddom"][:, 0]
    o["d_gdp"][:, 0] = o["D"][:, 0] / o["Y"][:, 0]
    o["OB"][:, 0] = c["ob0"] * c["y0"]; o["ob_gdp"][:, 0] = c["ob0"]
    o["PB"][:, 0] = c["pb0"] * c["y0"]; o["pb_gdp"][:, 0] = c["pb0"]
    o["fdi"][:, 0] = S["fdi"][0]; o["disb"][:, 0] = S["disb"][0]; o["amort"][:, 0] = S["amort"][0]
    o["priv"][:, 0] = S["priv"][0]
    o["priv_ext"][:, 0] = c["priv_ext0"]

    for t in range(1, T):
        proj = t >= 2
        f = lambda key: S[key][t]
        # precios efectivos
        o["P_prec"][:, t] = f("p_prec") * (1 + (p["sh_prec"] if proj else 0)) * np.exp(z("u_prec")[:, t])
        o["P_basem"][:, t] = f("p_basem") * (1 + (p["sh_basem"] if proj else 0)) * np.exp(z("u_basem")[:, t])
        o["P_oil"][:, t] = f("p_oil") * (1 + (p["sh_oil"] if proj else 0)) * np.exp(z("u_oil")[:, t])
        o["P_gold"][:, t] = f("p_gold") * (1 + (p["sh_prec"] if proj else 0)) * np.exp(z("u_prec")[:, t])
        o["rer"][:, t] = f("rer_base") + (p["rer_adj"] if t in (2, 3, 4) else 0)
        o["rd"][:, t] = o["rer"][:, t] + (1 - c["pt"]) * o["d"][:, t - 1]
        o["g_pre"][:, t] = f("g_base") + (p["g_shock"] if t in (2, 3) else 0) + z("u_g")[:, t]
        if not proj:
            for k, s in [("X_prec", "x_prec0"), ("X_basem", "x_basem0"), ("X_hydro", "x_hydro0"),
                         ("X_agro", "x_agro0"), ("X_other", "x_other0"), ("M_fuel", "m_fuel0"),
                         ("M_nf_pre", "m_nf0"), ("serv", "serv0"), ("sec", "sec0")]:
                o[k][:, t] = S[s][t]
        else:
            o["X_prec"][:, t] = o["X_prec"][:, t - 1] * o["P_prec"][:, t] / o["P_prec"][:, t - 1] * (1 + f("v_prec"))
            o["X_basem"][:, t] = o["X_basem"][:, t - 1] * o["P_basem"][:, t] / o["P_basem"][:, t - 1] * (1 + f("v_basem"))
            o["X_hydro"][:, t] = (o["X_hydro"][:, t - 1] * o["P_oil"][:, t - 1] / o["P_oil"][:, t - 2]
                                  * (1 + f("v_hydro") + p["hydro_extra"] + z("u_hv")[:, t]))
            o["X_agro"][:, t] = o["X_agro"][:, t - 1] * (1 + f("p_agro_g") + z("u_agro")[:, t]) * (1 + f("v_agro"))
            o["X_other"][:, t] = o["X_other"][:, t - 1] * (1 + c["w_nom"]) * (1 + c["eps_x"] * o["rd"][:, t])
            o["M_fuel"][:, t] = o["M_fuel"][:, t - 1] * o["P_oil"][:, t] / o["P_oil"][:, t - 1] * (1 + f("v_fuel"))
            o["M_nf_pre"][:, t] = (o["M_nf"][:, t - 1] * (1 + c["pi_world"]) * (1 + c["eta_y"] * o["g_pre"][:, t])
                                   * (1 - c["eps_m"] * o["rd"][:, t])
                                   - c["mu"] * (p["fis_slip"] + z("u_pb")[:, t]) * o["Yusd"][:, t - 1])
            o["serv"][:, t] = o["serv"][:, t - 1] * (1 + c["pi_world"])
            o["sec"][:, t] = o["sec"][:, t - 1] * (1 + c["remit_g"])
        o["X"][:, t] = o["X_prec"][:, t] + o["X_basem"][:, t] + o["X_hydro"][:, t] + o["X_agro"][:, t] + o["X_other"][:, t]
        o["prof"][:, t] = -c["prof_share"] * (o["X_prec"][:, t] + o["X_basem"][:, t] + o["X_hydro"][:, t])
        o["pubint"][:, t] = -f("i_ext") * o["Dext"][:, t - 1]
        o["imfch"][:, t] = -c["imf_rate"] * o["imf_out"][:, t - 1]
        o["CA_pre"][:, t] = (o["X"][:, t] - o["M_fuel"][:, t] - o["M_nf_pre"][:, t] + o["serv"][:, t]
                             + o["prof"][:, t] + o["pubint"][:, t] + o["imfch"][:, t] + o["sec"][:, t])
        o["fdi"][:, t] = f("fdi") * (p["fdi_mult"] if proj else 1)
        o["disb"][:, t] = f("disb")
        o["amort"][:, t] = f("amort") + p["reprof"] * (f("reprof_repay") - f("reprof_relief"))
        o["imf_disb"][:, t] = f("imf_disb")
        o["imf_rep"][:, t] = sum(S["imf_disb"][t - k] for k in range(5, 11) if t - k >= 0) / 6
        o["imf_out"][:, t] = o["imf_out"][:, t - 1] + o["imf_disb"][:, t] - o["imf_rep"][:, t]
        o["priv"][:, t] = f("priv") + (p["priv_shift"] if proj else 0) + z("u_priv")[:, t]
        o["gold"][:, t] = o["gold"][:, t - 1] * o["P_gold"][:, t] / o["P_gold"][:, t - 1]
        o["ngold_pre"][:, t] = (o["R"][:, t - 1] - o["gold"][:, t - 1] + o["CA_pre"][:, t] + o["fdi"][:, t]
                                + o["disb"][:, t] - o["amort"][:, t] + o["imf_disb"][:, t] - o["imf_rep"][:, t]
                                + o["priv"][:, t])
        o["R_pre"][:, t] = o["ngold_pre"][:, t] + o["gold"][:, t]
        o["floor"][:, t] = c["floor_m"] / 12 * (o["M_fuel"][:, t] + o["M_nf_pre"][:, t])
        o["target"][:, t] = c["target_m"] / 12 * (o["M_fuel"][:, t] + o["M_nf_pre"][:, t])
        o["gap"][:, t] = (np.maximum(0, np.maximum(o["floor"][:, t] - o["R_pre"][:, t],
                                                   c["lam"] * (o["target"][:, t] - o["R_pre"][:, t])))
                          if proj else 0)
        o["d"][:, t] = o["gap"][:, t] / (c["eps_sr"] * o["M_nf_pre"][:, t])
        o["M_nf"][:, t] = o["M_nf_pre"][:, t] - o["gap"][:, t]
        o["M"][:, t] = o["M_fuel"][:, t] + o["M_nf"][:, t]
        o["CA"][:, t] = o["CA_pre"][:, t] + o["gap"][:, t]
        o["R"][:, t] = o["R_pre"][:, t] + o["gap"][:, t]

        # ---- macro y fiscal ----
        o["g"][:, t] = o["g_pre"][:, t] - c["kappa"] * o["d"][:, t]
        o["pi"][:, t] = f("defl_base") + c["pt"] * o["d"][:, t]
        o["Y"][:, t] = o["Y"][:, t - 1] * (1 + o["g"][:, t]) * (1 + o["pi"][:, t])
        if proj:
            o["ER"][:, t] = (o["ER"][:, t - 1] * (1 + f("defl_base")) / (1 + f("pi_us"))
                             * (1 + o["rer"][:, t]) * (1 + o["d"][:, t]))
        else:
            o["ER"][:, t] = f("er_input")
        o["ERavg"][:, t] = (o["ER"][:, t - 1] + o["ER"][:, t]) / 2
        o["Yusd"][:, t] = o["Y"][:, t] / o["ERavg"][:, t]
        e = o["ERavg"][:, t]
        o["rev_h"][:, t] = c["theta_h"] * o["X_hydro"][:, t] * e
        o["rev_m"][:, t] = c["theta_m"] * (o["X_prec"][:, t] + o["X_basem"][:, t]) * e
        o["sub"][:, t] = (f("sub_share") + (p["sub_return"] if proj else 0)) * o["M_fuel"][:, t] * e
        o["pb_nc"][:, t] = S["pb_nc"][t] + (p["fis_slip"] if proj else 0) + z("u_pb")[:, t]
        o["PB"][:, t] = o["pb_nc"][:, t] * o["Y"][:, t] + o["rev_h"][:, t] + o["rev_m"][:, t] - o["sub"][:, t]
        o["int_ext"][:, t] = f("i_ext") * o["Dext"][:, t - 1] * e
        o["int_dom"][:, t] = (f("i_dom") + z("u_i")[:, t]) * o["Ddom"][:, t - 1]
        o["OB"][:, t] = o["PB"][:, t] - o["int_ext"][:, t] - o["int_dom"][:, t]
        o["netext"][:, t] = (o["disb"][:, t] - o["amort"][:, t]) * e
        o["cont"][:, t] = (p["cont"] if t == 2 else 0) * o["Y"][:, t]
        o["domfin"][:, t] = -o["OB"][:, t] - o["netext"][:, t] + o["cont"][:, t]
        o["Dext"][:, t] = o["Dext"][:, t - 1] + o["disb"][:, t] - o["amort"][:, t]
        o["Ddom_usd"][:, t] = f("ddom_usd")
        o["Ddom_lc"][:, t] = o["Ddom_lc"][:, t - 1] + o["domfin"][:, t]
        o["Ddom"][:, t] = o["Ddom_lc"][:, t] + o["Ddom_usd"][:, t] * o["ER"][:, t]
        o["D"][:, t] = o["Dext"][:, t] * o["ER"][:, t] + o["Ddom"][:, t]
        o["d_gdp"][:, t] = o["D"][:, t] / o["Y"][:, t]
        o["Dx_bcb"][:, t] = o["D"][:, t] - f("bcb_share") * o["Ddom_lc"][:, t]
        o["dx_gdp"][:, t] = o["Dx_bcb"][:, t] / o["Y"][:, t]
        o["dom_amort"][:, t] = o["Ddom"][:, t - 1] / c["dom_mat"]
        o["GFN"][:, t] = -o["OB"][:, t] + o["dom_amort"][:, t] + o["amort"][:, t] * e
        o["gfn_gdp"][:, t] = o["GFN"][:, t] / o["Y"][:, t]
        o["ob_gdp"][:, t] = o["OB"][:, t] / o["Y"][:, t]
        o["pb_gdp"][:, t] = o["PB"][:, t] / o["Y"][:, t]
        # descomposición de la variación de deuda/PIB
        gr = (1 + o["g"][:, t]) * (1 + o["pi"][:, t])
        dl = o["d_gdp"][:, t - 1]
        o["c_pb"][:, t] = -o["PB"][:, t] / o["Y"][:, t]
        o["c_int"][:, t] = (o["int_ext"][:, t] + o["int_dom"][:, t]) / o["Y"][:, t]
        o["c_g"][:, t] = -dl * o["g"][:, t] / gr
        o["c_pi"][:, t] = -dl * o["pi"][:, t] * (1 + o["g"][:, t]) / gr
        o["c_er"][:, t] = ((o["Dext"][:, t - 1] + o["Ddom_usd"][:, t - 1]) * (o["ER"][:, t] - o["ER"][:, t - 1])
                           + (o["disb"][:, t] - o["amort"][:, t]) * (o["ER"][:, t] - e)
                           + (o["Ddom_usd"][:, t] - o["Ddom_usd"][:, t - 1]) * o["ER"][:, t]) / o["Y"][:, t]
        o["c_cont"][:, t] = o["cont"][:, t] / o["Y"][:, t]
        o["c_res"][:, t] = (o["d_gdp"][:, t] - dl) - (o["c_pb"][:, t] + o["c_int"][:, t] + o["c_g"][:, t]
                                                     + o["c_pi"][:, t] + o["c_er"][:, t] + o["c_cont"][:, t])
        # indicadores externos
        o["priv_ext"][:, t] = o["priv_ext"][:, t - 1] * o["Yusd"][:, t] / o["Yusd"][:, t - 1] if t >= 2 else c["priv_ext0"]
        o["ED"][:, t] = o["Dext"][:, t] + o["imf_out"][:, t] + o["priv_ext"][:, t]
        o["ed_gdp"][:, t] = o["ED"][:, t] / o["Yusd"][:, t]
        o["ed_x"][:, t] = o["ED"][:, t] / o["X"][:, t]
        o["ds"][:, t] = -o["pubint"][:, t] - o["imfch"][:, t] + o["amort"][:, t] + o["imf_rep"][:, t]
        o["ds_x"][:, t] = o["ds"][:, t] / o["X"][:, t]
        o["res_m"][:, t] = o["R"][:, t] / (o["M"][:, t] / 12)
        o["gg"][:, t] = o["R"][:, t] / (o["amort"][:, t] + o["imf_rep"][:, t] + c["st_share"] * o["priv_ext"][:, t])
        o["ca_gdp"][:, t] = o["CA"][:, t] / o["Yusd"][:, t]
    o["ca_gdp"][:, 0] = o["CA"][:, 0] / o["Yusd"][:, 0]
    o["ED"][:, 0] = c["dext0"] + c["priv_ext0"]
    o["ed_gdp"][:, 0] = o["ED"][:, 0] / o["Yusd"][:, 0]
    o["ed_x"][:, 0] = o["ED"][:, 0] / o["X"][:, 0]
    o["res_m"][:, 0] = o["R"][:, 0] / (o["M"][:, 0] / 12)
    return o


def calibrate_pb_nc():
    """Calibra el balance primario no-commodity para replicar el sendero de déficit del programa."""
    S = dict(SERIES)
    pb = [-0.0715] + [0.0] * (T - 1)
    S["pb_nc"] = pb
    for t in range(1, T):
        y = YEARS[t]
        target = TARGET_OB.get(y, TARGET_OB[2030])
        lo, hi = -0.15, 0.10
        for _ in range(60):
            mid = (lo + hi) / 2
            pb[t] = mid
            for k in range(t + 1, T):
                pb[k] = mid
            ob = run(series=S)["ob_gdp"][0, t]
            if ob < target:
                lo = mid
            else:
                hi = mid
        pb[t] = round((lo + hi) / 2, 4)
    return pb


SERIES["pb_nc"] = calibrate_pb_nc()


def calibrate_imf_replica(target_2029=0.892):
    lo, hi = -0.20, 0.05
    for _ in range(60):
        mid = (lo + hi) / 2
        v = run({"rer_adj": mid})["d_gdp"][0, YEARS.index(2029)]
        if v > target_2029:
            hi = mid
        else:
            lo = mid
    return round((lo + hi) / 2, 4)


def scenario_params():
    out = []
    for name, prm in SCENARIOS:
        q = dict(PARAM_DEFAULT)
        q.update(prm)
        if name.startswith("Réplica"):
            q["rer_adj"] = calibrate_imf_replica()
        out.append((name, q))
    return out


# ---------------------------------------------------------------------------
# Monte Carlo
# ---------------------------------------------------------------------------
MC_SPEC = {
    # clave: (sigma, rho AR(1)) — choques sobre niveles en log (precios) o tasas
    "u_prec": (0.18, 0.80),
    "u_basem": (0.15, 0.75),
    "u_oil": (0.25, 0.70),
    "u_agro": (0.10, 0.0),
    "u_hv": (0.06, 0.0),
    "u_g": (0.015, 0.30),
    "u_priv": (0.50, 0.0),
    "u_pb": (0.008, 0.50),
    "u_i": (0.010, 0.70),
}
MC_CORR_KEYS = ["u_prec", "u_basem", "u_oil", "u_g"]
MC_CORR = np.array([
    [1.0, 0.5, 0.3, 0.3],
    [0.5, 1.0, 0.4, 0.3],
    [0.3, 0.4, 1.0, 0.1],
    [0.3, 0.3, 0.1, 1.0],
])


def draw_shocks(n, seed=2026):
    rng = np.random.default_rng(seed)
    L = np.linalg.cholesky(MC_CORR)
    eps = {}
    corr = rng.standard_normal((n, T, len(MC_CORR_KEYS))) @ L.T
    for i, k in enumerate(MC_CORR_KEYS):
        eps[k] = corr[:, :, i]
    for k in MC_SPEC:
        if k not in eps:
            eps[k] = rng.standard_normal((n, T))
    shocks = {}
    for k, (sig, rho) in MC_SPEC.items():
        u = np.zeros((n, T))
        for t in range(2, T):
            u[:, t] = rho * u[:, t - 1] + sig * eps[k][:, t]
        shocks[k] = u
    return shocks


def monte_carlo(n=5000, params=None):
    return run(params, shocks=draw_shocks(n), n=n)


if __name__ == "__main__":
    import sys
    np.set_printoptions(precision=3, suppress=True, linewidth=160)
    print("PB no-commodity calibrado:", SERIES["pb_nc"])
    for name, q in scenario_params():
        o = run(q)
        print(f"\n== {name}  rer_adj={q['rer_adj']}")
        for k in ["d_gdp", "ob_gdp", "gfn_gdp", "R", "ER", "gap", "d", "CA", "ca_gdp", "X", "ds_x", "Y", "Yusd"]:
            print(f"{k:8s}", o[k][0])
