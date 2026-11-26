"""
Modelo simplificado de simulación de eventos discretos (SimPy) para el
sistema de pago de parqueaderos del Centro Comercial Supercentro.

Actividad didáctica 2 - Módulo 4 (S30).
Integrantes:
- Raul Santiago Morales Henao
- David Octavio Castro Mora
"""

from pathlib import Path
from typing import Dict, List, Tuple
import numpy as np
import pandas as pd
from scipy import stats
import simpy


# =============================================================================
# CONFIGURACIÓN GENERAL Y PARÁMETROS DEL ENUNCIADO
# =============================================================================

EQUIPO_TRABAJO = [
    "Raul Santiago Morales Henao",
    "David Octavio Castro Mora",
]
URL_GITHUB = "https://github.com/shenao356/simulacion-parqueaderos-supercentro"
URL_COLAB = (
    "https://colab.research.google.com/github/shenao356/"
    "simulacion-parqueaderos-supercentro/blob/main/"
    "notebooks/Simulacion_Parqueaderos_Supercentro.ipynb"
)
NOMBRE_INFORME = "Morales_Raul_Castro_David_problemaParqueadero.docx"

ROOT_DIR = Path(__file__).resolve().parent.parent
RESULTADOS_DIR = ROOT_DIR / "resultados"
GRAFICAS_DIR = RESULTADOS_DIR / "graficas"
TABLAS_DIR = RESULTADOS_DIR / "tablas"
NOTEBOOKS_DIR = ROOT_DIR / "notebooks"
PLANTILLA_DOCX = ROOT_DIR / "plantilla de trabajo.docx"

# Perfiles de usuario: (nombre, probabilidad, media_servicio_min, media_llegada_min)
# 25% Rápido, 20% Normal y el 55% restante dividido equitativamente (27.5% Lento, 27.5% Muy lento)
PERFILES_USUARIO = [
    {"nombre": "Rápido", "prob": 0.25, "servicio": 1.0, "llegada": 3.0},
    {"nombre": "Normal", "prob": 0.20, "servicio": 3.0, "llegada": 3.0},
    {"nombre": "Lento", "prob": 0.275, "servicio": 4.0, "llegada": 5.0},
    {"nombre": "Muy lento", "prob": 0.275, "servicio": 6.0, "llegada": 7.0},
]
NOMBRES_TIPOS = [p["nombre"] for p in PERFILES_USUARIO]

NUM_CAJEROS_BASE = 3
HORIZONTE_MIN = 2400.0      # 40 horas por réplica para asegurar estado estable
MAX_REPLICAS_PILOTO = 200   # Corridas piloto para identificar R*
SEMILLA_BASE = 202604

# 3 criterios claros de decisión del equipo de trabajo (Punto C)
CRITERIOS_DECISION = {
    "max_wq": 5.0,    # Espera promedio en fila <= 5 minutos
    "max_w": 10.0,    # Tiempo total en el sistema <= 10 minutos
    "max_rho": 0.70,  # Ocupación promedio de cajeros <= 70 %
}


def parametros_teoricos(perfiles: List[Dict] = PERFILES_USUARIO) -> Dict[str, float]:
    """Calcula medias teóricas y valores analíticos M/M/1 por cajero."""
    e_s = sum(p["prob"] * p["servicio"] for p in perfiles)
    e_a = sum(p["prob"] * p["llegada"] for p in perfiles)
    lam = 1.0 / e_a
    mu = 1.0 / e_s
    rho = lam / mu
    wq_mm1 = rho / (mu * (1.0 - rho))
    w_mm1 = wq_mm1 + e_s
    lq_mm1 = (rho ** 2) / (1.0 - rho)
    l_mm1 = rho / (1.0 - rho)
    return {
        "E_S": e_s,
        "E_A": e_a,
        "lambda": lam,
        "mu": mu,
        "demanda_hora_total": lam * NUM_CAJEROS_BASE * 60.0,
        "rho": rho,
        "Wq_MM1": wq_mm1,
        "W_MM1": w_mm1,
        "Lq_MM1": lq_mm1,
        "L_MM1": l_mm1,
    }


# =============================================================================
# 1. MOTOR DE SIMULACIÓN EN SIMPY (M/M/1 INDEPENDIENTES Y FILA ÚNICA)
# =============================================================================

def simular_una_replica(
    replica: int,
    semilla,
    num_cajeros: int = NUM_CAJEROS_BASE,
    horizonte: float = HORIZONTE_MIN,
    warmup: float = 0.0,
    fila_unica: bool = False,
    perfiles: List[Dict] = PERFILES_USUARIO,
    escala_llegada: float = 1.0,
    mm1_puro: bool = False,
    guardar_usuarios: bool = False,
) -> Tuple[List[Dict], List[Dict]]:
    """
    Simula una réplica del punto de salida de parqueaderos.
    - fila_unica=False: `num_cajeros` independientes tipo M/M/1 (sin cambio ni deserción).
    - fila_unica=True: `num_cajeros` compartiendo una sola fila para la misma demanda de 3 flujos.
    """
    env = simpy.Environment()
    ss = semilla if isinstance(semilla, np.random.SeedSequence) else np.random.SeedSequence(semilla)

    probs = np.array([p["prob"] for p in perfiles], dtype=float)
    probs = probs / probs.sum()
    nombres = [p["nombre"] for p in perfiles]
    m_serv = {p["nombre"]: p["servicio"] for p in perfiles}
    m_lleg = {p["nombre"]: p["llegada"] * escala_llegada for p in perfiles}
    e_s_eff = sum(p["prob"] * p["servicio"] for p in perfiles)
    e_a_eff = sum(p["prob"] * p["llegada"] * escala_llegada for p in perfiles)

    n_estaciones = 1 if fila_unica else num_cajeros
    capacidad = num_cajeros if fila_unica else 1
    n_flujos = NUM_CAJEROS_BASE if fila_unica else num_cajeros
    child_seeds = ss.spawn(n_flujos)

    recursos = {i: simpy.Resource(env, capacity=capacidad) for i in range(1, n_estaciones + 1)}
    area_cola = {i: 0.0 for i in range(1, n_estaciones + 1)}
    area_sis = {i: 0.0 for i in range(1, n_estaciones + 1)}
    area_ocup = {i: 0.0 for i in range(1, n_estaciones + 1)}
    t_prev = {i: 0.0 for i in range(1, n_estaciones + 1)}

    datos = {
        i: {"esperas": [], "servicios": [], "sistemas": [], "conteo": {n: 0 for n in nombres}}
        for i in range(1, n_estaciones + 1)
    }
    registros_usr: List[Dict] = []

    def actualizar_areas(est_id: int) -> None:
        ini = max(t_prev[est_id], warmup)
        fin = min(env.now, horizonte)
        if fin > ini:
            dt = fin - ini
            rec = recursos[est_id]
            area_cola[est_id] += len(rec.queue) * dt
            area_sis[est_id] += (len(rec.queue) + rec.count) * dt
            area_ocup[est_id] += (rec.count / rec.capacity) * dt
        t_prev[est_id] = env.now

    def atender(est_id: int, tipo: str, t_serv: float):
        t_lleg = env.now
        actualizar_areas(est_id)
        with recursos[est_id].request() as req:
            yield req
            actualizar_areas(est_id)
            espera = env.now - t_lleg
            yield env.timeout(t_serv)
            actualizar_areas(est_id)
            sistema = env.now - t_lleg

            if t_lleg >= warmup:
                d = datos[est_id]
                d["esperas"].append(espera)
                d["servicios"].append(t_serv)
                d["sistemas"].append(sistema)
                d["conteo"][tipo] += 1

            if guardar_usuarios:
                registros_usr.append({
                    "replica": replica,
                    "cajero_id": est_id,
                    "tipo": tipo,
                    "t_llegada": t_lleg,
                    "espera": espera,
                    "servicio": t_serv,
                    "sistema": sistema,
                })

    def generador(est_id: int, rng: np.random.Generator):
        while True:
            tipo = nombres[int(rng.choice(len(nombres), p=probs))]
            dt_lleg = float(rng.exponential(e_a_eff if mm1_puro else m_lleg[tipo]))
            t_serv = float(rng.exponential(e_s_eff if mm1_puro else m_serv[tipo]))
            yield env.timeout(dt_lleg)
            if env.now >= horizonte:
                break
            env.process(atender(est_id, tipo, t_serv))

    for f_idx in range(n_flujos):
        est_dest = 1 if fila_unica else (f_idx + 1)
        env.process(generador(est_dest, np.random.default_rng(child_seeds[f_idx])))

    def cierre_horizonte():
        yield env.timeout(horizonte)
        for i in range(1, n_estaciones + 1):
            actualizar_areas(i)

    env.process(cierre_horizonte())
    env.run()  # Drena los usuarios que ya llegaron antes de `horizonte` (evita sesgo de borde)

    ventana = max(horizonte - warmup, 1e-9)
    resumen = []
    for i in range(1, n_estaciones + 1):
        d = datos[i]
        esp = np.array(d["esperas"], dtype=float)
        srv = np.array(d["servicios"], dtype=float)
        sis = np.array(d["sistemas"], dtype=float)
        fila = {
            "replica": replica,
            "cajero": "Fila Única" if fila_unica else f"Cajero {i}",
            "atendidos": len(esp),
            "Wq": float(esp.mean()) if len(esp) else 0.0,
            "S": float(srv.mean()) if len(srv) else 0.0,
            "W": float(sis.mean()) if len(sis) else 0.0,
            "Lq": (area_cola[i] / ventana) / (num_cajeros if fila_unica else 1),
            "L": (area_sis[i] / ventana) / (num_cajeros if fila_unica else 1),
            "rho": area_ocup[i] / ventana,
            "lambda_emp": (len(esp) / ventana) / (num_cajeros if fila_unica else 1),
            "P_gt_10": float(np.mean(esp > 10.0)) if len(esp) else 0.0,
        }
        for n in nombres:
            fila[f"usr_{n}"] = d["conteo"][n]
        resumen.append(fila)

    return resumen, registros_usr


def correr_replicas(
    n_replicas: int,
    semilla_base: int = SEMILLA_BASE,
    **kwargs,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Ejecuta `n_replicas` independientes con semillas reproducibles."""
    seeds = np.random.SeedSequence(semilla_base).spawn(n_replicas)
    filas, usuarios = [], []
    for r in range(1, n_replicas + 1):
        res, usr = simular_una_replica(replica=r, semilla=seeds[r - 1], **kwargs)
        filas.extend(res)
        if kwargs.get("guardar_usuarios", False):
            usuarios.extend(usr)
    return pd.DataFrame(filas), (pd.DataFrame(usuarios) if usuarios else pd.DataFrame())


# =============================================================================
# 2. ESTADO ESTABLE (R*) Y ELIMINACIÓN DEL ESTADO TRANSITORIO (PUNTO E)
# =============================================================================

def calcular_estado_estable(max_replicas: int = MAX_REPLICAS_PILOTO) -> Tuple[pd.DataFrame, Dict]:
    """Determina R* por la técnica del promedio acumulado sobre el tiempo en sistema W."""
    df_caj, _ = correr_replicas(n_replicas=max_replicas, semilla_base=SEMILLA_BASE)
    df_rep = df_caj.groupby("replica", as_index=False)[["W", "Wq", "S", "rho"]].mean()

    w_vals = df_rep["W"].to_numpy()
    wq_vals = df_rep["Wq"].to_numpy()
    w_final = float(w_vals.mean())

    filas = []
    for r in range(1, max_replicas + 1):
        mw = float(w_vals[:r].mean())
        mwq = float(wq_vals[:r].mean())
        if r >= 2:
            hw = float(stats.t.ppf(0.975, r - 1) * w_vals[:r].std(ddof=1) / np.sqrt(r))
            err_rel = hw / mw
        else:
            hw, err_rel = np.nan, np.nan
        filas.append({
            "replica": r,
            "W_replica": float(w_vals[r - 1]),
            "W_acum": mw,
            "Wq_acum": mwq,
            "ic95_inf": mw - hw if not np.isnan(hw) else np.nan,
            "ic95_sup": mw + hw if not np.isnan(hw) else np.nan,
            "semiancho": hw,
            "error_rel_pct": err_rel * 100.0 if not np.isnan(err_rel) else np.nan,
            "desv_final_pct": abs(mw - w_final) / w_final * 100.0,
        })
    df_conv = pd.DataFrame(filas)

    # Primer r >= 40 donde W_acum permanece dentro de ±1% del valor final y error IC <= 5%
    r_exacto = max_replicas
    for r in range(40, max_replicas - 19):
        if (df_conv.iloc[r - 1 :]["desv_final_pct"] <= 1.0).all() and df_conv.loc[r - 1, "error_rel_pct"] <= 5.0:
            r_exacto = r
            break
    r_optimo = min(max(int(np.ceil(r_exacto / 10.0) * 10), 60), max_replicas)
    fila_sel = df_conv.loc[df_conv["replica"] == r_optimo].iloc[0]

    return df_conv, {
        "r_exacto": r_exacto,
        "r_optimo": r_optimo,
        "W_final": w_final,
        "W_en_r": float(fila_sel["W_acum"]),
        "semiancho_en_r": float(fila_sel["semiancho"]),
        "error_rel_pct": float(fila_sel["error_rel_pct"]),
    }


def eliminar_estado_transitorio(n_replicas: int, bin_min: float = 30.0) -> Dict:
    """
    Algoritmo de eliminación del estado transitorio (Punto E):
    1. Agrupa las esperas Wq(t) en bloques de 30 minutos a través de las R* réplicas.
    2. Aplica media móvil de Welch (w=5) y la regla MSER para hallar el corte óptimo t0*.
    3. Re-evalúa las mismas réplicas descartando t < t0* para comparar Antes vs. Después.
    """
    df_antes, df_usr = correr_replicas(n_replicas=n_replicas, warmup=0.0, guardar_usuarios=True)
    n_bins = int(HORIZONTE_MIN // bin_min)
    df_usr["bin"] = np.clip((df_usr["t_llegada"] // bin_min).astype(int), 0, n_bins - 1)

    serie = (
        df_usr.groupby(["replica", "bin"], as_index=False)["espera"].mean()
        .groupby("bin", as_index=False)["espera"].mean()
        .rename(columns={"espera": "Wq_t"})
    )
    serie["t_min"] = (serie["bin"] + 0.5) * bin_min
    arr = serie["Wq_t"].to_numpy()
    m = len(arr)

    # Suavizado de Welch (semi-ventana w = 5)
    welch = [float(arr[max(0, i - min(5, i, m - 1 - i)) : i + min(5, i, m - 1 - i) + 1].mean()) for i in range(m)]
    serie["Wq_welch"] = welch

    # Regla MSER para encontrar el corte óptimo d* (t0* = d* * bin_min)
    mser = [float(np.sum((arr[d:] - arr[d:].mean()) ** 2) / (len(arr[d:]) ** 2)) for d in range(int(0.45 * m) + 1)]
    d_opt = int(np.argmin(mser[3:]) + 3)
    t0_min = float(d_opt * bin_min)

    df_despues, _ = correr_replicas(n_replicas=n_replicas, warmup=t0_min, guardar_usuarios=False)

    # Comparación Antes vs. Después por réplica (mismas semillas)
    a_rep = df_antes.groupby("replica").mean(numeric_only=True)
    d_rep = df_despues.groupby("replica").mean(numeric_only=True)
    metricas = [
        ("Wq", "Tiempo de espera en cola (Wq, min)"),
        ("W", "Tiempo total en sistema (W, min)"),
        ("S", "Tiempo de servicio en cajero (S, min)"),
        ("Lq", "Longitud promedio de cola (Lq, usuarios)"),
        ("rho", "Utilización del cajero (ρ)"),
    ]
    comp = []
    for col, nombre in metricas:
        ma, md = float(a_rep[col].mean()), float(d_rep[col].mean())
        _, pval = stats.ttest_rel(d_rep[col], a_rep[col])
        comp.append({
            "Metrica_cod": col,
            "Metrica": nombre,
            "Antes": ma,
            "Despues": md,
            "Cambio_pct": (md - ma) / ma * 100.0 if ma else 0.0,
            "p_valor": float(pval),
        })

    return {
        "t0_min": t0_min,
        "pct_descartado": t0_min / HORIZONTE_MIN * 100.0,
        "ventana_estable_min": HORIZONTE_MIN - t0_min,
        "serie": serie,
        "df_antes": df_antes,
        "df_despues": df_despues,
        "df_usuarios": df_usr,
        "df_comparacion": pd.DataFrame(comp),
    }


# =============================================================================
# 3. PUNTOS A, B, C Y D
# =============================================================================

def analizar_puntos_a_y_b(df_estable: pd.DataFrame, ventana_min: float) -> Tuple[Dict, pd.DataFrame]:
    """Calcula estadísticas por cajero (Punto A) y promedio de usuarios por tipo (Punto B)."""
    # Punto A: resumen por cajero + ANOVA
    cajeros = sorted(df_estable["cajero"].unique())
    filas_a = []
    for c in cajeros:
        sub = df_estable[df_estable["cajero"] == c]
        filas_a.append({
            "Cajero": c,
            "Servicio_S_min": float(sub["S"].mean()),
            "Espera_Wq_min": float(sub["Wq"].mean()),
            "Total_W_min": float(sub["W"].mean()),
            "Desv_W_min": float(sub["W"].std(ddof=1)),
            "Cola_Lq": float(sub["Lq"].mean()),
            "Utilizacion_pct": float(sub["rho"].mean() * 100.0),
            "Atendidos": float(sub["atendidos"].mean()),
        })
    df_a = pd.DataFrame(filas_a)
    _, p_anova_w = stats.f_oneway(*[df_estable.loc[df_estable["cajero"] == c, "W"] for c in cajeros])
    _, p_anova_s = stats.f_oneway(*[df_estable.loc[df_estable["cajero"] == c, "S"] for c in cajeros])

    res_a = {
        "tabla": df_a,
        "menor_cajero": str(df_a.loc[df_a["Total_W_min"].idxmin(), "Cajero"]),
        "menor_W": float(df_a["Total_W_min"].min()),
        "mayor_cajero": str(df_a.loc[df_a["Total_W_min"].idxmax(), "Cajero"]),
        "mayor_W": float(df_a["Total_W_min"].max()),
        "anova_p_W": float(p_anova_w),
        "anova_p_S": float(p_anova_s),
    }

    # Punto B: promedio de usuarios de cada tipo en la totalidad de los 3 cajeros
    cols_u = [f"usr_{n}" for n in NOMBRES_TIPOS]
    tot_rep = df_estable.groupby("replica", as_index=False)[cols_u].sum()
    tot_rep["total"] = tot_rep[cols_u].sum(axis=1)
    gran_total = float(tot_rep["total"].mean())
    horas = ventana_min / 60.0
    teo_pct = {p["nombre"]: p["prob"] * 100.0 for p in PERFILES_USUARIO}

    filas_b = []
    for n in NOMBRES_TIPOS:
        prom = float(tot_rep[f"usr_{n}"].mean())
        filas_b.append({
            "Tipo_usuario": n,
            "Promedio_3_cajeros": prom,
            "Promedio_por_cajero": prom / 3.0,
            "Usuarios_por_hora": prom / horas,
            "Porcentaje_simulado": prom / gran_total * 100.0,
            "Porcentaje_teorico": teo_pct[n],
        })
    filas_b.append({
        "Tipo_usuario": "TOTAL SISTEMA",
        "Promedio_3_cajeros": gran_total,
        "Promedio_por_cajero": gran_total / 3.0,
        "Usuarios_por_hora": gran_total / horas,
        "Porcentaje_simulado": 100.0,
        "Porcentaje_teorico": 100.0,
    })
    return res_a, pd.DataFrame(filas_b)


def evaluar_estrategias(n_replicas: int, t0_min: float) -> pd.DataFrame:
    """Punto C: compara 4 escenarios claros contra los 3 criterios del equipo de trabajo."""
    escenarios_cfg = [
        ("Actual: 3 cajeros (filas separadas)", 3, False, 1.0),
        ("Opción 1: 4 cajeros (filas separadas)", 4, False, 4.0 / 3.0),
        ("Opción 2: 3 cajeros (fila única)", 3, True, 1.0),
        ("Recomendada: 4 cajeros (fila única)", 4, True, 1.0),
    ]
    filas = []
    wq_base = None
    for nombre, k, fu, esc in escenarios_cfg:
        df, _ = correr_replicas(
            n_replicas=n_replicas, num_cajeros=k, warmup=t0_min, fila_unica=fu, escala_llegada=esc
        )
        rep = df.groupby("replica").mean(numeric_only=True)
        wq, w, lq, rho = float(rep["Wq"].mean()), float(rep["W"].mean()), float(rep["Lq"].mean()), float(rep["rho"].mean())
        p10 = float(rep["P_gt_10"].mean() * 100.0)
        if wq_base is None:
            wq_base = wq
        cumple = int(wq <= CRITERIOS_DECISION["max_wq"]) + int(w <= CRITERIOS_DECISION["max_w"]) + int(rho <= CRITERIOS_DECISION["max_rho"])
        filas.append({
            "Escenario": nombre,
            "Cajeros": k,
            "Espera_Wq_min": wq,
            "Total_W_min": w,
            "Cola_Lq_cajero": lq,
            "Utilizacion_pct": rho * 100.0,
            "Espera_gt_10_pct": p10,
            "Reduccion_Wq_pct": (wq_base - wq) / wq_base * 100.0,
            "Criterios_SLA": f"{cumple}/3",
            "Cumple_num": cumple,
        })
    return pd.DataFrame(filas)


def validar_modelo_vcv(df_estable: pd.DataFrame, t0_min: float) -> Dict:
    """Punto D: pruebas claras de Verificación, Calibración y Validación."""
    teo = parametros_teoricos()

    # 1. Verificación contra fórmulas analíticas M/M/1 (modo homogéneo puro)
    df_mm1, _ = correr_replicas(
        n_replicas=120, semilla_base=SEMILLA_BASE + 50000, horizonte=8000.0, warmup=1000.0, mm1_puro=True
    )
    verif_filas = []
    for col, k_teo, etq in [
        ("S", "E_S", "Servicio S (min)"),
        ("rho", "rho", "Utilización ρ"),
        ("Wq", "Wq_MM1", "Espera Wq (min)"),
        ("W", "W_MM1", "Sistema W (min)"),
        ("Lq", "Lq_MM1", "Cola Lq (usr)"),
    ]:
        sim_v = float(df_mm1[col].mean())
        teo_v = float(teo[k_teo])
        verif_filas.append({
            "Metrica": etq,
            "Teorico_MM1": teo_v,
            "Simulado_SimPy": sim_v,
            "Error_pct": abs(sim_v - teo_v) / teo_v * 100.0,
        })
    df_verif = pd.DataFrame(verif_filas)

    # Ley de Little en el modelo completo
    lam_e = float(df_estable["lambda_emp"].mean())
    err_little = abs(float(df_estable["Lq"].mean()) - lam_e * float(df_estable["Wq"].mean())) / float(df_estable["Lq"].mean()) * 100.0

    # 2. Calibración: sensibilidad del reparto del 55 % (Lento / Muy lento)
    cal_filas = []
    for i, (pl, pml) in enumerate([(0.35, 0.20), (0.275, 0.275), (0.20, 0.35)]):
        perf_alt = [
            {"nombre": "Rápido", "prob": 0.25, "servicio": 1.0, "llegada": 3.0},
            {"nombre": "Normal", "prob": 0.20, "servicio": 3.0, "llegada": 3.0},
            {"nombre": "Lento", "prob": pl, "servicio": 4.0, "llegada": 5.0},
            {"nombre": "Muy lento", "prob": pml, "servicio": 6.0, "llegada": 7.0},
        ]
        df_alt, _ = correr_replicas(
            n_replicas=40, semilla_base=SEMILLA_BASE + 70000 + i * 50, warmup=t0_min, perfiles=perf_alt
        )
        cal_filas.append({
            "Reparto": f"{pl*100:g}% Lento / {pml*100:g}% Muy lento",
            "rho_pct": float(df_alt["rho"].mean() * 100.0),
            "Wq_min": float(df_alt["Wq"].mean()),
        })
    df_cal = pd.DataFrame(cal_filas)

    # 3. Validación: prueba de carga (baja 70 %, base 100 %, alta 110 %)
    val_filas = []
    for i, (etq, factor) in enumerate([("Demanda baja (70%)", 1 / 0.70), ("Demanda base (100%)", 1.0), ("Demanda alta (110%)", 1 / 1.10)]):
        df_v, _ = correr_replicas(
            n_replicas=30, semilla_base=SEMILLA_BASE + 80000 + i * 50, warmup=t0_min, escala_llegada=factor
        )
        val_filas.append({
            "Escenario_carga": etq,
            "rho_pct": float(df_v["rho"].mean() * 100.0),
            "Wq_min": float(df_v["Wq"].mean()),
        })
    df_val = pd.DataFrame(val_filas)

    return {
        "df_verificacion": df_verif,
        "error_little_pct": err_little,
        "df_calibracion": df_cal,
        "df_validacion": df_val,
    }
