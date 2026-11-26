"""
Generación de 5 gráficas limpias y fáciles de explicar (Matplotlib, 300 DPI).

Integrantes:
- Raul Santiago Morales Henao
- David Octavio Castro Mora
"""

from pathlib import Path
import sys
from typing import Dict
import matplotlib
if "matplotlib.pyplot" not in sys.modules:
    matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np
import pandas as pd

from src.simulacion import CRITERIOS_DECISION, GRAFICAS_DIR, NOMBRES_TIPOS


AZUL = "#1B365D"
AZUL_MED = "#2B6CB0"
VERDE = "#2F855A"
ROJO = "#C53030"
NARANJA = "#DD6B20"
GRIS = "#4A5568"


def _estilo() -> None:
    plt.rcParams.update({
        "figure.facecolor": "white",
        "axes.facecolor": "white",
        "axes.edgecolor": "#CBD5E0",
        "axes.grid": True,
        "grid.color": "#E2E8F0",
        "grid.linestyle": "--",
        "font.family": "sans-serif",
        "axes.titlesize": 11.5,
        "axes.titleweight": "bold",
        "axes.labelsize": 10,
    })


def _guardar(fig: plt.Figure, nombre: str, carpeta: Path = GRAFICAS_DIR) -> Path:
    carpeta.mkdir(parents=True, exist_ok=True)
    ruta = carpeta / nombre
    fig.savefig(ruta, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return ruta


def graficar_estado_estable(df_conv: pd.DataFrame, res_est: Dict, carpeta: Path = GRAFICAS_DIR) -> Path:
    """Ilustración 1: Convergencia del promedio acumulado para elegir R* réplicas."""
    _estilo()
    fig, ax = plt.subplots(figsize=(10.5, 4.8))
    r = df_conv["replica"].to_numpy()
    r_opt = res_est["r_optimo"]
    w_fin = res_est["W_final"]

    ax.fill_between(r[4:], df_conv["ic95_inf"].iloc[4:], df_conv["ic95_sup"].iloc[4:],
                    color=AZUL_MED, alpha=0.18, label="Intervalo de confianza (95 %)")
    ax.axhspan(w_fin * 0.99, w_fin * 1.01, color="#9AE6B4", alpha=0.40, label="Banda de estabilidad (±1 %)")
    ax.plot(r, df_conv["W_acum"], color=AZUL, lw=2.5, label="Promedio acumulado W (sistema)")
    ax.plot(r, df_conv["Wq_acum"], color=NARANJA, lw=2.0, ls="-.", label="Promedio acumulado Wq (cola)")
    ax.axvline(r_opt, color=ROJO, ls="--", lw=2.0, label=f"Réplicas elegidas: R* = {r_opt}")

    ax.annotate(
        f"Estado estable alcanzado\nR* = {r_opt} réplicas\nW = {res_est['W_en_r']:.2f} min (error {res_est['error_rel_pct']:.1f} %)",
        xy=(r_opt, res_est["W_en_r"]),
        xytext=(72, 20.6),
        arrowprops=dict(arrowstyle="->", color=ROJO, lw=1.5),
        fontsize=9.5, fontweight="bold",
        bbox=dict(boxstyle="round,pad=0.4", fc="#FFF5F5", ec=ROJO),
    )

    ax.set_title("Identificación del Estado Estable mediante el Promedio Acumulado")
    ax.set_xlabel("Número de corridas / réplicas independientes")
    ax.set_ylabel("Tiempo promedio (minutos)")
    ax.set_xlim(1, len(r))
    ax.set_ylim(12.0, 26.5)
    ax.legend(loc="upper right", ncol=2, fontsize=8.8)
    fig.tight_layout()
    return _guardar(fig, "fig01_estado_estable.png", carpeta)


def graficar_cajeros_y_usuarios(res_a: Dict, df_b: pd.DataFrame, carpeta: Path = GRAFICAS_DIR) -> Path:
    """Ilustración 2 (Puntos A y B): Menor/mayor tiempo por cajero y promedio de usuarios por tipo."""
    _estilo()
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12.2, 4.7))

    # Panel izquierdo: Punto A (Tiempo promedio total W por cajero)
    df_a = res_a["tabla"]
    colores_a = [
        VERDE if c == res_a["menor_cajero"] else (ROJO if c == res_a["mayor_cajero"] else AZUL_MED)
        for c in df_a["Cajero"]
    ]
    x1 = np.arange(len(df_a))
    bars1 = ax1.bar(x1, df_a["Total_W_min"], color=colores_a, width=0.52, alpha=0.88, edgecolor="black", lw=0.7)
    for b, row in zip(bars1, df_a.itertuples()):
        etq = " (MENOR)" if row.Cajero == res_a["menor_cajero"] else (" (MAYOR)" if row.Cajero == res_a["mayor_cajero"] else "")
        ax1.text(b.get_x() + b.get_width() / 2, b.get_height() + 0.35,
                 f"{row.Total_W_min:.2f} min{etq}\n(Espera: {row.Espera_Wq_min:.2f} m)",
                 ha="center", va="bottom", fontsize=8.8, fontweight="bold", color=AZUL)
    ax1.set_xticks(x1)
    ax1.set_xticklabels(df_a["Cajero"])
    ax1.set_ylim(0, float(df_a["Total_W_min"].max()) * 1.25)
    ax1.set_ylabel("Tiempo promedio total en el sistema W (min)")
    ax1.set_title(f"Punto A: Tiempo Promedio por Cajero\n(ANOVA p = {res_a['anova_p_W']:.3f}: sin diferencia significativa)")

    # Panel derecho: Punto B (Promedio de usuarios por tipo en los 3 cajeros)
    sub_b = df_b[df_b["Tipo_usuario"].isin(NOMBRES_TIPOS)]
    x2 = np.arange(len(sub_b))
    colores_b = [VERDE, AZUL_MED, NARANJA, ROJO]
    bars2 = ax2.bar(x2, sub_b["Promedio_3_cajeros"], color=colores_b, width=0.55, alpha=0.88, edgecolor="black", lw=0.7)
    for b, row in zip(bars2, sub_b.itertuples()):
        ax2.text(b.get_x() + b.get_width() / 2, b.get_height() + 6,
                 f"{row.Promedio_3_cajeros:.1f} usr\n({row.Porcentaje_simulado:.1f} %)\n{row.Usuarios_por_hora:.1f}/h",
                 ha="center", va="bottom", fontsize=8.6, fontweight="bold", color=AZUL)
    ax2.set_xticks(x2)
    ax2.set_xticklabels(sub_b["Tipo_usuario"])
    ax2.set_ylim(0, float(sub_b["Promedio_3_cajeros"].max()) * 1.30)
    ax2.set_ylabel("Promedio de usuarios (total de los 3 cajeros)")
    ax2.set_title("Punto B: Promedio de Usuarios por Tipo\n(55 % son Lentos o Muy lentos)")

    fig.tight_layout()
    return _guardar(fig, "fig02_cajeros_y_usuarios.png", carpeta)


def graficar_estrategias(df_esc: pd.DataFrame, carpeta: Path = GRAFICAS_DIR) -> Path:
    """Ilustración 3 (Punto C): Comparación directa de las 4 estrategias frente a la meta de 5 min."""
    _estilo()
    fig, ax = plt.subplots(figsize=(10.5, 4.9))
    x = np.arange(len(df_esc))
    colores = [ROJO, NARANJA, AZUL_MED, VERDE]
    bars = ax.bar(x, df_esc["Espera_Wq_min"], color=colores, width=0.54, alpha=0.90, edgecolor="black", lw=0.8)

    ax.axhline(CRITERIOS_DECISION["max_wq"], color=ROJO, ls="--", lw=1.8,
               label=f"Meta máxima de espera: {CRITERIOS_DECISION['max_wq']:.0f} min")

    for i, (b, row) in enumerate(zip(bars, df_esc.itertuples())):
        red = "Base" if i == 0 else f"-{row.Reduccion_Wq_pct:.0f} % espera"
        ax.text(
            b.get_x() + b.get_width() / 2,
            b.get_height() + 0.45,
            f"{row.Espera_Wq_min:.2f} min\n({red})\nρ = {row.Utilizacion_pct:.1f} % | SLA {row.Criterios_SLA}",
            ha="center", va="bottom", fontsize=8.7, fontweight="bold", color=AZUL,
            bbox=dict(boxstyle="round,pad=0.18", fc="white", ec="none", alpha=0.85),
        )

    etiquetas = [
        "Actual\n3 cajeros\n(filas separadas)",
        "Opción 1\n4 cajeros\n(filas separadas)",
        "Opción 2\n3 cajeros\n(fila única)",
        "Recomendada\n4 cajeros\n(fila única)",
    ]
    ax.set_xticks(x)
    ax.set_xticklabels(etiquetas, fontsize=9.5)
    ax.set_ylim(0, float(df_esc["Espera_Wq_min"].max()) * 1.32)
    ax.set_ylabel("Tiempo promedio de espera en fila Wq (min)")
    ax.set_title("Punto C: Comparación de Estrategias de Mejora — Tiempo de Espera y Cumplimiento SLA")
    ax.legend(loc="upper right", fontsize=9.5)
    fig.tight_layout()
    return _guardar(fig, "fig03_estrategias_mejora.png", carpeta)


def graficar_vcv(res_vcv: Dict, carpeta: Path = GRAFICAS_DIR) -> Path:
    """Ilustración 4 (Punto D): Esquema claro del proceso V&C&V y evidencia numérica."""
    _estilo()
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12.6, 4.8), gridspec_kw={"width_ratios": [1.05, 1.0]})

    # Panel izquierdo: Diagrama simple de 3 pasos (Verificación -> Calibración -> Validación)
    ax1.set_xlim(0, 100)
    ax1.set_ylim(0, 100)
    ax1.axis("off")
    ax1.set_title("Proceso Desarrollado (Verificación, Calibración y Validación)")

    pasos = [
        (72, "1. VERIFICACIÓN (¿El código funciona bien?)",
         "• Comparación SimPy vs. fórmulas M/M/1 (error < 0.5 %)\n• Comprobación de la Ley de Little (error < 1 %)", "#FFF5F5", ROJO),
        (40, "2. CALIBRACIÓN (¿Los parámetros son robustos?)",
         "• Ajuste de réplicas R* = 160 y calentamiento t0* = 390 min\n• Sensibilidad del 55 % de lentos (Wq siempre > 12.5 min)", "#FFFAF0", NARANJA),
        (8, "3. VALIDACIÓN (¿Refleja el parqueadero real?)",
         "• Pruebas con demanda baja (70 %), normal (100 %) y alta (110 %)\n• Comportamiento creciente y coherente con teoría de colas", "#F0FFF4", VERDE),
    ]
    for y, tit, cuerpo, fc, ec in pasos:
        ax1.add_patch(patches.FancyBboxPatch((4, y), 92, 23, boxstyle="round,pad=0.6,rounding_size=2",
                                             facecolor=fc, edgecolor=ec, lw=1.8))
        ax1.text(8, y + 17.5, tit, fontsize=9.2, fontweight="bold", color=ec, va="center")
        ax1.text(8, y + 7.5, cuerpo, fontsize=8.3, color="#1A202C", va="center", linespacing=1.35)

    ax1.annotate("", xy=(50, 64.5), xytext=(50, 71.2), arrowprops=dict(arrowstyle="-|>", color=GRIS, lw=2))
    ax1.annotate("", xy=(50, 32.5), xytext=(50, 39.2), arrowprops=dict(arrowstyle="-|>", color=GRIS, lw=2))

    # Panel derecho: Comparación visual Teórico M/M/1 vs Simulado SimPy
    df_v = res_vcv["df_verificacion"]
    x = np.arange(len(df_v))
    w = 0.36
    ax2.bar(x - w / 2, df_v["Teorico_MM1"], width=w, color=AZUL, label="Fórmula teórica M/M/1")
    ax2.bar(x + w / 2, df_v["Simulado_SimPy"], width=w, color=VERDE, alpha=0.88, label="Simulación en SimPy")
    for i, row in enumerate(df_v.itertuples()):
        top = max(row.Teorico_MM1, row.Simulado_SimPy)
        ax2.text(i, top + 0.4, f"Err:\n{row.Error_pct:.1f} %", ha="center", va="bottom", fontsize=8.2, fontweight="bold")
    ax2.set_xticks(x)
    ax2.set_xticklabels(["Servicio\nS", "Ocupación\nρ", "Espera\nWq", "Sistema\nW", "Cola\nLq"])
    ax2.set_ylim(0, float(df_v[["Teorico_MM1", "Simulado_SimPy"]].to_numpy().max()) * 1.24)
    ax2.set_ylabel("Valor de la métrica")
    ax2.set_title("Evidencia de Exactitud: Teórico vs. Simulado")
    ax2.legend(loc="upper left", fontsize=8.8)

    fig.tight_layout()
    return _guardar(fig, "fig04_verificacion_calibracion_validacion.png", carpeta)


def graficar_transitorio_antes_despues(res_t: Dict, carpeta: Path = GRAFICAS_DIR) -> Path:
    """Ilustración 5 (Punto E): Eliminación del estado transitorio mostrando Antes vs. Después."""
    _estilo()
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12.4, 4.7), gridspec_kw={"width_ratios": [1.2, 1.0]})

    serie = res_t["serie"]
    t0 = res_t["t0_min"]
    df_c = res_t["df_comparacion"]
    wq_post = float(df_c.loc[df_c["Metrica_cod"] == "Wq", "Despues"].iloc[0])

    # Panel izquierdo: Corte del periodo transitorio
    ax1.axvspan(0, t0, color="#FEB2B2", alpha=0.35, label=f"Transitorio eliminado (0–{t0:.0f} min)")
    ax1.axvspan(t0, 2400, color="#9AE6B4", alpha=0.22, label=f"Estado estable ({t0:.0f}–2400 min)")
    ax1.plot(serie["t_min"], serie["Wq_t"], color="#63B3ED", lw=1.2, alpha=0.8, label="Espera promedio por bloques")
    ax1.plot(serie["t_min"], serie["Wq_welch"], color=AZUL, lw=2.4, label="Tendencia suavizada (Welch)")
    ax1.axvline(t0, color=ROJO, ls="--", lw=2.0)
    ax1.axhline(wq_post, color=VERDE, ls=":", lw=1.8)
    ax1.set_xlim(0, 2400)
    ax1.set_xlabel("Tiempo de simulación (minutos)")
    ax1.set_ylabel("Tiempo de espera promedio Wq (min)")
    ax1.set_title(f"Detección Automática del Corte (t0* = {t0:.0f} min)")
    ax1.legend(loc="lower right", fontsize=8.2)

    # Panel derecho: Barras Antes vs Después
    sub = df_c[df_c["Metrica_cod"].isin(["Wq", "W", "S"])].copy()
    x = np.arange(len(sub))
    w = 0.35
    ax2.bar(x - w / 2, sub["Antes"], width=w, color=ROJO, alpha=0.82, edgecolor="black", lw=0.7, label="ANTES (con transitorio)")
    ax2.bar(x + w / 2, sub["Despues"], width=w, color=VERDE, alpha=0.88, edgecolor="black", lw=0.7, label="DESPUÉS (sin transitorio)")
    for i, row in enumerate(sub.itertuples()):
        ax2.text(i - w / 2, row.Antes + 0.35, f"{row.Antes:.2f}", ha="center", va="bottom", fontsize=8.5, fontweight="bold", color=ROJO)
        ax2.text(i + w / 2, row.Despues + 0.35, f"{row.Despues:.2f}", ha="center", va="bottom", fontsize=8.5, fontweight="bold", color=VERDE)
        ax2.text(i, max(row.Antes, row.Despues) + 1.8, f"{row.Cambio_pct:+.1f} %", ha="center", va="bottom",
                 fontsize=8.6, fontweight="bold", color=AZUL,
                 bbox=dict(boxstyle="round,pad=0.18", fc="#EBF8FF", ec=AZUL_MED, lw=0.8))

    ax2.set_xticks(x)
    ax2.set_xticklabels(["Espera\nWq (min)", "Sistema\nW (min)", "Servicio\nS (min)"])
    ax2.set_ylim(0, float(sub["Despues"].max()) * 1.28)
    ax2.set_ylabel("Tiempo promedio (minutos)")
    ax2.set_title("Comparación Antes vs. Después")
    ax2.legend(loc="upper right", fontsize=8.5)

    fig.tight_layout()
    return _guardar(fig, "fig05_transitorio_antes_despues.png", carpeta)
