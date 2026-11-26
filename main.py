"""
Script principal reproducible — Actividad Didáctica 2, Módulo 4 (S30).
Centro Comercial Supercentro: Simulación M/M/1 de parqueaderos.

Integrantes:
- Raul Santiago Morales Henao
- David Octavio Castro Mora
"""

import time
from src.generar_entregables import exportar_excel, generar_notebook_colab, generar_word_apa
from src.graficas import (
    graficar_cajeros_y_usuarios,
    graficar_estado_estable,
    graficar_estrategias,
    graficar_transitorio_antes_despues,
    graficar_vcv,
)
from src.simulacion import (
    GRAFICAS_DIR,
    TABLAS_DIR,
    analizar_puntos_a_y_b,
    calcular_estado_estable,
    eliminar_estado_transitorio,
    evaluar_estrategias,
    parametros_teoricos,
    validar_modelo_vcv,
)


def main(ejecutar_notebook: bool = True) -> None:
    t_ini = time.time()
    GRAFICAS_DIR.mkdir(parents=True, exist_ok=True)
    TABLAS_DIR.mkdir(parents=True, exist_ok=True)

    print("1/6 Calculando parámetros teóricos...")
    teo = parametros_teoricos()

    print("2/6 Identificando estado estable (200 réplicas piloto)...")
    df_conv, res_est = calcular_estado_estable()
    r_opt = res_est["r_optimo"]

    print(f"3/6 Eliminando estado transitorio (R* = {r_opt})...")
    res_t = eliminar_estado_transitorio(r_opt)
    t0 = res_t["t0_min"]

    print("4/6 Calculando Puntos A, B, C y D...")
    res_a, df_b = analizar_puntos_a_y_b(res_t["df_despues"], res_t["ventana_estable_min"])
    df_esc = evaluar_estrategias(r_opt, t0)
    res_vcv = validar_modelo_vcv(res_t["df_despues"], t0)

    print("5/6 Exportando 6 tablas CSV, 5 gráficas PNG, Excel y Word (APA 7)...")
    df_conv.to_csv(TABLAS_DIR / "01_estado_estable.csv", index=False)
    res_a["tabla"].to_csv(TABLAS_DIR / "02_punto_a_cajeros.csv", index=False)
    df_b.to_csv(TABLAS_DIR / "03_punto_b_usuarios.csv", index=False)
    df_esc.to_csv(TABLAS_DIR / "04_punto_c_estrategias.csv", index=False)
    res_vcv["df_verificacion"].to_csv(TABLAS_DIR / "05_punto_d_validacion.csv", index=False)
    res_t["df_comparacion"].to_csv(TABLAS_DIR / "06_punto_e_transitorio.csv", index=False)

    graficas = {
        "fig01": graficar_estado_estable(df_conv, res_est),
        "fig02": graficar_cajeros_y_usuarios(res_a, df_b),
        "fig03": graficar_estrategias(df_esc),
        "fig04": graficar_vcv(res_vcv),
        "fig05": graficar_transitorio_antes_despues(res_t),
    }

    datos = {
        "teo": teo, "df_conv": df_conv, "res_est": res_est, "res_t": res_t,
        "res_a": res_a, "df_b": df_b, "df_esc": df_esc, "res_vcv": res_vcv,
    }
    ruta_xlsx = exportar_excel(datos)
    ruta_docx = generar_word_apa(datos, graficas)

    if ejecutar_notebook:
        print("6/6 Generando y ejecutando cuaderno de Google Colab...")
        ruta_nb = generar_notebook_colab(ejecutar=True)
    else:
        ruta_nb = None

    print(f"\nListo en {time.time() - t_ini:.1f} s | R* = {r_opt} | t0* = {t0:.0f} min")
    print(f"  Word APA 7: {ruta_docx}")
    print(f"  Excel:      {ruta_xlsx}")
    if ruta_nb:
        print(f"  Colab:      {ruta_nb}")


if __name__ == "__main__":
    main()
