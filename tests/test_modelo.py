"""
Pruebas básicas de verificación del modelo de simulación SimPy.

Integrantes:
- Raul Santiago Morales Henao
- David Octavio Castro Mora
"""

import pytest
from src.simulacion import (
    PERFILES_USUARIO,
    correr_replicas,
    parametros_teoricos,
    simular_una_replica,
)


def test_parametros_teoricos():
    assert abs(sum(p["prob"] for p in PERFILES_USUARIO) - 1.0) < 1e-9
    teo = parametros_teoricos()
    assert teo["E_S"] == pytest.approx(3.60, rel=1e-6)
    assert teo["E_A"] == pytest.approx(4.65, rel=1e-6)
    assert teo["rho"] < 1.0


def test_reproducibilidad_semilla():
    r1, _ = simular_una_replica(replica=1, semilla=123, horizonte=1000.0)
    r2, _ = simular_una_replica(replica=1, semilla=123, horizonte=1000.0)
    for a, b in zip(r1, r2):
        assert a["W"] == pytest.approx(b["W"], rel=1e-12)
        assert a["atendidos"] == b["atendidos"]


def test_verificacion_mm1_puro():
    teo = parametros_teoricos()
    df, _ = correr_replicas(n_replicas=30, semilla_base=9000, horizonte=3500.0, warmup=700.0, mm1_puro=True)
    assert df["S"].mean() == pytest.approx(teo["E_S"], rel=0.04)
    assert df["rho"].mean() == pytest.approx(teo["rho"], rel=0.04)
    assert df["Wq"].mean() == pytest.approx(teo["Wq_MM1"], rel=0.12)


def test_fila_unica_reduce_espera():
    df_ind, _ = correr_replicas(n_replicas=15, semilla_base=7000, num_cajeros=3, warmup=400.0, fila_unica=False)
    df_uni, _ = correr_replicas(n_replicas=15, semilla_base=7000, num_cajeros=3, warmup=400.0, fila_unica=True)
    assert df_uni["Wq"].mean() < df_ind["Wq"].mean() * 0.5
