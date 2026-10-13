"""Tests du module de calcul (sans interface graphique)."""

import numpy as np
import pandas as pd
import pytest

from smartdigicat.core import bayesian_optimization as bo


@pytest.fixture
def donnees_1d():
    x = pd.Series([0.5, 1.5, 2.5, 3.5, 4.5], name="temperature")
    y = pd.Series(-(x - 3.0) ** 2 + 10.0, name="yield")
    return x, y


def test_matrice_matern_symetrique_diagonale_unite():
    x = pd.DataFrame({"a": [0.0, 0.3, 0.9], "b": [1.0, 0.5, 0.2]})
    k = bo.corr_matern_matrix(bo.distance_pairwise(x, [0.5, 0.5]), 2.0)
    assert np.allclose(k, k.T)
    assert np.allclose(np.diag(k), 1.0)
    assert np.all(np.linalg.eigvalsh(k) > 0)


def test_yields_optimization_ajoute_des_points(donnees_1d):
    x, y = donnees_1d
    appels = []
    bo.progress_callback = lambda: appels.append(1)
    try:
        resultat, duree, iterations = bo.Yields_optimization(
            y, 0.0, 20.0, x, 0.0, 5.0,
            n_simul_1=20, n_simul_2=3, n_replication=3, n_pas_theta=10, random_states=42)
    finally:
        bo.progress_callback = None
    assert len(iterations) == 3
    assert len(appels) == 2                      # une notification par iteration terminee (sauf la derniere)
    assert list(resultat.columns) == ["yield", "temperature"]
    assert len(resultat) >= len(y)
    nouveaux = resultat.iloc[len(y):]
    assert nouveaux["temperature"].between(0.0, 5.0).all()


def test_yields_optimization_reproductible(donnees_1d):
    x, y = donnees_1d
    r1 = bo.Yields_optimization(y, 0.0, 20.0, x, 0.0, 5.0, n_simul_1=20, n_simul_2=3,
                                n_replication=2, random_states=7)[0]
    r2 = bo.Yields_optimization(y, 0.0, 20.0, x, 0.0, 5.0, n_simul_1=20, n_simul_2=3,
                                n_replication=2, random_states=7)[0]
    pd.testing.assert_frame_equal(r1, r2)


def test_bayesian_predict_intervalle(donnees_1d):
    x, y = donnees_1d
    x_new = np.linspace(0.0, 5.0, 11)
    prediction, tirages = bo.Bayesian_predict(y, 0.0, 20.0, x, 0.0, 5.0, x_new,
                                              n_simul=30, n_pas_theta=10, random_state=0)
    assert len(prediction) == len(x_new)
    assert (prediction["f_inf"] <= prediction["f_median"] + 1e-9).all()
    assert (prediction["f_median"] <= prediction["f_sup"] + 1e-9).all()
    assert len(tirages) == 3
