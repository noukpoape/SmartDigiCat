# Bayesian optimization: reference settings

Settings used with `Yields_optimization` (module `smartdigicat/core/bayesian_optimization.py`)
on the reference datasets. In the application they are entered in
**Plots → Optimization → Optimization settings**:

| Setting in the GUI | Argument          | Meaning                                                          |
|--------------------|-------------------|------------------------------------------------------------------|
| N_targ             | `n_simul_1`       | MCMC draws of the GP parameters (theta, nu)                      |
| N_feat             | `n_simul_2`       | MCMC draws of candidate points for each (theta, nu)              |
| N                  | `n_replication`   | number of optimization iterations (one new point per iteration)  |
| seed               | `random_states`   | random seed (results are reproducible for a given seed)          |
| —                  | `n_pas_theta`     | thinning of the (theta, nu) chain (fixed to 10 in the GUI)       |

## Synthetic data

| Case         | n_simul_1 | n_simul_2 | n_replication | n_pas_theta | random_states |
|--------------|----------:|----------:|--------------:|------------:|--------------:|
| Univariate   | 200       | 20        | 20            | 10          | 42            |
| Multivariate | 200       | 600       | 20            | 10          | 42            |

## Real data

| Dataset                     | Type  | n_simul_1 | n_simul_2 | n_replication | n_pas_theta | random_states |
|-----------------------------|-------|----------:|----------:|--------------:|------------:|--------------:|
| data_Catalyst_1             | univ  | 100       | 10        | 30            | 10          | 42            |
| data_Catalyst_1_280         | univ  | 100       | 20        | 30            | 10          | 42            |
| data_Catalyst_11            | univ  | 100       | 20        | 30            | 10          | 42            |
| data_Catalyst_1             | multi | 100       | 300       | 30            | 10          | 42            |
| data_Catalyst_1_280         | multi | 100       | 300       | 30            | 10          | 42            |
| data_Catalyst_1_285_290_295 | multi | 100       | 100       | 30            | 10          | 42            |

Computation time grows roughly with `n_simul_1 × n_simul_2` and with the number of
points already evaluated.
