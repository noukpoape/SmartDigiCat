# -*- coding: utf-8 -*-
"""
Optimisation bayesienne de SmartDigiCat (methode MHI : Metropolis-Hastings
avec Expected Improvement) et loi predictive du metamodele.

Modele : processus gaussien a noyau de Matern (portee ``theta``, regularite
``nu``), moyenne constante et variance (inverse-gamma) integrees. Les
parametres ``(theta, nu)`` et les points candidats sont tires par MCMC ; le
point suivant est celui qui maximise l'Expected Improvement.

Ce module ne depend d'aucune interface graphique : il peut etre importe et
teste seul.

Fonctions principales
---------------------
- ``Yields_optimization`` : boucle d'optimisation (``n_replication`` iterations).
- ``MHI_Bayesian_Optimisation`` : une iteration (propose un nouveau point).
- ``Bayesian_predict`` : courbe predite et intervalle de credibilite.

Les valeurs de reglage utilisees sur les jeux de donnees de reference sont
decrites dans ``docs/optimization_parameters.md``.
"""

import time
from math import exp, log, log10, sqrt

import numpy as np
import pandas as pd
import scipy
import scipy.linalg
from mpmath import gamma
from numpy import shape, std, var
from scipy.spatial.distance import cdist
from scipy.special import kv
from scipy.stats import gamma as gamma_random
from scipy.stats import mode, t

# Fonction appelee (sans argument) a la fin de chaque iteration de
# Yields_optimization ; l'interface s'en sert pour le temps restant estime.
progress_callback = None


''' ------ Standardiser les données'''
def data_stand (data_x, x_min, x_max):
    
    n_n = len(shape(data_x))
    
    if n_n != 1:
        
        x_col = data_x.columns
    
        data_stand_1 = pd.DataFrame()
        
        for m_col, col in enumerate(x_col):
            
            dx = data_x[col].to_numpy(dtype = float)
            
            data_stand_1[col] = (dx - x_min[m_col]) / (x_max[m_col] - x_min[m_col])
            
    else:
         
         data_arr = np.asarray(data_x, dtype = float)
         
         data_stand_1 = ((data_arr - x_min) / (x_max - x_min)).tolist()
              
    return data_stand_1


def nu_int_new (f_n, f_min, f_max, x_n, x_min, x_max):
    
    dim = np.shape(x_n)
    
    f_s = data_stand(f_n, f_min, f_max)
    
    x_s = data_stand(x_n, x_min, x_max)
    
    if len(dim) == 2:
        
        dim_1 = dim[1]
        
        nu_value = 0
        
        for i in range(dim[1]):
            
            nu_value += var(x_s.iloc[:,i])
            
    else:
        
        dim_1 = 1
        
        nu_value = var(x_s)
    
    alpha_n =  1 + 0.5*(len(f_n) - 1) 
    
    return float(sqrt(nu_value) / std(f_s))**(1/(sqrt(dim_1) + 3)) + dim_1**(1/(sqrt(dim_1) + 3))





def nu_int (f_n, f_min, f_max, x_n, x_min, x_max):
    
    dim = np.shape(x_n)
    
    if len(dim) == 2:
        
        dim_1 = dim[1]
        
        nu_value = 0
         
        for i in range(dim[1]):
            
            nu_value += var(x_n.iloc[:,i])
            
    else:
        
        dim_1 = 1
        
        nu_value = var(x_n)
        
    alpha_n =  1 + 0.5*(len(f_n) - 1) 
    
    return (sqrt(nu_value) / std(f_n))**(1/(sqrt(dim_1) + 3)) + dim_1**(1/(sqrt(dim_1) + 3))




def nu_int_2 (f_n, x_n):
    
    dim = np.shape(x_n)
    
    if len(dim) == 2:
        
        dim_1 = dim[1]
        
        nu_value = var(f_n)
         
        for i in range(dim[1]):
            
            nu_value += var(x_n.iloc[:,i])
            
    else:
        
        dim_1 = 1
        
        nu_value = var(x_n) + var(f_n)
        
    alpha_n =  1 + 0.5*(len(f_n) - 1) 
    
    return  (dim_1 + 1) / sqrt(nu_value) 


from scipy.stats import truncnorm


##truncnorm(225, 305, loc = 267, scale = 1).rvs(1)


''' ------ Loi instrumentalede l'optimise + 1ur'''
def proposal_law_value_x (n_dim, x_min, x_max) :
    
    if n_dim > 1:
        
        x_init = []
        
        for i in range(n_dim):
            
            x_init.append(float(np.random.uniform(low = x_min[i], high = x_max[i], size = 1)[0]))
        
    else:
        
        x_init = float(np.random.uniform(low = x_min, high = x_max, size = 1)[0])
        
    return x_init





''' ------ Loi a priori (et instrumentale) du paramètre theta'''
## Loi log-normale sur theta, calibree sur la geometrie du plan d'experience,
## independamment pour chaque dimension :
##   - borne basse l_k = max(plus petit ecart entre valeurs observees, etendue / (2(n-1)))
##     -> en dessous, les donnees ne renseignent pas la correlation (exploration pure)
##   - borne haute u_k = etendue du domaine (x_max - x_min)
##     -> au-dessus, le metamodele devient quasi lineaire (exploitation pure)
##   - mediane = sqrt(l_k * u_k) (moyenne geometrique : compromis exploration / exploitation)
##   - 95 % de la masse dans [l_k, u_k]
## Les arguments tau et std_v sont conserves pour la compatibilite des appels
## mais ne sont plus utilises.
def theta_prior_param (x, x_min, x_max):
    
    x_arr = np.asarray(x, dtype = float)
    
    if x_arr.ndim == 1:
        
        x_arr = x_arr.reshape(-1, 1)
        
    lo = np.atleast_1d(np.asarray(x_min, dtype = float))
    
    up = np.atleast_1d(np.asarray(x_max, dtype = float))
    
    n = x_arr.shape[0]
    
    mu, sigma = [], []
    
    for k in range(x_arr.shape[1]):
        
        etendue = up[k] - lo[k]
        
        valeurs = np.unique(x_arr[:, k])
        
        ecart_min = np.min(np.diff(valeurs)) if len(valeurs) > 1 else etendue
        
        ## Plancher : evite l'effondrement de la loi quand deux points sont quasi confondus
        borne_basse = np.maximum(ecart_min, etendue / (2*np.maximum(n - 1, 1)))
        
        borne_basse = np.minimum(borne_basse, etendue / 2)
        
        mu.append(0.5*(log10(borne_basse) + log10(etendue)))
        
        sigma.append((log10(etendue) - log10(borne_basse)) / 3.92)
        
    return np.array(mu), np.array(sigma)



def prior_theta (x, x_min, x_max):
    
    mu, sigma = theta_prior_param(x, x_min, x_max)
    
    log_theta = np.random.normal(loc = mu, scale = sigma)
    
    theta = [float(10**l_t) for l_t in log_theta]
    
    if len(shape(x)) > 1:
        
        return theta
    
    return theta[0]




''' ------ Calcul des distances par paires'''
def distance_pairwise (data_x, theta_x):
    
    dim = len(shape(data_x))
    
    if dim != 1:
        
        ## Chaque colonne est mise a l'echelle par son theta, puis la distance
        ## euclidienne entre toutes les paires de points est calculee d'un
        ## coup par cdist, au lieu d'une triple boucle Python (i, j, k)
        scaled = data_x.to_numpy(dtype = float) / np.asarray(theta_x, dtype = float)
        
        data_dist = cdist(scaled, scaled)
                    
    else:
        
        data_arr = np.asarray(data_x, dtype = float)
        
        data_dist = np.abs(data_arr[:, None] - data_arr[None, :]) / theta_x
    
    return data_dist


''' ------ Calcul de la matrice des corrélations (noyau Matern)'''
def corr_matern_matrix (dx, nu):
    
    dx_arr = np.asarray(dx, dtype = float)
    
    ## gamma(nu) ne depend pas de i/j : on ne le calcule plus qu'une seule
    ## fois au lieu de n*n fois dans la double boucle d'origine
    gamma_nu = float(gamma(nu))
    
    denom = gamma_nu * (2**(nu - 1))
    
    nu0 = sqrt(2*nu) * dx_arr
    
    with np.errstate(divide = 'ignore', invalid = 'ignore'):
        
        mat_corr = (nu0**nu) * kv(nu, nu0) / denom
    
    ## Distance nulle (diagonale ET points dupliques) : la formule donne
    ## 0 * inf = NaN, alors que la correlation vaut 1 par definition
    mat_corr[dx_arr == 0] = 1.0
    
    np.fill_diagonal(mat_corr, 1.0)
                
    return mat_corr



''' ------ Calcul de la matrice des corrélations (noyau exponentiel)'''
def corr_exponential_matrix (dx):
    
    return np.exp(-np.asarray(dx, dtype = float))



''' ------  Vecteur de corrélations du noyau Matern'''
def Matern_vect_corr (data_x, new_point, theta_x, nu):
    
    dim_x = len(shape(data_x))
    
    gamma_nu = float(gamma(nu))
    
    denom = gamma_nu * (2**(nu - 1))
    
    if dim_x != 1:
        
        data_arr = data_x.to_numpy(dtype = float)
        
        theta_arr = np.asarray(theta_x, dtype = float)
        
        new_arr = np.asarray(new_point, dtype = float)
        
        dist_p = np.sum(((data_arr - new_arr) / theta_arr)**2, axis = 1)
        
        nu0 = np.sqrt(2*nu*dist_p)
        
    else:
        
        data_arr = np.asarray(data_x, dtype = float)
        
        dist_p = np.abs(data_arr - new_point) / theta_x
        
        nu0 = sqrt(2*nu) * dist_p
    
    with np.errstate(divide = 'ignore', invalid = 'ignore'):
        
        corr_x = (nu0**nu) * kv(nu, nu0) / denom
    
    ## Point candidat confondu avec un point observe : correlation = 1 (et non NaN)
    corr_x = np.where(np.asarray(nu0) == 0, 1.0, corr_x).tolist()
    
    return corr_x



''' ------ Produit d'un vecteur par une matrice'''
def prod_vect_matrix_vect (v1, mat, v2):
    
    p_v_mat = np.dot(np.array([v1]), mat)[0]
    
    sum_mat = float(np.array(p_v_mat)@np.array(v2))
    
    return sum_mat



''' ------ Paramètre d'échelle de la distribution a posteriori de la variance '''
def Scale_parameter (mat, f_n, unit_v, b0):
        
    m = float(prod_vect_matrix_vect(unit_v, mat, f_n) / prod_vect_matrix_vect(unit_v, mat, unit_v))
        
    f_n_m = [f_m - m for f_m in f_n]
    
    R = prod_vect_matrix_vect(f_n_m, mat, f_n_m)
        
    b_n = b0 + 0.5*float(R)
        
    return b_n




def Scale_parameter_ (mat, f_n, unit_v):
    
    n = len(mat)
    
    m = float(prod_vect_matrix_vect(unit_v, mat, f_n) / prod_vect_matrix_vect(unit_v, mat, unit_v))
    
    f_n_m = [float(f_m - m) for f_m in f_n]
    
    R = prod_vect_matrix_vect(f_n_m, mat, f_n_m)
    
    return float(R)




def Optimal_function(f_n, x_n, points_x_opt, theta_x_opt, unit_vect, nu_x_opt, param_b0, param_n):
    
    etha = 2*param_n
                
    pdw_opt = distance_pairwise(x_n, theta_x_opt)
    
    K_mat_opt = corr_matern_matrix(pdw_opt, nu_x_opt)
    
    K_inv_opt = np.linalg.inv(K_mat_opt)
        
    r_opt = Matern_vect_corr(x_n, points_x_opt, theta_x_opt, nu_x_opt)
    
    value_all_opt =  prod_vect_matrix_vect(unit_vect, K_inv_opt, unit_vect)
    
    m_n_opt_1 = prod_vect_matrix_vect(unit_vect, K_inv_opt, f_n)
    
    m_n_opt = m_n_opt_1 / value_all_opt
   
    d_fm_opt = [v1_opt - m_n_opt for v1_opt in f_n]
    
    part_2 = prod_vect_matrix_vect(r_opt, K_inv_opt, d_fm_opt)
    
    f_next = float(m_n_opt +  part_2)
          
    return f_next
    




''' ------ Mise à jour du paramètre de porté'''
def Updat_theta (f_n, x_n, x_min, x_max, theta_old, param_b0, param_n, param_matern, unit_vect) :
    
    dim_x = len(shape(theta_old))
    
    pd_theta0 = distance_pairwise(x_n, theta_old)
    
    K_mat_theta0 = corr_matern_matrix(pd_theta0, param_matern)
        
    K_inv_theta0 = scipy.linalg.inv(K_mat_theta0)
    
    bn_theta0 = Scale_parameter(K_inv_theta0, f_n, unit_vect, param_b0)
    
    pr_0 = prod_vect_matrix_vect(unit_vect, K_inv_theta0, unit_vect)
    
    theta_new = prior_theta(x_n, x_min, x_max)
    
    pd_theta_new = distance_pairwise(x_n, theta_new)
    
    K_mat_theta_new = corr_matern_matrix(pd_theta_new, param_matern)
        
    K_inv_theta_new = scipy.linalg.inv(K_mat_theta_new)
    
    bn_theta_new = Scale_parameter(K_inv_theta_new, f_n, unit_vect, param_b0)
    
    pr_new = prod_vect_matrix_vect(unit_vect, K_inv_theta_new, unit_vect)
    
    pj_theta = 1
    
    if dim_x != 0:
        
        for i in range(len(theta_old)):
        
            pj_theta *= theta_new[i] / theta_old[i]
        
    else:
         
        pj_theta *= theta_new / theta_old  
        
    
    b_u = bn_theta_new / bn_theta0
    
    ## log-determinants : det() sous-deborde vers 0 quand K est mal conditionnee
    sign_new, logdet_th_new = np.linalg.slogdet(K_mat_theta_new)
        
    sign_old, logdet_th_old = np.linalg.slogdet(K_mat_theta0)
    
    fact3 = pr_0 / pr_new
    
    
    if (np.isfinite(b_u) and b_u > 0 and sign_new > 0 and sign_old > 0 and np.isfinite(fact3) and fact3 > 0 and np.isfinite(pj_theta) and pj_theta > 0):
        
        ## Meme rapport qu'avant : b_u^(-a_n) * sqrt(det_old/det_new) * sqrt(fact3) * pj_theta, calcule en log
        log_fact = (-param_n*log(b_u) + 0.5*(logdet_th_old - logdet_th_new) + 0.5*log(fact3) + log(pj_theta))
        
        fact = float(exp(np.minimum(log_fact, 0.0)))
        
    else:
        
        fact = 0
        
    u_theta = np.random.uniform(low= 0, high= 1, size= 1)  
    
    if fact > u_theta[0]:
        
        theta_val = theta_new
    
    else:
    
        theta_val = theta_old
    
    
    ##print("theta_val = " + str(theta_val))   
    
    return theta_val    



''' ------ Mise à jour du paramètre de régularité'''
def Updat_nu (f_n, x_n, theta_old, param_b0, param_n, nu_old, unit_vect, nu_param, pram_ig) :
    
    pd_nu = distance_pairwise(x_n, theta_old)
    
    K_mat_nu_old = corr_matern_matrix(pd_nu, nu_old)
        
    K_inv_nu_old = scipy.linalg.inv(K_mat_nu_old)
    
    bn_nu_old = Scale_parameter(K_inv_nu_old, f_n, unit_vect, param_b0)
    
    pr_nu_old = prod_vect_matrix_vect(unit_vect, K_inv_nu_old, unit_vect)
    
    nu_new = float(gamma_random.rvs(a = 1, loc = nu_param, scale = 0.01, size = 1)[0]) 
    
    K_mat_nu_new = corr_matern_matrix(pd_nu, nu_new)
    
    K_inv_nu_new = scipy.linalg.inv(K_mat_nu_new)
    
    bn_nu_new = Scale_parameter(K_inv_nu_new, f_n, unit_vect, param_b0)
    
    pr_nu_new = prod_vect_matrix_vect(unit_vect, K_inv_nu_new, unit_vect)
    
    b_u = bn_nu_new / bn_nu_old
    
    sign_new, logdet_nu_new = np.linalg.slogdet(K_mat_nu_new)
    
    sign_old, logdet_nu_old = np.linalg.slogdet(K_mat_nu_old)
    
    fact3 = pr_nu_old / pr_nu_new
    
    
    if (np.isfinite(b_u) and b_u > 0 and sign_new > 0 and sign_old > 0 and np.isfinite(fact3) and fact3 > 0):
        
        log_fact = (-param_n*log(b_u) + 0.5*(logdet_nu_old - logdet_nu_new) + 0.5*log(fact3))
        
        fact_nu = float(exp(np.minimum(log_fact, 0.0)))
        
    else:
        
        fact_nu = 0
        
    u_nu = np.random.uniform(low= 0, high= 1, size= 1)    
    
    if fact_nu > u_nu[0]:
        
        nu_val = nu_new
        
    else:
        
        nu_val = nu_old
       
    return nu_val    
    

    
''' ------ Log-vraisemblance integree de (theta, nu)'''
## log L(theta, nu) = -a_n log(b_n) - 0.5 log|K| - 0.5 log(1'K^-1 1)
## (moyenne constante et variance integrees ; meme quantite que les rapports
##  utilises dans Updat_theta et Updat_nu). Renvoie -inf si K n'est pas definie positive.
def log_vraisemblance (f_n, x_n, theta_x, nu, param_b0, param_n):
    
    f = np.asarray(f_n, dtype = float)
    
    un = np.ones(len(f))
    
    K = corr_matern_matrix(distance_pairwise(x_n, theta_x), nu)
    
    try:
        
        cK = scipy.linalg.cho_factor(K, lower = True)
        
    except (np.linalg.LinAlgError, ValueError):
        
        return -np.inf
    
    logdet_K = 2.0*float(np.sum(np.log(np.diag(cK[0]))))
    
    K_inv_1 = scipy.linalg.cho_solve(cK, un)
    
    s_11 = float(un @ K_inv_1)
    
    m = float(un @ scipy.linalg.cho_solve(cK, f)) / s_11
    
    R = float((f - m*un) @ scipy.linalg.cho_solve(cK, f - m*un))
    
    b_n = param_b0 + 0.5*R
    
    if not (np.isfinite(b_n) and b_n > 0 and np.isfinite(s_11) and s_11 > 0 and np.isfinite(logdet_K)):
        
        return -np.inf
    
    return float(-param_n*log(b_n) - 0.5*logdet_K - 0.5*log(s_11))



''' ------ Mise à jour conjointe de (theta, nu)'''
## Proposition independante du couple : theta' ~ prior_theta, nu' ~ nu_param + Gamma(1, 0.01)
## (memes lois que Updat_theta et Updat_nu). Les propositions etant les lois a priori, le
## rapport d'acceptation se reduit a :
##     L(theta', nu') / L(theta, nu) * prod_k theta'_k / theta_k
## (le dernier facteur est le pj_theta de Updat_theta, conserve pour viser la meme loi).
## loglik_old est la log-vraisemblance du couple courant : elle est reutilisee, si bien
## qu'une seule matrice K est factorisee par iteration.
## Renvoie (theta, nu, loglik) du couple retenu.
def Updat_theta_nu (f_n, x_n, x_min, x_max, theta_old, nu_old, loglik_old, param_b0, param_n, nu_param):
    
    theta_new = prior_theta(x_n, x_min, x_max)
    
    nu_new = float(gamma_random.rvs(a = 1, loc = nu_param, scale = 0.01, size = 1)[0])
    
    loglik_new = log_vraisemblance(f_n, x_n, theta_new, nu_new, param_b0, param_n)
    
    if not np.isfinite(loglik_new):
        
        return theta_old, nu_old, loglik_old
    
    if not np.isfinite(loglik_old):
        
        return theta_new, nu_new, loglik_new
    
    log_pj = float(np.sum(np.log(np.asarray(theta_new, dtype = float) / np.asarray(theta_old, dtype = float))))
    
    log_r = loglik_new - loglik_old + log_pj
    
    if log(np.random.uniform(low = 0, high = 1, size = 1)[0]) < log_r:
        
        return theta_new, nu_new, loglik_new
    
    return theta_old, nu_old, loglik_old



''' ------ Mise à jour des points d'évaluation'''
def Updat_x (f_n, x_n, x_min, x_max, dim_x, x_old, theta_x, param_b0, param_n, param_matern, unit_vect) :
    
    etah_n = 2*param_n
    
    pwd = distance_pairwise(x_n, theta_x)
    
    K_mat = corr_matern_matrix(pwd, param_matern)
    
    ##K_inv = np.linalg.inv(K_mat)
    
    K_inv = scipy.linalg.inv(K_mat)
    
    r_old = Matern_vect_corr(x_n, x_old, theta_x, param_matern)
    
    pr_all = prod_vect_matrix_vect(unit_vect, K_inv, unit_vect)
    
    m_n = float(prod_vect_matrix_vect(unit_vect, K_inv, f_n) / pr_all)
    
    d_fm = [float(f_i - m_n) for f_i in f_n]
    
    f_n_old = m_n +  prod_vect_matrix_vect(r_old, K_inv, d_fm) - np.max(f_n)
    
    ##bn_x = exp(Scale_parameter(K_inv, f_n, unit_vect, param_b0))
    
    bn_x = Scale_parameter(K_inv, f_n, unit_vect, param_b0)
    
    kapa1 = prod_vect_matrix_vect(r_old, K_inv, r_old)
    
    kapa2 = prod_vect_matrix_vect(r_old, K_inv, unit_vect)
    
    kapa_2 = float(1 - kapa1 + ((1 - kapa2)**2) / pr_all)
    
    x_new = proposal_law_value_x(dim_x, x_min, x_max) ## ---- Proposed new evaluation point
    
    r_new = Matern_vect_corr(x_n, x_new, theta_x, param_matern)
    
    f_n_new = m_n  +  prod_vect_matrix_vect(r_new, K_inv, d_fm) - np.max(f_n)
    
    kapa1_new = prod_vect_matrix_vect(r_new, K_inv, r_new)
    
    kapa2_new = prod_vect_matrix_vect(r_new, K_inv, unit_vect)
    
    kapa_2_new = 1 - kapa1_new + ((1 - kapa2_new)**2) / pr_all
    
    
    if (np.isfinite(kapa_2) and kapa_2 > 0 and np.isfinite(bn_x) and bn_x > 0 and np.isfinite(kapa_2_new) and kapa_2_new > 0):
    
        gamma_0 = sqrt((bn_x*kapa_2) / param_n)
        
        gamma_new = sqrt((bn_x * kapa_2_new)/ param_n)
        
        u_f_old = f_n_old / gamma_0
        
        u_f_new = f_n_new/ gamma_new
        
        cef_pdf_old = float((etah_n + u_f_old**2) / (etah_n - 1))
        
        cef_pdf_new = float((etah_n + u_f_new**2) / (etah_n - 1))
        
        ei_old = gamma_0*(cef_pdf_old*t.pdf(u_f_old, etah_n) + u_f_old*t.cdf(u_f_old, etah_n))
        
        ei_new = gamma_new*(cef_pdf_new*t.pdf(u_f_new, etah_n) + u_f_new*t.cdf(u_f_new, etah_n))
        
        if ei_new <= 0 or not np.isfinite(ei_new):
            
            r_x = 0
            
        elif ei_old <= 0 or not np.isfinite(ei_old):
            
            r_x = 1
            
        else:
            
            r_x = float(exp(np.minimum(log(ei_new) - log(ei_old), 0.0)))
    
    else:
        
        r_x = 0
    
    u_x = np.random.uniform(low = 0, high = 1, size = 1)
    
    if r_x > u_x[0]:
        
        x_value = x_new
    
    else:
        
        x_value = x_old
    
    ##print(x_value)
    
    return x_value



def Expected_Improvement(f_n, x_n, points_x, theta_x, unit_vect, param_matern, param_b0, param_n):
  
  a_n_2 = 2*param_n
  
  Estimator_EI = []
  
  for x_i in points_x:
      
      EI_theta = 0
      
      for theta_i in theta_x:
            
            pd_EI = distance_pairwise(x_n, theta_i)
            
            K_mat_EI = corr_matern_matrix(pd_EI, param_matern)
                        
            K_inv_EI = scipy.linalg.inv(K_mat_EI)
            
            r_EI = Matern_vect_corr(x_n, x_i, theta_i, param_matern)
            
            EI_all = prod_vect_matrix_vect(unit_vect, K_inv_EI, unit_vect)
            
            m_n_EI = float(prod_vect_matrix_vect(unit_vect, K_inv_EI, f_n)/EI_all)     
            
            d_fm_EI = [float(f_n_i - m_n_EI) for f_n_i in f_n]
            
            ## Pas de troncature a 0 ici : u doit pouvoir etre negatif dans la formule de l'EI
            f_n_EI = m_n_EI +  prod_vect_matrix_vect(r_EI, K_inv_EI, d_fm_EI) - np.max(f_n)
      
            bn_ei = Scale_parameter(K_inv_EI, f_n, unit_vect, param_b0)
            
            kapa_EI1 = prod_vect_matrix_vect(r_EI, K_inv_EI, r_EI)
            
            kapa_EI2 = prod_vect_matrix_vect(r_EI, K_inv_EI, unit_vect)
            
            p_t = (((1 - kapa_EI2)**2)/EI_all)
            
            kapa_EI_2 = float(1 - kapa_EI1 + p_t)
            
            if ((kapa_EI_2 >0) & (bn_ei > 0)):
                
                ## Echelle de la loi de Student predictive : sqrt(b_n * kappa / a_n)
                ## (comme dans Updat_x et Expected_Improvement_1)
                gamma_n = sqrt((bn_ei * kapa_EI_2) / param_n)
                
                u_n_t = f_n_EI / gamma_n
                
                coef_pdf = float((a_n_2 + u_n_t**2) / (a_n_2 - 1))
                
                EI_theta  += gamma_n*( coef_pdf*t.pdf(u_n_t, a_n_2) + u_n_t*t.cdf(u_n_t, a_n_2))/len(theta_x)
                
                        
      Estimator_EI.append(np.max([EI_theta , 0]))

            
  return Estimator_EI



''' ------ Fonction principale : Optimisation bayésienne'''
def MHI_Bayesian_Optimisation (f_n, lower_f, upper_f, x_n, lower_x, upper_x, n_simul_1, n_simul_2, n_pas_theta, nu_0 = 2):   
    
    ## ---------------------------------------------------------------------------
    ## Normalisation interne : x et f sont ramenes sur [0, 1] a partir des bornes
    ## (lower_x, upper_x) et (lower_f, upper_f). Tous les calculs (MCMC, EI,
    ## prediction) se font sur l'echelle normalisee ; les resultats sont ramenes
    ## sur l'echelle de depart en fin de fonction.
    ## ---------------------------------------------------------------------------
    
    f_orig = f_n
    
    x_orig = x_n
    
    lower_f_orig = float(lower_f)
    
    scale_f = float(upper_f) - float(lower_f)
    
    multi_dim = len(shape(x_n)) > 1
    
    if multi_dim:
        
        lower_x_orig = np.asarray(lower_x, dtype = float)
        
        scale_x = np.asarray(upper_x, dtype = float) - lower_x_orig
        
        x_n = (x_n.astype(float) - lower_x_orig) / scale_x      ## colonne par colonne
        
        lower_x = [0.0]*len(scale_x)
        
        upper_x = [1.0]*len(scale_x)
        
    else:
        
        lower_x_orig = float(lower_x)
        
        scale_x = float(upper_x) - float(lower_x)
        
        x_n = (x_n.astype(float) - lower_x_orig) / scale_x
        
        lower_x, upper_x = 0.0, 1.0
        
    f_n = (f_n.astype(float) - lower_f_orig) / scale_f
    
    lower_f, upper_f = 0.0, 1.0
    
    
    ## Retour a l'echelle de depart
    def x_to_orig (x_s):
        
        if multi_dim:
            
            return [float(v) for v in lower_x_orig + scale_x*np.asarray(x_s, dtype = float)]
        
        return float(lower_x_orig + scale_x*float(x_s))
    
    def theta_to_orig (th_s):
        
        if multi_dim:
            
            return [float(v) for v in scale_x*np.asarray(th_s, dtype = float)]
        
        return float(scale_x*float(th_s))
    

    n = len(f_n)
    
    unit_vect = [1 for i in range(n)]
    
    param_a0 =  1 
    
    a_n = param_a0 + 0.5*(n - 1)
    
    param_b0 =  1 
    
    etah_n = 2*a_n
    
    dim = shape(x_n)
        
    if len(dim) != 1:
        
        dim_x = dim[1]
        
    else:
        
        dim_x = 1
        
            
    theta = []
    
    theta.append(prior_theta(x_n, lower_x, upper_x))
                
    nu_sim = []
    
    nu_sim.append(float(gamma_random.rvs(a = 1, loc = nu_0, scale = 0.01, size = 1)[0]))
    
    loglik = log_vraisemblance(f_n, x_n, theta[0], nu_sim[0], param_b0, a_n)
      
    data_simul = []
    
    xx_0 = []
    
    xx_0.append(proposal_law_value_x(dim_x, lower_x, upper_x))
    
    for k in range(1, n_simul_2):
        
        xx_0.append(Updat_x(f_n, x_n, lower_x, upper_x, dim_x, xx_0[k - 1], theta[0], param_b0, a_n, nu_sim[0], unit_vect))
    
    data_simul.append(xx_0)
    
    for i in range(1, n_simul_1):
        
        ## Mise a jour conjointe de (theta, nu)
        theta_i, nu_i, loglik = Updat_theta_nu(f_n, x_n, lower_x, upper_x, theta[i-1], nu_sim[i-1], loglik, param_b0, a_n, nu_0)
        
        theta.append(theta_i)
        
        nu_sim.append(nu_i)
        
        x_simul = []
        
        x_simul.append(proposal_law_value_x(dim_x, lower_x, upper_x))
        
        for j in range(1, n_simul_2):
            
            x_simul.append(Updat_x(f_n, x_n, lower_x, upper_x, dim_x, x_simul[j-1], theta[i], param_b0, a_n, nu_sim[i], unit_vect))
        
        data_simul.append(x_simul)
        
    theta_filt = []
    
    data_simul_fit = []
    
    nu_sim_fit = []
    
    for pas_theta in range(0, n_simul_1, n_pas_theta):
        
        theta_filt.append(theta[pas_theta])
        
        nu_sim_fit.append(nu_sim[pas_theta])
        
        ## Tous les points de la chaine x associee a ce theta sont conserves
        data_simul_fit.extend(data_simul[pas_theta])
    
    n_theta = len(theta_filt)
        
    nu_posterior_mode = mode(nu_sim_fit, keepdims=True).mode[0]  ##  np.mean(nu_sim_fit) ##             
        
    Estimator_EI = Expected_Improvement(f_n, x_n, data_simul_fit, theta_filt, unit_vect, nu_posterior_mode, param_b0, a_n)
    
    theta_all = []
    
    nu_sim_all = []
    
    for i in range(n_theta):
        
        for j in range(len(data_simul[0])):
            
            theta_all.append(theta_filt[i])
            
            nu_sim_all.append(nu_sim_fit[i])
            
    
    data_EI  = pd.DataFrame({"x": data_simul_fit,"theta": theta_all, "nu": nu_sim_all, "EI": Estimator_EI})     
    
    optimal_value_EI = data_EI.sort_values(by = 'EI', ascending = False) 
        
    optimal_value =  np.array(optimal_value_EI.iloc[0])  
    
    points_opt = optimal_value[0]
            
    if dim_x != 1:
        
        d_theta = pd.DataFrame(theta_filt)
                
        theta_mode = mode(theta_filt, keepdims = True).mode[0]    
   
    else:
         
        theta_mode = mode(theta_filt, keepdims = True).mode[0] ## np.mean(theta_filt) ##      
        
        
    f_n_opt = Optimal_function(f_n, x_n, points_opt, theta_mode, unit_vect, nu_posterior_mode, param_b0, a_n)
        
    f_x_opt_s = float(f_n_opt)      ## valeur predite sur l'echelle normalisee
    
    
    ## ---- Retour a l'echelle de depart
    
    f_x_opt = lower_f_orig + scale_f*f_x_opt_s
    
    updat_Y =  f_orig.copy() 
    
    update_x = x_orig.copy()
    
    theta_filt = [theta_to_orig(th) for th in theta_filt]
    
    data_EI["x"] = [x_to_orig(x_i) for x_i in data_EI["x"]]
    
    data_EI["theta"] = [theta_to_orig(th) for th in data_EI["theta"]]
    
    data_EI["EI"] = scale_f*data_EI["EI"]          ## l'EI s'exprime dans l'unite de f
    
    if ((f_x_opt_s >= lower_f) & (f_x_opt_s <= upper_f)):
        
        updat_Y[len(updat_Y)] =  f_x_opt
        
        x_next = x_to_orig(points_opt)
        
        if multi_dim:          ## DataFrame (y compris une seule colonne) : ajout d'une ligne
            
            data_x_new = pd.DataFrame([x_next], columns = x_orig.columns)
            
            update_x = pd.concat([update_x, data_x_new], ignore_index=True)
            
        else:
            
            update_x[len(update_x)] = x_next
                        

    return updat_Y, update_x, theta_filt, data_EI



''' ------ Simulation des valeurs de la fonction'''
def Yields_optimization (f_n, lower_f, upper_f, x_n, lower_x, upper_x,  n_simul_1 = 200, n_simul_2 = 10, n_replication = 10, n_pas_theta = 10, random_states = None):
    
    
    n_0 = len(f_n)
    
    dim = shape(x_n)
        
    if len(dim) != 1:
        
        dim_x = dim[1]
       
    else:
        
        dim_x = 1
    
    start_time = time.time()
    
    val_nu_int = nu_int(f_n, lower_f, upper_f, x_n, lower_x, upper_x)
    
    alpha_val = (1 + 0.5*(n_0 - 1))
            
    np.random.seed(seed = random_states)
    
    data_update = []
    
    data_update.append(MHI_Bayesian_Optimisation(f_n= f_n, lower_f= lower_f, upper_f= upper_f, x_n= x_n, lower_x = lower_x, upper_x = upper_x, n_simul_1 = n_simul_1, n_simul_2 = n_simul_2, n_pas_theta= n_pas_theta, nu_0 = val_nu_int))
    
    for i in range(1, n_replication):
        
        if random_states is not None:
            
            random_states_i = int(random_states + i)
            
        else:
            
            random_states_i = None
            
        np.random.seed(seed = random_states_i)
        
        signaler = progress_callback
        if signaler is not None:
            signaler()                     # iteration precedente terminee : mise a jour du temps restant
        
        d_x = data_update[i-1]
        
        px = MHI_Bayesian_Optimisation (f_n= d_x[0], lower_f= lower_f, upper_f= upper_f, x_n= d_x[1], lower_x = lower_x, upper_x = upper_x, n_simul_1 =  n_simul_1, n_simul_2 = n_simul_2, n_pas_theta= n_pas_theta, nu_0 = val_nu_int)        
                
        data_update.append(px)
        
    end_time = time.time()
        
    execution_time = time.strftime("%H:%M:%S", time.gmtime(end_time - start_time)) ##f" Execution time : {iteration_time:.2f} seconds"
        
    data_update_n = pd.concat([data_update[-1][0], data_update[-1][1]], axis = 1)
    
    return data_update_n, execution_time, data_update




''' ------ Prédiction bayésienne : loi prédictive du modèle (pour la représentation graphique)'''
## Meme modele que MHI_Bayesian_Optimisation :
##   f(x) | theta, nu  ~  processus gaussien, noyau de Matern (portee theta, regularite nu),
##   moyenne constante m integree, variance ~ inverse-gamma(a0, b0) integree.
## Pour (theta, nu) fixes, la loi predictive en x est une Student :
##   f(x) ~ m_n + gamma_n(x) * T_{2 a_n}
##   m_n(x)     = m + r(x)' K^-1 (f - m 1),        m = 1'K^-1 f / 1'K^-1 1
##   gamma_n(x) = sqrt(b_n * kappa(x) / a_n),     b_n = b0 + 0.5 (f - m 1)' K^-1 (f - m 1)
##   kappa(x)   = 1 - r'K^-1 r + (1 - 1'K^-1 r)^2 / 1'K^-1 1
## (theta, nu) sont tires par la meme chaine MCMC (Updat_theta, Updat_nu) ; la prediction
## finale est le melange des lois de Student sur les tirages retenus.


''' ------ Correlations de Matern entre les points a predire et les points observes'''
def correlation_croisee (x_new, x_n, theta_x, nu):
    
    xa = np.asarray(x_new, dtype = float)
    
    xb = np.asarray(x_n, dtype = float)
    
    if xa.ndim == 1:
        
        xa = xa.reshape(-1, 1)
        
    if xb.ndim == 1:
        
        xb = xb.reshape(-1, 1)
    
    ## Meme distance que distance_pairwise : chaque colonne divisee par son theta
    th = np.atleast_1d(np.asarray(theta_x, dtype = float))
    
    nu0 = sqrt(2*nu) * cdist(xa / th, xb / th)
    
    denom = float(gamma(nu)) * (2**(nu - 1))
    
    with np.errstate(divide = 'ignore', invalid = 'ignore'):
        
        corr = (nu0**nu) * kv(nu, nu0) / denom
    
    corr[nu0 == 0] = 1.0
    
    return corr                                     ## matrice (n_new x n)



''' ------ Loi predictive de Student pour un couple (theta, nu) : position et echelle'''
def loi_predictive (f_n, x_n, x_new, theta_x, nu, param_b0, param_n):
    
    f = np.asarray(f_n, dtype = float)
    
    n = len(f)
    
    un = np.ones(n)
    
    K = corr_matern_matrix(distance_pairwise(x_n, theta_x), nu)
    
    try:
        
        cK = scipy.linalg.cho_factor(K, lower = True)
        
    except np.linalg.LinAlgError:
        
        ## Points quasi confondus : tres leger terme diagonal pour la stabilite numerique
        cK = scipy.linalg.cho_factor(K + 1e-10*np.eye(n), lower = True)
    
    K_inv_1 = scipy.linalg.cho_solve(cK, un)
    
    s_11 = float(un @ K_inv_1)                                      ## 1'K^-1 1
    
    m = float(un @ scipy.linalg.cho_solve(cK, f)) / s_11            ## moyenne estimee
    
    K_inv_fm = scipy.linalg.cho_solve(cK, f - m*un)                 ## K^-1 (f - m 1)
    
    b_n = param_b0 + 0.5*float((f - m*un) @ K_inv_fm)
    
    r = correlation_croisee(x_new, x_n, theta_x, nu)               ## (n_new x n)
    
    position = m + r @ K_inv_fm
    
    q_rr = np.sum(r.T * scipy.linalg.cho_solve(cK, r.T), axis = 0) ## r'K^-1 r
    
    q_r1 = r @ K_inv_1                                              ## r'K^-1 1
    
    kappa = 1 - q_rr + ((1 - q_r1)**2) / s_11
    
    echelle = np.sqrt(np.maximum(b_n*kappa / param_n, 0.0))
    
    return position, echelle



''' ------ Quantile d'un melange (a poids egaux) de lois de Student, par dichotomie'''
def quantile_melange (M, G, df, p, n_iter = 60):
    
    G = np.maximum(G, 1e-12)
    
    q_p = t.ppf(p, df)
    
    ## Le quantile du melange est compris entre le plus petit et le plus grand
    ## quantile des composantes au meme niveau
    borne_inf = np.min(M + G*q_p, axis = 0)
    
    borne_sup = np.max(M + G*q_p, axis = 0)
    
    for it in range(n_iter):
        
        milieu = 0.5*(borne_inf + borne_sup)
        
        F = np.mean(t.cdf((milieu - M) / G, df), axis = 0)
        
        borne_inf = np.where(F < p, milieu, borne_inf)
        
        borne_sup = np.where(F < p, borne_sup, milieu)
        
    return 0.5*(borne_inf + borne_sup)



''' ------ Fonction de prediction'''
## f_n, x_n            : donnees (echelle de depart) ; x_n Series (1 variable) ou DataFrame
## lower_f, upper_f,
## lower_x, upper_x    : bornes utilisees pour la normalisation (comme MHI_Bayesian_Optimisation)
## x_new               : points a predire, echelle de depart (tableau 1D, ou tableau 2D / DataFrame
##                       avec les colonnes dans le meme ordre que x_n)
## n_simul, n_pas_theta, burn_in : chaine MCMC sur (theta, nu) et eclaircissement
## niveau              : niveau de l'intervalle de credibilite
## methode             : 'melange' -> moyenne sur les tirages de (theta, nu) (defaut)
##                       'mode'    -> un seul couple : mode a posteriori (comme Optimal_function)
## Sortie : DataFrame (x, f_inf, f_mean, f_median, f_sup) sur l'echelle de depart,
##          et DataFrame des tirages (theta, nu) retenus (theta sur l'echelle de depart)
def Bayesian_predict (f_n, lower_f, upper_f, x_n, lower_x, upper_x, x_new, n_simul = 200, n_pas_theta = 10, burn_in = 0, niveau = 0.95, nu_0 = None, methode = 'melange', random_state = None):
    
    if random_state is not None:
        
        np.random.seed(seed = random_state)
    
    if nu_0 is None:
        
        nu_0 = float(nu_int(f_n, lower_f, upper_f, x_n, lower_x, upper_x))
    
    ## ---- Normalisation (identique a MHI_Bayesian_Optimisation)
    multi_dim = len(shape(x_n)) > 1
    
    lo_x = np.asarray(lower_x, dtype = float)
    
    scale_x = np.asarray(upper_x, dtype = float) - lo_x
    
    lower_f_orig = float(lower_f)
    
    scale_f = float(upper_f) - float(lower_f)
    
    x_s = (x_n.astype(float) - lo_x) / scale_x
    
    f_s = (f_n.astype(float) - lower_f_orig) / scale_f
    
    x_new_arr = np.asarray(x_new, dtype = float)
    
    x_new_s = (x_new_arr - lo_x) / scale_x
    
    if multi_dim:
        
        d = x_s.shape[1]
        
        lo_s, up_s = [0.0]*d, [1.0]*d
        
    else:
        
        d = 1
        
        lo_s, up_s = 0.0, 1.0
    
    ## ---- Parametres de la loi a priori (memes valeurs que MHI_Bayesian_Optimisation)
    n = len(f_s)
    
    unit_vect = [1 for i in range(n)]
    
    param_a0 = 1
    
    param_b0 = 1
    
    a_n = param_a0 + 0.5*(n - 1)
    
    ## ---- Chaine MCMC sur (theta, nu)
    theta = [prior_theta(x_s, lo_s, up_s)]
    
    nu_sim = [float(gamma_random.rvs(a = 1, loc = nu_0, scale = 0.01, size = 1)[0])]
    
    for i in range(1, n_simul):
        
        nu_sim.append(Updat_nu(f_s, x_s, theta[i-1], param_b0, a_n, nu_sim[i-1], unit_vect, nu_0, d))
        
        theta.append(Updat_theta(f_s, x_s, lo_s, up_s, theta[i-1], param_b0, a_n, nu_sim[i-1], unit_vect))
    
    indices = list(range(int(np.minimum(burn_in, n_simul - 1)), n_simul, n_pas_theta))
    
    theta_ret = [theta[k] for k in indices]
    
    nu_ret = [nu_sim[k] for k in indices]
    
    if methode == 'mode':
        
        theta_ret = [mode(theta_ret, keepdims = True).mode[0]]
        
        nu_ret = [float(mode(nu_ret, keepdims = True).mode[0])]
    
    ## ---- Lois predictives pour chaque tirage retenu
    positions, echelles = [], []
    
    for th_k, nu_k in zip(theta_ret, nu_ret):
        
        pos_k, ech_k = loi_predictive(f_s, x_s, x_new_s, th_k, nu_k, param_b0, a_n)
        
        positions.append(pos_k)
        
        echelles.append(ech_k)
    
    M = np.array(positions)                         ## (n_tirages x n_new)
    
    G = np.array(echelles)
    
    df_t = 2*a_n                                    ## degres de liberte de la Student
    
    p_inf = 0.5*(1 - niveau)
    
    f_mean = np.mean(M, axis = 0)                   ## esperance du melange
    
    f_median = quantile_melange(M, G, df_t, 0.5)
    
    f_inf = quantile_melange(M, G, df_t, p_inf)
    
    f_sup = quantile_melange(M, G, df_t, 1 - p_inf)
    
    ## ---- Retour a l'echelle de depart
    def f_to_orig (v):
        
        return lower_f_orig + scale_f*np.asarray(v)
    
    if multi_dim:
        
        prediction = pd.DataFrame(x_new_arr, columns = x_n.columns)
        
    else:
        
        prediction = pd.DataFrame({(x_n.name if x_n.name is not None else "x"): x_new_arr.ravel()})
    
    prediction["f_inf"] = f_to_orig(f_inf)
    
    prediction["f_mean"] = f_to_orig(f_mean)
    
    prediction["f_median"] = f_to_orig(f_median)
    
    prediction["f_sup"] = f_to_orig(f_sup)
    
    tirages = pd.DataFrame({"theta": [(scale_x*np.asarray(th)).tolist() if multi_dim else float(scale_x*th) for th in theta_ret],
                            "nu": nu_ret})
    
    return prediction, tirages
