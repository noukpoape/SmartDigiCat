#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SmartDigiCat : interface graphique (Tkinter).

Logiciel de criblage catalytique : import et nettoyage de donnees, statistiques
descriptives, graphiques, modeles de regression (scikit-learn), prediction et
optimisation bayesienne.

Chaire industrielle - Centrale Lille.
Auteurs : Sebastien Paul, Sophie Dabo-Niang, Komlan Noukpoape,
Wahba Saweiras, Svetlana Heyte.

Organisation du fichier
-----------------------
L'interface partage son etat entre les callbacks par des variables globales
de ce module (instructions ``global``, ``globals()``). Elle reste donc dans un
seul module, decoupe en sections numerotees (bandeaux ``# ==  N.``) :

 1. Imports                              14. Resultats des modeles et Run
 2. Etat global                          15. Prediction (bouton Log)
 3. Langue et police                     16. Configuration d'un modele
 4. Utilitaires                          17. Graphiques descriptifs
 5. Projet, figures, ressources          18. Optimisation
 6. Utilitaires d'export                 19. Zoom et nuage de points
 7. Execution arretable                  20. Filtres de Edit data
 8. Selection de variables / CV          21. Detail des modeles
 9. Modeles de regression                22. Tableaux statistiques
10. Plusieurs modeles, Undo / Redo       23. Panneaux Edit data et Plots
11. Entrainement (Run)                   24. Actions de la barre du haut
12. Enregistrement du projet             25. Fenetre principale
13. Lecture des donnees, ouverture       26. Autosave, boutons, manuel

Le calcul de l'optimisation bayesienne est dans
``smartdigicat/core/bayesian_optimization.py`` et les tables de traduction
dans ``smartdigicat/translations.py``.

Lancement : ``python -m smartdigicat`` (ou ``python run_smartdigicat.py``).
"""


# ==========================================================================
# ==  1. Imports                                                          ==
# ==========================================================================

import os, sys
import os.path
import tkinter
from tkinter import *
import tkinter as tk
from tkinter import ttk
from PIL import ImageTk, Image, ImageGrab
from tkmacosx import Button
from tkinter.messagebox import showinfo
from tkinter import messagebox
from tkinter import StringVar
from tkinter import simpledialog
from tkinter import colorchooser
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
from tkinter.filedialog import asksaveasfile 
import tkmacosx as tkm
import webbrowser
import numpy
from numpy import*
import numpy as np
import matplotlib
import matplotlib.pyplot as plt ## package graphique
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
import pandas as pd
from tkmacosx import SaveVar
from math import*
import csv
import json
import time
import openpyxl
from plotnine import *
import plotnine
from fpdf import FPDF
from mpl_toolkits import mplot3d
import seaborn as sns
import sklearn
from sklearn.linear_model import LinearRegression
from sklearn.linear_model import Lasso
from sklearn.model_selection import train_test_split
from sklearn.metrics import r2_score
from sklearn.gaussian_process.kernels import DotProduct, WhiteKernel
from sklearn.model_selection import KFold, StratifiedKFold
from sklearn.model_selection import cross_val_score ## Pour la validation croisée
from sklearn.model_selection import cross_val_predict
from sklearn.tree import DecisionTreeRegressor
from sklearn.svm import SVR
from sklearn import svm
from sklearn.kernel_ridge import KernelRidge
from sklearn.ensemble import RandomForestRegressor
from sklearn.gaussian_process import GaussianProcessRegressor
from sklearn.ensemble import GradientBoostingRegressor
from sklearn import ensemble
from sklearn.metrics import mean_squared_error
from sklearn.feature_selection import RFE
from mlxtend.feature_selection import SequentialFeatureSelector as my_select
from scipy import stats
from scipy.stats import shapiro 
from enum import Enum
from random import sample
import colorsys
from docx import Document
import itertools
from pandastable import Table 
from tableone import TableOne 
import re
from sklearn.model_selection import GridSearchCV
from sklearn.model_selection import RandomizedSearchCV
from sklearn.gaussian_process.kernels import Matern as R_n_Matern
from sklearn.preprocessing import MinMaxScaler
from scipy.stats import qmc
from scipy.stats import gamma as gamma_random
from scipy.stats import norm
from scipy.stats import invgamma
from scipy.stats import chi
from scipy.stats import chi2
from mpmath import gamma
from scipy.special import kv
from scipy.spatial.distance import cdist
from scipy.stats import t
from scipy.stats import uniform as my_uniform
import scipy
from scipy.stats import mode 



# ---------------------------------------------------------------------------
# La bibliotheque tierce pandastable (utilisee ici pour tous les tableaux,
# ex. "Overview stat") fait en interne un `fillna` sur des colonnes object
# lorsqu'elle dessine ou fait defiler un tableau (Table.redrawVisible), ce qui
# declenche ce FutureWarning de pandas a chaque fois. Ce n'est pas notre code
# qui est en cause (voir https://github.com/dmnfarrell/pandastable/issues/251)
# et l'avertissement est sans consequence : on le masque, sans toucher aux
# autres avertissements pandas qui pourraient etre utiles.
# ---------------------------------------------------------------------------
import warnings
warnings.filterwarnings(
    "ignore",
    message = "Downcasting object dtype arrays on .fillna, .ffill, .bfill is deprecated.*",
    category = FutureWarning,
)


# ==========================================================================
# ==  2. Etat global partage par les callbacks de l'interface             ==
# ==========================================================================
# ---- Global state variables shared across GUI callbacks ----
data_test_graph = None
data = None  
file_path  = None
file_sheets = None          # feuille(s) Excel importee(s) pour file_path (None : premiere feuille / fichier non Excel)
Target_column = None
Feature_column  = None
Calculation_method  = None
codage_method = None
Test_size  = None
Target_list = None
header_target = None
Name_model = None
nom_modele_actif = None  

nom_projet = None  
chemin_projet = None 

graph_stat_all_pred_widget = None  # reference vers le panneau "Overview stat" actuellement affiche (Edit data > Overview stat)
graph_stat_all_widget = None  # reference vers le panneau "Overview stat" actuellement affiche (donnees principales)
data_stat_all_exp = None      # tableau affiche par "Overview stat" (stat_data_all), exporte par "Export stat"

## Derniere selection valide (X, Y, Couleur) faite dans "Scatter plot", pour la
## restaurer automatiquement quand on quitte l'onglet "Plots" puis qu'on y revient
dernier_x_scatter = None
dernier_y_scatter = None
dernier_color_scatter = None

## Meme principe pour "Box plot" : methode (Bivariate/Trivariate), colonnes X/Y
## et, en trivarie, la colonne de couleur
dernier_methode_boxplot = None
dernier_x_boxplot = None
dernier_y_boxplot = None
dernier_color_boxplot = None

## Meme principe pour "Histogram" : colonne cible, couleur, couleur de densite
dernier_var_histogram = None
dernier_couleur_histogram = None
dernier_densite_histogram = None

## Meme principe pour "plot 3D" : colonnes X/Y/Z et couleur
dernier_x_plot3d = None
dernier_y_plot3d = None
dernier_z_plot3d = None
dernier_color_plot3d = None

## Meme principe pour "Plot corr" (correlation lineaire) : colonne cible et
## liste des colonnes "feature" cochees (selection multiple)
dernier_target_correlation = None
dernier_features_correlation = None
linear_correlation_widgets = (None, None)  # reference vers le panneau (windows_graphe, windows_param) actuellement affiche par "Plot corr"

## Meme principe pour "Other analysis" (min/mean/max/median) : dimension,
## methode et liste des colonnes "feature" cochees
dernier_dimension_min_max = None
dernier_methode_min_max = None
dernier_features_min_max = None
mean_min_max_widgets = (None, None)  # reference vers le panneau (windows_graphe, windows_param) actuellement affiche par "Other analysis"
calcul_stat_widgets = (None, None)  # reference vers le panneau (windows_text, features) actuellement affiche par "Statistical data" (calcul_stat)
stat_tableaux_widget = None  # panneau des tableaux affiche par "View stat" (check_statistis), detruit au clic suivant
stat_tableaux_export = []    # tableaux affiches par "View stat" : liste de (titre, DataFrame, lignes a mettre en evidence), exportes par "Export stat"


class _CaseNonCochee:
    """ Remplace une BooleanVar pas encore creee : .get() renvoie False.
    (une vraie BooleanVar ne peut pas etre creee avant la fenetre Tk) """
    def get(self):
        return False

## Case "Categorizing variables" : creee dans my_Edit_all_data (Edit data) ou dans la
## fenetre de prediction. Tant que ces fenetres n'ont pas ete ouvertes, les fonctions
## qui la lisent (my_TableOne...) la considerent comme non cochee (evite un NameError).
categorial_var_def = _CaseNonCochee()
categorial_var_def_pred = _CaseNonCochee()
result_recod_stat = []
data_sort_stat = []

windows_plot_widgets = (None, None)  # reference vers le panneau (graph, features) actuellement affiche dans "Plots" (scatter/box/hist/3D)

modeles_projet = {} 
modeles_supprimes = {} 
positions_boites_modeles = {}
couleurs_boites_modeles = {}
positions_boites_supprimees = {}
couleurs_boites_supprimees = {}
taille_boites_modeles = 88
zoom_graphique_modeles = 1.0
decalage_zoom_x = 0.0
decalage_zoom_y = 0.0
ZOOM_MIN_GRAPHIQUE = 0.4
ZOOM_MAX_GRAPHIQUE = 3.0

## Zoom applique au graphique de "Scatter plot" via la barre verticale gauche
## Zoom et decalage (pan) des graphiques du panneau "Plots" et de "Plot corr",
## pilotes par la barre verticale gauche. Chaque type de graphique (scatter,
## boxplot, histogram, plot3d, correlation) a son propre niveau de zoom et sa
## propre position : zoomer sur l'un ne doit pas affecter les autres.
ZOOM_MIN_REPRESENTATION = 0.4
ZOOM_MAX_REPRESENTATION = 3.0
zoom_representation_par_type = {}       # {cle: niveau de zoom}
decalage_zoom_x_representation_par_type = {}  # {cle: decalage x}
decalage_zoom_y_representation_par_type = {}  # {cle: decalage y}
graph_window_representation_widget = None  # dernier graphique dessine (scatter, box plot, histogramme, plot 3D...), pour le detruire avant d'en redessiner un autre

## Memorise, pour chacun des 4 parametres de filtrage de la fenetre "Edit data",
## la colonne choisie et l'intervalle [min, max] applique, afin de pre-remplir
## automatiquement ces champs la prochaine fois que la fenetre est rouverte.
filtre_edit_parametre_1 = {"colonne": None, "min": 0, "max": 100}
filtre_edit_parametre_2 = {"colonne": None, "min": 0, "max": 100}
filtre_edit_parametre_3 = {"colonne": None, "min": 0, "max": 100}
filtre_edit_parametre_4 = {"colonne": None, "min": 0, "max": 100}

# ======================================================================
# ==  3. Langue et police (bouton "langue" de la barre gauche)            ==
# ======================================================================
# Principe : le code existant n'est pas reecrit message par message.
#  - les boites de dialogue (messagebox, askstring, selecteurs de fichiers,
#    choix de couleur) sont "enveloppees" : titre et message passent par
#    traduire() juste avant d'etre affiches ;
#  - les textes d'instruction des panneaux (Label rouges / gris) sont crees
#    avec text = traduire("..."), et sont retraduits a chaud quand on change
#    de langue (cf. _retraduire_labels).
# Langues disponibles : francais et anglais (tables ci-dessous, hors ligne).
# Police (meme boite de dialogue) : la police et la taille choisies s'appliquent
#  - aux polices nommees de Tk (TkDefaultFont...), donc aux widgets qui n'en
#    precisent pas ;
#  - aux widgets deja affiches (_appliquer_police_partout) ;
#  - aux widgets crees ensuite : la creation de tout widget tkinter est
#    interceptee (_widget_init_avec_police), car l'application donne une police
#    explicite a presque tous ses widgets, ex. font = ('arial', 15).
# La taille choisie est celle du texte courant (TAILLE_POLICE_REFERENCE) : les
# autres tailles (titres, sous-titres...) sont mises a l'echelle proportionnellement.
# Pour modifier/ajouter un message : ajouter un tuple (francais, anglais) dans
# _TRADUCTIONS_EXACTES (phrase entiere) ou _TRADUCTIONS_FRAGMENTS (morceau d'un
# message construit avec un nom, un chemin, un nombre...).
import re as _re_langue
import builtins as _builtins_langue   # `from numpy import *` masque max/min/sum... du fichier
import tkinter.font as _tkfont_langue

LANGUES_DISPONIBLES = {"fr": "French", "en": "English"}
_ORDRE_LANGUES = ("fr", "en")        # colonnes des tuples des tables ci-dessous
langue_courante = "en"               # langue par defaut au lancement du logiciel ("fr" ou "en")
fenetre_langue = None                # reference vers la boite de dialogue, si ouverte

TAILLE_POLICE_REFERENCE = 15         # taille du texte courant de l'application (le plus frequent : ('arial', 15))
police_famille_choisie = None        # None = polices d'origine de l'application (arial / Helvetica)
police_taille_choisie = TAILLE_POLICE_REFERENCE
_dialogue_langue_en_construction = False   # les widgets de la boite de choix ne suivent pas la police choisie

## Textes complets (comparaison exacte) : titres de fenetres et instructions
## affichees dans les panneaux.   (francais, anglais)
## Tables (francais, anglais) : voir smartdigicat/translations.py
from smartdigicat.translations import _TRADUCTIONS_EXACTES, _TRADUCTIONS_FRAGMENTS

_traducteurs_en_cache = {}


def _obtenir_traducteur(code_langue):
    """Construit (une seule fois par langue) les tables de traduction vers
    `code_langue` : chaque texte de l'AUTRE langue est associe a son equivalent."""
    if code_langue in _traducteurs_en_cache:
        return _traducteurs_en_cache[code_langue]
    indice = _ORDRE_LANGUES.index(code_langue)
    tables = []
    for entrees in (_TRADUCTIONS_EXACTES, _TRADUCTIONS_FRAGMENTS):
        correspondance = {}
        for entree in entrees:
            if len(entree) <= indice or not entree[indice]:
                continue
            for j, source in enumerate(entree):
                if j != indice and source and source != entree[indice]:
                    correspondance.setdefault(source, entree[indice])
        tables.append(correspondance)
    exactes, fragments = tables
    motif = None
    if len(fragments) != 0:
        ## les fragments les plus longs sont essayes en premier
        motif = _re_langue.compile("|".join(_re_langue.escape(cle) for cle in sorted(fragments, key = len, reverse = True)))
    _traducteurs_en_cache[code_langue] = (exactes, fragments, motif)
    return _traducteurs_en_cache[code_langue]


def _adapter_guillemets(texte, code_langue):
    """Hors francais, remplace les guillemets francais « ... » par des guillemets anglais."""
    if code_langue == "fr":
        return texte
    texte = _re_langue.sub(r"«\s?", "\u201c", texte)
    texte = _re_langue.sub(r"\s?»", "\u201d", texte)
    return texte


def traduire(texte, exact_seulement = False):
    """Renvoie `texte` dans la langue courante (langue_courante).
    Un texte inconnu est renvoye tel quel. Avec exact_seulement = True, seuls
    les textes complets sont traduits (utilise pour les Label deja affiches)."""
    if not isinstance(texte, str) or texte == "" or langue_courante not in _ORDRE_LANGUES:
        return texte
    exactes, fragments, motif = _obtenir_traducteur(langue_courante)
    if texte in exactes:
        return _adapter_guillemets(exactes[texte], langue_courante)
    if exact_seulement or motif is None:
        return texte
    resultat, nombre = motif.subn(lambda m: fragments[m.group(0)], texte)
    if nombre == 0:
        return texte
    return _adapter_guillemets(resultat, langue_courante)


def _envelopper_dialogue(fonction, cles_texte = ("title", "message", "prompt"), positionnels = True):
    """Renvoie une version de `fonction` (une boite de dialogue tkinter) dont
    le titre / message / invite sont traduits avant affichage."""
    if getattr(fonction, "_traduit", False):
        return fonction

    def enveloppe(*args, **kwargs):
        if positionnels:
            args = tuple(traduire(a) if isinstance(a, str) else a for a in args)
        for cle in cles_texte:
            if isinstance(kwargs.get(cle), str):
                kwargs[cle] = traduire(kwargs[cle])
        ## un message affiche pendant une execution (fin de « Run », erreur...) : l'encadre de progression
        ## disparait AVANT (il restait sinon affiche et anime derriere le message jusqu'au clic sur OK)
        terminer = globals().get("_terminer_indicateur_execution")
        etat_execution = globals().get("_execution_run")
        if terminer is not None and etat_execution is not None and etat_execution.get("en_cours"):
            terminer()
        return fonction(*args, **kwargs)

    enveloppe.__name__ = getattr(fonction, "__name__", "dialogue")
    enveloppe._traduit = True
    return enveloppe


for _nom_fonction in ("showinfo", "showwarning", "showerror", "askquestion", "askokcancel",
                      "askyesno", "askyesnocancel", "askretrycancel"):
    if hasattr(messagebox, _nom_fonction):
        setattr(messagebox, _nom_fonction, _envelopper_dialogue(getattr(messagebox, _nom_fonction)))
for _nom_fonction in ("askstring", "askinteger", "askfloat"):
    if hasattr(simpledialog, _nom_fonction):
        setattr(simpledialog, _nom_fonction, _envelopper_dialogue(getattr(simpledialog, _nom_fonction)))
for _nom_fonction in ("askopenfilename", "asksaveasfilename", "askopenfilenames", "askdirectory"):
    if hasattr(filedialog, _nom_fonction):
        setattr(filedialog, _nom_fonction, _envelopper_dialogue(getattr(filedialog, _nom_fonction), ("title",), False))
colorchooser.askcolor = _envelopper_dialogue(colorchooser.askcolor, ("title",), False)
showinfo = messagebox.showinfo   # meme nom importe directement en haut du fichier
del _nom_fonction


def _retraduire_labels(widget):
    """Parcourt tous les widgets et retraduit les Label dont le texte est une
    instruction connue, pour que le changement de langue soit visible tout de
    suite sur les panneaux deja affiches."""
    try:
        enfants = widget.winfo_children()
    except Exception:
        return
    for enfant in enfants:
        try:
            if enfant.winfo_class() in ("Label", "TLabel"):
                texte = str(enfant.cget("text"))
                nouveau_texte = traduire(texte, exact_seulement = True)
                if nouveau_texte != texte:
                    enfant.configure(text = nouveau_texte)
        except Exception:
            pass
        _retraduire_labels(enfant)


## Boites de modeles (Canvas) : le texte "nom + etat" est un element de canevas, pas un widget
## a texte ; on garde donc (canevas, id du texte, nom) pour le mettre a jour au changement de langue.
_textes_etat_modeles = []


def _enregistrer_texte_etat_modele(canevas, id_texte, nom):
    """Appelee a la creation du texte d'une boite de modele. Oublie au passage les boites detruites."""
    vivants = []
    for entree in _textes_etat_modeles:
        try:
            if entree[0].winfo_exists():
                vivants.append(entree)
        except Exception:
            pass
    vivants.append((canevas, id_texte, nom))
    _textes_etat_modeles[:] = vivants


def _rafraichir_etat_modeles():
    """Reecrit "nom + etat" dans les boites de modeles affichees, dans la langue courante,
    ainsi que l'algorithme choisi, en haut de chaque boite."""
    try:
        _rafraichir_algo_boites_modeles()
    except Exception:
        pass
    vivants = []
    for canevas, id_texte, nom in _textes_etat_modeles:
        try:
            if not canevas.winfo_exists():
                continue
            canevas.itemconfig(id_texte, text = nom + _statut_affichage_modele(nom))
            vivants.append((canevas, id_texte, nom))
        except Exception:
            pass
    _textes_etat_modeles[:] = vivants


def appliquer_langue(code_langue):
    """Change la langue des instructions et met a jour ce qui est deja affiche."""
    global langue_courante
    if code_langue not in LANGUES_DISPONIBLES:
        return
    langue_courante = code_langue
    _retraduire_labels(SmartDigicat)
    _rafraichir_etat_modeles()
    ## manuel affiche (bouton Help) : il passe dans la nouvelle langue
    try:
        if _manuel_affiche["widget"] is not None and _manuel_affiche["widget"].winfo_exists():
            afficher_manuel(None, _manuel_affiche["parent"])
    except Exception:
        pass


## ----------------------------------------------------------------------
## Police
## ----------------------------------------------------------------------
_NOMS_POLICES_NOMMEES = ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont",
                         "TkCaptionFont", "TkSmallCaptionFont", "TkIconFont", "TkTooltipFont")
_polices_nommees_origine = {}      # {nom: (famille, taille)} avant toute modification


def _police_modifiee():
    """Vrai si l'utilisateur a choisi autre chose que les polices d'origine."""
    return police_famille_choisie is not None or police_taille_choisie != TAILLE_POLICE_REFERENCE


def _taille_mise_a_l_echelle(taille_origine):
    """Taille d'origine -> taille proportionnelle a la taille de police choisie."""
    return _builtins_langue.max(6, int(taille_origine * police_taille_choisie / TAILLE_POLICE_REFERENCE + 0.5))


def _police_reelle(spec):
    """Caracteristiques reelles (famille, taille, weight, slant...) d'une police
    quelconque : tuple ('arial', 15), texte 'Helvetica 16 bold italic'..."""
    brut = [str(x) for x in SmartDigicat.tk.splitlist(SmartDigicat.tk.call("font", "actual", spec))]
    return {brut[i].lstrip("-"): brut[i + 1] for i in range(0, len(brut) - 1, 2)}


def _police_transformee(origine):
    """Police d'origine d'un widget -> police a utiliser avec le choix actuel.
    Les polices nommees (TkDefaultFont...) sont laissees telles quelles : on modifie
    directement ces polices nommees (cf. appliquer_police)."""
    if not _police_modifiee():
        return origine
    try:
        if str(origine) in _tkfont_langue.names(SmartDigicat):
            return origine
        reel = _police_reelle(origine)
        police = [police_famille_choisie or reel["family"],
                  _taille_mise_a_l_echelle(abs(int(float(reel["size"]))))]
        if reel.get("weight") == "bold":
            police.append("bold")
        if reel.get("slant") == "italic":
            police.append("italic")
        if str(reel.get("underline")) == "1":
            police.append("underline")
        if str(reel.get("overstrike")) == "1":
            police.append("overstrike")
        return tuple(police)
    except Exception:
        return origine


## Interception de la creation des widgets : ils recoivent la police choisie
## des leur creation (les panneaux de l'application sont construits a chaque clic).
_widget_init_sans_police = tkinter.Widget.__init__


def _widget_init_avec_police(self, master, widgetName, cnf = {}, kw = {}, extra = ()):
    origine = None
    if not _dialogue_langue_en_construction:
        for nom_dico in ("kw", "cnf"):
            dico = kw if nom_dico == "kw" else cnf
            if isinstance(dico, dict) and dico.get("font") is not None:
                origine = dico["font"]
                if _police_modifiee():
                    dico = dict(dico)
                    dico["font"] = _police_transformee(origine)
                    if nom_dico == "kw":
                        kw = dico
                    else:
                        cnf = dico
                break
    _widget_init_sans_police(self, master, widgetName, cnf, kw, extra)
    if origine is not None:
        self._police_origine = origine      # police d'origine, pour pouvoir la recalculer


if not getattr(tkinter.Widget.__init__, "_police_patch", False):
    _widget_init_avec_police._police_patch = True
    tkinter.Widget.__init__ = _widget_init_avec_police


def _capturer_polices_nommees():
    """Memorise (une fois, avant toute modification) les polices nommees de Tk."""
    if len(_polices_nommees_origine) != 0:
        return
    for nom in _NOMS_POLICES_NOMMEES:
        try:
            reel = _tkfont_langue.nametofont(nom, root = SmartDigicat).actual()
            _polices_nommees_origine[nom] = (reel["family"], abs(int(reel["size"])))
        except Exception:
            pass


def _appliquer_police_partout(widget):
    """Recalcule la police de tous les widgets deja affiches (fenetres comprises)."""
    try:
        enfants = widget.winfo_children()
    except Exception:
        return
    for enfant in enfants:
        try:
            origine = getattr(enfant, "_police_origine", None)
            if origine is None and "font" in enfant.keys():
                courante = str(enfant.cget("font"))
                if courante not in _tkfont_langue.names(SmartDigicat):
                    origine = courante
                    enfant._police_origine = courante
            if origine is not None:
                enfant.configure(font = _police_transformee(origine))
        except Exception:
            pass
        _appliquer_police_partout(enfant)


def appliquer_police(famille, taille):
    """Applique la police (famille = None : polices d'origine) et la taille du
    texte courant a toute l'application."""
    global police_famille_choisie, police_taille_choisie
    _capturer_polices_nommees()
    police_famille_choisie = famille
    police_taille_choisie = int(taille)
    for nom, (famille_origine, taille_origine) in _polices_nommees_origine.items():
        try:
            _tkfont_langue.nametofont(nom, root = SmartDigicat).configure(
                family = famille or famille_origine, size = _taille_mise_a_l_echelle(taille_origine))
        except Exception:
            pass
    _appliquer_police_partout(SmartDigicat)


## ----------------------------------------------------------------------
## Boite de dialogue : langue + police
## ----------------------------------------------------------------------
def choisir_langue_instructions():
    """Boite de dialogue (appelee par le bouton `button_langue` de la barre
    verticale gauche) : langue des instructions, police et taille de police."""
    global fenetre_langue, _dialogue_langue_en_construction

    ## une seule boite a la fois : si elle est deja ouverte, on la ramene devant
    if fenetre_langue is not None:
        try:
            if fenetre_langue.winfo_exists():
                fenetre_langue.lift()
                fenetre_langue.focus_force()
                return
        except Exception:
            pass

    _dialogue_langue_en_construction = True
    try:
        _construire_boite_langue()
    finally:
        _dialogue_langue_en_construction = False


def _construire_boite_langue():
    global fenetre_langue
    _capturer_polices_nommees()

    fenetre = tk.Toplevel(SmartDigicat)
    fenetre_langue = fenetre
    fenetre.title("Langue et police / Language and font")
    fenetre.resizable(False, False)
    fenetre.transient(SmartDigicat)
    fenetre.option_add("*TCombobox*Listbox.font", ('arial', 13))

    ## textes bilingues : lisibles quelle que soit la langue actuellement active
    tk.Label(fenetre, text = "Langue et police\nLanguage and font", font = ('arial', 14, 'bold'), justify = tk.CENTER).pack(padx = 30, pady = (18, 8))

    ## --- langue : liste deroulante (langue active preselectionnee)
    tk.Label(fenetre, text = "Votre langue préférée / Your preferred language", font = ('arial', 12)).pack(pady = (4, 4))
    codes_langues = list(LANGUES_DISPONIBLES)
    liste_langues = ttk.Combobox(fenetre, values = [LANGUES_DISPONIBLES[c] for c in codes_langues], state = "readonly", width = 34, font = ('arial', 13))
    liste_langues.current(codes_langues.index(langue_courante) if langue_courante in codes_langues else 0)
    liste_langues.pack(padx = 30, pady = (0, 8))

    ## --- police : famille
    familles = sorted({f for f in _tkfont_langue.families(SmartDigicat) if not f.startswith("@")}, key = lambda f: f.lower())
    choix_familles = [None] + familles       # None = polices d'origine de l'application
    tk.Label(fenetre, text = "Type de police / Font", font = ('arial', 12)).pack(pady = (4, 4))
    liste_polices = ttk.Combobox(fenetre, values = ["Par défaut / Default (arial)"] + familles, state = "readonly", width = 34, height = 15, font = ('arial', 13))
    liste_polices.current(choix_familles.index(police_famille_choisie) if police_famille_choisie in choix_familles else 0)
    liste_polices.pack(padx = 30, pady = (0, 8))

    ## --- police : taille du texte courant
    tailles = list(range(8, 33))
    tk.Label(fenetre, text = "Taille de police / Font size", font = ('arial', 12)).pack(pady = (4, 4))
    liste_tailles = ttk.Combobox(fenetre, values = [(str(t) + ("  (par défaut / default)" if t == TAILLE_POLICE_REFERENCE else "")) for t in tailles], state = "readonly", width = 34, height = 15, font = ('arial', 13))
    liste_tailles.current(tailles.index(police_taille_choisie) if police_taille_choisie in tailles else tailles.index(TAILLE_POLICE_REFERENCE))
    liste_tailles.pack(padx = 30, pady = (0, 8))

    def police_par_defaut():
        liste_polices.current(0)
        liste_tailles.current(tailles.index(TAILLE_POLICE_REFERENCE))

    def valider(event = None):
        code_langue = codes_langues[liste_langues.current()]
        famille = choix_familles[liste_polices.current()]
        taille = tailles[liste_tailles.current()]
        fenetre.destroy()                        # d'abord fermer la boite : elle ne suit pas la police choisie
        appliquer_langue(code_langue)
        appliquer_police(famille, taille)

    def annuler(event = None):
        fenetre.destroy()

    cadre_boutons = tk.Frame(fenetre)
    cadre_boutons.pack(pady = (12, 18))
    ttk.Button(cadre_boutons, text = "OK", command = valider).pack(side = tk.LEFT, padx = 6)
    ttk.Button(cadre_boutons, text = "Police par défaut / Default font", command = police_par_defaut).pack(side = tk.LEFT, padx = 6)
    ttk.Button(cadre_boutons, text = "Annuler / Cancel", command = annuler).pack(side = tk.LEFT, padx = 6)

    fenetre.bind("<Return>", valider)
    fenetre.bind("<Escape>", annuler)
    fenetre.after(50, fenetre.focus_force)       # le clavier va a la boite (sinon Entree arrive a la fenetre principale)
    fenetre.bind("<Escape>", annuler)
    fenetre.protocol("WM_DELETE_WINDOW", annuler)

    ## centrage par rapport a la fenetre principale
    fenetre.update_idletasks()
    x = SmartDigicat.winfo_rootx() + (SmartDigicat.winfo_width() - fenetre.winfo_reqwidth()) // 2
    y = SmartDigicat.winfo_rooty() + (SmartDigicat.winfo_height() - fenetre.winfo_reqheight()) // 2
    fenetre.geometry("+%d+%d" % (_builtins_langue.max(x, 0), _builtins_langue.max(y, 0)))

    fenetre.grab_set()
    fenetre.focus_force()


# ======================================================================
# ==  3b. Fermeture propre de l'application                               ==
# ======================================================================
# Les taches planifiees avec after() (enregistrement automatique, mise a jour des
# boutons...) doivent etre annulees AVANT la destruction de la fenetre. Sinon Tk les
# execute sur une application detruite et affiche a la fermeture :
#   bgerror failed to handle background error.
#   Original error: invalid command name "...autosave_projet"
_taches_application = {}            # {nom: identifiant retourne par after()}
_application_en_fermeture = False


def _planifier_tache_application(nom, delai_ms, fonction):
    """Equivalent de SmartDigicat.after(delai_ms, fonction), mais l'identifiant est
    conserve pour pouvoir annuler la tache a la fermeture."""
    if _application_en_fermeture:
        return
    try:
        _taches_application[nom] = SmartDigicat.after(delai_ms, fonction)
    except Exception:
        pass                        # fenetre deja detruite


def _annuler_taches_application(event = None):
    """Annule toutes les taches planifiees. Appelee a la fermeture (bouton de la fenetre)
    et a la destruction de la fenetre principale."""
    global _application_en_fermeture
    ## <Destroy> lie a la fenetre principale est aussi recu pour chacun de ses widgets
    if event is not None and str(event.widget) != str(SmartDigicat):
        return
    _application_en_fermeture = True
    for identifiant in list(_taches_application.values()):
        try:
            SmartDigicat.after_cancel(identifiant)
        except Exception:
            pass
    _taches_application.clear()


## ------ Enregistrement du travail a la fermeture du logiciel ------
## Quand l'utilisateur ferme la fenetre (croix, Alt+F4) ou quitte (Cmd+Q, menu « Quitter » du Dock sous macOS), une boite
## lui propose : Enregistrer / Ne pas enregistrer / Annuler. Elle n'apparait que si le projet a ete modifie depuis
## son dernier enregistrement (manuel « Save » ou automatique toutes les 5 minutes). Mettre la constante a True pour
## la proposer a chaque fermeture, meme sans modification.
TOUJOURS_DEMANDER_A_LA_FERMETURE = False

_id_etat_enregistre = {"signature": None}        # « photo » du projet au dernier enregistrement
_fermeture_dialogue_ouvert = False


def _signature_projet():
    """Resume comparable de l'etat du projet : point courant de l'historique (modeles, donnees, prediction, filtres...)
    + elements non suivis par l'historique mais enregistres avec le projet (parametres d'optimisation, dernieres
    selections de graphiques, graphiques affiches)."""
    signature = []
    try:
        signature.append(_historique["etats"][_historique["index"]].get("id") if _historique["etats"] else None)
    except Exception:
        signature.append(None)
    for nom in ("opt_parameters", "opt_affichage"):
        try:
            signature.append(repr(globals().get(nom)))
        except Exception:
            signature.append(None)
    for nom in ("data_opt", "opt_plot_data"):
        valeur = globals().get(nom)
        signature.append(None if valeur is None else id(valeur))
    try:
        signature.append(tuple(sorted(str(k) for k in figures_ouvertes_courantes)))
    except Exception:
        signature.append(None)
    for nom in sorted(n for n in list(globals()) if n.startswith("dernier_")):
        try:
            signature.append((nom, repr(globals().get(nom))))
        except Exception:
            signature.append((nom, None))
    return tuple(signature)


def _projet_marquer_enregistre():
    """A appeler apres tout enregistrement reussi (ou a l'ouverture / creation d'un projet)."""
    try:
        _id_etat_enregistre["signature"] = _signature_projet()
    except Exception:
        _id_etat_enregistre["signature"] = None


def projet_a_des_modifications():
    """Vrai si un projet est ouvert et differe de son dernier enregistrement."""
    if nom_projet is None:
        return False
    try:
        return _signature_projet() != _id_etat_enregistre.get("signature")
    except Exception:
        return True                     # dans le doute, on propose d'enregistrer


def _boite_enregistrer_avant_quitter(nom, modifie):
    """Boite modale a trois choix. Renvoie « enregistrer », « quitter » (sans enregistrer) ou « annuler »."""
    choix = {"valeur": "annuler"}
    fenetre = tk.Toplevel(SmartDigicat)
    fenetre.title(traduire("Enregistrer avant de quitter"))
    fenetre.transient(SmartDigicat)
    fenetre.resizable(False, False)

    def choisir(valeur, event = None):
        choix["valeur"] = valeur
        fenetre.destroy()

    tk.Label(fenetre, text = traduire("Voulez-vous enregistrer votre travail avant de quitter SmartDigiCat ?"),
             font = ("arial", 15, "bold"), wraplength = 480, justify = "left").pack(padx = 24, pady = (20, 6), anchor = "w")
    if modifie:
        detail = (traduire("Le projet « " + str(nom) + " » contient des modifications non enregistrees.") + "\n" +
                  traduire("Vos modifications seront perdues si vous ne les enregistrez pas."))
    else:
        detail = traduire("Le projet « " + str(nom) + " » est ouvert.")
    tk.Label(fenetre, text = detail, font = ("arial", 13), wraplength = 480, justify = "left",
             fg = "gray30").pack(padx = 24, pady = (0, 14), anchor = "w")
    boutons = tk.Frame(fenetre)
    boutons.pack(pady = (0, 18))
    bouton_enregistrer = ttk.Button(boutons, text = traduire("Enregistrer"), command = lambda: choisir("enregistrer"))
    bouton_enregistrer.pack(side = tk.LEFT, padx = 6)
    ttk.Button(boutons, text = traduire("Ne pas enregistrer"), command = lambda: choisir("quitter")).pack(side = tk.LEFT, padx = 6)
    ttk.Button(boutons, text = traduire("Annuler"), command = lambda: choisir("annuler")).pack(side = tk.LEFT, padx = 6)
    fenetre.bind("<Return>", lambda event: choisir("enregistrer"))
    fenetre.bind("<Escape>", lambda event: choisir("annuler"))
    fenetre.after(50, fenetre.focus_force)
    fenetre.protocol("WM_DELETE_WINDOW", lambda: choisir("annuler"))     # croix de la boite = Annuler (on reste dans le logiciel)
    # centrage sur la fenetre principale
    fenetre.update_idletasks()
    try:
        x = SmartDigicat.winfo_rootx() + (SmartDigicat.winfo_width() - fenetre.winfo_reqwidth()) // 2
        y = SmartDigicat.winfo_rooty() + (SmartDigicat.winfo_height() - fenetre.winfo_reqheight()) // 2
        fenetre.geometry("+%d+%d" % (_builtins_langue.max(x, 0), _builtins_langue.max(y, 0)))
    except Exception:
        pass
    fenetre.grab_set()
    bouton_enregistrer.focus_set()
    fenetre.focus_force()
    SmartDigicat.wait_window(fenetre)
    return choix["valeur"]


def _demander_enregistrement_avant_fermeture():
    """Renvoie True si la fermeture peut continuer, False si l'utilisateur veut rester dans le logiciel."""
    if nom_projet is None:
        return True                     # aucun projet ouvert : rien a enregistrer
    modifie = projet_a_des_modifications()
    if not modifie and not TOUJOURS_DEMANDER_A_LA_FERMETURE:
        return True
    choix = _boite_enregistrer_avant_quitter(nom_projet, modifie)
    if choix == "annuler":
        return False
    if choix == "quitter":
        return True
    resultat = save_current_project(silencieux = True)        # True : ok, False : erreur (deja signalee), None : « Enregistrer sous » annule
    if resultat is True:
        return True
    if resultat is None:
        return False                    # l'utilisateur a renonce a choisir un fichier : on reste dans le logiciel
    return messagebox.askyesno("Enregistrer avant de quitter", "L'enregistrement a echoue. Quitter quand meme, sans enregistrer ?")


def fermer_application(event = None):
    global _fermeture_dialogue_ouvert
    if _fermeture_dialogue_ouvert:
        return                          # boite deja affichee (ex. Cmd+Q repete)
    _fermeture_dialogue_ouvert = True
    try:
        try:
            peut_fermer = _demander_enregistrement_avant_fermeture()
        except Exception:
            peut_fermer = True          # une erreur de la boite ne doit jamais empecher de fermer le logiciel
    finally:
        _fermeture_dialogue_ouvert = False
    if not peut_fermer:
        return
    _annuler_taches_application()
    SmartDigicat.destroy()


def installer_fermeture_propre():
    SmartDigicat.protocol("WM_DELETE_WINDOW", fermer_application)
    if sys.platform == "darwin":
        try:                            # Cmd+Q et « Quitter » du menu du Dock passent par la meme demande d'enregistrement
            SmartDigicat.createcommand("::tk::mac::Quit", fermer_application)
        except Exception:
            pass
    SmartDigicat.bind("<Destroy>", _annuler_taches_application, add = "+")



# ==========================================================================
# ==  4. Utilitaires (filtres, meilleur modele, optimisation)             ==
# ==========================================================================
def _appliquer_filtre_parametre_edit(data_frame, reglage):
    colonne = reglage.get("colonne")
    if colonne is None or colonne not in data_frame.columns:
        return data_frame
    valeur_min = reglage.get("min", 0)
    valeur_max = reglage.get("max", 100)
    if valeur_min < valeur_max:
        data_frame = data_frame[data_frame[colonne] > valeur_min]
    if valeur_max > valeur_min:
        data_frame = data_frame[data_frame[colonne] <= valeur_max]
    return data_frame

def _r2_modele_entraine (nom_modele):
    ## renvoie le R-squared enregistre pour ce modele entraine, ou None si absent
    etat = modeles_projet.get(nom_modele, {})
    donnees_modele = etat.get("data_model")
    if donnees_modele is None:
        return None
    try:
        return donnees_modele[0]
    except Exception:
        return None

def _modele_utilisable_pour_prediction (nom_modele):
    ## CORRECTIF : un modele « trained » n'est pas forcement utilisable pour predire (projet recharge depuis
    ## le JSON : reg = None ; colonnes retenues absentes). Sans ce test, « View pred » plantait.
    etat = modeles_projet.get(nom_modele, {})
    if etat.get("data_model") is None or etat.get("reg") is None:
        return False
    if etat.get("columns_select") is None or etat.get("result_model") is None or etat.get("selectet_target") is None:
        return False
    if not hasattr(etat.get("reg"), "predict"):
        return False
    return True

def _meilleur_modele_entraine (noms_modeles, pour_prediction = False):
    ## compare le R-squared de chaque modele entraine et renvoie le nom et le
    ## R-squared du meilleur (evite volontairement max() : ecrase par numpy.max
    ## a cause du "from numpy import *" en tete de fichier)
    ## pour_prediction = True : ne considere que les modeles utilisables pour predire
    meilleur_nom = None
    meilleur_r2 = None
    for nom in noms_modeles:
        if pour_prediction and not _modele_utilisable_pour_prediction(nom):
            continue
        r2 = _r2_modele_entraine(nom)
        if r2 is None:
            continue
        try:
            if not np.isfinite(r2):      ## R2 = nan / inf : ne peut pas etre compare
                continue
        except Exception:
            continue
        if meilleur_r2 is None or r2 > meilleur_r2:
            meilleur_r2 = r2
            meilleur_nom = nom
    return meilleur_nom, meilleur_r2

def _filtre_edit_vers_json (reglage):
    return {
        "colonne": _valeur_reglage_vers_json(reglage.get("colonne")),
        "min": _valeur_reglage_vers_json(reglage.get("min")),
        "max": _valeur_reglage_vers_json(reglage.get("max")),
    }

def _filtre_edit_depuis_json (reglage_json):
    if not isinstance(reglage_json, dict):
        return {"colonne": None, "min": 0, "max": 100}
    return {
        "colonne": reglage_json.get("colonne"),
        "min": reglage_json.get("min", 0),
        "max": reglage_json.get("max", 100),
    }
Alpha_size = None
Alpha_lasso_size = None
C_size = None
gamma_size = None
Random_size = None
trees_size = None
estimators_size = None
depth_size = None
loss_size = None
rate_size = None
reg = None
data_analy = None
scatter_title_value = None
cor_title_value = None
min_max_title_value = None
boxplot_title_value = None
boxplot_label_x = None
boxplot_label_y = None
histogram_title_value = None
color_column = None
scatter_x_value = None
scatter_y_value = None
scatter_color_value = None
plot3d_title_value = None
plot3d_x_value = None
plot3d_y_value = None
plot3d_z_value = None
plot3d_color_value = None
cor_x_value = None
cor_y_value = None
min_max_x_value = None
min_max_y_value = None
histogram_x_value = None
histogram_y_value  = None
file_path_pred = None 
file_sheets_pred = None     # feuille(s) Excel de file_path_pred
data_pred = None            # donnees de prediction importees (memorisees dans le projet)
data_pred_new = None        # donnees de prediction + colonne predite par le meilleur modele (memorisees)
prediction_infos = {}       # modele utilise, R2, cible, colonne, fichier source (memorises)
data_opt = None 
def parametres_optimisation_par_defaut ():
    ''' Parametres de "Optimization settings" avant toute saisie de l'utilisateur. '''
    return {"method": None, "target": None, "features": [], "min_y": 0.0, "max_y": 100.0, "min_x": [], "max_x": [], "n_replication": 20, "seed": 42, "n_simul_1": 100, "n_simul_2": 10}


def parametres_optimisation_depuis_json (donnees_json):
    ''' Reconstruit les parametres d'optimisation a partir du JSON d'un projet.
    Toute valeur absente ou invalide est remplacee par sa valeur par defaut
    (un ancien projet, sans ces parametres, donne donc les valeurs par defaut). '''

    parametres = parametres_optimisation_par_defaut()
    if not isinstance(donnees_json, dict):
        return parametres

    def est_reel (valeur):
        return type(valeur) in (int, float) and np.isfinite(valeur)

    def est_entier (valeur):
        return type(valeur) is int

    for cle in ("method", "target"):
        if isinstance(donnees_json.get(cle), str):
            parametres[cle] = donnees_json[cle]
    features = donnees_json.get("features")
    if isinstance(features, list) and len([nom for nom in features if not isinstance(nom, str)]) == 0:
        parametres["features"] = list(features)
    min_y, max_y = donnees_json.get("min_y"), donnees_json.get("max_y")
    if est_reel(min_y) and est_reel(max_y) and min_y < max_y:
        parametres["min_y"], parametres["max_y"] = float(min_y), float(max_y)
    min_x, max_x = donnees_json.get("min_x"), donnees_json.get("max_x")
    if isinstance(min_x, list) and isinstance(max_x, list) and len(min_x) == len(max_x) and len([v for v in min_x + max_x if not est_reel(v)]) == 0:
        parametres["min_x"], parametres["max_x"] = [float(v) for v in min_x], [float(v) for v in max_x]
    n_replication = donnees_json.get("n_replication")
    if est_entier(n_replication) and n_replication >= 2:
        parametres["n_replication"] = n_replication
    if est_entier(donnees_json.get("seed")):
        parametres["seed"] = donnees_json["seed"]
    n_simul_1 = donnees_json.get("n_simul_1")
    if est_entier(n_simul_1) and n_simul_1 >= 10:
        parametres["n_simul_1"] = n_simul_1
    n_simul_2 = donnees_json.get("n_simul_2")
    if est_entier(n_simul_2) and n_simul_2 >= 1:
        parametres["n_simul_2"] = n_simul_2
    return parametres


opt_parameters = parametres_optimisation_par_defaut()   # parametres definis dans Edit_opt_parameters (memorises dans le projet)


def affichage_optimisation_par_defaut ():
    ''' Elements affiches sur le graphique d'optimisation (cases a cocher) et reglages du trace. '''
    return {"given": True, "opt": True, "curve": True, "error": True, "table": False, "table_given": False, "line_width": "1.5", "point_size": "15", "x_axis": None}


def affichage_optimisation_depuis_json (donnees_json):
    ''' Relit les choix d'affichage du graphique d'optimisation depuis le JSON d'un
    projet ; toute valeur absente ou invalide reprend sa valeur par defaut. '''

    affichage = affichage_optimisation_par_defaut()
    if not isinstance(donnees_json, dict):
        return affichage
    for cle in ("given", "opt", "curve", "error", "table", "table_given"):
        if donnees_json.get(cle) is True or donnees_json.get(cle) is False:
            affichage[cle] = donnees_json[cle]
    for cle in ("line_width", "point_size"):
        valeur = donnees_json.get(cle)
        if type(valeur) in (str, int, float):
            try:
                if float(str(valeur).replace(",", ".")) > 0:
                    affichage[cle] = str(valeur)
            except ValueError:
                pass
    if isinstance(donnees_json.get("x_axis"), str):
        affichage["x_axis"] = donnees_json["x_axis"]
    return affichage


opt_affichage = affichage_optimisation_par_defaut()   # derniers choix d'affichage (memorises dans le projet)


def donnees_graphique_optimisation_depuis_json (donnees_json):
    ''' Reconstruit opt_plot_data (donnees du graphique d'optimisation) depuis le
    JSON d'un projet. Retourne None si les donnees sont absentes ou invalides. '''

    if not isinstance(donnees_json, dict):
        return None
    try:
        def tableau (cle):
            return np.array(donnees_json[cle], dtype = float)

        def table (cle):
            contenu_table = donnees_json[cle]
            if not (isinstance(contenu_table, dict) and contenu_table.get("__type__") == "DataFrame"):
                raise ValueError(cle)
            return pd.DataFrame(contenu_table["data"])

        donnees = {"multi": donnees_json["multi"] is True, "xlabel": str(donnees_json["xlabel"]), "ylabel": str(donnees_json["ylabel"]),
                   "y_init": tableau("y_init"), "y_opt": tableau("y_opt"), "x_init": tableau("x_init"), "x_opt": tableau("x_opt"),
                   "table": table("table"), "table_init": table("table_init")}
        if not donnees["multi"]:
            for cle in ("x_plot", "y_plot", "y_low", "y_up"):
                donnees[cle] = tableau(cle)
            if len({len(donnees[cle]) for cle in ("x_plot", "y_plot", "y_low", "y_up")}) != 1:
                return None
        if len(donnees["x_init"]) != len(donnees["y_init"]) or len(donnees["x_opt"]) != len(donnees["y_opt"]):
            return None
        return donnees
    except (KeyError, ValueError, TypeError):
        return None
opt_result = None   # dernier resultat de l'optimisation (sortie de Yields_optimization)
opt_plot_data = None       # donnees du dernier graphique d'optimisation (calculees par check_optim_parameters ; memorisees dans le projet)
opt_canvas_widget = None   # widget Tk du graphique affiche dans le panneau
opt_table_widget = None    # widget Tk du tableau des points optimises
my_scatter_plot = None  


figures_ouvertes_courantes = {}



# ==========================================================================
# ==  5. Projet ouvert, figures et chemins des ressources                 ==
# ==========================================================================
def verifier_projet_ouvert ():
  
    if nom_projet is None:
        messagebox.showwarning("Aucun projet ouvert", "Veuillez d'abord creer un nouveau projet (bouton « New ») ou en ouvrir un existant (bouton « Open ») avant de continuer.")
        return False
    return True


def _enregistrer_figure_courante (nom, fig):
  
    if fig is not None:
        figures_ouvertes_courantes[nom] = fig

def my_path(true_path):

    ## Ressources (dossier « figures ») : sys._MEIPASS quand le programme est « gele » par PyInstaller (>= 6 :
    ## sous-dossier _internal), sinon dossier du script.
    if getattr(sys, 'frozen', False):
        base_ressources = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    else:
        base_ressources = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_ressources, true_path)

my_missing_values = ["n/a", "na", "--"]

type_style = theme(axis_line=element_line(size = 1, colour="black"), panel_grid_major=element_line(colour="#d3d3d3"), panel_grid_minor=element_blank(),  panel_border=element_blank(), panel_background=element_blank(), plot_title=element_text(size = 15, family="Tahoma", face="bold"), text=element_text(family="Tahoma", size = 11), axis_text_x=element_text(colour="black", size = 10), axis_text_y=element_text(colour="black", size = 10))

kernel = DotProduct() + WhiteKernel()

colors_palettes = ["maroon", "green", "blue", "red", "gray", "lime", "aqua", "yellow", "pink","orange","purple","violet", "brown", "magenta", "indianred", "olive", "lightskyblue", "royalblue", "navy", "deeppink", "crimson", "peru", "chocolate", "firebrick", "darkolivegreen", "beige"]

my_colors_palettes = ["saddlebrown", "olivedrab", "darkturquoise", "gold" , "cadetblue", "lime", "aqua", "yellow", "pink","lightsteelblue","orange","purple","violet", "brown", "magenta", "indianred", "olive", "lightskyblue", "royalblue", "navy", "deeppink", "crimson", "peru", "chocolate", "firebrick", "darkolivegreen", "beige", "red"]





# ==========================================================================
# ==  5b. Calcul : optimisation bayesienne (module core)                  ==
# ==========================================================================
from smartdigicat.core import bayesian_optimization as _optimisation_bayesienne
from smartdigicat.core.bayesian_optimization import (
    data_stand,
    nu_int_new,
    nu_int,
    nu_int_2,
    proposal_law_value_x,
    theta_prior_param,
    prior_theta,
    distance_pairwise,
    corr_matern_matrix,
    corr_exponential_matrix,
    Matern_vect_corr,
    prod_vect_matrix_vect,
    Scale_parameter,
    Scale_parameter_,
    Optimal_function,
    Updat_theta,
    Updat_nu,
    log_vraisemblance,
    Updat_theta_nu,
    Updat_x,
    Expected_Improvement,
    MHI_Bayesian_Optimisation,
    Yields_optimization,
    correlation_croisee,
    loi_predictive,
    quantile_melange,
    Bayesian_predict,
)


# ==========================================================================
# ==  6. Utilitaires d'export (graphiques, tableaux statistiques)         ==
# ==========================================================================

def souris_clic(event, parent):
        #Met en avant l'objet sélectionné
        x1=event.x
        y1=event.y
        objet=parent.find_closest(x1,y1)
        parent.itemconfig(objet,width=5)
        parent.lift(objet)
        parent.bind("<Button1-Motion>",lambda event :souris_deplace(event,x1,y1,parent,objet))   


def souris_deplace(event,x1,y1,parent,objet):
       #Effectue le déplacement des objets
       x2=event.x
       y2=event.y
       dx,dy=x2-x1,y2-y1
       x1,y1=x2,y2
       parent.coords(objet,x2,y2,x2+40,y2+40)
       x1,y1=x2,y2
       parent.bind("<Button1-ButtonRelease>",lambda event:souris_releve(event,x2,y2,parent,objet))  


def souris_releve(event,x2,y2,parent,objet):
       #Remet l'objet dans sa configuration initiale
       parent.itemconfig(objet,width=1)
       parent.coords(objet,x2,y2,x2+30,y2+30)
       parent.bind("<B1-Motion>", lambda event:souris_clic(event,parent))


def my_save_plot_file (img_save):

    my_file = filedialog.asksaveasfilename(filetypes=[("PNG file", ".png"), ('PDF file', ".pdf"), ('JPEG file', ".jpeg")], defaultextension = ".png")
    file_impor = my_file
    img_save.savefig(file_impor)


def my_exporting_data (my_data):

    my_file = filedialog.asksaveasfilename(filetypes = [("csv file", "*.csv")], defaultextension = "*.csv")
    my_data.to_csv(my_file, sep = ";", index = False, decimal = ".")


def _nom_feuille_excel (titre, noms_utilises):
    """ Nom de feuille Excel valide (31 caracteres max, sans caracteres interdits) et unique. """
    nom = "".join("-" if c in ':\\/?*[]' else c for c in str(titre)).strip()[:31] or "Table"
    base, numero = nom, 2
    while nom.lower() in noms_utilises:
        suffixe = " (" + str(numero) + ")"
        nom = base[:31 - len(suffixe)] + suffixe
        numero += 1
    noms_utilises.add(nom.lower())
    return nom


def _exporter_tableaux_stat_excel (chemin, tableaux):
    """ Un tableau par feuille : titre en ligne 1, en-tetes en gras sur fond
    gris, lignes de variable en gras sur fond bleu clair, largeurs ajustees. """
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    noms_utilises = set()
    with pd.ExcelWriter(chemin, engine = "openpyxl") as fichier:
        for titre, tableau, lignes_groupe in tableaux:
            feuille_nom = _nom_feuille_excel(titre, noms_utilises)
            tableau_export = tableau.rename(columns = {c: "Variable" for c in tableau.columns if str(c).strip() == ""})
            tableau_export.to_excel(fichier, sheet_name = feuille_nom, index = False, startrow = 2)
            feuille = fichier.sheets[feuille_nom]
            feuille.cell(row = 1, column = 1, value = titre).font = Font(bold = True, size = 13)
            for cellule in feuille[3]:
                cellule.font = Font(bold = True)
                cellule.fill = PatternFill("solid", fgColor = "DDDDDD")
                cellule.alignment = Alignment(horizontal = "center")
            for numero_ligne in lignes_groupe:
                for cellule in feuille[4 + numero_ligne]:
                    cellule.font = Font(bold = True)
                    cellule.fill = PatternFill("solid", fgColor = "E2EBF4")
            for numero_colonne, nom_colonne in enumerate(tableau_export.columns, start = 1):
                largeur = len(str(nom_colonne))
                for valeur in tableau_export.iloc[:, numero_colonne - 1].head(300):
                    largeur = _builtins_langue.max(largeur, len(_formater_valeur_stat(valeur)))
                feuille.column_dimensions[get_column_letter(numero_colonne)].width = _builtins_langue.min(largeur + 3, 60)
            feuille.freeze_panes = "B4"


def _exporter_tableaux_stat_word (chemin, tableaux):
    """ Tous les tableaux dans un seul document Word (paysage) : un titre puis
    le tableau (grille, en-tete en gras, lignes de variable en gras). """
    from docx.shared import Pt, Cm
    from docx.enum.section import WD_ORIENT
    document = Document()
    section = document.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    for marge in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, marge, Cm(1.5))
    for numero_tableau, (titre, tableau, lignes_groupe) in enumerate(tableaux):
        if numero_tableau != 0:
            document.add_paragraph()
        document.add_heading(titre, level = 2)
        entetes = [str(c) if str(c).strip() != "" else "Variable" for c in tableau.columns]
        table = document.add_table(rows = 1, cols = len(entetes))
        table.style = "Table Grid"
        for cellule, entete in zip(table.rows[0].cells, entetes):
            cellule.text = entete
            cellule.paragraphs[0].runs[0].font.bold = True
        for numero_ligne, ligne in enumerate(tableau.itertuples(index = False, name = None)):
            cellules = table.add_row().cells
            for cellule, valeur in zip(cellules, ligne):
                cellule.text = _formater_valeur_stat(valeur)
                if numero_ligne in lignes_groupe and cellule.paragraphs[0].runs:
                    cellule.paragraphs[0].runs[0].font.bold = True
        for ligne_table in table.rows:
            for cellule in ligne_table.cells:
                for paragraphe in cellule.paragraphs:
                    for morceau in paragraphe.runs:
                        morceau.font.size = Pt(9)
    document.save(chemin)


def _exporter_tableaux_stat_csv (chemin, tableaux):
    """ Tous les tableaux dans un seul CSV (separateur ";"), chacun precede de
    son titre et separe du suivant par une ligne vide. """
    with open(chemin, "w", encoding = "utf-8-sig", newline = "") as fichier:
        for numero_tableau, (titre, tableau, lignes_groupe) in enumerate(tableaux):
            if numero_tableau != 0:
                fichier.write("\n")
            fichier.write(str(titre) + "\n")
            tableau.rename(columns = {c: "Variable" for c in tableau.columns if str(c).strip() == ""}).to_csv(fichier, sep = ";", index = False, decimal = ".")


def my_exporting_data_stat (event, my_data = None):
    """ "Export stat" : exporte TOUS les tableaux affiches par "View stat"
    (stat_tableaux_export) en Excel (une feuille par tableau), Word ou CSV. """

    if len(stat_tableaux_export) == 0:
        messagebox.showwarning("Export stat", "Please display the statistics first (button 'View stat').")
        return
    chemin = filedialog.asksaveasfilename(filetypes = [("Excel file", "*.xlsx"), ("Word file", "*.docx"), ("csv file", "*.csv")], defaultextension = ".xlsx")
    if not chemin:
        return
    extension = os.path.splitext(chemin)[1].lower()
    try:
        if extension == ".docx":
            _exporter_tableaux_stat_word(chemin, stat_tableaux_export)
        elif extension == ".csv":
            _exporter_tableaux_stat_csv(chemin, stat_tableaux_export)
        else:
            if extension != ".xlsx":
                chemin += ".xlsx"
            _exporter_tableaux_stat_excel(chemin, stat_tableaux_export)
    except Exception as erreur:
        messagebox.showerror("Export stat", "The tables could not be exported:\n" + str(erreur))
        return
    messagebox.showinfo("Export stat", str(len(stat_tableaux_export)) + " table(s) exported to:\n" + chemin)


def my_save_sat_view (my_data):

   my_file = filedialog.asksaveasfilename(filetypes = [('Word file', '.docx'), ('Excel file', '*.xlsx')], defaultextension = ".docx")
   extension = os.path.splitext(my_file)
   if extension[1] == ".docx":
       document = Document()
       df_columns = my_data.columns 
       table = document.add_table(rows = 1, cols = len(df_columns))
       table_header = table.rows[0].cells
       for idx, this_col in enumerate(df_columns): table_header[idx].text = this_col
       for this_row in my_data.values.tolist():
           row_cells = table.add_row().cells
           for idx, this_cell in enumerate(this_row):
               row_cells[idx].text = str(this_cell)
       document.save(my_file)
   else:
       my_data.to_excel(my_file)


def recup (parent):

    global result_model_recod
    result_model_recod = [recoding_param.get(i) for i in recoding_param.curselection()]
    parent.destroy()


def r2_squared (y_true, y_pred):

    r2 = 1 - sum((y_true - y_pred)**2)/ sum((y_true - mean(y_true))**2)
    return r2


def _colonnes_qualitatives (data_user, colonnes):
    """ Colonnes non numeriques (texte, categories...) parmi "colonnes" : elles
    doivent etre recodees en indicatrices (0/1) avant d'entrainer un modele,
    sinon scikit-learn echoue ("could not convert string to float"). """
    return [c for c in colonnes if not pd.api.types.is_numeric_dtype(data_user[c])]


def _colonnes_a_recoder (data_user, features_parameter, features_recod):
    """ Colonnes a recoder en indicatrices selon le codage choisi. Les colonnes
    qualitatives sont TOUJOURS recodees, meme avec le codage "None" ou si
    elles n'ont pas ete cochees dans la liste du codage "Disjunctive". """
    if selectet_codage == "Disjunctive":
        x_recod = list(features_recod) if features_recod is not None else []
    elif selectet_codage == "None":
        x_recod = []
    else:
        x_recod = list(features_parameter)
    return x_recod + [c for c in _colonnes_qualitatives(data_user, features_parameter) if c not in x_recod]



# ==========================================================================
# ==  7. Execution arretable (bouton d'arret, progression, blocage)       ==
# ==========================================================================
## ------ Arret de l'execution (bouton button_arret_model, actif uniquement pendant « Run ») ------

class ExecutionArretee (BaseException):
    """ Levee quand l'utilisateur clique sur le bouton d'arret pendant « Run ».
    Derive de BaseException (et non d'Exception) pour ne PAS etre interceptee
    par les « except Exception » du code d'entrainement : elle remonte jusqu'a
    lancer_run, qui arrete proprement l'execution. """


_execution_run = {"en_cours": False, "arret_demande": False}

## Fils de calcul en cours (y compris ceux abandonnes apres un arret) et etat du ramasse-miettes.
## Le ramasse-miettes cyclique de Python peut se declencher dans N'IMPORTE QUEL fil : s'il se
## declenche dans un fil de calcul, il detruit des objets Tk (PhotoImage, StringVar...) hors du
## fil principal -> « RuntimeError: main thread is not in main loop ». Il est donc suspendu tant
## qu'un fil de calcul tourne, puis relance DANS le fil principal.
_fils_calcul = set()
_gc_etat = {"suspendu": False, "etait_actif": True}


## Protection definitive : si un objet Tk (image, variable, police) est malgre tout libere dans
## un fil de calcul (ex. bibliotheque appelant gc.collect() elle-meme, ce que gc.disable()
## n'empeche pas), sa destruction est REPORTEE et faite plus tard dans le fil principal.
import threading as _threading_tk
_objets_tk_a_detruire = []


def _proteger_destruction_tk (classe):
    origine = classe.__dict__.get("__del__")
    if origine is None or getattr(origine, "_protege_fil", False):
        return

    def __del__ (self):
        if _threading_tk.current_thread() is _threading_tk.main_thread():
            origine(self)
        else:
            _objets_tk_a_detruire.append((origine, self))     # garde l'objet en vie jusqu'a la destruction differee

    __del__._protege_fil = True
    classe.__del__ = __del__


for _classe_tk in (tkinter.Image, tkinter.Variable, _tkfont_langue.Font, getattr(ImageTk, "PhotoImage", None), getattr(ImageTk, "BitmapImage", None)):
    if _classe_tk is not None:
        _proteger_destruction_tk(_classe_tk)


def _detruire_objets_tk_differes ():
    """ Detruit, dans le fil principal, les objets Tk liberes par un fil de calcul. """
    if _threading_tk.current_thread() is not _threading_tk.main_thread():
        return
    while len(_objets_tk_a_detruire) != 0:
        origine, objet = _objets_tk_a_detruire.pop()
        try:
            origine(objet)
        except Exception:
            pass
        del objet


def _suspendre_ramasse_miettes ():
    import gc
    if not _gc_etat["suspendu"]:
        _gc_etat["etait_actif"] = gc.isenabled()
        _gc_etat["suspendu"] = True
        gc.collect()        # detruit ICI (fil principal) les objets Tk deja inutilises, avant de lancer le calcul
    gc.disable()


def _reprendre_ramasse_miettes ():
    """ Relance le ramasse-miettes (dans le fil principal) quand plus aucun fil de
    calcul ne tourne ; sinon, reessaie un peu plus tard. """
    import gc
    _detruire_objets_tk_differes()
    for fil in [f for f in _fils_calcul if not f.is_alive()]:
        _fils_calcul.discard(fil)
        _fils_abandonnes.discard(fil)
    if len(_fils_calcul) != 0:
        try:
            SmartDigicat.after(500, _reprendre_ramasse_miettes)
        except Exception:
            pass
        return
    if _gc_etat["suspendu"]:
        _gc_etat["suspendu"] = False
        if _gc_etat["etait_actif"]:
            gc.enable()
        gc.collect()        # nettoyage fait ici, dans le fil principal
    _detruire_objets_tk_differes()


_fils_abandonnes = set()     # fils de calcul abandonnes apres un arret : ils doivent s'arreter au prochain point de controle


def _verifier_arret_demande ():
    """ Point de controle : leve ExecutionArretee si l'arret a ete demande. Appele dans
    un fil de calcul abandonne (arret deja traite par l'interface), il l'arrete aussi. """
    import threading
    if threading.current_thread() in _fils_abandonnes:
        raise ExecutionArretee()
    if _execution_run["en_cours"] and _execution_run["arret_demande"]:
        raise ExecutionArretee()


def _calcul_interruptible (fonction):
    """ Execute un calcul long (entrainement, recherche d'hyperparametres,
    selection de features) dans un fil secondaire pendant que l'interface
    reste reactive, afin qu'un clic sur le bouton d'arret soit pris en compte
    immediatement. Le calcul ne doit PAS toucher aux widgets Tk.
    Hors « Run », le calcul est simplement execute directement.
    Remarque : un calcul scikit-learn ne peut pas etre interrompu de l'exterieur ;
    apres un arret, son resultat est ignore (le fil se termine seul en arriere-plan). """

    if not _execution_run["en_cours"]:
        return fonction()
    _verifier_arret_demande()
    import threading
    resultat = {}

    def travail ():
        try:
            resultat["valeur"] = fonction()
        except BaseException as erreur:
            resultat["erreur"] = erreur

    fil = threading.Thread(target = travail, daemon = True)
    _suspendre_ramasse_miettes()
    _fils_calcul.add(fil)
    fil.start()
    try:
        while fil.is_alive():
            try:
                SmartDigicat.update()          # traite les clics (dont celui du bouton d'arret)
            except tk.TclError:                # fenetre fermee pendant le calcul
                _execution_run["arret_demande"] = True
                _fils_abandonnes.add(fil)
                raise ExecutionArretee()
            _detruire_objets_tk_differes()
            _maintenir_indicateur_execution()
            if _execution_run["arret_demande"]:
                _fils_abandonnes.add(fil)      # le calcul s'arretera a son prochain point de controle
                raise ExecutionArretee()
            fil.join(0.03)
    finally:
        _reprendre_ramasse_miettes()           # immediat si le calcul est fini, differe s'il a ete abandonne
    if "erreur" in resultat:
        raise resultat["erreur"]
    return resultat.get("valeur")


_icones_bouton_arret = {"normale": None, "grisee": None}     # renseignees apres la creation du bouton


def _etat_bouton_arret (actif):
    """ Active / desactive le bouton d'arret. tkmacosx ne change pas l'apparence d'un
    bouton-image desactive : on affiche donc aussi une icone grisee quand il est desactive. """
    try:
        button_arret_model.configure(state = NORMAL if actif else DISABLED)
    except Exception:
        pass
    icone = _icones_bouton_arret["normale"] if actif else (_icones_bouton_arret["grisee"] or _icones_bouton_arret["normale"])
    if icone is not None:
        try:
            button_arret_model.configure(image = icone)
        except Exception:
            pass


def _desactiver_bouton_arret_fin_execution ():
    """ Fin de l'execution : desactive le bouton d'arret (etat + icone grisee + fond
    de la barre), puis le refait apres le traitement des evenements en attente, au
    cas ou tkmacosx redessinerait le bouton avec son aspect actif (survol, clic). """

    def desactiver ():
        # une NOUVELLE execution a demarre entre-temps (et aucun arret n'est demande) : on ne touche a rien
        if _execution_run["en_cours"] and not _execution_run["arret_demande"]:
            return
        try:
            button_arret_model.configure(background = "gray8", activebackground = "gray8", bordercolor = "gray8", focuscolor = "gray8", highlightbackground = "gray8", cursor = "")
        except Exception:
            pass
        _etat_bouton_arret(False)

    desactiver()
    for delai_ms in (50, 300):
        try:
            SmartDigicat.after(delai_ms, desactiver)
        except Exception:
            pass


TEXTE_BARRE_BASSE = "Industrial chair by Centrale Lille"
COULEUR_BOUTON_ARRET_ACTIF = "cornflowerblue"     # fond du bouton d'arret pendant l'execution (meme bleu que le curseur de zoom)
_indicateur_execution = {"cadre": None, "barre": None, "zone": None, "cle_position": "central",
                         "positions": {"central": (0.5, 0.5), "optimisation": (0.5, 0.5)}}
## zone : widget dans lequel l'encadre est affiche (None = panneau central ; zone du graphique pour
## l'optimisation) ; positions : centre de l'encadre (deplacable), memorise separement pour chaque zone


def _zone_indicateur ():
    zone = _indicateur_execution["zone"]
    try:
        if zone is not None and zone.winfo_exists():
            return zone
    except Exception:
        pass
    return Center_frame


class _BarreProgression (tk.Canvas):
    """ Barre de progression « indeterminee » dessinee par le programme (bloc bleu qui
    va et vient). Remplace ttk.Progressbar : sur macOS, la barre native est animee par
    le systeme et peut rester dessinee a l'ecran apres sa destruction (l'encadre
    semblait alors ne pas disparaitre a la fin de l'optimisation). """

    def __init__ (self, parent, longueur = 320, hauteur = 10):
        tk.Canvas.__init__(self, parent, width = longueur, height = hauteur, background = "gray88", highlightthickness = 0)
        self._longueur = longueur
        self._largeur_bloc = int(longueur * 0.25)
        self._position = 0
        self._sens = 1
        self._delai = 15
        self._tache = None
        self._bloc = self.create_rectangle(0, 0, self._largeur_bloc, hauteur, fill = COULEUR_BOUTON_ARRET_ACTIF, width = 0)

    def start (self, delai_ms = 15):
        self._delai = delai_ms
        if self._tache is None:
            self._avancer()

    def stop (self):
        if self._tache is not None:
            try:
                self.after_cancel(self._tache)
            except Exception:
                pass
            self._tache = None

    def _avancer (self):
        self._tache = None
        try:
            if not self.winfo_exists():
                return
            self._position += 6 * self._sens
            if self._position <= 0 or self._position + self._largeur_bloc >= self._longueur:
                self._sens = -self._sens
                self._position = _builtins_langue.max(0, _builtins_langue.min(self._position, self._longueur - self._largeur_bloc))
            self.coords(self._bloc, self._position, 0, self._position + self._largeur_bloc, int(self.cget("height")))
            self._tache = self.after(self._delai, self._avancer)
        except Exception:
            self._tache = None

    def destroy (self):
        self.stop()
        tk.Canvas.destroy(self)


## ------ Temps restant estime, affiche dans l'encadre d'execution ------
## Le travail est decoupe en « unites » : un modele a entrainer (Run) ou une iteration (optimisation). Duree
## prevue d'une unite : sa duree lors du precedent entrainement si elle est connue, sinon la moyenne des unites
## deja terminees pendant cette execution. Le temps restant diminue chaque seconde ; s'il atteint zero alors que
## le calcul continue, l'encadre affiche « presque termine ».
_progression = {"actif": False, "total": 0, "fait": 0, "debut_unite": 0.0, "durees_faites": [], "durees_prevues": [], "duree_par_defaut": None,
                "apriori": []}
## optimisation : duree moyenne d'une iteration et duree du calcul du graphique lors de la derniere optimisation
_duree_iteration_optimisation = {"moyenne": None, "graphique": None}


def _progression_demarrer (total, durees_prevues = None, duree_par_defaut = None, apriori = None):
    """ durees_prevues : duree attendue de chaque unite (None si inconnue). duree_par_defaut : duree d'une
    unite inconnue tant qu'aucune unite n'est terminee (par defaut : moyenne des durees connues).
    apriori : pour chaque unite, True si sa duree prevue est une ESTIMATION (premier lancement) : elle est
    corrigee au fil de l'execution par le rapport « duree reelle / duree estimee » des unites deja terminees. """
    durees = list(durees_prevues) if durees_prevues is not None else []
    durees = (durees + [None] * int(total))[:int(total)]
    drapeaux = (list(apriori) if apriori is not None else []) + [False] * int(total)
    _progression["apriori"] = drapeaux[:int(total)]
    connues = [d for d in durees if d is not None and d > 0]
    if duree_par_defaut is None and len(connues) != 0:
        duree_par_defaut = _builtins_langue.sum(connues) / len(connues)
    _progression.update(actif = int(total) > 0, total = int(total), fait = 0, debut_unite = time.time(),
                        durees_faites = [], durees_prevues = durees, duree_par_defaut = duree_par_defaut)


def _progression_unite_terminee ():
    """ Une unite (modele / iteration) vient de se terminer ; renvoie sa duree en secondes. """
    if not _progression["actif"]:
        return None
    maintenant = time.time()
    duree = maintenant - _progression["debut_unite"]
    _progression["durees_faites"].append(duree)
    _progression["fait"] = _progression["fait"] + 1
    _progression["debut_unite"] = maintenant
    return duree



# l'optimisation (module core) signale la fin de chaque iteration a l'interface
_optimisation_bayesienne.progress_callback = _progression_unite_terminee


def _progression_arreter ():
    _progression["actif"] = False


def _progression_restant ():
    """ Secondes restantes estimees ; None tant qu'aucune estimation n'est possible. """
    p = _progression
    if not p["actif"] or p["fait"] >= p["total"]:
        return 0.0
    faites = p["durees_faites"]
    ## unite de duree inconnue : la plus grande valeur entre la moyenne et la derniere unite terminee (en
    ## optimisation, chaque iteration est un peu plus longue que la precedente : on evite de sous-estimer)
    moyenne = _builtins_langue.max(_builtins_langue.sum(faites) / len(faites), faites[-1]) if len(faites) != 0 else p["duree_par_defaut"]
    ## correction des estimations a priori : rapport duree reelle / duree estimee des unites estimees deja terminees
    reel, estime = 0.0, 0.0
    for i in range(_builtins_langue.min(p["fait"], len(faites))):
        prevue = p["durees_prevues"][i]
        if p["apriori"][i] and prevue is not None and prevue > 0:
            reel, estime = reel + faites[i], estime + prevue
    facteur = (reel / estime) if estime > 0 else 1.0
    restantes = []
    for i in range(p["fait"], p["total"]):
        d = p["durees_prevues"][i]
        if d is not None and d > 0:
            restantes.append(d * facteur if p["apriori"][i] else d)
        else:
            restantes.append(moyenne)
    if _builtins_langue.any(d is None for d in restantes):     # any natif (numpy.any ne gere pas un generateur)
        return None
    ecoule_unite = time.time() - p["debut_unite"]
    return _builtins_langue.max(restantes[0] - ecoule_unite, 0.0) + _builtins_langue.sum(restantes[1:])


def _progression_texte ():
    if not _progression["actif"]:
        return ""
    restant = _progression_restant()
    if restant is None:
        valeur = traduire("estimation en cours...")
    elif restant < 1:
        valeur = traduire("presque termine...")
    else:
        secondes = int(restant + 0.999)
        if secondes >= 3600:
            valeur = "%d h %02d min" % (secondes // 3600, (secondes % 3600) // 60)
        elif secondes >= 60:
            valeur = "%d min %02d s" % (secondes // 60, secondes % 60)
        else:
            valeur = "%d s" % secondes
    return traduire("Temps restant estime :") + " " + valeur


def _actualiser_temps_restant ():
    label = _indicateur_execution.get("label_temps")
    try:
        if label is not None and label.winfo_exists():
            texte = _progression_texte()
            if label.cget("text") != texte:
                label.configure(text = texte)
    except Exception:
        pass


def _creer_indicateur_execution ():
    """ Encadre centre dans le panneau central : message + barre de progression animee. """
    _supprimer_tous_les_indicateurs()        # jamais deux encadres a la fois
    zone = _zone_indicateur()
    cadre = tk.Frame(zone, background = "white", highlightthickness = 2, highlightbackground = COULEUR_BOUTON_ARRET_ACTIF, highlightcolor = COULEUR_BOUTON_ARRET_ACTIF, padx = 30, pady = 20)     # contour bleu, comme le bouton d'arret actif
    Label(cadre, text = "Execution in progress...", font = ('arial', 20, 'bold'), fg = "gray20", bg = "white").pack(pady = (0, 6))
    Label(cadre, text = "Click the stop button (left bar) to interrupt it.", font = ('arial', 14), fg = "gray40", bg = "white").pack()
    label_temps = Label(cadre, text = _progression_texte(), font = ('arial', 15, 'bold'), fg = COULEUR_BOUTON_ARRET_ACTIF, bg = "white")
    label_temps.pack(pady = (10, 0))
    _indicateur_execution["label_temps"] = label_temps
    barre = None            # l'encadre n'affiche plus de barre de progression (message seul)
    cadre._indicateur_execution_smartdigicat = True     # marque : retrouve et supprime a coup sur en fin d'execution
    print("[SmartDigicat] encadre de progression cree")
    _indicateur_execution["cadre"] = cadre
    _indicateur_execution["barre"] = barre
    _rendre_indicateur_deplacable(cadre)
    relx, rely = _indicateur_execution["positions"][_indicateur_execution["cle_position"]]
    cadre.place(relx = relx, rely = rely, anchor = "center")


def _supprimer_tous_les_indicateurs ():
    """ Supprime TOUS les encadres de progression presents dans l'application (a
    n'importe quel niveau, y compris un encadre qui ne serait plus reference), chaque
    etape etant protegee separement : une erreur sur l'une n'empeche pas la suppression. """
    try:
        enfants = _descendants_widget(SmartDigicat)
    except Exception:
        enfants = []
    zones = []
    for widget in enfants:
        if getattr(widget, "_indicateur_execution_smartdigicat", False):
            try:
                zones.append((widget.master, widget.winfo_x(), widget.winfo_y(), widget.winfo_width(), widget.winfo_height()))
            except Exception:
                pass
            for enfant in _descendants_widget(widget):
                if isinstance(enfant, (ttk.Progressbar, _BarreProgression)):
                    try:
                        enfant.stop()
                    except Exception:
                        pass
            try:
                widget.destroy()
                print("[SmartDigicat] encadre de progression supprime")
            except Exception as erreur:
                print("[SmartDigicat] suppression de l'encadre impossible :", erreur)
    _indicateur_execution["cadre"] = None
    _indicateur_execution["barre"] = None
    for zone in zones:
        _forcer_redessin_zone(*zone)


def _forcer_redessin_zone (parent, x, y, largeur, hauteur):
    """ Force le redessin de la zone qu'occupait l'encadre : un voile temporaire est
    pose puis retire, ce qui oblige les widgets situes dessous (graphique de
    l'optimisation...) a se redessiner. Evite qu'une image de l'encadre reste a l'ecran. """
    if largeur <= 1 or hauteur <= 1:
        return
    try:
        if parent is None or not parent.winfo_exists():
            return
        voile = tk.Frame(parent, background = parent.cget("background"), highlightthickness = 0)
        voile.place(x = x, y = y, width = largeur, height = hauteur)
        voile.update_idletasks()
        SmartDigicat.after(20, voile.destroy)
    except Exception:
        pass


## ------ Touche Entree dans toutes les boites de dialogue ------

def _bouton_par_defaut (fenetre):
    """ Bouton a declencher par la touche Entree dans une boite de dialogue : celui marque
    comme bouton par defaut (default = "active"), sinon le premier bouton de la boite. """
    boutons = []
    for widget in _descendants_widget(fenetre):
        try:
            if widget.winfo_class() in ("Button", "TButton") and str(widget.cget("state")) != "disabled":
                boutons.append(widget)
        except Exception:
            pass
    for bouton in boutons:
        try:
            if str(bouton.cget("default")) == "active":
                return bouton
        except Exception:
            pass
    return boutons[0] if len(boutons) != 0 else None


def _entree_boite_dialogue (event):
    """ Touche Entree dans une boite de dialogue (fenetre Toplevel) qui ne la gere pas
    elle-meme : declenche son bouton par defaut. Sans effet dans la fenetre principale. """
    try:
        fenetre = event.widget.winfo_toplevel()
    except Exception:
        return None
    if fenetre is SmartDigicat or not isinstance(fenetre, tk.Toplevel):
        return None
    try:
        if fenetre.bind("<Return>"):                  # la boite a deja sa propre action pour Entree
            return None
        if event.widget.winfo_class() == "Text":      # Entree = retour a la ligne dans un texte multiligne
            return None
    except Exception:
        return None
    bouton = _bouton_par_defaut(fenetre)
    if bouton is not None:
        try:
            bouton.invoke()
        except Exception:
            pass
        return "break"
    return None


def _entree_pave_numerique (event):
    """ Touche Entree du pave numerique : meme effet que la touche Entree principale. """
    try:
        if event.widget.winfo_exists():
            event.widget.event_generate("<Return>")
    except Exception:
        pass
    return "break"


def installer_touche_entree_boites_dialogue ():
    """ Toutes les boites de dialogue reagissent a la touche Entree :
    - boites systeme (messages, questions Oui/Non, fichiers, couleurs) : Entree = bouton par defaut (natif) ;
    - boites de saisie (askstring...) et boites du logiciel : Entree = OK / bouton par defaut ;
    - toute autre boite (Toplevel) : Entree = son bouton par defaut ou, a defaut, son premier bouton ;
    - la touche Entree du pave numerique fait partout la meme chose que la touche Entree. """
    SmartDigicat.bind_all("<Return>", _entree_boite_dialogue, add = "+")
    SmartDigicat.bind_all("<KP_Enter>", _entree_pave_numerique, add = "+")
    ## boites de saisie tkinter (askstring, askinteger...) : clavier dirige vers la boite
    ## (la touche Entree du pave numerique y est convertie en Entree par _entree_pave_numerique)
    try:
        classe_dialogue = simpledialog.Dialog
        if not getattr(classe_dialogue.buttonbox, "_entree_installee", False):
            buttonbox_origine = classe_dialogue.buttonbox

            def buttonbox (self):
                buttonbox_origine(self)

                def donner_focus ():
                    ## CORRECTIF : focus_force() seul donnait le clavier a la FENETRE et retirait le curseur du
                    ## champ de saisie (nom du projet, nom du modele...). On active la fenetre PUIS on remet le
                    ## curseur dans le champ, texte propose selectionne : taper le nom le remplace directement.
                    try:
                        self.lift()
                        self.focus_force()
                    except Exception:
                        pass
                    champ = getattr(self, "entry", None) or getattr(self, "initial_focus", None)
                    if champ is None:
                        return
                    try:
                        champ.focus_force()
                        if hasattr(champ, "select_range"):
                            champ.select_range(0, "end")
                            champ.icursor("end")
                    except Exception:
                        pass

                self.after(50, donner_focus)

            buttonbox._entree_installee = True
            classe_dialogue.buttonbox = buttonbox
    except Exception:
        pass


def _surveiller_indicateur_execution ():
    """ Surveillance permanente (toutes les 500 ms) : en dehors d'une execution, aucun
    encadre de progression ne doit exister. S'il en reste un (quelle qu'en soit la
    cause), il est supprime. """
    try:
        if not _execution_run["en_cours"]:
            restants = [w for w in _descendants_widget(SmartDigicat) if getattr(w, "_indicateur_execution_smartdigicat", False)]
            if len(restants) != 0:
                print("[SmartDigicat] surveillance : encadre de progression encore present apres l'execution -> suppression")
                _supprimer_tous_les_indicateurs()
    except Exception:
        pass
    try:
        SmartDigicat.after(500, _surveiller_indicateur_execution)
    except Exception:
        pass


def _rendre_indicateur_deplacable (cadre):
    """ L'encadre peut etre deplace a la souris (clic-glisse sur n'importe quelle partie),
    en restant dans le panneau central. Sa position est conservee s'il est recree. """
    depart = {"dx": 0, "dy": 0}

    def debut (event):
        # ecart entre le pointeur et le centre de l'encadre
        depart["dx"] = event.x_root - (cadre.winfo_rootx() + cadre.winfo_width() / 2)
        depart["dy"] = event.y_root - (cadre.winfo_rooty() + cadre.winfo_height() / 2)
        return "break"

    def glisser (event):
        zone = cadre.master          # panneau central, ou zone du graphique pour l'optimisation
        largeur, hauteur = zone.winfo_width(), zone.winfo_height()
        if largeur <= 1 or hauteur <= 1:
            return "break"
        demi_l, demi_h = cadre.winfo_width() / 2, cadre.winfo_height() / 2
        centre_x = event.x_root - depart["dx"] - zone.winfo_rootx()
        centre_y = event.y_root - depart["dy"] - zone.winfo_rooty()
        # l'encadre reste entierement dans sa zone
        centre_x = _builtins_langue.min(_builtins_langue.max(centre_x, demi_l), _builtins_langue.max(demi_l, largeur - demi_l))
        centre_y = _builtins_langue.min(_builtins_langue.max(centre_y, demi_h), _builtins_langue.max(demi_h, hauteur - demi_h))
        position = (centre_x / largeur, centre_y / hauteur)
        _indicateur_execution["positions"][_indicateur_execution["cle_position"]] = position
        cadre.place(relx = position[0], rely = position[1], anchor = "center")
        return "break"

    for widget in [cadre] + _descendants_widget(cadre):
        try:
            widget.configure(cursor = "fleur")          # curseur « deplacer »
        except Exception:
            pass
        widget.bind("<ButtonPress-1>", debut)
        widget.bind("<B1-Motion>", glisser)


## ------ Blocage du panneau central pendant « Run » ------
## Chaque widget du panneau central recoit en tete de ses « bindtags » une etiquette qui
## intercepte (return "break") clics, glisser et clavier : ses propres reactions (boutons,
## listes, champs, cases a cocher, tableaux...) ne sont plus declenchees. Le contenu reste
## visible et le defilement a la molette reste possible (sauf sur les listes deroulantes et
## les champs numeriques, ou la molette changerait la valeur).
_TAG_BLOCAGE_RUN = "SmartDigicatBlocageRun"
_TAG_BLOCAGE_MOLETTE_RUN = "SmartDigicatBlocageMoletteRun"
_CLASSES_MOLETTE_BLOQUEE = ("TCombobox", "Spinbox", "TSpinbox", "Scale", "TScale", "Listbox")
_EVENEMENTS_BLOQUES_RUN = ("<ButtonPress>", "<ButtonRelease>", "<Double-ButtonPress>", "<Triple-ButtonPress>",
                           "<B1-Motion>", "<B2-Motion>", "<B3-Motion>", "<KeyPress>", "<KeyRelease>")
_EVENEMENTS_MOLETTE_RUN = ("<MouseWheel>", "<Shift-MouseWheel>", "<Button-4>", "<Button-5>")
_blocage_run = {"classes_liees": False, "dernier_passage": 0.0}


def _descendants_widget (widget):
    try:
        enfants = widget.winfo_children()
    except Exception:
        return []
    resultat = []
    for enfant in enfants:
        resultat.append(enfant)
        resultat.extend(_descendants_widget(enfant))
    return resultat


def _bloquer_panneau_central (bloquer):
    """ bloquer = True : pendant « Run », plus aucun widget de l'application n'est
    accessible (boutons des barres, panneau central, curseur de zoom...), SAUF le
    bouton d'arret button_arret_model ; curseur « montre ». False : retour a la
    normale. Peut etre rappelee pendant l'execution pour bloquer aussi les widgets
    crees entre-temps. (Le nom est conserve : le blocage couvrait d'abord le seul
    panneau central.) """

    if not _blocage_run["classes_liees"]:
        for evenement in _EVENEMENTS_BLOQUES_RUN:
            SmartDigicat.bind_class(_TAG_BLOCAGE_RUN, evenement, lambda event: "break")
        for evenement in _EVENEMENTS_MOLETTE_RUN:
            SmartDigicat.bind_class(_TAG_BLOCAGE_MOLETTE_RUN, evenement, lambda event: "break")
        _blocage_run["classes_liees"] = True

    try:
        bouton_arret = button_arret_model
        exclus = set(str(w) for w in [bouton_arret] + _descendants_widget(bouton_arret))
    except Exception:
        exclus = set()
    # l'encadre de progression reste deplacable a la souris
    cadre_indicateur = _indicateur_execution.get("cadre")
    try:
        if cadre_indicateur is not None and cadre_indicateur.winfo_exists():
            exclus.update(str(w) for w in [cadre_indicateur] + _descendants_widget(cadre_indicateur))
    except Exception:
        pass

    for widget in [SmartDigicat] + _descendants_widget(SmartDigicat):
        if str(widget) in exclus:
            continue
        try:
            etiquettes = [t for t in widget.bindtags() if t not in (_TAG_BLOCAGE_RUN, _TAG_BLOCAGE_MOLETTE_RUN)]
            if bloquer:
                ajout = [_TAG_BLOCAGE_RUN]
                if widget.winfo_class() in _CLASSES_MOLETTE_BLOQUEE:
                    ajout.append(_TAG_BLOCAGE_MOLETTE_RUN)
                etiquettes = ajout + etiquettes
            widget.bindtags(tuple(etiquettes))
        except Exception:
            pass
    try:
        SmartDigicat.configure(cursor = "watch" if bloquer else "")
        Center_frame.configure(cursor = "")
        button_arret_model.configure(cursor = "hand2" if bloquer else "")     # seul bouton cliquable
    except Exception:
        pass
    if bloquer:
        try:
            SmartDigicat.focus_set()        # retire le focus clavier d'un champ de saisie
        except Exception:
            pass


def _terminer_indicateur_execution ():
    """ Les calculs sont termines : l'encadre de progression disparait IMMEDIATEMENT, sans attendre la
    fin des operations qui suivent (reconstruction du panneau « Model », historique Undo/Redo, messages
    de fin) ; il n'est plus recree jusqu'a la fin de cette execution. """
    _execution_run["indicateur_termine"] = True
    try:
        _supprimer_tous_les_indicateurs()
        SmartDigicat.update_idletasks()        # effacement visible tout de suite
    except Exception:
        pass


def _maintenir_indicateur_execution ():
    """ Pendant l'execution : recree l'indicateur s'il a ete detruit (panneau central
    vide par un affichage) et le garde au premier plan, au centre du panneau.
    Apres un clic sur le bouton d'arret, ou une fois les calculs termines, il n'est plus recree. """
    if not _execution_run["en_cours"] or _execution_run["arret_demande"] or _execution_run.get("indicateur_termine"):
        return
    try:
        cadre = _indicateur_execution["cadre"]
        if cadre is None or not cadre.winfo_exists():
            _creer_indicateur_execution()
            if _indicateur_execution["barre"] is not None:
                _indicateur_execution["barre"].start(15)
        _indicateur_execution["cadre"].lift()
    except Exception:
        pass
    _actualiser_temps_restant()             # decompte du temps restant (ce passage a lieu plusieurs fois par seconde)
    # widgets apparus pendant l'execution : bloques eux aussi (au plus 3 fois / s)
    if time.time() - _blocage_run["dernier_passage"] > 0.3:
        _blocage_run["dernier_passage"] = time.time()
        _bloquer_panneau_central(True)


def _afficher_execution_en_cours (en_cours):
    """ Indique visuellement qu'une execution est en cours : bouton d'arret sur fond
    bleu et message avec barre de progression animee au centre du panneau central. """
    try:
        couleur = COULEUR_BOUTON_ARRET_ACTIF if en_cours else "gray8"
        button_arret_model.configure(background = couleur, activebackground = couleur, bordercolor = couleur, focuscolor = couleur, highlightbackground = couleur)
    except Exception:
        pass
    try:
        if en_cours:
            _maintenir_indicateur_execution()
        else:
            _supprimer_tous_les_indicateurs()
            # nouvelle verification juste apres : aucun encadre ne doit survivre a la fin de l'execution
            SmartDigicat.after(100, lambda: None if _execution_run["en_cours"] else _supprimer_tous_les_indicateurs())
    except Exception:
        pass
    try:
        SmartDigicat.update_idletasks()        # affichage immediat
    except Exception:
        pass


def demander_arret_execution ():
    """ Commande du bouton d'arret : demande l'arret de l'execution en cours. """
    if _execution_run["en_cours"]:
        _execution_run["arret_demande"] = True
        _afficher_execution_en_cours(False)     # la barre de progression disparait immediatement
        # le bouton se desactive tout de suite (etat + icone grisee + fond sombre), puis a nouveau
        # apres 50 et 300 ms : tkmacosx redessine le bouton clique APRES l'appel de sa commande
        # et pouvait lui redonner son aspect actif
        _desactiver_bouton_arret_fin_execution()


def _executer_avec_bouton_arret (fonction, titre, preparation = None, zone_indicateur = None, cle_position = "central"):
    """ Execute "fonction" (entrainement des modeles ou optimisation) en rendant le
    bouton d'arret actif : tous les autres widgets sont inaccessibles, l'indicateur
    d'execution est affiche, puis l'interface est retablie et le bouton d'arret
    desactive. "preparation" (optionnelle) est appelee juste avant le blocage (ex. :
    afficher le panneau « Model »). "zone_indicateur" : widget ou afficher l'encadre de
    progression (par defaut le panneau central). Retourne True si l'execution a ete arretee. """

    if _execution_run["en_cours"]:
        messagebox.showinfo(titre, "An execution is already in progress.\n\nWait for it to finish, or click the stop button (left bar) to interrupt it.")
        return False
    _execution_run["en_cours"] = True
    _execution_run["arret_demande"] = False
    _execution_run["indicateur_termine"] = False
    _progression_arreter()                  # chaque execution declare elle-meme ses unites de travail
    _indicateur_execution["zone"] = zone_indicateur
    _indicateur_execution["cle_position"] = cle_position
    _etat_bouton_arret(True)
    if preparation is not None:
        try:
            preparation()
        except Exception:
            pass
    _bloquer_panneau_central(True)          # plus aucun bouton accessible, sauf le bouton d'arret
    _afficher_execution_en_cours(True)
    arrete = False
    try:
        fonction()
    except ExecutionArretee:
        arrete = True
    finally:
        _progression_arreter()
        _terminer_indicateur_execution()            # l'encadre disparait en premier
        _execution_run["en_cours"] = False
        _execution_run["arret_demande"] = False
        _afficher_execution_en_cours(False)         # (re-verification : aucun encadre ne doit subsister)
        _indicateur_execution["zone"] = None
        _bloquer_panneau_central(False)
        # desactivation EN DERNIER : les reconfigurations precedentes (fond, curseur) du bouton
        # tkmacosx ne doivent pas lui redonner un aspect actif
        _desactiver_bouton_arret_fin_execution()
    return arrete


def lancer_run (event):
    """ Clic sur « Run » : entrainement des modeles (run_linear_model), arretable
    par le bouton d'arret. Le panneau central affiche le panneau du bouton « Model ». """

    arrete = _executer_avec_bouton_arret(
        lambda: run_linear_model(event, Center_frame, Target_column, Calculation_method, Test_size, Target_list, header_target, loss_size, rate_size, codage_method),
        "Run", preparation = lambda: _afficher_panneau_model_apres_run(Center_frame))
    if arrete:
        # le modele interrompu garde sa configuration et ses resultats precedents
        try:
            if nom_modele_actif is not None and nom_modele_actif in modeles_projet:
                restore_modele(nom_modele_actif)
            _rafraichir_etat_modeles()
        except Exception:
            pass
        try:
            messagebox.showinfo("Run", "Execution stopped.\n\nThe model being trained keeps its previous configuration and results; the models already trained before the stop are kept.")
        except tk.TclError:
            pass


def _reinitialiser_zone_graphique_optimisation (graph):
    """ Clic sur « Run optim » : la zone du graphique (panneau gauche) est videe (ancien
    graphique et ancien tableau retires) avant le debut de l'optimisation. Si
    l'optimisation n'aboutit pas, le panneau reconstruit a la fin reaffiche le resultat
    precedent. """
    global opt_canvas_widget, opt_table_widget
    if graph is None:
        return
    try:
        if not graph.winfo_exists():
            return
        for enfant in list(graph.winfo_children()):
            try:
                enfant.destroy()
            except Exception:
                pass
        graph.delete("all")        # (graph est un Canvas : on efface aussi ce qui y serait dessine)
    except Exception:
        pass
    opt_canvas_widget = None
    opt_table_widget = None


def lancer_optimisation (*arguments):
    """ Clic sur « Run optim » : optimisation (check_optim_parameters), arretable par
    le bouton d'arret. En cas d'arret, le resultat precedent est conserve. """

    graph = arguments[0] if len(arguments) > 0 else None
    # encadre de progression dans la zone du graphique (arguments[0] = graph), reinitialisee au prealable
    arrete = _executer_avec_bouton_arret(lambda: check_optim_parameters(*arguments), "Optimization",
                                         preparation = lambda: _reinitialiser_zone_graphique_optimisation(graph),
                                         zone_indicateur = graph, cle_position = "optimisation")
    _rafraichir_panneau_optimisation(arguments)
    if arrete:
        try:
            messagebox.showinfo("Optimization", "Optimization stopped.\n\nThe previous optimization result (if any) is kept.")
        except tk.TclError:
            pass


_conteneur_optimisation = {"widget": None}


def _rafraichir_panneau_optimisation (arguments):
    """ Fin de chaque optimisation (terminee, en erreur ou arretee) : comme a la fin de
    « Run » (ou le panneau « Model » est recree), un panneau NEUF qui couvre TOUT le
    panneau central est cree, et le panneau d'optimisation est reconstruit dedans.
    Cela oblige l'ecran a redessiner entierement le panneau central (l'encadre de
    progression ne peut plus y rester affiche). Le dernier resultat et les choix
    d'affichage sont repris automatiquement (opt_plot_data, opt_affichage). """
    graph = arguments[0] if len(arguments) > 0 else None
    x_axis_opt = arguments[9] if len(arguments) > 9 else None
    anciens = []
    for widget in (graph, getattr(x_axis_opt, "master", None), _conteneur_optimisation["widget"]):
        if widget is not None and widget not in anciens:
            anciens.append(widget)
    conteneur = Canvas(Center_frame, background = "white", highlightthickness = 0)
    conteneur.place(relx = 0, rely = 0, relwidth = 1, relheight = 1)
    _conteneur_optimisation["widget"] = conteneur
    try:
        my_optimization(None, conteneur)
    except Exception as erreur:
        messagebox.showwarning("Optimization", "The optimization panel could not be refreshed:\n" + str(erreur))
        return
    # anciens panneaux detruits APRES le traitement du clic en cours (le bouton « Run optim »
    # qui a lance l'optimisation se trouve dans l'un d'eux)
    _detruire_widgets_differe([w for w in anciens if w is not conteneur])


# ==========================================================================
# ==  8. Selection de variables et validation croisee                     ==
# ==========================================================================
def _ajuster_features (my_reg_select, X_train, Y_train, cv, n_jobs = 1):
    """ Appelle mlxtend.SequentialFeatureSelector (« my_select »). En cas d'echec (frequemment une
    incompatibilite de version entre mlxtend et scikit-learn), l'erreur est reformulee pour identifier
    clairement cette etape au lieu d'un message technique obscur. """
    try:
        reg_result = my_select(my_reg_select, k_features = 'best', forward = True, floating = False, cv = cv, n_jobs = n_jobs)
        if n_jobs == 1:
            return _calcul_interruptible(lambda: reg_result.fit(X_train, Y_train))

        def ajuster_en_parallele ():
            with _contexte_parallele_cv():
                return reg_result.fit(X_train, Y_train)

        return _calcul_interruptible(ajuster_en_parallele)
    except ExecutionArretee:
        raise
    except Exception as erreur:
        try:
            import mlxtend
            version_mlxtend = mlxtend.__version__
        except Exception:
            version_mlxtend = "inconnue"
        try:
            import sklearn
            version_sklearn = sklearn.__version__
        except Exception:
            version_sklearn = "inconnue"
        raise RuntimeError("La selection des variables (mlxtend " + str(version_mlxtend) + ", scikit-learn " +
                           str(version_sklearn) + ") a echoue : " + type(erreur).__name__ + ": " + str(erreur) +
                           "\n\nCause frequente : versions de mlxtend / scikit-learn incompatibles entre elles.") from erreur


def my_select_all_features(data_user, Y_user, features_parameter, features_recod, my_reg_select, value_test):

    global my_modality_find
    X = data_user[features_parameter].copy()
    x_recod = _colonnes_a_recoder(X, features_parameter, features_recod)
    if len(x_recod) == 0:
        XX_train, XX_test, Y_train_select, Y_test_select = train_test_split(X, Y_user, test_size = value_test, random_state = 0)
        reg_result = _ajuster_features(my_reg_select, XX_train, Y_train_select, cv = 5, n_jobs = N_JOBS_CV)
        my_columns = np.array(reg_result.k_feature_names_)
        ## CORRECTIF : sans recodage, my_modality_find n'etait jamais mis a jour (None ou valeur
        ## d'un autre modele) -> « Best features used » plantait / affichait les mauvaises variables.
        my_modality_find = my_columns
        d_XX_train = XX_train[my_columns]
        d_XX_test = XX_test[my_columns]
    else:
        for x_col in x_recod:
            dx = np.array(X[x_col])
            if type(dx[0]) != str:
                xm = []
                for i in range(len(dx)):
                    xm.append(str(dx[i]))
                X[x_col] = xm
        modality_ref_name = []
        for x_col in x_recod:
            modality_ref = X[x_col].value_counts().index[0]
            modality_ref_name.append(x_col + "_" +  modality_ref)
        X_train_select, X_test_select, Y_train_select, Y_test_select = train_test_split(X, Y_user, test_size = value_test, random_state = 0)
        XX = pd.get_dummies(data = X_train_select, columns = x_recod, dtype = float)
        data_X = XX.drop(modality_ref_name, axis = 1, errors = "ignore")
        reg_result = _ajuster_features(my_reg_select, data_X, Y_train_select, cv = 5, n_jobs = N_JOBS_CV)
        my_columns = np.array(reg_result.k_feature_names_)
        modality_find = []
        for colx in my_columns:
            colx = re.sub(r'[()]', '', colx)
            for colxx in features_parameter:
                colxxx = re.sub(r'[()]', '', colxx)
                my_motif = re.search(colxxx, colx)
                if my_motif:
                    modality_find.append(colxx)
        my_modality_find = unique(modality_find)
        my_int_sec = [int_sec for int_sec in x_recod if int_sec in my_modality_find]
        X_best = X[my_modality_find]
        modality_ref_name_all = []
        for x_col in my_int_sec:
                modality_ref = X_best[x_col].value_counts().index[0]
                modality_ref_name_all.append(x_col + "_" +  modality_ref)
        # colonnes retenues par la selection ; seules celles a recoder (my_int_sec) deviennent des indicatrices
        # (avant : toutes les colonnes retenues etaient recodees, y compris les numeriques, qui donnaient
        # une indicatrice par valeur ; et les colonnes non retenues etaient conservees)
        XX_best_train = pd.get_dummies(data = X_train_select[list(my_modality_find)], columns = my_int_sec, dtype = float)
        XX_best_test = pd.get_dummies(data = X_test_select[list(my_modality_find)], columns = my_int_sec, dtype = float)
        d_XX_train = XX_best_train.drop(modality_ref_name_all, axis = 1, errors = "ignore")
        d_XX_test = XX_best_test.drop(modality_ref_name_all, axis = 1, errors = "ignore")
        d_XX_test = d_XX_test.reindex(columns = d_XX_train.columns, fill_value = 0)     # memes colonnes indicatrices qu'a l'entrainement
    return d_XX_train, d_XX_test, Y_train_select, Y_test_select



## Hyperparametres optimises par la cross-validation, pour chaque algorithme
## (les valeurs saisies dans le panneau servent de BORNES SUPERIEURES aux grilles)
_CLES_HYPERPARAMETRES_CV = {
    "Lasso": ["alpha"],
    "Kernel ridge (KR)": ["alpha"],
    "Support Vector Regression (SVR)": ["kernel", "C", "gamma"],
    "Random Forest (RF)": ["n_estimators"],
    "Generalized Boosted Regression Modeling (GBM)": ["n_estimators", "max_depth", "learning_rate", "loss"],
    "Decision Tree (DT)": ["max_depth"],
}


## Cross-validation : part des donnees mise de cote pour TESTER le modele (prediction).
## Les 80 % restants servent a optimiser les hyperparametres (K plis) et a entrainer le modele final.
PART_TEST_CV = 0.20


def _nombre_plis_cv (valeur, n_lignes, defaut = 5):
    """ Nombre K de plis de la cross-validation : valeur saisie ("Split size"),
    ramenee entre 2 et le nombre de lignes d'entrainement. """
    try:
        k = int(float(valeur))
    except (TypeError, ValueError):
        k = defaut
    if k < 2:
        k = defaut
    return int(_builtins_langue.max(2, _builtins_langue.min(k, n_lignes)))


## ------ Optimisation de la cross-validation (vitesse) ------
## Au-dela de MAX_COMBINAISONS_CV combinaisons, la grille n'est plus parcourue en entier : on en teste
## MAX_COMBINAISONS_CV tirees au hasard (RandomizedSearchCV, reproductible : random_state = 0).
MAX_COMBINAISONS_CV = 30
## Nombre de calculs menes en parallele pendant la recherche (-1 : tous les coeurs du processeur).
## Execution par FILS (threads) : pas de nouveau processus (compatible avec l'application empaquetee).
N_JOBS_CV = -1
## Garde-fou SVR : sans normalisation des donnees, les noyaux « poly » (et « linear » avec un grand C)
## peuvent tourner des heures sur UNE combinaison. Au-dela de ce nombre d'iterations, l'ajustement s'arrete ;
## la combinaison, alors mal notee, n'est pas retenue.
SVR_MAX_ITER_CV = 1000000


def _valeurs_geometriques_entieres (maximum, nombre = 6, minimum = 1):
    """ Entiers repartis geometriquement de minimum a maximum (inclus), sans doublon. """
    maximum = int(_builtins_langue.max(1, int(round(float(maximum)))))
    minimum = int(minimum) if int(minimum) < maximum else 1      # borne saisie tres petite : grille depuis 1
    minimum = int(_builtins_langue.max(1, minimum))
    if maximum - minimum + 1 <= nombre:
        return list(range(minimum, maximum + 1))
    valeurs = sorted(set(int(round(v)) for v in np.geomspace(minimum, maximum, nombre)))
    if valeurs[-1] != maximum:
        valeurs.append(maximum)
    return valeurs


def _valeurs_geometriques_reelles (maximum, minimum, nombre = 6):
    """ Reels repartis geometriquement de minimum a maximum (inclus). """
    maximum = float(maximum)
    if maximum <= 0:
        return [maximum]
    minimum = float(_builtins_langue.min(minimum, maximum))
    if minimum <= 0 or minimum == maximum:
        return [maximum]
    return [float(v) for v in np.geomspace(minimum, maximum, nombre)]


def _grille_cv (nom_algo, reg_base):
    """ (estimateur de base, grille d'hyperparametres) pour la recherche par cross-validation, construits
    a partir de l'estimateur configure : ses valeurs sont les BORNES SUPERIEURES (incluses) des grilles.
    Grille None : algorithme sans hyperparametre a optimiser (Linear, GP).
    OPTIMISATION : grilles a nombre de points FIXE, reparties geometriquement (avant : pas fixe depuis 1,
    ex. GBM par defaut = 39 200 combinaisons, soit environ 10 h de calcul ; SVR = 360 combinaisons). """

    parametres = reg_base.get_params()

    if nom_algo == "Lasso":
        return Lasso(), {"alpha": _valeurs_geometriques_reelles(parametres["alpha"], 0.001, 10)}
    if nom_algo == "Kernel ridge (KR)":
        return KernelRidge(), {"alpha": _valeurs_geometriques_reelles(parametres["alpha"], 0.001, 10)}
    if nom_algo == "Support Vector Regression (SVR)":
        C_max, gamma_max = float(parametres["C"]), float(parametres["gamma"]) if not isinstance(parametres["gamma"], str) else 0.1
        return svm.SVR(max_iter = SVR_MAX_ITER_CV), {"kernel": ["linear", "rbf", "poly", "sigmoid"],
                                                     "C": _valeurs_geometriques_reelles(C_max, _builtins_langue.min(1.0, C_max / 10), 6),
                                                     "gamma": _valeurs_geometriques_reelles(gamma_max, _builtins_langue.min(0.001, gamma_max / 10), 5)}
    if nom_algo == "Random Forest (RF)":
        return RandomForestRegressor(random_state = parametres.get("random_state")), {"n_estimators": _valeurs_geometriques_entieres(parametres["n_estimators"], 6, 10)}
    if nom_algo == "Generalized Boosted Regression Modeling (GBM)":
        return GradientBoostingRegressor(random_state = 0), {"n_estimators": _valeurs_geometriques_entieres(parametres["n_estimators"], 6, 10),
                                                             "max_depth": _valeurs_geometriques_entieres(parametres["max_depth"] or 3, 5),
                                                             "learning_rate": _valeurs_geometriques_reelles(parametres["learning_rate"], 0.01, 6),
                                                             "loss": ["squared_error", "absolute_error", "huber", "quantile"]}
    if nom_algo == "Decision Tree (DT)":
        return DecisionTreeRegressor(random_state = parametres.get("random_state")), {"max_depth": list(range(1, int(parametres["max_depth"] or 20) + 1))}
    return reg_base, None


def _parametres_principaux (nom_algo, reg):
    """ Hyperparametres principaux d'un estimateur (affiches dans « Detail of model »). """
    parametres = reg.get_params()
    return {cle: parametres[cle] for cle in _CLES_HYPERPARAMETRES_CV.get(nom_algo, []) if cle in parametres}


def _ajuster_hyperparametres_cv (nom_algo, reg_base, data_X, Y_train, k):
    """ Recherche des meilleurs hyperparametres par cross-validation a K plis.
    Retourne (estimateur non entraine avec ces hyperparametres, dict des hyperparametres). """
    from sklearn.base import clone as _cloner_estimateur
    from sklearn.model_selection import ParameterGrid
    base, grille = _grille_cv(nom_algo, reg_base)
    if grille is None:
        return _cloner_estimateur(reg_base), _parametres_principaux(nom_algo, reg_base)

    def rechercher ():
        scoreur = _scoreur_cv_interruptible()       # cree dans le fil de calcul : il le reconnait s'il est abandonne
        if len(ParameterGrid(grille)) <= MAX_COMBINAISONS_CV:
            recherche = GridSearchCV(base, grille, cv = k, scoring = scoreur, n_jobs = N_JOBS_CV)
        else:
            recherche = RandomizedSearchCV(base, grille, n_iter = MAX_COMBINAISONS_CV, cv = k, scoring = scoreur,
                                           n_jobs = N_JOBS_CV, random_state = 0)
        with _contexte_parallele_cv():
            recherche.fit(data_X, Y_train)
        return recherche

    recherche = _calcul_interruptible(rechercher)
    return _cloner_estimateur(base).set_params(**recherche.best_params_), dict(recherche.best_params_)


def _contexte_parallele_cv ():
    """ Calculs paralleles par FILS (joblib « threading ») : aucun nouveau processus n'est lance (un
    processus relancerait le logiciel lui-meme, et poserait probleme dans l'application empaquetee). """
    try:
        from joblib import parallel_config
        return parallel_config(backend = "threading")
    except ImportError:                     # joblib < 1.3
        from joblib import parallel_backend
        return parallel_backend("threading")


def _scoreur_cv_interruptible ():
    """ Score = -MSE (comme scoring = "neg_mean_squared_error"). Il verifie aussi, apres chaque ajustement,
    si l'execution a ete arretee (bouton d'arret) : la recherche abandonnee s'arrete alors au lieu de
    continuer a occuper le processeur en arriere-plan jusqu'a la fin de la grille. """
    import threading
    fil_de_calcul = threading.current_thread()

    def score (estimateur, X, y):
        if fil_de_calcul in _fils_abandonnes or (_execution_run["en_cours"] and _execution_run["arret_demande"]):
            raise ExecutionArretee()
        return -mean_squared_error(y, estimateur.predict(X))

    return score


def my_select_all_features_CV (data_user, Y_user, features_parameter, features_recod, k = 5, reg_base = None, nom_algo = None, reg_deja_ajuste = False):
    """ Mode "Cross-validation" : recodage des colonnes (qualitatives toujours
    recodees), separation 80 % / 20 %, recherche des hyperparametres par
    cross-validation a K plis sur la partie entrainement, puis selection des
    meilleures features.
    - reg_base / nom_algo : estimateur configure et nom de l'algorithme (par
      defaut : widgets du panneau et Name_model du modele actif) ;
    - reg_deja_ajuste = True : reg_base contient deja les meilleurs
      hyperparametres (modele rejoue par Run) -> pas de nouvelle recherche.
    Retourne (X_train, X_test, Y_train, Y_test, estimateur, hyperparametres). """

    global my_modality_find
    from sklearn.base import clone as _cloner_estimateur
    if nom_algo is None:
        nom_algo = Name_model
    if reg_base is None:
        reg_base = _instancier_reg_depuis_widgets(nom_algo)
    if reg_base is None:
        raise ValueError("unknown model: " + str(nom_algo))

    X = data_user[list(features_parameter)].copy()
    x_recod = _colonnes_a_recoder(X, features_parameter, features_recod)
    for x_col in x_recod:
        dx = np.array(X[x_col])
        if type(dx[0]) != str:
            X[x_col] = [str(valeur) for valeur in dx]
    # 1) 80 % entrainement (optimisation des hyperparametres + modele final) / 20 % test (prediction)
    X_train_select, X_test_select, Y_train_select, Y_test_select = train_test_split(X, Y_user, test_size = PART_TEST_CV, random_state = 0)
    # modalite de reference de chaque colonne recodee : la plus frequente dans la partie ENTRAINEMENT
    # (les 20 % de test ne doivent servir a rien d'autre qu'a evaluer le modele)
    modalite_reference = {x_col: str(X_train_select[x_col].value_counts().index[0]) for x_col in x_recod}
    modality_ref_name = [x_col + "_" + modalite_reference[x_col] for x_col in x_recod]
    data_X = pd.get_dummies(data = X_train_select, columns = x_recod, dtype = float).drop(modality_ref_name, axis = 1, errors = "ignore")
    k = _nombre_plis_cv(k, len(data_X))

    # 2) hyperparametres : cross-validation a K plis sur les 80 % d'entrainement uniquement
    if reg_deja_ajuste:
        my_reg, meilleurs_parametres = _cloner_estimateur(reg_base), _parametres_principaux(nom_algo, reg_base)
    else:
        my_reg, meilleurs_parametres = _ajuster_hyperparametres_cv(nom_algo, reg_base, data_X, Y_train_select, k)

    # 3) selection des meilleures features, elle aussi sur les 80 % d'entrainement uniquement
    reg_result = _ajuster_features(my_reg, data_X, Y_train_select, cv = k, n_jobs = N_JOBS_CV)
    my_columns = np.array(reg_result.k_feature_names_)
    modality_find = []
    for colx in my_columns:
        colx = re.sub(r'[()]', '', colx)
        for colxx in features_parameter:
            colxxx = re.sub(r'[()]', '', colxx)
            if re.search(colxxx, colx):
                modality_find.append(colxx)
    my_modality_find = unique(modality_find)
    my_int_sec = [int_sec for int_sec in x_recod if int_sec in my_modality_find]
    modality_ref_name_all = [x_col + "_" + modalite_reference[x_col] for x_col in my_int_sec]
    # 4) jeux finaux : le modele est entraine sur les 80 % puis evalue (prediction) sur les 20 % de test
    d_XX_train = pd.get_dummies(data = X_train_select[list(my_modality_find)], columns = my_int_sec, dtype = float).drop(modality_ref_name_all, axis = 1, errors = "ignore")
    d_XX_test = pd.get_dummies(data = X_test_select[list(my_modality_find)], columns = my_int_sec, dtype = float).drop(modality_ref_name_all, axis = 1, errors = "ignore")
    d_XX_test = d_XX_test.reindex(columns = d_XX_train.columns, fill_value = 0)
    return d_XX_train, d_XX_test, Y_train_select, Y_test_select, my_reg, meilleurs_parametres


def my_select_all_features_pred(data_user, features_parameter, features_recod):

    X = data_user[features_parameter].copy()
    for x_col in features_recod:
        dx = np.array(X[x_col])
        if type(dx[0]) != str:
            xm = []
            for i in range(len(dx)):
                xm.append(str(dx[i]))
            X[x_col] = xm
    XX = pd.get_dummies(data = X, columns = features_recod, dtype = float)
    return XX


def recup_my_feature (parent):

    global result_model_1
    result_model_1 = [Feature_column.get(i) for i in Feature_column.curselection()]
    parent.destroy()


# ==========================================================================
# ==  9. Modeles de regression                                            ==
# ==========================================================================
##--------- Implementing regression models
def regression_model (XX = None, Y = None, k = 5, methode = None, value_test = 0.20, loss_var = None):

    global reg, methode_pred, kf_pred, columns_select, quali_columns ## ------ Global variables
    methode_pred = methode
    quali_columns = []
    for var in result_model:
        ## CORRECTIF : on lit data_analy (jeu reellement utilise pour l'entrainement) et .iloc[0]
        ## (l'etiquette d'index 0 n'existe plus apres un filtrage/suppression de lignes).
        feature_var = data_analy[var]
        if type(feature_var.iloc[0]) == str:
            quali_columns.append(var)
    ## ------ The models executed
    m1 = "Linear"  
    minit = "Choose model"
    m1b = "Lasso"
    m2 = "Kernel ridge (KR)"
    m3 = "Support Vector Regression (SVR)"
    m4 = "Random Forest (RF)"
    m5 = "Generalized Boosted Regression Modeling (GBM)"
    m6 = "Decision Tree (DT)"
    if methode == "TrainTestSplit data":
            if Name_model != minit:
                  if Name_model == m1:
                      reg = LinearRegression()
                  elif Name_model == m1b:
                      reg = Lasso(alpha = float(Alpha_lasso_size.get()))
                  elif Name_model == m2:
                      reg = KernelRidge(alpha = float(Alpha_size.get()))
                  elif  Name_model == m3:
                        reg = svm.SVR(kernel = "rbf", C = float(C_size.get()), gamma = float(gamma_size.get()))
                  elif Name_model == m4:
                      reg = RandomForestRegressor(n_estimators =  int(trees_size.get()), random_state =  int(Random_size.get()))
                  elif Name_model == m5:
                      reg = ensemble.GradientBoostingRegressor(n_estimators = int(estimators_size.get()), max_depth = int(depth_size.get()), loss = loss_var, learning_rate = float(rate_size.get()))
                  elif Name_model == m6:
                      reg =  DecisionTreeRegressor(splitter = 'best', max_depth = int(depth_size_r.get()), random_state = int(Random_size_r.get()))
                  else:
                      reg = GaussianProcessRegressor(kernel = kernel, n_targets = 1, random_state = int(Random_size_gpr.get()))
            ##---- Selection and split data
            X_reg_data = my_select_all_features(XX, Y, result_model, result_model, reg, value_test)
            ##---- Train data
            X_train = X_reg_data[0]
            Y_train = X_reg_data[2]
            ##---- Test data
            X_test = X_reg_data[1]
            Y_test = X_reg_data[3]
            columns_select = np.array(X_train.columns)
            _calcul_interruptible(lambda: reg.fit(X_train, Y_train))
            pred_test = reg.predict(X_test)
            pred_train = reg.predict(X_train)
            dp = pd.DataFrame(pred_test)
            xpred = np.array(dp[0].tolist())
            R2 = r2_squared(Y_test, xpred) 
            mse = mean_squared_error(Y_test, pred_test)
            if Name_model == "Linear":
                my_coef = reg.coef_ 
            elif Name_model == "Lasso":
                my_coef = reg.coef_ 
            elif Name_model == "Random Forest (RF)":
                my_coef = reg.feature_importances_
            elif Name_model == "Generalized Boosted Regression Modeling (GBM)":
                my_coef = reg.feature_importances_
            elif Name_model == "Decision Tree (DT)":
                my_coef = reg.feature_importances_
            else:
                my_coef = None
            return R2, pred_test, Y_test, mse, pred_train, Y_train, my_coef
    else:
        X_reg_data = my_select_all_features_CV(XX, Y, result_model, result_model, k = value_test)     # value_test = "Split size" (K) en cross-validation
        ##---- Train data
        X_train = X_reg_data[0]
        Y_train = X_reg_data[2]
        columns_select = np.array(X_train.columns)    
        ##---- Test data
        X_test = X_reg_data[1]
        Y_test = X_reg_data[3]
        reg = X_reg_data[4]
        _calcul_interruptible(lambda: reg.fit(X_train, Y_train))
        pred_test = reg.predict(X_test)
        pred_train = reg.predict(X_train)
        dp = pd.DataFrame(pred_test)
        xpred = np.array(dp[0].tolist())
        R2 = r2_squared(Y_test, xpred) 
        mse = mean_squared_error(Y_test, pred_test)
        return R2, pred_test, Y_test, mse, pred_train, Y_train, X_reg_data[5]


class Titles_model(Enum):
    LSLR = "Linear"
    LSR = "Lasso"
    KRR = "Kernel ridge (KR)"
    SVR = "Support Vector Regression (SVR)"
    RFR = "Random Forest (RF)"
    GBM = "Generalized Boosted Regression Modeling (GBM)"
    DTR = "Decision Tree (DT)"

data_model_name = pd.DataFrame({"Linear": ["LSLR"], "Lasso": ["LSR"], "Kernel ridge (KR)": ["KRR"], "Support Vector Regression (SVR)": ["SVR"], "Random Forest (RF)": ["RFR"],"Generalized Boosted Regression Modeling (GBM)": ["GBM"],  "Decision Tree (DT)": ["DTR"], "Gaussian process (GP)": ["GPR"]})      



# ==========================================================================
# ==  10. Plusieurs modeles par projet, historique Undo / Redo            ==
# ==========================================================================
## ------ Gestion de plusieurs modeles/projets distincts ------

_MODEL_STATE_KEYS = ["Name_model", "selectet_target", "selectet_codage", "selectet_methode",
                       "selectet_methode_imp", "result_model", "result_model_recod", "reg",
                       "columns_select", "quali_columns", "my_data_coef", "my_modality_find",
                       "data_model", "my_y_min", "my_y_max", "size_data_value"]


## CORRECTIF (plusieurs modeles) : le panneau de configuration (widgets Target_column, Spinbox, ...) est
## partage par tous les modeles via des variables globales. Rien n'indiquait a quel modele il appartenait :
## _capturer_configuration_widgets_actifs() pouvait donc enregistrer les valeurs d'un panneau resté
## ouvert pour le modele A dans la configuration du modele B (ou ecraser un modele deja entraine).
## On memorise desormais le proprietaire du panneau ouvert.
_panneau_config = {"nom": None, "conteneur": None, "param": None, "reconfiguration": False}


def _widget_existe (widget):
    try:
        return widget is not None and widget.winfo_exists() == 1
    except Exception:
        return False


def _fermer_panneau_config ():
    for cle in ("param", "conteneur"):
        widget = _panneau_config.get(cle)
        if _widget_existe(widget):
            try:
                widget.destroy()
            except Exception:
                pass
        _panneau_config[cle] = None
    _panneau_config["nom"] = None
    _panneau_config["reconfiguration"] = False


## ------ Undo / Redo ------
## Historique d'etats du PROJET (modeles et leur configuration, modeles supprimes, apparence des boites,
## donnees importees / filtrees, donnees et resultats de prediction, filtres). Un « point de sauvegarde »
## est enregistre apres chaque action (_historique_valider) ; Undo / Redo se deplacent dans cette liste.
## Les objets `reg` sont traites comme IMMUABLES une fois entraines (tout re-entrainement cree un nouvel
## objet), ils sont donc partages entre les etats sans copie profonde.
HISTORIQUE_TAILLE_MAX = 20
_historique = {"etats": [], "index": -1, "restauration": False}


def _copie_etat_modele_historique (etat):
    copie = {}
    for cle, valeur in etat.items():
        if isinstance(valeur, list):
            copie[cle] = list(valeur)
        elif isinstance(valeur, dict):
            copie[cle] = dict(valeur)
        else:
            copie[cle] = valeur
    return copie


def _copie_tableau_historique (tableau):
    return None if tableau is None else tableau.copy()


def _historique_etat_courant (etiquette):
    filtres = [globals().get("filtre_edit_parametre_" + str(k)) for k in (1, 2, 3, 4)]
    _historique["compteur"] = _historique.get("compteur", 0) + 1
    return {
        "id": _historique["compteur"],                 # identifiant unique du point (pour savoir s'il a ete enregistre)
        "etiquette": etiquette,
        "modeles_projet": {nom: _copie_etat_modele_historique(etat) for nom, etat in modeles_projet.items()},
        "modeles_supprimes": {nom: _copie_etat_modele_historique(etat) for nom, etat in modeles_supprimes.items()},
        "positions": dict(positions_boites_modeles),
        "couleurs": dict(couleurs_boites_modeles),
        "positions_sup": dict(positions_boites_supprimees),
        "couleurs_sup": dict(couleurs_boites_supprimees),
        "nom_modele_actif": nom_modele_actif,
        "file_path": file_path,
        "file_sheets": list(file_sheets) if file_sheets else None,
        "data": data,
        "data_analy": _copie_tableau_historique(data_analy),
        "file_path_pred": file_path_pred,
        "data_pred": _copie_tableau_historique(data_pred),
        "data_pred_new": _copie_tableau_historique(data_pred_new),
        "prediction_infos": dict(prediction_infos) if isinstance(prediction_infos, dict) else {},
        "filtres": [dict(f) if isinstance(f, dict) else None for f in filtres],
    }


def _historique_valider (etiquette):
    """Enregistre l'etat courant comme nouveau point de l'historique (apres une action)."""
    if _historique["restauration"]:
        return
    try:
        etat = _historique_etat_courant(etiquette)
    except Exception:
        return
    etats = _historique["etats"]
    del etats[_historique["index"] + 1:]          ## une nouvelle action invalide les « Redo » en attente
    etats.append(etat)
    while len(etats) > HISTORIQUE_TAILLE_MAX:
        del etats[0]
    _historique["index"] = len(etats) - 1


def _historique_reinitialiser (etiquette = "Initial state"):
    _historique["etats"] = []
    _historique["index"] = -1
    _historique_valider(etiquette)
    _projet_marquer_enregistre()        # projet neuf ou tout juste ouvert : rien a enregistrer pour l'instant


def _historique_peut_annuler ():
    return _historique["index"] > 0


def _historique_peut_refaire ():
    return _historique["index"] < len(_historique["etats"]) - 1


def _historique_appliquer (etat):
    global nom_modele_actif, file_path, file_sheets, data, data_analy, file_path_pred, data_pred, data_pred_new, prediction_infos
    _historique["restauration"] = True
    try:
        _fermer_panneau_config()
        modeles_projet.clear()
        modeles_projet.update({nom: _copie_etat_modele_historique(e) for nom, e in etat["modeles_projet"].items()})
        modeles_supprimes.clear()
        modeles_supprimes.update({nom: _copie_etat_modele_historique(e) for nom, e in etat["modeles_supprimes"].items()})
        for dictionnaire, cle in ((positions_boites_modeles, "positions"), (couleurs_boites_modeles, "couleurs"),
                                  (positions_boites_supprimees, "positions_sup"), (couleurs_boites_supprimees, "couleurs_sup")):
            dictionnaire.clear()
            dictionnaire.update(etat[cle])
        for k, valeur_filtre in zip((1, 2, 3, 4), etat["filtres"]):
            filtre = globals().get("filtre_edit_parametre_" + str(k))
            if isinstance(filtre, dict) and isinstance(valeur_filtre, dict):
                filtre.clear()
                filtre.update(valeur_filtre)
        file_path = etat["file_path"]
        file_sheets = etat.get("file_sheets")
        data = etat["data"]
        data_analy = _copie_tableau_historique(etat["data_analy"])
        file_path_pred = etat["file_path_pred"]
        data_pred = _copie_tableau_historique(etat["data_pred"])
        data_pred_new = _copie_tableau_historique(etat["data_pred_new"])
        prediction_infos = dict(etat["prediction_infos"])
        nom_actif = etat["nom_modele_actif"]
        if nom_actif is not None and nom_actif in modeles_projet:
            restore_modele(nom_actif)
        else:
            nom_modele_actif = None
            for cle in _MODEL_STATE_KEYS:
                globals()[cle] = None
    finally:
        _historique["restauration"] = False


def _historique_rafraichir_interface ():
    """Redessine le panneau central (boites de modeles) et l'etat des boutons apres Undo / Redo."""
    try:
        for enfant in list(Center_frame.winfo_children()):
            enfant.destroy()
    except Exception:
        pass
    try:
        choice_model(None, Center_frame)
    except Exception:
        pass


def _historique_message (texte):
    # pendant une execution (« Run »), la barre du bas affiche l'indicateur d'execution : on ne l'ecrase pas
    if _execution_run["en_cours"]:
        return
    try:
        bottom_frame.configure(text = texte)
        SmartDigicat.after(3500, lambda: None if _execution_run["en_cours"] else bottom_frame.configure(text = TEXTE_BARRE_BASSE))
    except Exception:
        pass


def annuler_derniere_action ():
    if not analyse_autorisee() or not _historique_peut_annuler():
        return
    etiquette = _historique["etats"][_historique["index"]]["etiquette"]
    _historique["index"] -= 1
    _historique_appliquer(_historique["etats"][_historique["index"]])
    _historique_rafraichir_interface()
    _historique_message("Undo : " + str(etiquette))


def refaire_derniere_action ():
    if not analyse_autorisee() or not _historique_peut_refaire():
        return
    _historique["index"] += 1
    etat = _historique["etats"][_historique["index"]]
    _historique_appliquer(etat)
    _historique_rafraichir_interface()
    _historique_message("Redo : " + str(etat["etiquette"]))


def snapshot_modele_actif():
  
    if nom_modele_actif is None:
        return

    etat = {}
    for cle in _MODEL_STATE_KEYS:
        valeur = globals().get(cle)
        if cle in ("result_model", "result_model_recod") and valeur is not None:
            valeur = list(valeur)
        etat[cle] = valeur
    ## signature du dernier entrainement conservee (sauvegarde automatique, enregistrement...) : si l'etat a
    ## change entre-temps, elle ne correspondra plus et le modele sera re-entraine au prochain « Run »
    for cle_conservee in ("signature_entrainement", "duree_entrainement"):
        valeur_conservee = modeles_projet.get(nom_modele_actif, {}).get(cle_conservee)
        if valeur_conservee is not None:
            etat[cle_conservee] = valeur_conservee

    modeles_projet[nom_modele_actif] = etat
    try:
        _rafraichir_etat_modeles()      ## CORRECTIF : libelle « (trained) »
    except Exception:
        pass


def restore_modele(nom):

    global nom_modele_actif, Name_model, selectet_target, selectet_codage, selectet_methode
    global selectet_methode_imp, result_model, result_model_recod, reg, columns_select
    global quali_columns, my_data_coef, my_modality_find, data_model, my_y_min, my_y_max, size_data_value

    etat = modeles_projet.get(nom)
    if etat is None:
        return

    nom_modele_actif = nom
    Name_model = etat["Name_model"]
    selectet_target = etat["selectet_target"]
    selectet_codage = etat["selectet_codage"]
    selectet_methode = etat["selectet_methode"]
    selectet_methode_imp = etat["selectet_methode_imp"]
    result_model = etat["result_model"]
    result_model_recod = etat["result_model_recod"]
    reg = etat["reg"]
    columns_select = etat["columns_select"]
    quali_columns = etat["quali_columns"]
    my_data_coef = etat["my_data_coef"]
    my_modality_find = etat["my_modality_find"]
    data_model = etat["data_model"]
    my_y_min = etat["my_y_min"]
    my_y_max = etat["my_y_max"]
    size_data_value = etat["size_data_value"]


def creer_nouveau_modele(nom):

    _capturer_configuration_widgets_actifs()
    _fermer_panneau_config()      ## CORRECTIF : ne pas laisser le panneau du modele precedent ouvert
    modeles_projet[nom] = {cle: None for cle in _MODEL_STATE_KEYS}
    restore_modele(nom)


def _statut_affichage_modele (nom):
  
    etat = modeles_projet.get(nom, {})
    if etat.get("data_model") is not None:
        return traduire("  (trained)")
    elif etat.get("reg") is not None:
        return traduire("  (configured)")
    else:
        return traduire("  (new)")


def _assurer_data_analy (silencieux = False):
    """CORRECTIF : data_analy (jeu de donnees utilise pour l'entrainement) n'etait
    renseigne que par « Edit data > View & Edit » ou par l'ouverture d'un projet
    existant. Dans un NOUVEAU projet (New > Import data > Analysis > Run), il valait
    donc None et l'entrainement echouait sans message. Ici on le construit a la
    demande a partir des donnees importees (data)."""
    global data_analy

    if data_analy is not None:
        return True

    if data is None:
        if silencieux and not (isinstance(file_path, str) and os.path.isfile(file_path)):
            return False
        if not _assurer_donnees_source_chargees():
            return False

    if data is None:
        return False

    data_analy = data.copy()
    return True



# ==========================================================================
# ==  11. Entrainement (Run) des modeles du projet                        ==
# ==========================================================================
## ------ « Run » : pas de recalcul des modeles inchanges ------
## Un modele deja entraine n'est pas re-entraine par « Run » si RIEN n'a change depuis : meme algorithme et
## hyperparametres, meme cible, memes variables, meme recodage, meme methode / taille de test, et memes donnees.
## Ses resultats sont conserves tels quels. Mettre a True pour tout recalculer a chaque « Run ».
RELANCER_MODELES_INCHANGES = False


def _empreinte_donnees (colonnes):
    """ Empreinte (dimensions + contenu) des colonnes de data_analy utilisees par un modele. """
    try:
        sous_tableau = data_analy[list(colonnes)]
        return (tuple(sous_tableau.shape), tuple(str(c) for c in sous_tableau.columns),
                int(pd.util.hash_pandas_object(sous_tableau, index = True).sum()))
    except Exception:
        return None


def _signature_entrainement (etat):
    """ Tout ce dont depend le resultat de l'entrainement d'un modele. """
    reg_modele = etat.get("reg")
    try:
        parametres = repr(sorted((cle, repr(valeur)) for cle, valeur in reg_modele.get_params().items())) if reg_modele is not None else None
    except Exception:
        parametres = repr(reg_modele)
    features = [str(v) for v in (etat.get("result_model") if etat.get("result_model") is not None else [])]
    cible = etat.get("selectet_target")
    return repr((etat.get("Name_model"), cible, etat.get("selectet_codage"), etat.get("selectet_methode"), etat.get("size_data_value"),
                 tuple(features), tuple(str(v) for v in (etat.get("result_model_recod") if etat.get("result_model_recod") is not None else [])),
                 parametres, _empreinte_donnees(features + ([cible] if cible is not None else []))))


def _modele_a_jour (etat):
    """ Vrai si le modele est deja entraine avec exactement sa configuration et les donnees actuelles. """
    if RELANCER_MODELES_INCHANGES or etat.get("data_model") is None or etat.get("reg") is None:
        return False
    signature = etat.get("signature_entrainement")
    return signature is not None and signature == _signature_entrainement(etat)


def _estimer_duree_modele (etat):
    """ Duree estimee de l'entrainement d'un modele jamais entraine (premier lancement) : UN entrainement de
    l'estimateur est chronometre sur les donnees d'entrainement, puis multiplie par le nombre d'entrainements
    que realise « Run » :
      - selection des variables (mlxtend, ajout pas a pas) : K plis x p(p+1)/2 sous-ensembles de variables ;
      - en Cross-validation : recherche des hyperparametres (combinaisons testees x K plis) ;
      - entrainement final.
    Retourne None si l'estimation est impossible. """
    from sklearn.base import clone as _cloner_estimateur
    from sklearn.model_selection import ParameterGrid
    try:
        features = [f for f in (etat.get("result_model") if etat.get("result_model") is not None else [])]
        cible = etat.get("selectet_target")
        if cible is None or len(features) == 0 or data_analy is None:
            return None
        reg_modele = etat.get("reg")
        if reg_modele is None:
            reg_modele = _reconstruire_estimateur(None, etat)
        if reg_modele is None:
            return None
        X = data_analy[features].copy()
        Y = pd.to_numeric(data_analy[cible], errors = "coerce")
        codage = etat.get("selectet_codage")
        a_recoder = list(features) if codage == "All-disjunctive" else [c for c in features if not pd.api.types.is_numeric_dtype(X[c])]
        if codage == "Disjunctive" and etat.get("result_model_recod") is not None:
            a_recoder = sorted(set(a_recoder) | (set(etat.get("result_model_recod")) & set(features)))
        if len(a_recoder) != 0:
            X = pd.get_dummies(X, columns = a_recoder, drop_first = True, dtype = float)
        lignes_valides = Y.notna().to_numpy() & X.notna().all(axis = 1).to_numpy()
        X, Y = X[lignes_valides], Y[lignes_valides]
        n_lignes, n_variables = X.shape
        if n_lignes < 3 or n_variables == 0:
            return None
        methode = etat.get("selectet_methode")
        taille = etat.get("size_data_value")
        if methode == "Cross-validation":
            plis = _nombre_plis_cv(taille, n_lignes)
            part_entrainement = 1 - PART_TEST_CV
        else:
            plis = 5
            part_entrainement = 1 - (float(taille) if taille is not None and 0 < float(taille) < 1 else 0.2)
        n_entrainement = _builtins_langue.max(2, int(n_lignes * part_entrainement))
        mesure = {}

        def chronometrer ():
            debut = time.perf_counter()
            _cloner_estimateur(reg_modele).fit(X.iloc[:n_entrainement], Y.iloc[:n_entrainement])
            mesure["duree"] = time.perf_counter() - debut

        _calcul_interruptible(chronometrer)
        duree_un = _builtins_langue.max(mesure.get("duree", 0.0), 0.005)     # 5 ms : cout minimal d'un entrainement (Python, mlxtend)
        p = n_variables
        ## selection des variables : sous-ensembles de taille moyenne (p + 2) / 3, sur (K - 1) / K des lignes
        equivalents = plis * p * (p + 1) / 2.0 * ((p + 2) / (3.0 * p)) * ((plis - 1) / float(plis))
        if methode == "Cross-validation":
            base, grille = _grille_cv(etat.get("Name_model"), reg_modele)
            if grille is not None and not etat.get("signature_entrainement"):
                combinaisons = _builtins_langue.min(len(ParameterGrid(grille)), MAX_COMBINAISONS_CV)
                equivalents += combinaisons * plis * ((plis - 1) / float(plis)) * 0.5     # grilles : en moyenne moitie moins couteuses
        equivalents += 2                                                                   # entrainement final + prediction
        coeurs = os.cpu_count() or 1
        acceleration = _builtins_langue.max(1.0, _builtins_langue.min(coeurs, plis) * 0.5)  # calculs en parallele (fils)
        ## 1,7 : surcout mesure de la validation croisee (copies de l'estimateur, decoupages, scores...)
        return 1.7 * duree_un * equivalents / acceleration + 0.2
    except ExecutionArretee:
        raise
    except Exception:
        return None


def _estimer_durees_optimisation (n_points, n_simul_1, n_simul_2, n_iterations):
    """ Durees estimees de chaque iteration de l'optimisation (puis du calcul du graphique), avant toute mesure.
    Cout mesure d'une iteration : proportionnel a n_simul_1 x n_simul_2, et croissant avec le nombre de points
    deja connus (un point de plus a chaque iteration). Le rapport reel / estime corrige ensuite ces valeurs. """
    durees = []
    for i in range(int(n_iterations)):
        durees.append(1.06e-3 * float(n_simul_1) * float(n_simul_2) * (0.7 + 0.033 * (n_points + i)))
    durees.append(durees[-1] if len(durees) != 0 else 1.0)          # graphique : environ une iteration
    return durees


def _reg_panneau_sans_erreur (nom_algo):
    """ Estimateur du panneau ouvert, pour l'estimation du temps uniquement (None en cas de probleme). """
    try:
        return _instancier_reg_depuis_widgets(nom_algo)
    except Exception:
        return None


DUREE_PAR_DEFAUT_MODELE = 10.0     # secondes : derniere solution si aucune estimation n'est possible


def _planifier_progression_run (nom_actif_panneau = None, etat_panneau = None):
    """ Declare les modeles que « Run » va reellement entrainer (modele actif via son panneau, puis les modeles
    configures et modifies), avec la duree de leur precedent entrainement quand elle est connue. """
    a_entrainer = []
    if nom_actif_panneau is not None:
        a_entrainer.append(etat_panneau if etat_panneau is not None else modeles_projet.get(nom_actif_panneau, {}))
    for nom, etat in modeles_projet.items():
        if nom == nom_actif_panneau:
            continue
        features = etat.get("result_model")
        if etat.get("selectet_target") is None or features is None or len(features) == 0 or etat.get("selectet_methode") is None:
            continue
        if _modele_a_jour(etat):
            continue
        a_entrainer.append(etat)
    ## duree de chaque modele : celle de son precedent entrainement si elle est connue, sinon une ESTIMATION
    ## (premier lancement), pour que le temps restant s'affiche des le debut
    durees, apriori = [], []
    for etat in a_entrainer:
        duree = etat.get("duree_entrainement")
        if duree is not None and duree > 0:
            durees.append(duree)
            apriori.append(False)
        else:
            estimation = _estimer_duree_modele(etat)
            durees.append(estimation if estimation is not None else DUREE_PAR_DEFAUT_MODELE)
            apriori.append(True)
    _progression_demarrer(len(durees), durees, apriori = apriori)


def _executer_autres_modeles_du_projet (nom_a_exclure):

    global selectet_codage

    reussis = []
    ignores = []
    codage_sauvegarde = selectet_codage

    if not _assurer_data_analy(silencieux = True):
        ignores.append("(erreur : jeu de donnees indisponible ; importez les donnees puis reessayez)")
        return reussis, ignores

    # try/finally : le codage global est restaure meme si l'execution est arretee (bouton d'arret)
    try:
        for nom, etat in list(modeles_projet.items()):
            _verifier_arret_demande()      # arret demande : on ne commence pas le modele suivant
            if nom == nom_a_exclure:
                continue

            cible = etat.get("selectet_target")
            features = etat.get("result_model")
            methode = etat.get("selectet_methode")
            reg_modele = etat.get("reg")
            taille_test = etat.get("size_data_value")
            codage = etat.get("selectet_codage")

            if cible is None or features is None or len(features) == 0 or methode is None:
                ignores.append(nom + " (pas encore configure)")
                continue

            if _modele_a_jour(etat):
                reussis.append(nom)         # inchange depuis son dernier entrainement : resultats conserves
                continue

            if reg_modele is None:
                ## CORRECTIF : apres reouverture d'un projet, l'estimateur n'est pas en memoire (le JSON
                ## ne garde que ses hyperparametres) -> ces modeles etaient IGNORES par « Run », sans
                ## message. On reconstruit l'estimateur a partir de l'algorithme et des hyperparametres.
                reg_modele = _reconstruire_estimateur(nom, etat)
                if reg_modele is None:
                    _progression_unite_terminee()
                    ignores.append(nom + " (algorithme inconnu : ouvrez le modele, configurez-le puis cliquez sur « Save configure »)")
                    continue


            try:
                selectet_codage = codage if codage is not None else "None"
                from sklearn.base import clone as _cloner_estimateur
                reg_modele = _cloner_estimateur(reg_modele)     ## nouvel objet : l'ancien reste intact (Undo / Redo)
                XX = data_analy[list(features)]
                Y = data_analy[cible]
                ## CORRECTIF : on memorise le recodage effectivement utilise (necessaire ensuite
                ## pour la prediction et l'affichage du detail du modele)
                recod_memorise = etat.get("result_model_recod")
                if codage == "Disjunctive" and recod_memorise is not None and len(recod_memorise) > 0:
                    recod_utilise = list(recod_memorise)
                else:
                    recod_utilise = list(features)
                if methode == "Cross-validation":
                    # modele deja entraine (resultats presents) : ses hyperparametres sont deja optimises ;
                    # sinon (jamais execute ou reconfigure) : recherche par cross-validation, les valeurs
                    # configurees servant de bornes superieures
                    deja_ajuste = etat.get("data_model") is not None
                    d_XX_train, d_XX_test, Y_train_select, Y_test_select, reg_modele, meilleurs_parametres = my_select_all_features_CV(
                        XX, Y, list(features), recod_utilise, k = taille_test, reg_base = reg_modele, nom_algo = etat.get("Name_model"), reg_deja_ajuste = deja_ajuste)
                else:
                    d_XX_train, d_XX_test, Y_train_select, Y_test_select = my_select_all_features(XX, Y, list(features), recod_utilise, reg_modele, taille_test if taille_test else 0.20)
                _calcul_interruptible(lambda: reg_modele.fit(d_XX_train, Y_train_select))
                pred_test = reg_modele.predict(d_XX_test)
                pred_train = reg_modele.predict(d_XX_train)
                dp = pd.DataFrame(pred_test)
                xpred = np.array(dp[0].tolist())
                R2 = r2_squared(Y_test_select, xpred)
                mse = mean_squared_error(Y_test_select, pred_test)
                type_modele = etat.get("Name_model")
                if methode == "Cross-validation":
                    my_coef = None          # comme a l'entrainement manuel : data_model[6] = meilleurs hyperparametres
                elif type_modele in ("Linear", "Lasso"):
                    my_coef = reg_modele.coef_
                elif type_modele in ("Random Forest (RF)", "Generalized Boosted Regression Modeling (GBM)", "Decision Tree (DT)"):
                    my_coef = reg_modele.feature_importances_
                else:
                    my_coef = None

                y_min = Y.min() - 5
                y_max = Y.max() + 5

                etat["reg"] = reg_modele
                etat["data_model"] = [R2, pred_test, Y_test_select, mse, pred_train, Y_train_select, meilleurs_parametres if methode == "Cross-validation" else my_coef]
                ## CORRECTIF : ces champs n'etaient pas renseignes pour les modeles rejoues
                ## automatiquement (None) -> erreurs dans « Detail of Model » et en prediction.
                etat["columns_select"] = np.array(d_XX_train.columns)
                etat["my_modality_find"] = my_modality_find
                etat["quali_columns"] = [v for v in features if type(data_analy[v].iloc[0]) == str]
                etat["result_model_recod"] = recod_utilise if codage != "None" else None
                etat["selectet_methode_imp"] = methode
                etat["my_data_coef"] = my_coef
                etat["my_y_min"] = y_min
                etat["my_y_max"] = y_max
                etat["signature_entrainement"] = _signature_entrainement(etat)
                etat["duree_entrainement"] = _progression_unite_terminee()     # sert a estimer le temps du prochain « Run »
                modeles_projet[nom] = etat
                reussis.append(nom)
            except Exception as erreur:
                _progression_unite_terminee()
                ignores.append(nom + " (erreur : " + str(erreur) + ")")
    finally:
        selectet_codage = codage_sauvegarde

    try:
        _rafraichir_etat_modeles()      ## CORRECTIF : libelles « (trained) » des modeles rejoues
    except Exception:
        pass

    return reussis, ignores


## Hyperparametres des modeles lus dans le fichier du projet (le JSON ne stocke pas l'estimateur
## scikit-learn lui-meme, seulement ses hyperparametres) : {nom du modele: {parametre: valeur}}.
## Ils servent a reconstruire l'estimateur des modeles rouverts, pour que « Run » les entraine.
_hyperparametres_charges = {}


def _estimateur_par_defaut (nom_algo):
    """ Estimateur de l'algorithme avec les valeurs par defaut du panneau de configuration
    (sans lire les widgets, qui peuvent appartenir a un autre modele). """
    if nom_algo == "Linear":
        return LinearRegression()
    if nom_algo == "Lasso":
        return Lasso(alpha = 1.0)
    if nom_algo == "Kernel ridge (KR)":
        return KernelRidge(alpha = 1.0)
    if nom_algo == "Support Vector Regression (SVR)":
        return svm.SVR(kernel = "rbf", C = 100.0, gamma = 0.1)
    if nom_algo == "Random Forest (RF)":
        return RandomForestRegressor(n_estimators = 100, random_state = 0)
    if nom_algo == "Generalized Boosted Regression Modeling (GBM)":
        return ensemble.GradientBoostingRegressor(n_estimators = 500, max_depth = 5, loss = "squared_error", learning_rate = 0.5)
    if nom_algo == "Decision Tree (DT)":
        return DecisionTreeRegressor(splitter = 'best', max_depth = 20, random_state = 0)
    if nom_algo == "Gaussian process (GP)":
        return GaussianProcessRegressor(kernel = kernel, n_targets = 1, random_state = 0)
    return None


def _reconstruire_estimateur (nom_modele, etat):
    """ Modele configure dont l'estimateur n'est pas en memoire (projet rouvert) : estimateur
    reconstruit a partir de son algorithme (Name_model) et des hyperparametres enregistres
    dans le projet. Retourne None si l'algorithme est inconnu. """
    estimateur = _estimateur_par_defaut(etat.get("Name_model"))
    if estimateur is None:
        return None
    parametres = _hyperparametres_charges.get(nom_modele) or {}
    valides = estimateur.get_params()
    simples = {cle: valeur for cle, valeur in parametres.items()
               if cle in valides and (valeur is None or isinstance(valeur, (bool, int, float, str)))}
    for cle, valeur in simples.items():
        try:
            estimateur.set_params(**{cle: valeur})
        except Exception:
            pass
    return estimateur


def _modeles_projet_pour_json ():
    """ Modeles a enregistrer : un modele rouvert et pas encore re-entraine (estimateur absent)
    garde les hyperparametres lus dans le projet, au lieu de les perdre au prochain enregistrement. """
    resultat = {}
    for nom, etat in modeles_projet.items():
        etat_json = _valeur_vers_json(etat)
        if etat.get("reg") is None and _hyperparametres_charges.get(nom):
            etat_json["reg"] = {"__type__": str(etat.get("Name_model")), "hyperparametres": _valeur_vers_json(_hyperparametres_charges[nom]), "__non_reentraine__": True}
        resultat[str(nom)] = etat_json
    return resultat


def _instancier_reg_depuis_widgets (nom_algo):
  
    def lire (nom_widget, defaut, cast = float):
        try:
            widget = globals().get(nom_widget)
            if widget is not None:
                return cast(widget.get())
        except Exception:
            pass
        return defaut

    try:
        if nom_algo == "Linear":
            return LinearRegression()
        elif nom_algo == "Lasso":
            return Lasso(alpha = lire("Alpha_lasso_size", 1.0))
        elif nom_algo == "Kernel ridge (KR)":
            return KernelRidge(alpha = lire("Alpha_size", 1.0))
        elif nom_algo == "Support Vector Regression (SVR)":
            return svm.SVR(kernel = "rbf", C = lire("C_size", 100.0), gamma = lire("gamma_size", 0.1))
        elif nom_algo == "Random Forest (RF)":
            return RandomForestRegressor(n_estimators = int(lire("trees_size", 100, int)), random_state = int(lire("Random_size", 0, int)))
        elif nom_algo == "Generalized Boosted Regression Modeling (GBM)":
            loss_defaut = "squared_error"
            try:
                widget_loss = globals().get("loss_size")
                if widget_loss is not None:
                    loss_defaut = widget_loss.get() or loss_defaut
            except Exception:
                pass
            return ensemble.GradientBoostingRegressor(n_estimators = int(lire("estimators_size", 500, int)), max_depth = int(lire("depth_size", 5, int)), loss = loss_defaut, learning_rate = lire("rate_size", 0.5))
        elif nom_algo == "Decision Tree (DT)":
            return DecisionTreeRegressor(splitter = 'best', max_depth = int(lire("depth_size_r", 20, int)), random_state = int(lire("Random_size_r", 0, int)))
        elif nom_algo == "Gaussian process (GP)":
            return GaussianProcessRegressor(kernel = kernel, n_targets = 1, random_state = int(lire("Random_size_gpr", 0, int)))
    except Exception:
        return None
    return None


def _capturer_configuration_widgets_actifs ():
   
    if nom_modele_actif is None:
        return False

    ## CORRECTIF : ne capturer que si le panneau ouvert appartient bien au modele actif
    if _panneau_config.get("nom") != nom_modele_actif or not _widget_existe(_panneau_config.get("param")):
        return False

    ## CORRECTIF : un modele deja entraine n'est modifie que si l'utilisateur a choisi de le reconfigurer
    etat_courant = modeles_projet.get(nom_modele_actif)
    if etat_courant is not None and etat_courant.get("data_model") is not None and not _panneau_config.get("reconfiguration"):
        return False

    try:
        target_widget = globals().get("Target_column")
        if target_widget is None:
            return False
        target_index = target_widget.current()
        if target_index <= 0:
            return False
        cible = header_target[target_index]

        feature_widget = globals().get("feature_select")
        if feature_widget is None:
            return False
        feature_index = feature_widget.current()
        if feature_index <= 0:
            return False
        if feature_index == 1:
            features = globals().get("result_model_1")
            if not features:
                return False
        else:
            ## CORRECTIF : data_analy peut etre None dans un nouveau projet
            if not _assurer_data_analy(silencieux = True):
                return False
            d_X = data_analy.copy()
            d_X = d_X.drop(cible, axis = 1)
            features = list(d_X.columns)

        methode_widget = globals().get("Calculation_method")
        if methode_widget is None:
            return False
        methode_index = methode_widget.current()
        if methode_index < 0:
            return False
        methode = ["TrainTestSplit data", "Cross-validation"][methode_index]

        codage_widget = globals().get("codage_method")
        if codage_widget is None:
            return False
        codage_index = codage_widget.current()
        if codage_index < 0:
            return False
        codage = ["None", "Disjunctive", "All-disjunctive"][codage_index]

        taille_test = 0.20
        try:
            test_widget = globals().get("Test_size")
            if test_widget is not None:
                taille_test = float(test_widget.get())
        except Exception:
            pass

        reg_instance = _instancier_reg_depuis_widgets(Name_model)
        if reg_instance is None:
            return False

        etat = modeles_projet.get(nom_modele_actif, {cle: None for cle in _MODEL_STATE_KEYS})
        etat["Name_model"] = Name_model
        etat["selectet_target"] = cible
        etat["result_model"] = features
        etat["selectet_methode"] = methode
        etat["selectet_methode_imp"] = methode
        etat["selectet_codage"] = codage
        etat["size_data_value"] = taille_test
        etat["reg"] = reg_instance
        if codage == "Disjunctive":
            recod_choisi = globals().get("result_model_recod")
            if recod_choisi is not None:
                etat["result_model_recod"] = list(recod_choisi)
        if _panneau_config.get("reconfiguration"):
            ## reconfiguration d'un modele entraine : anciens resultats invalides tant que Run n'a pas ete relance
            etat["data_model"] = None
            etat["my_data_coef"] = None
        modeles_projet[nom_modele_actif] = etat
        ## CORRECTIF : resynchroniser les variables globales du modele actif avec sa configuration enregistree.
        ## Sinon snapshot_modele_actif() (sauvegarde du projet, autosave, Run) recopiait les anciennes valeurs
        ## globales (reg = None, cible = None...) par-dessus la configuration qui venait d'etre enregistree.
        restore_modele(nom_modele_actif)
        ## CORRECTIF : rafraichir le libelle « (configured) » des boites de modeles deja affichees
        try:
            _rafraichir_etat_modeles()
        except Exception:
            pass
        return True
    except Exception:
        return False


def sauvegarder_configuration_modele_bouton (boite_a_fermer = None):

    if nom_modele_actif is None:
        messagebox.showwarning("Aucun modele actif", "Aucun modele n'est actuellement selectionne.")
        return

    if _capturer_configuration_widgets_actifs():
        messagebox.showinfo("Configuration enregistree", "La configuration du modele « " + str(nom_modele_actif) + " » a ete enregistree.\n\nElle sera automatiquement executee au prochain clic sur « Run », meme sans repasser par ce panneau.")
        if boite_a_fermer is not None:
            try:
                boite_a_fermer.destroy()
            except Exception:
                pass
        _fermer_panneau_config()
        _historique_valider("Configure model « " + str(nom_modele_actif) + " »")
    else:
        messagebox.showwarning("Configuration incomplete", "Impossible d'enregistrer la configuration de « " + str(nom_modele_actif) + " ».\n\nVerifiez que la cible, les features, la methode de calcul et le recodage sont tous bien renseignes.")



# ==========================================================================
# ==  12. Enregistrement du projet (JSON + dossier associe)               ==
# ==========================================================================
## ------ Sauvegarde du projet (etat des donnees / modeles / reglages) en JSON ------

def _valeur_vers_json (valeur):
  
    if valeur is None or isinstance(valeur, (bool, int, float, str)):
        return valeur

    if isinstance(valeur, (list, tuple)):
        return [_valeur_vers_json(v) for v in valeur]

    if isinstance(valeur, dict):
        return {str(cle): _valeur_vers_json(v) for cle, v in valeur.items()}

    if isinstance(valeur, np.ndarray):
        return _valeur_vers_json(valeur.tolist())

    if isinstance(valeur, (np.integer, np.floating, np.bool_)):
        return valeur.item()

    if isinstance(valeur, pd.DataFrame):
        return {"__type__": "DataFrame", "data": valeur.to_dict(orient = "list")}

    if isinstance(valeur, pd.Series):
        return {"__type__": "Series", "data": valeur.tolist()}

    if hasattr(valeur, "get_params"):
        try:
            return {"__type__": type(valeur).__name__, "hyperparametres": _valeur_vers_json(valeur.get_params()), "__non_reentraine__": True}
        except Exception:
            pass

    return {"__type__": type(valeur).__name__, "repr": str(valeur)[:300]}


def _maj_titre_fenetre ():
    ''' Met a jour le titre de la fenetre principale pour afficher le projet actif. '''
    try:
        if nom_projet:
            SmartDigicat.title("SmartDigiCat v1.0 - Projet : " + nom_projet)
        else:
            SmartDigicat.title("SmartDigiCat v1.0")
    except Exception:
        pass


def _valeur_reglage_vers_json (valeur):
   
    if valeur is None or isinstance(valeur, (bool, int, float, str)):
        return valeur

    if isinstance(valeur, (list, tuple)):
        return [_valeur_reglage_vers_json(v) for v in valeur]

    if isinstance(valeur, tk.Listbox):
        try:
            return list(valeur.get(0, tk.END))
        except Exception:
            return {"__type__": "Listbox", "repr": "contenu illisible"}

    ## Combobox / Spinbox / Entry / autres widgets dotes d'un .get() : valeur actuelle
    if hasattr(valeur, "get"):
        try:
            return _valeur_reglage_vers_json(valeur.get())
        except Exception:
            pass

    ## Dernier recours : jamais l'objet widget lui-meme, juste une trace lisible
    return {"__type__": type(valeur).__name__, "repr": str(valeur)[:300]}


def _dossier_associe_au_projet (chemin_fichier_json):
  
    base, _ext = os.path.splitext(chemin_fichier_json)
    return base + "_fichiers"


## Memorisation du reg (modele scikit-learn entraine) : le JSON ne peut pas le contenir, il est donc ecrit
## dans un fichier joblib du dossier du projet. Par defaut seul le MEILLEUR modele (R-squared le plus eleve
## parmi les modeles utilisables) est memorise ; passer a True pour memoriser tous les modeles entraines
## (fichiers plus volumineux, surtout avec Random Forest / GBM).
MEMORISER_REG_DE_TOUS_LES_MODELES = False

_cache_reg_enregistres = {}     ## evite de re-ecrire un gros fichier a chaque sauvegarde automatique


def _noms_modeles_reg_a_memoriser ():
    entraines = [nom for nom in modeles_projet if _modele_utilisable_pour_prediction(nom)]
    nom_meilleur, _r2 = _meilleur_modele_entraine(entraines, pour_prediction = True)
    if MEMORISER_REG_DE_TOUS_LES_MODELES:
        return entraines, nom_meilleur
    return ([nom_meilleur] if nom_meilleur is not None else []), nom_meilleur


def _enregistrer_reg_modeles (dossier_projet, nom_dossier, noms, info):
    import joblib, hashlib
    dossier_modeles = os.path.join(dossier_projet, "modeles")
    if len(noms) != 0:
        os.makedirs(dossier_modeles, exist_ok = True)
    fichiers_attendus = set()
    for nom in noms:
        etat = modeles_projet[nom]
        reg_modele = etat.get("reg")
        try:
            donnees = etat.get("data_model")
            empreinte_etat = (id(reg_modele), _nombre_pour_detail(donnees[0], 12), _nombre_pour_detail(donnees[3], 12), len(etat.get("columns_select")))
        except Exception:
            empreinte_etat = None
        nom_fichier = "reg_" + re.sub(r"[^0-9A-Za-z]+", "_", str(nom))[:40] + "_" + hashlib.sha256(str(nom).encode("utf-8")).hexdigest()[:8] + ".joblib"
        chemin_reg = os.path.join(dossier_modeles, nom_fichier)
        fichiers_attendus.add(nom_fichier)
        cache = _cache_reg_enregistres.get(nom)
        if cache is not None and empreinte_etat is not None and cache[0] == empreinte_etat and cache[1] == chemin_reg and os.path.isfile(chemin_reg):
            info["modeles_entraines"][nom] = dict(cache[2])
            continue
        try:
            joblib.dump(reg_modele, chemin_reg, compress = 3)
            with open(chemin_reg, "rb") as f_reg:
                contenu_reg = f_reg.read()
            infos_reg = {"fichier": nom_dossier + "/modeles/" + nom_fichier,
                         "sha256": hashlib.sha256(contenu_reg).hexdigest(),
                         "sklearn": sklearn.__version__,
                         "taille_mo": round(len(contenu_reg) / 1048576, 2)}
            info["modeles_entraines"][nom] = infos_reg
            _cache_reg_enregistres[nom] = (empreinte_etat, chemin_reg, dict(infos_reg))
        except Exception as erreur:
            info["erreurs"].append("Modele « " + str(nom) + " » (reg) non enregistre : " + str(erreur))
    ## suppression des anciens fichiers qui ne correspondent plus a un modele memorise
    if os.path.isdir(dossier_modeles):
        for nom_existant in os.listdir(dossier_modeles):
            if nom_existant.endswith(".joblib") and nom_existant not in fichiers_attendus:
                try:
                    os.remove(os.path.join(dossier_modeles, nom_existant))
                except Exception:
                    pass


def _exporter_donnees_et_graphiques (chemin_fichier_json):
  
    dossier_projet = _dossier_associe_au_projet(chemin_fichier_json)
    nom_dossier = os.path.basename(dossier_projet)

    info = {"dossier": None, "donnees_filtrees": None, "donnees_optimisation": None, "donnees_prediction": None, "resultats_prediction": None, "modeles_entraines": {}, "meilleur_modele": None, "graphiques": {}, "erreurs": []}

    noms_reg_a_memoriser, nom_meilleur_reg = _noms_modeles_reg_a_memoriser()
    info["meilleur_modele"] = nom_meilleur_reg

    ## Rien a exporter : ni donnees filtrees, ni donnees d'optimisation, ni donnees/resultats de prediction, ni graphique ouvert -> on ne cree pas de dossier
    if data_analy is None and data_opt is None and data_pred is None and data_pred_new is None and len(noms_reg_a_memoriser) == 0 and len(figures_ouvertes_courantes) == 0:
        return info

    try:
        os.makedirs(dossier_projet, exist_ok = True)
    except Exception as erreur:
        info["erreurs"].append("Impossible de creer le dossier « " + dossier_projet + " » : " + str(erreur))
        return info

    info["dossier"] = nom_dossier

    if data_analy is not None:
        try:
            nom_fichier_donnees = "donnees_filtrees.csv"
            data_analy.to_csv(os.path.join(dossier_projet, nom_fichier_donnees), sep = ";", index = False, decimal = ".")
            info["donnees_filtrees"] = nom_dossier + "/" + nom_fichier_donnees
        except Exception as erreur:
            info["erreurs"].append("Jeu de donnees filtre non enregistre : " + str(erreur))

    if data_opt is not None:
        try:
            nom_fichier_donnees_optimisation = "donnees_optimisation.csv"
            data_opt.to_csv(os.path.join(dossier_projet, nom_fichier_donnees_optimisation), sep = ";", index = False, decimal = ".")
            info["donnees_optimisation"] = nom_dossier + "/" + nom_fichier_donnees_optimisation
        except Exception as erreur:
            info["erreurs"].append("Donnees d'optimisation non enregistrees : " + str(erreur))

    ## reg du meilleur modele (ou de tous, voir MEMORISER_REG_DE_TOUS_LES_MODELES)
    try:
        _enregistrer_reg_modeles(dossier_projet, nom_dossier, noms_reg_a_memoriser, info)
    except Exception as erreur:
        info["erreurs"].append("Modele(s) entraine(s) (reg) non enregistre(s) : " + str(erreur))

    ## Donnees de prediction importees et resultats de la prediction
    if data_pred is not None:
        try:
            nom_fichier_pred = "donnees_prediction.csv"
            data_pred.to_csv(os.path.join(dossier_projet, nom_fichier_pred), sep = ";", index = False, decimal = ".")
            info["donnees_prediction"] = nom_dossier + "/" + nom_fichier_pred
        except Exception as erreur:
            info["erreurs"].append("Donnees de prediction non enregistrees : " + str(erreur))

    if data_pred_new is not None:
        try:
            nom_fichier_resultats_pred = "resultats_prediction.csv"
            data_pred_new.to_csv(os.path.join(dossier_projet, nom_fichier_resultats_pred), sep = ";", index = False, decimal = ".")
            info["resultats_prediction"] = nom_dossier + "/" + nom_fichier_resultats_pred
        except Exception as erreur:
            info["erreurs"].append("Resultats de prediction non enregistres : " + str(erreur))

    if len(figures_ouvertes_courantes) != 0:
        dossier_graphiques = os.path.join(dossier_projet, "graphiques")
        try:
            os.makedirs(dossier_graphiques, exist_ok = True)
        except Exception as erreur:
            info["erreurs"].append("Impossible de creer le dossier « graphiques » : " + str(erreur))
        else:
          
            for nom_graphique, fig in list(figures_ouvertes_courantes.items()):
                try:
                    nom_fichier_image = nom_graphique + ".png"
                    fig.savefig(os.path.join(dossier_graphiques, nom_fichier_image))
                    info["graphiques"][nom_graphique] = nom_dossier + "/graphiques/" + nom_fichier_image
                except Exception as erreur:
                    info["erreurs"].append("Graphique « " + str(nom_graphique) + " » non enregistre : " + str(erreur))

    return info


def save_current_project (silencieux = False):
    """Enregistre le projet. Renvoie True (enregistre), False (echec : erreur deja affichee) ou None (annule / aucun projet)."""
  
    global nom_projet, chemin_projet

    if nom_projet is None:
        messagebox.showwarning("Aucun projet ouvert", "Veuillez d'abord creer un nouveau projet (bouton « New ») ou en ouvrir un existant (bouton « Open ») avant de continuer.")
        return None

    snapshot_modele_actif()

    if chemin_projet:
        chemin_fichier = chemin_projet
    else:
        chemin_fichier = filedialog.asksaveasfilename(defaultextension = ".json", filetypes = [("Fichier JSON", "*.json"), ("Tous les fichiers", "*.*")], initialfile = nom_projet.replace(" ", "_") + ".json", title = "Enregistrer le projet")
        if not chemin_fichier:
            return None
        chemin_projet = chemin_fichier

  
    try:
        info_fichiers = _exporter_donnees_et_graphiques(chemin_fichier)
    except Exception as erreur:
        info_fichiers = {"dossier": None, "donnees_filtrees": None, "graphiques": {}, "erreurs": ["Export des donnees/graphiques : " + str(erreur)]}

    contenu = {
        "reglages_actifs": {
            "nom_projet": nom_projet,
            "nom_modele_actif": nom_modele_actif,
            "file_path": file_path,
            "file_sheets": list(file_sheets) if file_sheets else None,
            "Name_model": Name_model,
            "Target_column": _valeur_reglage_vers_json(Target_column),
            "Feature_column": _valeur_reglage_vers_json(Feature_column),
            "Calculation_method": _valeur_reglage_vers_json(Calculation_method),
            "codage_method": _valeur_reglage_vers_json(codage_method),
            "Test_size": _valeur_reglage_vers_json(Test_size),
            "Target_list": _valeur_reglage_vers_json(Target_list),
            "header_target": _valeur_reglage_vers_json(header_target),
            "dernier_x_scatter": _valeur_reglage_vers_json(dernier_x_scatter),
            "dernier_y_scatter": _valeur_reglage_vers_json(dernier_y_scatter),
            "dernier_color_scatter": _valeur_reglage_vers_json(dernier_color_scatter),
            "dernier_methode_boxplot": _valeur_reglage_vers_json(dernier_methode_boxplot),
            "dernier_x_boxplot": _valeur_reglage_vers_json(dernier_x_boxplot),
            "dernier_y_boxplot": _valeur_reglage_vers_json(dernier_y_boxplot),
            "dernier_color_boxplot": _valeur_reglage_vers_json(dernier_color_boxplot),
            "dernier_var_histogram": _valeur_reglage_vers_json(dernier_var_histogram),
            "dernier_couleur_histogram": _valeur_reglage_vers_json(dernier_couleur_histogram),
            "dernier_densite_histogram": _valeur_reglage_vers_json(dernier_densite_histogram),
            "dernier_x_plot3d": _valeur_reglage_vers_json(dernier_x_plot3d),
            "dernier_y_plot3d": _valeur_reglage_vers_json(dernier_y_plot3d),
            "dernier_z_plot3d": _valeur_reglage_vers_json(dernier_z_plot3d),
            "dernier_color_plot3d": _valeur_reglage_vers_json(dernier_color_plot3d),
            "dernier_target_correlation": _valeur_reglage_vers_json(dernier_target_correlation),
            "dernier_features_correlation": _valeur_reglage_vers_json(dernier_features_correlation),
            "dernier_dimension_min_max": _valeur_reglage_vers_json(dernier_dimension_min_max),
            "dernier_methode_min_max": _valeur_reglage_vers_json(dernier_methode_min_max),
            "dernier_features_min_max": _valeur_reglage_vers_json(dernier_features_min_max),
        },
        "modeles_projet": _modeles_projet_pour_json(),
        "modeles_supprimes": _valeur_vers_json(modeles_supprimes),
        "apparence_boites_modeles": {
            "couleurs": {nom: _valeur_reglage_vers_json(couleur) for nom, couleur in couleurs_boites_modeles.items()},
            "positions": {nom: [_valeur_reglage_vers_json(x), _valeur_reglage_vers_json(y)] for nom, (x, y) in positions_boites_modeles.items()},
        },
        "apparence_boites_modeles_supprimes": {
            "couleurs": {nom: _valeur_reglage_vers_json(couleur) for nom, couleur in couleurs_boites_supprimees.items()},
            "positions": {nom: [_valeur_reglage_vers_json(x), _valeur_reglage_vers_json(y)] for nom, (x, y) in positions_boites_supprimees.items()},
        },
        "filtres_edit_data": {
            "parametre_1": _filtre_edit_vers_json(filtre_edit_parametre_1),
            "parametre_2": _filtre_edit_vers_json(filtre_edit_parametre_2),
            "parametre_3": _filtre_edit_vers_json(filtre_edit_parametre_3),
            "parametre_4": _filtre_edit_vers_json(filtre_edit_parametre_4),
        },
        "prediction": {"file_path_pred": file_path_pred, "infos": _valeur_vers_json(prediction_infos)},
        "parametres_optimisation": _valeur_vers_json(opt_parameters),
        "graphique_optimisation": {"donnees": _valeur_vers_json(opt_plot_data), "affichage": _valeur_vers_json(opt_affichage)},
        "fichiers_associes": info_fichiers,
    }

    try:
        with open(chemin_fichier, "w", encoding = "utf-8") as f:
            json.dump(contenu, f, indent = 2, ensure_ascii = False)
        _maj_titre_fenetre()
        if not silencieux:
            message = "Le projet « " + nom_projet + " » a ete enregistre :\n" + chemin_fichier + "\n\nIl contient " + str(len(modeles_projet)) + " modele(s)."
            if info_fichiers["donnees_filtrees"] is not None:
                message += "\n\nJeu de donnees filtre enregistre : " + info_fichiers["donnees_filtrees"]
            if info_fichiers.get("donnees_optimisation") is not None:
                message += "\n\nDonnees d'optimisation enregistrees : " + info_fichiers["donnees_optimisation"]
            if len(info_fichiers.get("modeles_entraines", {})) != 0:
                message += "\n\nModele(s) memorise(s) pour la prediction : " + ", ".join([str(n) + " (" + str(i.get("taille_mo")) + " Mo)" for n, i in info_fichiers["modeles_entraines"].items()])
            if info_fichiers.get("donnees_prediction") is not None:
                message += "\n\nDonnees de prediction enregistrees : " + info_fichiers["donnees_prediction"]
            if info_fichiers.get("resultats_prediction") is not None:
                message += "\n\nResultats de prediction enregistres : " + info_fichiers["resultats_prediction"]
            if len(info_fichiers["graphiques"]) != 0:
                message += "\n" + str(len(info_fichiers["graphiques"])) + " graphique(s) actuellement affiche(s) enregistre(s) dans : " + info_fichiers["dossier"] + "/graphiques"
            if len(info_fichiers.get("erreurs", [])) != 0:
                message += "\n\nAttention, certains elements n'ont pas pu etre enregistres :\n- " + "\n- ".join(info_fichiers["erreurs"])
            message += "\n\nRemarque : les reglages, hyperparametres, donnees filtrees, graphiques affiches et le meilleur modele entraine (fichier joblib du dossier du projet) sont sauvegardes. Les autres modeles entraines devront etre re-entraines (bouton Run) apres rechargement pour etre relances ou utilises en prediction."
            messagebox.showinfo("Projet enregistre", message)
        _projet_marquer_enregistre()
        return True
    except Exception as erreur:
        messagebox.showerror("Enregistrement impossible", "Impossible d'enregistrer le projet :\n" + str(erreur))
        return False


def _etat_modele_depuis_json (etat_json):
 
    def vers_array (v):
      
        if isinstance(v, dict):
            type_v = v.get("__type__")
            if type_v == "Series":
                return np.array(v.get("data", []))
            if type_v == "DataFrame":
                return pd.DataFrame(v.get("data", {}))
            return v
        return np.array(v) if isinstance(v, list) else v

    etat = {cle: etat_json.get(cle) for cle in _MODEL_STATE_KEYS}
    etat["result_model"] = vers_array(etat.get("result_model"))
    etat["columns_select"] = vers_array(etat.get("columns_select"))
    etat["my_modality_find"] = vers_array(etat.get("my_modality_find"))
    etat["my_data_coef"] = vers_array(etat.get("my_data_coef"))
    if isinstance(etat.get("data_model"), list):
        etat["data_model"] = tuple(vers_array(v) for v in etat["data_model"])
    etat["reg"] = None  ## non reconstituable depuis JSON, voir docstring

    return etat



# ==========================================================================
# ==  13. Lecture des fichiers de donnees, ouverture d'un projet          ==
# ==========================================================================
## ---------------------------------------------------------------------------------------------
## Lecture des fichiers de donnees (CSV, texte, Excel a une ou plusieurs feuilles)
## ---------------------------------------------------------------------------------------------
EXTENSIONS_EXCEL = (".xlsx", ".xlsm", ".xls")


def _choisir_feuilles_excel (chemin, classeur):
    """ Fenetre modale de choix des feuilles d'un classeur Excel qui en contient plusieurs.
    classeur : dictionnaire {nom de feuille: DataFrame} (ordre du fichier).
    Renvoie la liste des feuilles choisies (dans l'ordre du fichier), ou None si l'utilisateur annule. """
    noms = list(classeur)
    resultat = {"feuilles": None}

    fenetre = tk.Toplevel(SmartDigicat)
    fenetre.title(traduire("Excel sheets"))
    fenetre.resizable(False, False)
    fenetre.transient(SmartDigicat)

    tk.Label(fenetre, text = os.path.basename(chemin), font = ('arial', 14, 'bold')).pack(padx = 20, pady = (15, 4))
    tk.Label(fenetre, text = traduire("This file contains several sheets. Select the sheet(s) to import.") + "\n"
             + traduire("Several sheets with the same columns are stacked into a single table."),
             font = ('arial', 12), justify = tk.LEFT).pack(padx = 20, pady = (0, 8))

    cadre = tk.Frame(fenetre)
    cadre.pack(padx = 20, pady = 4, fill = tk.BOTH, expand = True)
    barre = tk.Scrollbar(cadre, orient = tk.VERTICAL)
    liste = tk.Listbox(cadre, selectmode = tk.EXTENDED, exportselection = False, width = 55, height = int(np.clip(len(noms), 3, 12)),
                       font = ('arial', 13), yscrollcommand = barre.set, activestyle = "none")
    barre.config(command = liste.yview)
    liste.pack(side = tk.LEFT, fill = tk.BOTH, expand = True)
    barre.pack(side = tk.RIGHT, fill = tk.Y)
    for nom in noms:
        tableau = classeur[nom]
        liste.insert(tk.END, "%s   (%d %s x %d %s)" % (nom, tableau.shape[0], traduire("rows"), tableau.shape[1], traduire("columns")))
    liste.selection_set(0)

    message = tk.Label(fenetre, text = "", font = ('arial', 12), fg = "red")
    message.pack(padx = 20, pady = (4, 0))

    def valider (*_):
        indices = liste.curselection()
        if len(indices) == 0:
            message.config(text = traduire("Select at least one sheet."))
            return
        resultat["feuilles"] = [noms[i] for i in indices]
        fenetre.destroy()

    def tout_selectionner ():
        liste.selection_set(0, tk.END)

    boutons = tk.Frame(fenetre)
    boutons.pack(padx = 20, pady = (8, 15))
    tk.Button(boutons, text = traduire("Select all"), command = tout_selectionner, width = 12).pack(side = tk.LEFT, padx = 6)
    tk.Button(boutons, text = traduire("Cancel"), command = fenetre.destroy, width = 12).pack(side = tk.LEFT, padx = 6)
    tk.Button(boutons, text = traduire("Import"), command = valider, width = 12).pack(side = tk.LEFT, padx = 6)
    liste.bind("<Double-Button-1>", valider)
    fenetre.bind("<Return>", valider)
    fenetre.bind("<Escape>", lambda *_: fenetre.destroy())

    fenetre.update_idletasks()
    try:
        fenetre.grab_set()
    except Exception:
        pass
    liste.focus_set()
    SmartDigicat.wait_window(fenetre)
    return resultat["feuilles"]


def _empiler_feuilles_excel (classeur, feuilles):
    """ Une feuille : renvoyee telle quelle. Plusieurs feuilles : empilees (lignes les unes sous les autres),
    a condition qu'elles aient les memes colonnes (l'ordre des colonnes de la premiere feuille est conserve). """
    manquantes = [f for f in feuilles if f not in classeur]
    if manquantes:
        raise ValueError(traduire("Sheet(s) not found in the Excel file: ") + ", ".join(str(f) for f in manquantes))
    if len(feuilles) == 1:
        return pd.DataFrame(classeur[feuilles[0]])
    colonnes_ref = list(classeur[feuilles[0]].columns)
    for f in feuilles[1:]:
        if set(classeur[f].columns) != set(colonnes_ref):
            raise ValueError(traduire("The selected sheets do not have the same columns, they cannot be stacked:") + "\n"
                             + str(feuilles[0]) + " : " + ", ".join(str(c) for c in colonnes_ref) + "\n"
                             + str(f) + " : " + ", ".join(str(c) for c in classeur[f].columns))
    return pd.concat([classeur[f][colonnes_ref] for f in feuilles], ignore_index = True)


def lire_fichier_donnees (chemin, feuilles = None, na_values = None, choisir_feuilles = False):
    """ Lit un fichier de donnees CSV (;), Excel (.xlsx, .xlsm, .xls) ou texte (,).
    Excel : feuilles = liste des feuilles a lire (None : premiere feuille) ; si choisir_feuilles est vrai et que
    le classeur contient plusieurs feuilles, l'utilisateur les choisit dans une fenetre.
    Renvoie (DataFrame, feuilles utilisees) ; feuilles vaut None pour un fichier non Excel.
    Renvoie (None, None) si l'utilisateur annule le choix des feuilles. """
    extension = os.path.splitext(chemin)[1].lower()
    if extension == ".csv":
        return pd.DataFrame(pd.read_csv(chemin, delimiter = ";", na_values = na_values)), None
    if extension in EXTENSIONS_EXCEL:
        classeur = pd.read_excel(chemin, sheet_name = None, na_values = na_values)
        noms = list(classeur)
        if len(noms) == 0:
            raise ValueError(traduire("The Excel file does not contain any sheet."))
        if choisir_feuilles:
            if len(noms) > 1:
                feuilles = _choisir_feuilles_excel(chemin, classeur)
                if feuilles is None:
                    return None, None
            else:
                feuilles = noms
        elif not feuilles:
            feuilles = [noms[0]]
        return _empiler_feuilles_excel(classeur, list(feuilles)), list(feuilles)
    return pd.DataFrame(pd.read_table(chemin, delimiter = ',', na_values = na_values)), None


def _assurer_donnees_source_chargees ():
 
    global data, file_path, file_sheets

    if data is not None and file_path is not None and os.path.isfile(file_path):
        return True

    if file_path is None or not os.path.isfile(file_path):
        messagebox.showwarning("Fichier de donnees introuvable", "Le fichier de donnees d'origine de ce projet est introuvable a son emplacement enregistre" + ((traduire(" :\n") + str(file_path)) if file_path else "") + ".\n\nVeuillez le resélectionner pour pouvoir configurer les modeles.")
        my_fill = [('CSV files', "*.csv"), ('Excel file', '*.xlsx *.xlsm *.xls'), ('Texte file', '*.txt'), ('All', "*.*")]
        nouveau_chemin = filedialog.askopenfilename(title = "Resélectionner le fichier de donnees du projet", filetypes = my_fill)
        if not nouveau_chemin:
            return False
        file_path = nouveau_chemin
        file_sheets = None
        ## nouveau fichier Excel : les feuilles sont choisies a nouveau
        if os.path.splitext(nouveau_chemin)[1].lower() in EXTENSIONS_EXCEL:
            try:
                donnees_lues, feuilles = lire_fichier_donnees(nouveau_chemin, na_values = my_missing_values, choisir_feuilles = True)
            except Exception as erreur:
                messagebox.showerror("Lecture impossible", "Impossible de lire le fichier de donnees :\n" + str(erreur))
                data = None
                return False
            if donnees_lues is None:
                data = None
                return False
            file_sheets = feuilles

    try:
        data, _ = lire_fichier_donnees(file_path, file_sheets, na_values = my_missing_values)
        return True
    except Exception as erreur:
        messagebox.showerror("Lecture impossible", "Impossible de lire le fichier de donnees :\n" + str(erreur))
        data = None
        return False


def ouvrir_projet_depuis_fichier ():

    global nom_projet, chemin_projet, nom_modele_actif, file_path, file_sheets, Name_model, Target_column, Feature_column
    global Calculation_method, codage_method, Test_size, Target_list, header_target, data_analy
    global dernier_x_scatter, dernier_y_scatter, dernier_color_scatter
    global dernier_methode_boxplot, dernier_x_boxplot, dernier_y_boxplot, dernier_color_boxplot
    global dernier_var_histogram, dernier_couleur_histogram, dernier_densite_histogram
    global dernier_x_plot3d, dernier_y_plot3d, dernier_z_plot3d, dernier_color_plot3d
    global dernier_target_correlation, dernier_features_correlation
    global dernier_dimension_min_max, dernier_methode_min_max, dernier_features_min_max
    global opt_plot_data, data_opt, data, data_pred, data_pred_new, file_path_pred, prediction_infos

    chemin_fichier = filedialog.askopenfilename(defaultextension = ".json", filetypes = [("Fichier JSON", "*.json"), ("Tous les fichiers", "*.*")], title = "Ouvrir un projet")
    if not chemin_fichier:
        return

    try:
        with open(chemin_fichier, "r", encoding = "utf-8") as f:
            contenu = json.load(f)
    except Exception as erreur:
        messagebox.showerror("Ouverture impossible", "Impossible de lire le fichier projet :\n" + str(erreur))
        return

    try:
        reglages = contenu.get("reglages_actifs", {})
        modeles_json = contenu.get("modeles_projet", {})
        corbeille_json = contenu.get("modeles_supprimes", {})

        _fermer_panneau_config()
        ## CORRECTIF : ne pas garder les donnees d'un projet precedent (data / data_analy / prediction)
        data = None
        data_analy = None
        data_pred = None
        data_pred_new = None
        file_path_pred = None
        prediction_infos = {}
        modeles_projet.clear()
        _hyperparametres_charges.clear()
        for nom, etat_json in modeles_json.items():
            modeles_projet[nom] = _etat_modele_depuis_json(etat_json)
            reg_json = etat_json.get("reg") if isinstance(etat_json, dict) else None
            if isinstance(reg_json, dict) and isinstance(reg_json.get("hyperparametres"), dict):
                _hyperparametres_charges[nom] = reg_json["hyperparametres"]

        modeles_supprimes.clear()
        for nom, etat_json in corbeille_json.items():
            modeles_supprimes[nom] = _etat_modele_depuis_json(etat_json)

        ## Couleurs/positions des boites de modeles : celles enregistrees dans
        ## CE fichier sont prioritaires, mais une couleur/position deja definie
        ## pendant la session pour un modele de ce meme projet est conservee
        ## (utile si le projet n'a pas encore ete resauvegarde). En revanche,
        ## les couleurs/positions d'un AUTRE projet precedemment ouvert dans la
        ## meme session sont ecartees pour ne pas "deteindre" sur celui-ci.
        apparence_json = contenu.get("apparence_boites_modeles", {})

        couleurs_a_conserver = {nom: couleur for nom, couleur in couleurs_boites_modeles.items() if nom in modeles_projet}
        couleurs_a_conserver.update(apparence_json.get("couleurs", {}) or {})
        couleurs_boites_modeles.clear()
        couleurs_boites_modeles.update(couleurs_a_conserver)

        positions_a_conserver = {nom: position for nom, position in positions_boites_modeles.items() if nom in modeles_projet}
        for nom, position in (apparence_json.get("positions", {}) or {}).items():
            try:
                positions_a_conserver[nom] = (position[0], position[1])
            except Exception:
                pass
        positions_boites_modeles.clear()
        positions_boites_modeles.update(positions_a_conserver)

        ## Meme logique pour l'apparence des modeles se trouvant dans la corbeille
        apparence_corbeille_json = contenu.get("apparence_boites_modeles_supprimes", {})

        couleurs_boites_supprimees.clear()
        couleurs_boites_supprimees.update({nom: couleur for nom, couleur in (apparence_corbeille_json.get("couleurs", {}) or {}).items() if nom in modeles_supprimes})

        positions_boites_supprimees.clear()
        for nom, position in (apparence_corbeille_json.get("positions", {}) or {}).items():
            if nom in modeles_supprimes:
                try:
                    positions_boites_supprimees[nom] = (position[0], position[1])
                except Exception:
                    pass

        filtres_edit_json = contenu.get("filtres_edit_data", {})
        filtre_edit_parametre_1.clear()
        filtre_edit_parametre_1.update(_filtre_edit_depuis_json(filtres_edit_json.get("parametre_1")))
        filtre_edit_parametre_2.clear()
        filtre_edit_parametre_2.update(_filtre_edit_depuis_json(filtres_edit_json.get("parametre_2")))
        filtre_edit_parametre_3.clear()
        filtre_edit_parametre_3.update(_filtre_edit_depuis_json(filtres_edit_json.get("parametre_3")))
        filtre_edit_parametre_4.clear()
        filtre_edit_parametre_4.update(_filtre_edit_depuis_json(filtres_edit_json.get("parametre_4")))

        opt_parameters.clear()
        opt_parameters.update(parametres_optimisation_depuis_json(contenu.get("parametres_optimisation")))
        graphique_optimisation_json = contenu.get("graphique_optimisation")
        if not isinstance(graphique_optimisation_json, dict):
            graphique_optimisation_json = {}
        opt_plot_data = donnees_graphique_optimisation_depuis_json(graphique_optimisation_json.get("donnees"))
        opt_affichage.clear()
        opt_affichage.update(affichage_optimisation_depuis_json(graphique_optimisation_json.get("affichage")))

        nom_projet = reglages.get("nom_projet") or "Projet importe"
        chemin_projet = chemin_fichier
        file_path = reglages.get("file_path")
        file_sheets = reglages.get("file_sheets") if isinstance(reglages.get("file_sheets"), list) else None
        Name_model = reglages.get("Name_model")
        Target_column = reglages.get("Target_column")
        Feature_column = reglages.get("Feature_column")
        Calculation_method = reglages.get("Calculation_method")
        codage_method = reglages.get("codage_method")
        Test_size = reglages.get("Test_size")
        Target_list = reglages.get("Target_list")
        header_target = reglages.get("header_target")
        dernier_x_scatter = reglages.get("dernier_x_scatter")
        dernier_y_scatter = reglages.get("dernier_y_scatter")
        dernier_color_scatter = reglages.get("dernier_color_scatter")
        dernier_methode_boxplot = reglages.get("dernier_methode_boxplot")
        dernier_x_boxplot = reglages.get("dernier_x_boxplot")
        dernier_y_boxplot = reglages.get("dernier_y_boxplot")
        dernier_color_boxplot = reglages.get("dernier_color_boxplot")
        dernier_var_histogram = reglages.get("dernier_var_histogram")
        dernier_couleur_histogram = reglages.get("dernier_couleur_histogram")
        dernier_densite_histogram = reglages.get("dernier_densite_histogram")
        dernier_x_plot3d = reglages.get("dernier_x_plot3d")
        dernier_y_plot3d = reglages.get("dernier_y_plot3d")
        dernier_z_plot3d = reglages.get("dernier_z_plot3d")
        dernier_color_plot3d = reglages.get("dernier_color_plot3d")
        dernier_target_correlation = reglages.get("dernier_target_correlation")
        dernier_features_correlation = reglages.get("dernier_features_correlation")
        dernier_dimension_min_max = reglages.get("dernier_dimension_min_max")
        dernier_methode_min_max = reglages.get("dernier_methode_min_max")
        dernier_features_min_max = reglages.get("dernier_features_min_max")

        nom_modele_a_activer = reglages.get("nom_modele_actif")
        if nom_modele_a_activer is not None and nom_modele_a_activer in modeles_projet:
            restore_modele(nom_modele_a_activer)
        else:
            nom_modele_actif = None

      
        fichiers_associes = contenu.get("fichiers_associes", {})
        dossier_du_projet = os.path.dirname(chemin_fichier)
        donnees_rechargees = False
        chemin_donnees_filtrees = fichiers_associes.get("donnees_filtrees")
        if chemin_donnees_filtrees:
            chemin_donnees_absolu = os.path.join(dossier_du_projet, chemin_donnees_filtrees)
            if os.path.isfile(chemin_donnees_absolu):
                try:
                    data_analy = pd.read_csv(chemin_donnees_absolu, sep = ";")
                    donnees_rechargees = True
                except Exception:
                    pass
        data_opt = None
        donnees_optimisation_rechargees = False
        chemin_donnees_optimisation = fichiers_associes.get("donnees_optimisation")
        if chemin_donnees_optimisation:
            chemin_donnees_optimisation_absolu = os.path.join(dossier_du_projet, chemin_donnees_optimisation)
            if os.path.isfile(chemin_donnees_optimisation_absolu):
                try:
                    data_opt = pd.DataFrame(pd.read_csv(chemin_donnees_optimisation_absolu, sep = ";", na_values = my_missing_values))
                    donnees_optimisation_rechargees = True
                except Exception:
                    data_opt = None
        ## reg memorise(s) : rend les modeles utilisables en prediction sans re-entrainement
        modeles_reg_recharges = []
        avertissements_reg = []
        for nom_modele_reg, infos_reg in (fichiers_associes.get("modeles_entraines") or {}).items():
            if nom_modele_reg not in modeles_projet or not isinstance(infos_reg, dict):
                continue
            chemin_reg = os.path.join(dossier_du_projet, infos_reg.get("fichier") or "")
            if not os.path.isfile(chemin_reg):
                avertissements_reg.append("« " + str(nom_modele_reg) + " » : fichier du modele introuvable (" + str(chemin_reg) + ")")
                continue
            try:
                import joblib, hashlib, io
                with open(chemin_reg, "rb") as f_reg:
                    contenu_reg = f_reg.read()
                empreinte_attendue = infos_reg.get("sha256")
                if empreinte_attendue and hashlib.sha256(contenu_reg).hexdigest() != empreinte_attendue:
                    avertissements_reg.append("« " + str(nom_modele_reg) + " » : fichier du modele modifie ou corrompu, il n'a pas ete charge")
                    continue
                reg_charge = joblib.load(io.BytesIO(contenu_reg))
                if not hasattr(reg_charge, "predict"):
                    avertissements_reg.append("« " + str(nom_modele_reg) + " » : le fichier ne contient pas un modele valide")
                    continue
                modeles_projet[nom_modele_reg]["reg"] = reg_charge
                if nom_modele_reg == nom_modele_actif:
                    globals()["reg"] = reg_charge
                modeles_reg_recharges.append(str(nom_modele_reg))
                version_enregistree = infos_reg.get("sklearn")
                if version_enregistree and version_enregistree != sklearn.__version__:
                    avertissements_reg.append("« " + str(nom_modele_reg) + " » a ete enregistre avec scikit-learn " + str(version_enregistree) + " (version installee : " + str(sklearn.__version__) + ") : les predictions peuvent etre erronees ; re-entrainez le modele si besoin")
            except Exception as erreur_reg:
                avertissements_reg.append("« " + str(nom_modele_reg) + " » : chargement impossible (" + str(erreur_reg) + ")")

        ## Donnees de prediction et resultats memorises
        infos_prediction_json = contenu.get("prediction", {})
        if isinstance(infos_prediction_json, dict):
            file_path_pred = infos_prediction_json.get("file_path_pred") or None
            infos_lues = infos_prediction_json.get("infos")
            prediction_infos = infos_lues if isinstance(infos_lues, dict) else {}
        for cle_fichier, nom_variable in (("donnees_prediction", "data_pred"), ("resultats_prediction", "data_pred_new")):
            chemin_relatif = fichiers_associes.get(cle_fichier)
            if chemin_relatif:
                chemin_absolu = os.path.join(dossier_du_projet, chemin_relatif)
                if os.path.isfile(chemin_absolu):
                    try:
                        tableau = pd.DataFrame(pd.read_csv(chemin_absolu, sep = ";"))
                        if nom_variable == "data_pred":
                            data_pred = tableau
                        else:
                            data_pred_new = tableau
                    except Exception:
                        pass
        nombre_graphiques = len(fichiers_associes.get("graphiques", {}))
        dossier_graphiques = fichiers_associes.get("dossier")

    except Exception as erreur:
        messagebox.showerror("Ouverture impossible", "Le fichier ne correspond pas au format attendu :\n" + str(erreur))
        return

    _maj_titre_fenetre()

  
    donnees_source_disponibles = True
    if file_path is not None:
        donnees_source_disponibles = _assurer_donnees_source_chargees()

    _historique_reinitialiser("Project opened")
    graphiques_affiches = _afficher_toutes_les_representations_graphiques(Center_frame)

    ## Graphique d'optimisation memorise dans le projet : affiche directement le panneau
    ## "Optimization" (sauf si les graphiques des modeles entraines occupent deja le panneau
    ## central : il apparait alors dans Plots > Optimization).
    graphique_optimisation_affiche = False
    if opt_plot_data is not None and not graphiques_affiches:
        try:
            my_optimization(None, Center_frame)
            graphique_optimisation_affiche = True
        except Exception as erreur:
            messagebox.showwarning("Optimization", "The saved optimization plot could not be displayed:\n" + str(erreur))

    message = "Projet « " + nom_projet + " » charge depuis :\n" + chemin_fichier + "\n\n" + str(len(modeles_projet)) + " modele(s) dans ce projet."
    if donnees_rechargees:
        message += "\n\nLe jeu de donnees filtre a ete recharge (" + str(len(data_analy)) + " lignes)."
    if donnees_optimisation_rechargees:
        message += "\n\nLes donnees d'optimisation ont ete rechargees (" + str(len(data_opt)) + " lignes)."
    if len(modeles_reg_recharges) != 0:
        message += "\n\nModele(s) rechargé(s) et utilisable(s) directement pour la prediction (bouton Log) : " + ", ".join(modeles_reg_recharges)
    if len(avertissements_reg) != 0:
        message += "\n\nAttention :\n- " + "\n- ".join(avertissements_reg)
    if nombre_graphiques != 0:
        message += "\n" + str(nombre_graphiques) + " graphique(s) enregistre(s) sous forme d'image dans : " + dossier_graphiques + "/graphiques"
    if graphiques_affiches:
        message += "\n\nLes graphiques des modeles deja entraines sont affiches (un onglet par modele)."
    if graphique_optimisation_affiche:
        message += "\n\nLe graphique d'optimisation enregistre est affiche (panneau « Optimization »)."
    elif opt_plot_data is not None:
        message += "\n\nUn graphique d'optimisation est enregistre dans ce projet : il s'affiche dans Plots > Optimization."
    message += "\n\nRemarque : seul le meilleur modele entraine est memorise pour la prediction ; les autres modeles entraines devront etre re-entraines (bouton Run) pour etre relances automatiquement ou utilises en prediction. Les resultats et graphiques deja obtenus restent consultables tels quels."
    if not donnees_source_disponibles:
        message += "\n\nAttention : le fichier de donnees d'origine n'a pas pu etre recharge. Il vous sera de nouveau demande de le resélectionner des que vous tenterez de reconfigurer un modele (bouton « Model » -> « Setting »)."
    _projet_marquer_enregistre()        # apres l'affichage des graphiques : le projet ouvert n'est pas « modifie »
    messagebox.showinfo("Projet ouvert", message)



# ==========================================================================
# ==  14. Resultats des modeles et commande Run                           ==
# ==========================================================================
## ------------------------------------------------------------


def _afficher_fenetre_resultats_modele (parent):
   
    window_model_run = Canvas(parent, relief= GROOVE ,  background = 'gray80' , highlightthickness=0)
    window_button = Canvas(parent, relief= GROOVE, background = 'gray90', highlightthickness = 1)
    my_check_model_detail = BooleanVar(value = False)
    check_bouton_model_detail = Checkbutton(window_button,  text ="Detail of Model ", variable = my_check_model_detail, command = lambda : check_model_run(), font=('arial', 15), background = 'gray90')
    check_bouton_model_detail.place(relx = 0.05, rely = 0.15, relwidth= 0.6)
    color_model_label = Label(window_button, text="Point color", font=('arial', 15), bg= "gray90")
    color_model_label.place(relx=0.05, rely= 0.23, relwidth= 0.4)
    color_value_model = ttk.Combobox(window_button, values = colors_palettes)
    color_value_model.current(3)
    color_value_model.place(relx=0.45, rely= 0.23, relwidth= 0.5)
    label_size_model = Label(window_button, text="Point size  ", font=('arial', 15), bg= "gray90")
    label_size_model.place(relx=0.05, rely= 0.31, relwidth= 0.4)
    row_edit_scatter_default_model = StringVar()
    row_edit_scatter_default_model.set("15")
    point_size_scatter_model = Spinbox(window_button, from_ = 1 , to = 100000 , increment = 1, textvariable = row_edit_scatter_default_model, command = lambda : check_model_run())
    point_size_scatter_model.place(relx = 0.45, rely= 0.31, relwidth= 0.5) 
    color_model_linear = Label(window_button, text = "Linear color", font=('arial', 15), bg= "gray90")
    color_model_linear.place(relx=0.05, rely= 0.39, relwidth= 0.4)
    color_value_linear = ttk.Combobox(window_button, values = colors_palettes)
    color_value_linear.current(4)
    color_value_linear.place(relx=0.45, rely= 0.39, relwidth= 0.5)
    label_size_linear = Label(window_button, text = "Linear size", font=('arial', 15), bg= "gray90")
    label_size_linear.place(relx=0.05, rely= 0.47, relwidth= 0.4)
    row_edit_scatter_default_linear = StringVar()
    row_edit_scatter_default_linear.set("2")
    point_size_scatter_linear = Spinbox(window_button, from_ = 1 , to = 100000 , increment = 1, textvariable = row_edit_scatter_default_linear, command = lambda : check_model_run())
    point_size_scatter_linear.place(relx = 0.45, rely= 0.47, relwidth= 0.5) 
    my_check_model_train = BooleanVar(value = False)
    my_check_model_train_data = Checkbutton(window_button,  text ="Train data", variable = my_check_model_train, command = lambda : check_model_run(), font=('arial', 15), background = 'gray90')
    my_check_model_train_data.place(relx = 0.05, rely = 0.55, relwidth= 0.45)
    label_train = Label(window_button, text="Train color  ", font=('arial', 15), bg= "gray90")
    label_train.place(relx=0.05, rely= 0.63, relwidth= 0.4)
    color_value_train = ttk.Combobox(window_button, values = colors_palettes)
    color_value_train.current(17)
    color_value_train.place(relx=0.45, rely= 0.63, relwidth= 0.5)
    label_size_train = Label(window_button, text="Train size    ", font=('arial', 15), bg= "gray90")
    label_size_train.place(relx=0.05, rely= 0.71, relwidth= 0.4)
    row_edit_scatter_default_train = StringVar()
    row_edit_scatter_default_train.set("15")
    row_edit_scatter_train = Spinbox(window_button, from_ = 1 , to = 100000 , increment = 1, textvariable = row_edit_scatter_default_train, command = lambda : check_model_run())
    row_edit_scatter_train.place(relx = 0.45, rely= 0.71, relwidth= 0.5) 

    def _redessiner_resultats_modele():
        check_model_run(window_model_run, color_value_model, row_edit_scatter_default_model, color_value_linear, point_size_scatter_linear, color_value_train, row_edit_scatter_default_linear, row_edit_scatter_default_train, my_check_model_detail, my_check_model_train)

    check_bouton_model_detail.configure(command = _redessiner_resultats_modele)
    point_size_scatter_model.configure(command = _redessiner_resultats_modele)
    point_size_scatter_linear.configure(command = _redessiner_resultats_modele)
    my_check_model_train_data.configure(command = _redessiner_resultats_modele)
    row_edit_scatter_train.configure(command = _redessiner_resultats_modele)
    color_value_model.bind("<<ComboboxSelected>>", lambda event: _redessiner_resultats_modele())
    color_value_linear.bind("<<ComboboxSelected>>", lambda event: _redessiner_resultats_modele())
    color_value_train.bind("<<ComboboxSelected>>", lambda event: _redessiner_resultats_modele())
    save_graph_p = PhotoImage(file = my_path('figures/save_graph_p.png'))
    image_box_view = PhotoImage(file=my_path('figures/apercu_p.png'))
    bouton_plot_model_view = Button(window_button, command = _redessiner_resultats_modele, image = image_box_view , font=('arial', 14), text = "View plot",  compound = tk.TOP,  borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
    bouton_plot_model_view.place(relx= 0.27, rely= 0.03, width= 100, height= 60)
    bouton_save_model = Button(window_button, image = save_graph_p, command = lambda : my_save_plot_file(model_plot) , font=('arial', 14), text = "Save plot",  compound = tk.TOP,  borderwidth = 0, relief = FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
    bouton_save_model.place(relx= 0.27, rely= 0.87, width= 100, height= 60)
    window_model_run.place(relx=0, rely=0, relwidth= 0.80, relheight= 1)
    window_button.place(relx=0.80, rely = 0, relwidth= 0.2, relheight = 1)

    return _redessiner_resultats_modele


def _afficher_toutes_les_representations_graphiques (parent, modele_a_selectionner = None):
   

    modeles_entraines_noms = [nom for nom, etat in modeles_projet.items() if etat.get("data_model") is not None]

    if len(modeles_entraines_noms) == 0:
        return False

    notebook_modeles = ttk.Notebook(parent)
    notebook_modeles.place(relx = 0, rely = 0, relwidth = 1, relheight = 1)

    onglets_par_nom = {}
    redessiner_par_nom = {}
    onglet_a_selectionner = None
    for nom in modeles_entraines_noms:
        onglet = tk.Frame(notebook_modeles, background = "white")
        notebook_modeles.add(onglet, text = str(nom))
        onglets_par_nom[str(onglet)] = nom
        if nom == modele_a_selectionner:
            onglet_a_selectionner = onglet
        restore_modele(nom)
        redessiner_par_nom[nom] = _afficher_fenetre_resultats_modele(onglet)

    def _au_changement_onglet (event):
       
        nom = onglets_par_nom.get(notebook_modeles.select())
        if nom is not None:
            restore_modele(nom)
            if nom in redessiner_par_nom:
                redessiner_par_nom[nom]()

    notebook_modeles.bind("<<NotebookTabChanged>>", _au_changement_onglet)

    ## ---- Barre verticale gauche : zoom sur le graphique du modele affiche ----
    def _curseur_vers_zoom(valeur_curseur):
        return ZOOM_MIN_GRAPHIQUE + (float(valeur_curseur) / 100) * (ZOOM_MAX_GRAPHIQUE - ZOOM_MIN_GRAPHIQUE)

    def _zoom_vers_curseur(valeur_zoom):
        valeur = (valeur_zoom - ZOOM_MIN_GRAPHIQUE) / (ZOOM_MAX_GRAPHIQUE - ZOOM_MIN_GRAPHIQUE) * 100
        if valeur < 0:
            valeur = 0
        elif valeur > 100:
            valeur = 100
        return int(valeur)

    def _sur_changement_zoom(valeur_curseur):
        global zoom_graphique_modeles
        zoom_graphique_modeles = _curseur_vers_zoom(valeur_curseur)
        nom_actif = onglets_par_nom.get(notebook_modeles.select())
        if nom_actif is not None and nom_actif in redessiner_par_nom:
            redessiner_par_nom[nom_actif]()

    scale.configure(command = _sur_changement_zoom)
    scale.set(_zoom_vers_curseur(zoom_graphique_modeles))

    if onglet_a_selectionner is not None:
        notebook_modeles.select(onglet_a_selectionner)
        restore_modele(modele_a_selectionner)
    else:
        restore_modele(modeles_entraines_noms[0])

    return True


def _signaler_erreurs_execution (ignores):
    """CORRECTIF : les erreurs d'entrainement des modeles rejoues automatiquement etaient
    collectees puis jetees (message_run_tous n'etait jamais affiche) : Run semblait ne rien faire."""
    ## CORRECTIF : seuls les modeles en erreur etaient signales ; les autres modeles non entraines
    ## (pas encore configures, algorithme inconnu...) etaient ignores sans aucun message.
    if len(ignores) != 0:
        messagebox.showwarning("Run", "Certains modeles n'ont pas ete entraines :\n- " + "\n- ".join(ignores))


def _afficher_panneau_model_apres_run (parent):
    """ Affiche dans le panneau central le panneau du bouton « Model » (liste des
    modeles et leur etat : configure, entraine...). """
    try:
        choice_model(None, parent)
    except Exception:
        pass


def run_linear_model (event, parent, target_var,  methode_var, test_var, list_vart, check_target, loss_var, test_rate, check_codage):

    if not analyse_autorisee():      # bouton desactive : un evenement lie a la main peut quand meme arriver
        return

    if not verifier_projet_ouvert():
        return

    global file_path, color_value_model, point_size_scatter_model, result_model, color_value_linear, point_size_scatter_linear, model_plot, result_run, selectet_target, selectet_codage, my_data_coef, selectet_methode_imp, data_model, my_y_min, my_y_max, selectet_methode, size_data_value
    file_path = file_path
    data_model = None
    if file_path is not None:

        ## CORRECTIF : nouveau projet -> data_analy vaut None tant que « Edit data » n'a pas ete utilise
        if not _assurer_data_analy():
            return

        ## CORRECTIF (plusieurs modeles) : le panneau de configuration du modele ACTIF n'est utilise que s'il est
        ## COMPLET. Avant, quand le modele actif avait son panneau ouvert mais incomplet (modele que l'on vient de
        ## creer ou d'ouvrir : cible non choisie, variables non choisies...), « Run » ne faisait RIEN, sans message
        ## (ou plantait sur len(None)) et les AUTRES modeles, pourtant configures, n'etaient pas entraines non plus.
        raison_panneau = None           # pourquoi le panneau ouvert du modele actif ne peut pas etre utilise
        panneau_complet = False
        try:
            if _panneau_config.get("nom") == nom_modele_actif and _widget_existe(_panneau_config.get("param")):
                target_index = target_var.current()
                selection_features = globals().get("feature_select")
                if target_index <= 0:
                    raison_panneau = "colonne cible non choisie"
                elif test_var is None:
                    raison_panneau = "methode de calcul / taille de test non definie"
                elif selection_features is None or selection_features.current() <= 0:
                    raison_panneau = "colonnes explicatives (features) non choisies"
                elif selection_features.current() == 1 and len(globals().get("result_model_1") or []) == 0:
                    raison_panneau = "aucune colonne explicative cochee dans la liste"
                elif Name_model is None or Name_model == "Choose model":
                    raison_panneau = "algorithme non choisi"
                else:
                    panneau_complet = True
        except Exception:
            panneau_complet = False     # widgets detruits : on s'appuie sur les configurations enregistrees

        if not panneau_complet:
            ## tous les modeles sont entraines a partir de leur configuration ENREGISTREE (« Save configure »)
            _planifier_progression_run(None)
            reussis_autres, ignores_autres = _executer_autres_modeles_du_projet(None)
            _terminer_indicateur_execution()        # entrainement termine : l'encadre disparait tout de suite
            if raison_panneau is not None and nom_modele_actif not in reussis_autres:
                ignores_autres = [texte for texte in ignores_autres if not texte.startswith(str(nom_modele_actif) + " (")]
                ignores_autres.append(str(nom_modele_actif) + " (configuration en cours incomplete : " + raison_panneau +
                                      " ; terminez-la puis cliquez sur « Save configure »)")

            ## fin de « Run » : le panneau central affiche le panneau du bouton « Model »
            ## (les graphiques des modeles entraines sont accessibles depuis « Log » > Model plots)
            affichage_ok = len(reussis_autres) != 0
            _afficher_panneau_model_apres_run(parent)
            if not affichage_ok:
                if len(ignores_autres) != 0:
                    messagebox.showwarning("Run", "Aucun modele n'a pu etre entraine :\n- " + "\n- ".join(ignores_autres))
                else:
                    messagebox.showinfo("Run", "Aucun modele a executer. Creez un modele (bouton « Analysis »), configurez-le puis cliquez sur « Save configure ».")
            else:
                _signaler_erreurs_execution(ignores_autres)
                _historique_valider("Run")
            return

        ## panneau complet : le modele actif est entraine avec les valeurs du panneau, puis tous les autres
        selectet_target = check_target[target_index]
        if feature_select.current() == 1:
            result_model = list(result_model_1)
        else:
            d_X = data_analy.copy()
            d_X = d_X.drop(selectet_target, axis = 1)
            result_model = np.array(d_X.columns)
        erreur_modele_actif = None
        try:
            size_data_value = float(test_var.get())
        except Exception:
            size_data_value = 0.20      ## widget detruit / valeur illisible : valeur par defaut
        method_list = ["TrainTestSplit data", "Cross-validation"]
        methode_index = methode_var.current()
        selectet_methode = method_list[methode_index]
        selectet_methode_imp = selectet_methode
        method_recode = ["None", "Disjunctive", "All-disjunctive"]
        codage_index = check_codage.current()
        selectet_codage = method_recode[codage_index]
        loss_list = ["squared_error", "absolute_error", "huber", "quantile"]
        _planifier_progression_run(nom_modele_actif, dict(modeles_projet.get(nom_modele_actif, {}), Name_model = Name_model,
                                   selectet_target = selectet_target, result_model = list(result_model), selectet_codage = selectet_codage,
                                   selectet_methode = selectet_methode, size_data_value = size_data_value,
                                   result_model_recod = globals().get("result_model_recod"), reg = _reg_panneau_sans_erreur(Name_model),
                                   duree_entrainement = None if _panneau_config.get("reconfiguration") or modeles_projet.get(nom_modele_actif, {}).get("Name_model") != Name_model
                                   else modeles_projet.get(nom_modele_actif, {}).get("duree_entrainement")))
        data_XX = data_analy[result_model]
        Y = data_analy[selectet_target]
        my_y_min = Y.min() - 5
        my_y_max = Y.max() + 5
        try:
            ## CORRECTIF : loss_var (widget « loss ») n'existe que pour le GBM ; pour les autres
            ## algorithmes il pouvait pointer vers un widget detruit d'un modele precedent.
            if loss_var is not None and Name_model == "Generalized Boosted Regression Modeling (GBM)":
                loss_index = loss_var.current()
                selectet_loss = loss_list[loss_index]
                data_model = regression_model(data_XX, Y, k = 5, methode = selectet_methode, value_test = size_data_value, loss_var =  selectet_loss)
            else:
                data_model = regression_model(data_XX, Y, k = 5, methode = selectet_methode, value_test = size_data_value)
        except Exception as erreur_entrainement:
            ## CORRECTIF (plusieurs modeles) : l'echec du modele actif arretait tout (return) : les autres
            ## modeles n'etaient pas entraines. On le signale et on continue avec les autres.
            data_model = None
            erreur_modele_actif = str(erreur_entrainement)
            restore_modele(nom_modele_actif)          # le modele actif garde sa configuration enregistree
            _progression_unite_terminee()
        if erreur_modele_actif is None:
            if selectet_methode_imp != "Cross-validation":
                my_data_coef = data_model[6]
            else:
                my_data_coef = None
            snapshot_modele_actif()
            modeles_projet[nom_modele_actif]["signature_entrainement"] = _signature_entrainement(modeles_projet[nom_modele_actif])
            modeles_projet[nom_modele_actif]["duree_entrainement"] = _progression_unite_terminee()
        reussis_autres, ignores_autres = _executer_autres_modeles_du_projet(nom_modele_actif)
        _terminer_indicateur_execution()            # entrainement termine : l'encadre disparait tout de suite
        if erreur_modele_actif is not None:
            ignores_autres = [str(nom_modele_actif) + " (erreur : " + erreur_modele_actif + ")"] + list(ignores_autres)

        if erreur_modele_actif is None:
            _fermer_panneau_config()     ## la configuration est deja memorisee par snapshot_modele_actif()
        _afficher_panneau_model_apres_run(parent)       # panneau du bouton « Model »
        if erreur_modele_actif is not None and len(reussis_autres) == 0:
            messagebox.showwarning("Run", "Aucun modele n'a pu etre entraine :\n- " + "\n- ".join(ignores_autres))
        else:
            _signaler_erreurs_execution(ignores_autres)
            _historique_valider("Run")
    else:
        warning_inf = Label(parent, text = traduire("Please create a template after loading the data"), font=('Helvetica 16 bold italic'), fg = "red", background= "white")
        warning_inf.place(relx = 0, rely= 0, relwidth=1, relheight= 1)


label_Test_size = None

def data_size (event, parent, var_size, valeur_initiale = None):

    global Test_size, label_Test_size
    ## CORRECTIF : on retire l'ancien libelle / champ avant d'en creer de nouveaux (ils s'empilaient),
    ## et le champ peut desormais etre cree des l'ouverture du panneau (voir ConfirmWindow).
    for ancien_widget in (Test_size, label_Test_size):
        if ancien_widget is not None and hasattr(ancien_widget, "destroy"):
            try:
                ancien_widget.destroy()
            except Exception:
                pass
    if  var_size.current() != 1:
        default_value= StringVar()
        default_value.set(str(valeur_initiale) if valeur_initiale is not None else "0.2")
        label_Test_size = Label(parent, text="Test size          ", font=('arial', 15), fg = "gray30", bg= "white")
        label_Test_size.place(relx=0.01, rely= 0.25, relwidth= 0.4)
        Test_size = Spinbox(parent, from_ = 0 , to = 1 , increment = 0.1, textvariable = default_value)
        Test_size.place(relx=0.47, rely= 0.25, relwidth= 0.5) 
    else:
        default_value= StringVar()
        default_value.set("10")
        label_Test_size = Label(parent, text="Split size         ", font=('arial', 15), fg = "gray30", bg= "white")
        label_Test_size.place(relx=0.01, rely= 0.25, relwidth= 0.4)
        Test_size = Spinbox(parent, from_ = 0 , to = 10000 , increment = 1, textvariable = default_value)
        Test_size.place(relx=0.47, rely= 0.25, relwidth= 0.5) 


def variable_recoding_function (event, parent, check_recod, recod_list, var_recod):

    global recoding_param
    recod_index = check_recod.current()
    recod_value = recod_list[recod_index]
    if recod_value == "Disjunctive":
       window_model_recod = Canvas(parent, relief= GROOVE ,  background = 'cadetblue' , highlightthickness=0)
       recoding_param = tk.Listbox(window_model_recod, fg = "gray10", font = ('arial', 12), relief = GROOVE , selectmode = tk.MULTIPLE, exportselection = False)
       recoding_param.place(relx= 0.02, rely = 0.02, relwidth = 0.96, relheight = 0.85)
       for re_value in result_model_1:
           recoding_param.insert(END, re_value)     
       close_window_recod = Button(window_model_recod, text = "Save", command = lambda : recup(window_model_recod) , font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
       close_window_recod.place(relx= 0.4, rely= 0.9, width= 50, height= 20)
       window_model_recod.place(relx = 0.1, rely = 0.1, relwidth = 0.8, relheight = 0.6)



# ==========================================================================
# ==  15. Prediction (bouton Log)                                         ==
# ==========================================================================
def destroy_window_pred (parent, Feature_categorial_check, value_check):

    global result_recod_stat_pred, data_sort_stat_pred
    result_recod_stat_pred = [Feature_categorial_check.get(i) for i in Feature_categorial_check.curselection()]
    '''if len(result_recod_stat_pred) > 0:
        data_sort_stat = []
        for i in range(len(result_recod_stat_pred)):
            data_sort_stat.append(sort(unique((data_analy[result_recod_stat[i]]))))'''
    parent.destroy()


_graphiques_log = {"notebook": None, "case": None, "bande": None, "place_bande": None, "panneaux_log": []}
## panneaux_log : les deux zones de l'interface principale du bouton « Log » (a detruire avant de la reconstruire)
## notebook : onglets des graphiques affiches depuis le panneau « Log » ; case : copie de la case
## « Model plots » ; bande : colonne de droite du panneau « Log », cachee pendant l'affichage


def _cacher_bande_log (bande):
    """ Cache la colonne de droite du panneau « Log » en memorisant sa position. """
    if bande is None:
        return
    try:
        if bande.winfo_exists() and bande.winfo_manager() == "place":
            _graphiques_log["place_bande"] = {cle: valeur for cle, valeur in bande.place_info().items() if cle != "in"}
            bande.place_forget()
        _graphiques_log["bande"] = bande
    except Exception:
        pass


def _reafficher_bande_log ():
    """ Remet la colonne de droite du panneau « Log » a sa position d'origine. """
    bande, position = _graphiques_log["bande"], _graphiques_log["place_bande"]
    _graphiques_log["bande"] = None
    _graphiques_log["place_bande"] = None
    if bande is None or position is None:
        return
    try:
        if bande.winfo_exists():
            bande.place(**position)
            bande.lift()
    except Exception:
        pass


def _masquer_graphiques_modeles_log ():
    for cle in ("notebook", "case"):
        widget = _graphiques_log[cle]
        if widget is not None:
            try:
                if widget.winfo_exists():
                    widget.destroy()
            except Exception:
                pass
        _graphiques_log[cle] = None
    _reafficher_bande_log()


def afficher_graphiques_modeles_log (panneau, case_cochee = None, bande = None):
    """ Panneau « Log » : affiche sur TOUT le panneau central ("panneau") les
    representations graphiques de tous les modeles entraines (un onglet par modele),
    comme apres un « Run ». Une copie de la case « Model plots » reste visible en haut
    a droite (dans la barre des onglets) pour revenir au panneau « Log ». """

    _masquer_graphiques_modeles_log()
    enfants_avant = set(str(w) for w in panneau.winfo_children())
    if not _afficher_toutes_les_representations_graphiques(panneau):
        messagebox.showinfo("Model plots", "No trained model yet: click « Run » to train the models.")
        return
    nouveaux = [w for w in panneau.winfo_children() if str(w) not in enfants_avant]
    if len(nouveaux) == 0:
        return
    _cacher_bande_log(bande)                # la colonne contenant « Model plots » est cachee
    _graphiques_log["notebook"] = nouveaux[-1]
    nouveaux[-1].lift()
    if case_cochee is not None:
        case = Checkbutton(panneau, text = "Model plots", variable = case_cochee, command = lambda : basculer_graphiques_modeles_log(panneau, case_cochee, bande), fg = "gray30", font=('arial', 15), background = 'gray90')
        case.place(relx = 1.0, x = -8, y = 2, anchor = "ne")
        case.lift()
        _graphiques_log["case"] = case


def basculer_graphiques_modeles_log (panneau, case_cochee, bande = None):
    """ Case « Model plots » du panneau « Log » : cochee -> graphiques des modeles
    entraines sur tout le panneau central ; decochee -> retour a l'interface principale du bouton « Log ». """
    if case_cochee.get():
        afficher_graphiques_modeles_log(panneau, case_cochee, bande)
        if _graphiques_log["notebook"] is None:      # aucun modele entraine : on decoche
            case_cochee.set(False)
    else:
        _masquer_graphiques_modeles_log()
        ## EVOLUTION : retour a l'interface principale du bouton « Log », reconstruite comme par un clic sur
        ## « Log » (colonne de droite, donnees et resultats de prediction, curseur de zoom qui ne pilote plus
        ## les graphiques des modeles). Differe : la case decochee (la copie affichee sur les graphiques) vient
        ## d'etre detruite pendant son propre clic.
        try:
            SmartDigicat.after_idle(lambda: _revenir_interface_log(panneau))
        except Exception:
            _revenir_interface_log(panneau)


def _revenir_interface_log (panneau):
    """ Reaffiche l'interface principale du bouton « Log » dans le panneau central. """
    for widget in _graphiques_log.get("panneaux_log", []):
        try:
            if widget.winfo_exists():
                widget.destroy()
        except Exception:
            pass
    _graphiques_log["panneaux_log"] = []
    try:
        scale.configure(command = "")        # le curseur de zoom ne pilote plus les graphiques des modeles
    except Exception:
        pass
    try:
        my_predict_function(None, panneau)
    except Exception as erreur:
        messagebox.showwarning("Log", "The « Log » panel could not be displayed again:\n" + str(erreur))


def my_predict_function (event, parent):

    if not analyse_autorisee():      # bouton desactive : un evenement lie a la main peut quand meme arriver
        return

    if not verifier_projet_ouvert():
        return

    global nom_modele_actif
    modeles_entraines = [nom for nom in modeles_projet if modeles_projet[nom].get("data_model") is not None]
    if len(modeles_entraines) > 0:
        if nom_modele_actif not in modeles_entraines:
            _capturer_configuration_widgets_actifs()
            restore_modele(modeles_entraines[0])
        window_model_pred = Canvas(parent, relief= GROOVE ,  background = 'gray100' , highlightthickness=0)
        window_param_pred = Canvas(parent, relief= GROOVE, background = 'gray90', highlightthickness = 1)
        _graphiques_log["panneaux_log"] = [window_model_pred, window_param_pred]     # (retour depuis « Model plots »)

        def my_carbon_balance_pred (parent, feature_list, check_carbon_balance):
            carbon_balance_index = check_carbon_balance.current()
            if carbon_balance_index != 0:
                carbon_balance_selected = feature_list[carbon_balance_index]
                height_window_carbon_balance = Canvas(parent, relief= GROOVE ,  background = 'cadetblue' , highlightthickness=0)
                height_window_carbon_balance.place(relx=0.2, rely= 0.1, relwidth= 0.5, relheight = 0.7)
                label_arbon_balance = Label(height_window_carbon_balance, text="Window to filter from Parameter - 1", font=('arial', 15), fg = "gray10", background= "cadetblue")
                label_arbon_balance.place(relx= 0.2, rely= 0.12, relwidth= 0.6)
                label_Carbon_balance_min = Label(height_window_carbon_balance, text="Min", font=('arial', 15), fg = "gray10", background= "cadetblue")
                label_Carbon_balance_min.place(relx=0.15, rely= 0.4, relwidth= 0.2)
                label_Carbon_balance_max = Label(height_window_carbon_balance, text="Max", font=('arial', 15), fg = "gray10", background= "cadetblue")
                label_Carbon_balance_max.place(relx=0.15, rely= 0.6, relwidth= 0.2)
                carbon_min_default = DoubleVar()
                carbon_min_default.set(0)
                carbon_min_entry = Entry(height_window_carbon_balance, textvariable = carbon_min_default, justify = CENTER )
                carbon_min_entry.place(relx=0.46, rely = 0.4, width = 150,  height = 25)
                carbon_max_default = DoubleVar()
                carbon_max_default.set(100)
                carbon_max_entry = Entry(height_window_carbon_balance, textvariable = carbon_max_default, justify = CENTER )
                carbon_max_entry.place(relx=0.46, rely = 0.6, width = 150,  height = 25)
                destroy_window_carbon = Button(height_window_carbon_balance, text = "Save", command = lambda : destroy_my_carbon_balance_pred(height_window_carbon_balance, carbon_min_default, carbon_max_default, carbon_balance_selected) , font=('arial', 15), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
                destroy_window_carbon.place(relx= 0.46, rely= 0.85, width= 50, height= 20)  
                _entree_declenche_bouton(destroy_window_carbon)      # touche Entree = Save

        def destroy_my_carbon_balance_pred (parent, check_min_carbon, check_max_corbon, check_value_corbon):
              global data_pred, carbon_balance_min_value
              carbon_balance_min_value = check_min_carbon.get()
              carbon_balance_max_value = check_max_corbon.get()
              if carbon_balance_min_value < carbon_balance_max_value:
                   data_pred = data_pred[data_pred[check_value_corbon] > carbon_balance_min_value]
              if carbon_balance_max_value > carbon_balance_min_value:
                    data_pred = data_pred[data_pred[check_value_corbon] <= carbon_balance_max_value]
              parent.destroy()

        def my_conversion_pred (parent, feature_list, check_conversion_filter):
            conversion_index = check_conversion_filter.current()
            if conversion_index != 0:
                conversion_selected = feature_list[conversion_index]
                height_window_conversion = Canvas(parent, relief= GROOVE ,  background = 'cadetblue' , highlightthickness=0)
                height_window_conversion.place(relx=0.2, rely= 0.1, relwidth= 0.5, relheight = 0.7)
                label_conversion = Label(height_window_conversion, text="Window to filter from Parameter - 2", font=('arial', 15), fg = "gray10", background= "cadetblue")
                label_conversion.place(relx= 0.2, rely= 0.12, relwidth= 0.6)
                label_conversion_min = Label(height_window_conversion, text="Min", font=('arial', 15), fg = "gray10", background= "cadetblue")
                label_conversion_min.place(relx=0.15, rely= 0.4, relwidth= 0.2)
                label_conversion_max = Label(height_window_conversion, text="Max", font=('arial', 15), fg = "gray10", background= "cadetblue")
                label_conversion_max.place(relx=0.15, rely= 0.6, relwidth= 0.2)
                conversion_min_default = DoubleVar()
                conversion_min_default.set(0)
                conversion_min_entry = Entry(height_window_conversion, textvariable = conversion_min_default, justify = CENTER )
                conversion_min_entry.place(relx=0.46, rely = 0.4, width = 150,  height = 25)
                conversion_max_default =DoubleVar()
                conversion_max_default.set(100)
                conversion_max_entry = Entry(height_window_conversion, textvariable = conversion_max_default, justify = CENTER )
                conversion_max_entry.place(relx=0.46, rely = 0.6, width = 150,  height = 25)
                destroy_window_conversion = Button(height_window_conversion, text = "Save", command = lambda : destroy_my_conversion_pred(height_window_conversion, conversion_min_default, conversion_max_default, conversion_selected) , font=('arial', 15), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
                destroy_window_conversion.place(relx= 0.46, rely= 0.85, width= 50, height= 20)        
                _entree_declenche_bouton(destroy_window_conversion)      # touche Entree = Save

        def destroy_my_conversion_pred (parent, check_min_conv, check_max_conv, check_value_conversion):
              global data_pred
              conversion_min_value = check_min_conv.get()
              conversion_max_value = check_max_conv.get()
              if conversion_min_value < conversion_max_value:
                   data_pred = data_pred[data_pred[check_value_conversion] > conversion_min_value]
              if conversion_max_value > conversion_min_value:
                    data_pred = data_pred[data_pred[check_value_conversion] <= conversion_max_value]
              parent.destroy()

        def my_yields_pred (parent, feature_list, check_yields_filter):
            yields_index = check_yields_filter.current()
            if yields_index != 0:
                yields_selected = feature_list[yields_index]
                height_window_yields = Canvas(parent, relief= GROOVE ,  background = 'cadetblue' , highlightthickness=0)
                height_window_yields.place(relx=0.2, rely= 0.1, relwidth= 0.5, relheight = 0.7)
                label_yields = Label(height_window_yields, text="Window to filter from Parameter - 3", font=('arial', 15), fg = "gray10", background= "cadetblue")
                label_yields.place(relx= 0.2, rely= 0.12, relwidth= 0.6)
                label_yields_min = Label(height_window_yields, text="Min", font=('arial', 15), fg = "gray10", background= "cadetblue")
                label_yields_min.place(relx=0.15, rely= 0.4, relwidth= 0.2)
                label_yields_max = Label(height_window_yields, text="Max", font=('arial', 15), fg = "gray10", background= "cadetblue")
                label_yields_max.place(relx=0.15, rely= 0.6, relwidth= 0.2)
                yields_min_default = DoubleVar()
                yields_min_default.set(0)
                yields_min_entry = Entry(height_window_yields, textvariable = yields_min_default, justify = CENTER )
                yields_min_entry.place(relx=0.46, rely = 0.4, width = 150,  height = 25)
                yields_max_default = DoubleVar()
                yields_max_default.set(100)
                yields_max_entry = Entry(height_window_yields, textvariable = yields_max_default, justify = CENTER )
                yields_max_entry.place(relx=0.46, rely = 0.6, width = 150,  height = 25)
                destroy_window_yields = Button(height_window_yields, text = "Save", command = lambda : destroy_my_yields_pred(height_window_yields, yields_min_default, yields_max_default, yields_selected) , font=('arial', 15), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
                destroy_window_yields.place(relx= 0.46, rely= 0.85, width= 50, height= 20)        
                _entree_declenche_bouton(destroy_window_yields)      # touche Entree = Save

        def destroy_my_yields_pred (parent, check_min_yied, check_max_yied, check_value_yield):
              global data_pred
              yields_min_value = check_min_yied.get()
              yields_max_value = check_max_yied.get()
              if yields_min_value < yields_max_value:
                   data_pred = data_pred[data_pred[check_value_yield] > yields_min_value]
              if yields_max_value > yields_min_value:
                   data_pred = data_pred[data_pred[check_value_yield] <= yields_max_value]
              parent.destroy()

        def my_selectivity_pred (parent,  feature_list, check_selectivity_filter):
            selectivity_index = check_selectivity_filter.current()
            if selectivity_index != 0:
                selectivity_selected = feature_list[selectivity_index]
                height_window_selectivity = Canvas(parent, relief= GROOVE ,  background = 'cadetblue' , highlightthickness=0)
                height_window_selectivity.place(relx=0.2, rely= 0.1, relwidth= 0.5, relheight = 0.7)
                label_selectivity = Label(height_window_selectivity, text="Window to filter from Parameter - 4", font=('arial', 15), fg = "gray10", background= "cadetblue")
                label_selectivity.place(relx= 0.2, rely= 0.12, relwidth= 0.6)
                label_selectivity_min = Label(height_window_selectivity, text="Min", font=('arial', 15), fg = "gray10", background= "cadetblue")
                label_selectivity_min.place(relx=0.15, rely= 0.4, relwidth= 0.2)
                label_selectivity_max = Label(height_window_selectivity, text="Max", font=('arial', 15), fg = "gray10", background= "cadetblue")
                label_selectivity_max.place(relx=0.15, rely= 0.6, relwidth= 0.2)
                selectivity_min_default = DoubleVar()
                selectivity_min_default.set(0)
                selectivity_min_entry = Entry(height_window_selectivity, textvariable = selectivity_min_default, justify = CENTER )
                selectivity_min_entry.place(relx=0.46, rely = 0.4, width = 150,  height = 25)
                selectivity_max_default = DoubleVar()
                selectivity_max_default.set(100)
                selectivity_max_entry = Entry(height_window_selectivity, textvariable = selectivity_max_default, justify = CENTER )
                selectivity_max_entry.place(relx=0.46, rely = 0.6, width = 150,  height = 25)
                destroy_window_selectivity = Button(height_window_selectivity, text = "Save", command = lambda : destroy_my_selectivity_pred(height_window_selectivity, selectivity_min_default, selectivity_max_default, selectivity_selected) , font=('arial', 15), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
                destroy_window_selectivity.place(relx= 0.46, rely= 0.85, width= 50, height= 20)        
                _entree_declenche_bouton(destroy_window_selectivity)      # touche Entree = Save

        def destroy_my_selectivity_pred (parent, check_min_select, check_max_select, check_value_selct):
              global data_pred
              selectivity_min_value = check_min_select.get()
              selectivity_max_value = check_max_select.get()
              if selectivity_min_value < selectivity_max_value:
                    data_pred = data_pred[data_pred[check_value_selct] > selectivity_min_value]
              if selectivity_max_value > selectivity_min_value:
                    data_pred = data_pred[data_pred[check_value_selct] <= selectivity_max_value]
              parent.destroy()

        def window_prediction (check_parent, check_parent_button):
            global label_data_pred2, rapport_missing_value_all_pred, change_modality_pred, data_pred, categorial_var_def_pred
            if data_pred is not None or file_path_pred:
                ## CORRECTIF / EVOLUTION : les donnees memorisees sont reutilisees (avant, le fichier etait relu a
                ## chaque « Show all » : modifications perdues, et impossible si le fichier d'origine avait disparu).
                if data_pred is None:
                    try:
                        data_pred, _ = _lire_donnees_prediction(file_path_pred, file_sheets_pred)
                    except Exception as erreur_lecture_pred:
                        messagebox.showerror("Lecture impossible", "Impossible de lire le fichier de prediction :\n" + str(erreur_lecture_pred))
                        return
                ## CORRECTIF : header_clean_value(_all) n'existaient qu'apres un passage par « Edit data » (NameError
                ## dans un nouveau projet) ; on les construit a partir des donnees de prediction elles-memes.
                header_clean_value_all = ["Select column name"] + [c for c in data_pred.columns]
                header_clean_value = ["Select column name"] + [c for c in data_pred.columns if pd.api.types.is_numeric_dtype(data_pred[c])]
                window_model_pred_info = Canvas(check_parent, relief= GROOVE ,  background = 'gray100' , highlightthickness=0)
                low_window_pred = Canvas(check_parent, relief= GROOVE , background = 'gray100', highlightthickness=0)
                separator_view_data1_pred = tk.Frame(window_model_pred_info, bd = 10, relief = 'sunken', width = 4)
                separator_view_data1_pred.place(relx=0.37 , rely = 0.05 , relheight= 0.8)
                separator_view_data2_pred = tk.Frame(window_model_pred_info, bd = 10, relief = 'sunken', width = 4)
                separator_view_data2_pred.place(relx=0.64 , rely = 0.05 , relheight= 0.8)
                ## ********************************************
                label_missing_value_pred = Label(window_model_pred_info, text="Missing value imputation", font=('arial', 15), fg = "saddlebrown", background= "white")
                label_missing_value_pred.place(relx=0.05, rely= 0.03, relwidth= 0.25)
                list_quanti = ["select method" , "mean", "median", "delete"]
                label_quanti_pred = Label(window_model_pred_info, text="Quantitative  ", font=('arial', 14), fg = "gray30", background= "white")
                label_quanti_pred.place(relx=0.05, rely= 0.17, relwidth= 0.1)
                imputation_quanti_pred = ttk.Combobox(window_model_pred_info, values = list_quanti, width= 10)
                imputation_quanti_pred.place(relx=0.05, rely= 0.23, relwidth= 0.25)
                imputation_quanti_pred.current(0)
                list_quali = ["select method" , "modality reference", "delete"]
                label_quali_pred = Label(window_model_pred_info, text="Qualitative  ", font=('arial', 14), fg = "gray30", background= "white")
                label_quali_pred.place(relx=0.05, rely= 0.37, relwidth= 0.1)
                imputation_quali_pred = ttk.Combobox(window_model_pred_info, values = list_quali, width= 10)
                imputation_quali_pred.place(relx=0.05, rely= 0.43, relwidth= 0.25)
                imputation_quali_pred.current(0)
                Categorizing_variables = header_clean_value_all.copy()
                categorial_var_def_pred  = BooleanVar(value = False)
                categorial_var_pred = Checkbutton(window_model_pred_info,  text ="Categorizing variables", variable = categorial_var_def_pred, command = lambda : my_categorial_function(window_model_pred_info, Categorizing_variables, categorial_var_def_pred, Edit_data_pred_new) , fg = "gray30", font=('arial', 15), background = 'gray100')
                categorial_var_pred.place(relx=0.03, rely= 0.53, relwidth= 0.22)
                model_detail_def_pred  = BooleanVar(value = False)
                model_detail_pred = Checkbutton(window_model_pred_info,  text ="Detail of model           ", variable = model_detail_def_pred, command = lambda : my_model_detail_function(window_model_pred_info, model_detail_def_pred) , fg = "gray30", font=('arial', 15), background = 'gray100')
                model_detail_pred.place(relx=0.7, rely= 0.8, relwidth= 0.22)
                ##window_model_plot
                #view_info_data_all_pred = tk.Label(window_model_pred_info, text = "", font=('arial', 15), fg = "gray30", bg = "gray100")
                #view_info_data_all_pred.place(relx=0.01, rely= 0.94, relwidth= 0.4)
                label_data_pred2 = Label(window_model_pred_info, text = "" , font = ('arial', 15), fg = "gray30", bg = "gray100")
                label_data_pred2.place(relx=0.01, rely= 0.94, relwidth= 0.4)
                rapport_missing_value_all_pred = Label(window_model_pred_info, text = "" , font=('arial', 15), fg = "gray30", bg = "gray100")
                rapport_missing_value_all_pred.place(relx=0.32, rely= 0.94, relwidth= 0.4)
                change_modality_pred = tk.Label(window_model_pred_info, text = "", font=('arial', 15), fg = "gray30", bg = "gray100")
                change_modality_pred.place(relx = 0.67, rely = 0.94, relwidth= 0.25)
                Edit_data_init_pred_new = []
                def  Edit_data_pred(check_data_pred, check_modality_col, old_modality, new_modality, check_rename, check_rename_value, check_missing_quanti, check_missing_quali, Check_list_quanti, Check_list_quali):
                    global  Edit_data_pred_new, Edit_data_init_pred_new
                    header_columns = check_data_pred.columns
                    header2 =  np.concatenate((["Choice value"], header_columns))
                    if check_rename.current() != 0:
                        check_x = header2[check_rename.current()]
                        if check_rename_value.get() !=  "Enter new name column":
                            check_data_pred = check_data_pred.rename(columns={check_x: check_rename_value.get()})
                    if check_modality_col.current() !=0 :
                        check_x = header2[check_modality_col.current()]
                        NN = len(check_data_pred[check_data_pred[check_x] == old_modality.get()])
                        if old_modality.get() != "old modality":
                            if new_modality.get() != "new modality":
                                check_data_pred[check_x]= check_data_pred[check_x].replace(old_modality.get(), new_modality.get()) 
                    quanti_missing_value = check_missing_quanti.get()
                    quali_missing_value = check_missing_quali.get()
                    if quanti_missing_value != "select method":
                        if quanti_missing_value == "mean":
                            for i in range(len(header_columns)):
                                if type(check_data_pred[header_columns[i]][0]) != str:
                                    check_data_pred[header_columns[i]] = check_data_pred[header_columns[i]].fillna(mean(check_data_pred[header_columns[i]])).infer_objects(copy=False)
                        elif  quanti_missing_value == "median":
                            for i in range(len(header_columns)):
                                if type(check_data_pred[header_columns[i]][0]) != str:
                                    check_data_pred[header_columns[i]] = check_data_pred[header_columns[i]].fillna(check_data_pred[header_columns[i]].median()).infer_objects(copy=False)
                        else:
                             data_new = data_new.dropna(axis = 0, how = 'any') 
                    if quali_missing_value != "select method":
                        if  quali_missing_value == "modality reference":
                            for i in range(len(header_columns)):
                                if type(check_data_pred[header_columns[i]][0]) == str:
                                    quali_value = check_data_pred[header_columns[i]].describe()
                                    check_data_pred[header_columns[i]] = check_data_pred[header_columns[i]].fillna(quali_value[2]).infer_objects(copy=False)
                        else:
                            check_data_pred = check_data_pred.dropna(axis = 0, how = 'any')
                    Edit_data_pred_new = check_data_pred.copy()
                    Edit_data_init_pred_new = list(Edit_data_pred_new) 
                    info_data_all_pred = "Rows: " + str(Edit_data_pred_new.shape[0])+ " , Columns: " + str(Edit_data_pred_new.shape[1])
                    label_data_pred2.configure(text = info_data_all_pred +  "                                          ")
                    rapport_missing_value_all_pred.configure(text = "Missing value: " +  " N = " + str(Edit_data_pred_new.isnull().sum().sum()))
                    if check_modality_col.current() !=0 :
                        if old_modality.get() != "old modality":
                            if new_modality.get() != "new modality":
                                check_x = header2[check_modality_col.current()]
                                NN = len(Edit_data_pred_new[Edit_data_pred_new[check_x] == new_modality.get()])
                                change_modality_pred.config(text =  " Number of "  +  old_modality.get() + ": " +   str(NN))
                    Edit_refresh_data_pred()

                def Edit_refresh_data_pred(): 
                    global window_data_pred_new_edit, Edit_data_init_pred_new, Edit_data_pred_new
                    Edit_my_data_pred_new = Edit_data_pred_new.to_numpy().tolist() 
                    window_data_pred_new_edit = ttk.Treeview(check_parent, selectmode = 'extended', show = "headings", columns = Edit_data_init_pred_new)
                    window_data_pred_new_edit.place(relx = 0, rely = 0.7, relwidth = 1, relheight = 0.3)
                    for i in Edit_data_init_pred_new:
                        window_data_pred_new_edit.column(i, width = 130, anchor = 'c')
                        window_data_pred_new_edit.heading(i, text = " " + str(i))
                    for info_value in Edit_my_data_pred_new:
                        value_pred = [x for x in info_value]
                        window_data_pred_new_edit.insert("", 'end', values = value_pred)
                ## ********************************************
                label_modality_info_pred = Label(window_model_pred_info, text="Search & change modality", font=('arial', 15), fg = "saddlebrown", bg = "white")
                label_modality_info_pred.place(relx=0.36, rely= 0.03, relwidth= 0.3)
                label_rename_columns_perd = Label(window_model_pred_info, text = "Rename column", font=('arial', 15), fg = "saddlebrown", bg = "white")
                label_rename_columns_perd.place(relx=0.4, rely= 0.6, relwidth= 0.15)
                label_modality_edit_pred = Label(window_model_pred_info, text = "Characteristic", font=('arial', 14), fg = "gray30", background= "white")
                label_modality_edit_pred.place(relx=0.38, rely= 0.15, relwidth= 0.1)
                modality_edit_pred = ttk.Combobox(window_model_pred_info, values = header_clean_value_all, width= 10)
                modality_edit_pred.place(relx=0.38, rely= 0.2, relwidth= 0.25)
                modality_edit_pred.current(0)
                label_search_pred = Label(window_model_pred_info, text = "Search          " , font=('arial', 14), fg = "gray30", background= "white")
                label_search_pred.place(relx=0.38, rely= 0.3, relwidth= 0.1)
                zone_entry_search_default_pred = StringVar()
                zone_entry_search_default_pred.set("old modality")
                zone_entry_search_pred = Entry(window_model_pred_info, textvariable = zone_entry_search_default_pred)
                zone_entry_search_pred.place(relx=0.38, rely = 0.35, width = 215,  height = 25)
                label_replace_pred = Label(window_model_pred_info, text = "Replace        " , font=('arial', 14), fg = "gray30", background= "white")
                label_replace_pred.place(relx=0.38, rely= 0.45, relwidth= 0.1)
                zone_entry_replace_default_pred = StringVar()
                zone_entry_replace_default_pred.set("new modality")
                zone_entry_replace_pred = Entry(window_model_pred_info, textvariable = zone_entry_replace_default_pred)
                zone_entry_replace_pred.place(relx=0.38, rely = 0.5, width = 215,  height = 25)
                rename_columns_pred = ttk.Combobox(window_model_pred_info, values = header_clean_value_all, width= 10)
                rename_columns_pred.place(relx = 0.38, rely = 0.7, relwidth= 0.25)
                rename_columns_pred.current(0)
                zone_rename_columns_default_pred = StringVar()
                zone_rename_columns_default_pred.set("Enter new name column")
                zone_rename_columns_pred = Entry(window_model_pred_info, textvariable = zone_rename_columns_default_pred)
                zone_rename_columns_pred.place(relx=0.38, rely = 0.8, width = 215,  height = 25)
                ## ********************************************* 
                label_filter_pred = Label(window_model_pred_info, text="Filtering parameters", font=('arial', 15), fg = "saddlebrown", bg = "white")
                label_filter_pred.place(relx=0.65, rely= 0.03, relwidth= 0.3)
                label_Carbon_balance_pre = Label(window_model_pred_info, text="Parameter - 1           ", font=('arial', 14), fg = "gray30", background= "white")
                label_Carbon_balance_pre.place(relx=0.65, rely= 0.15, relwidth= 0.25)
                carbon_balance_pred = ttk.Combobox(window_model_pred_info, values = header_clean_value, width= 10)
                carbon_balance_pred.place(relx=0.7, rely= 0.2, relwidth= 0.25)
                carbon_balance_pred.current(0)
                label_conversion_pred = Label(window_model_pred_info, text="Parameter - 2           ", font=('arial', 14), fg = "gray30", background= "white")
                label_conversion_pred.place(relx=0.65, rely= 0.3, relwidth= 0.25)
                conversion_value_pred = ttk.Combobox(window_model_pred_info, values = header_clean_value, width= 10)
                conversion_value_pred.place(relx=0.7, rely= 0.35, relwidth= 0.25)
                conversion_value_pred.current(0)
                label_yield_pred = Label(window_model_pred_info, text= "Parameter - 3            ", font=('arial', 14), fg = "gray30", background= "white")
                label_yield_pred.place(relx=0.65, rely= 0.45, relwidth= 0.25)
                yield_value_pred = ttk.Combobox(window_model_pred_info, values = header_clean_value, width= 10)
                yield_value_pred.place(relx=0.7, rely= 0.5, relwidth= 0.25)
                yield_value_pred.current(0)
                label_selectivity_pred = Label(window_model_pred_info, text="Parameter - 4           ", font=('arial', 14), fg = "gray30", background= "white")
                label_selectivity_pred.place(relx=0.65, rely= 0.6, relwidth= 0.25)
                selectivity_value_pred = ttk.Combobox(window_model_pred_info, values = header_clean_value, width= 10)
                selectivity_value_pred.place(relx=0.7, rely= 0.65, relwidth= 0.25)
                selectivity_value_pred.current(0)
                View_data_pred_img = PhotoImage(file = my_path('figures/apercu_p.png'))
                View_pred = PhotoImage(file = my_path('figures/View_pred.png'))
                Export_data_pred_img = PhotoImage(file = my_path('figures/Export_data_pred.png'))
                Edit_data_p = PhotoImage(file = my_path('figures/Edit_data_p.png'))
                Overview_stat_pred = Button(check_parent_button, image = View_data_pred_img, font=('arial', 14), text = "Overview stat", command = lambda : stat_data_pred(window_model_pred, data_pred) , compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
                Overview_stat_pred.place(relx= 0.3, rely= 0.53, width= 110, height= 60)
                View_data_pred = Button(check_parent_button, image = View_pred, text = "View  pred", command = lambda : my_data_pred(data_pred) , font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
                View_data_pred.place(relx= 0.3, rely= 0.67, width= 110, height= 60)
                Export_data_pred = Button(check_parent_button, image = Export_data_pred_img, text = "Export  pred", font=('arial', 14), command = lambda : my_exporting_data(data_pred_new) , compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
                Export_data_pred.place(relx= 0.3, rely= 0.81, width= 110, height= 60)
                my_save_all_def  = BooleanVar(value = False)
                my_save_all = Checkbutton(check_parent_button,  text ="Summary for export", variable = my_save_all_def, command = lambda : summary_for_export(window_model_pred_info, my_save_all_def) , fg = "gray30", font=('arial', 15), background = 'gray90')
                my_save_all.place(relx=0.15, rely= 0.93, relwidth= 0.7)
                button_Edit_data_pred = Button(window_model_pred_info, image = Edit_data_p, text = "View & Edit", command = lambda : Edit_data_pred(data_pred, modality_edit_pred, zone_entry_search_default_pred, zone_entry_replace_default_pred, rename_columns_pred, zone_rename_columns_default_pred, imputation_quanti_pred, imputation_quali_pred, list_quanti, list_quali) , font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
                button_Edit_data_pred.place(relx= 0.12, rely= 0.69, width= 110, height= 60)
                carbon_balance_pred.bind("<<ComboboxSelected>>", lambda event: my_carbon_balance_pred(window_model_pred_info, header_clean_value, carbon_balance_pred))
                conversion_value_pred.bind("<<ComboboxSelected>>", lambda event: my_conversion_pred(window_model_pred_info, header_clean_value, conversion_value_pred))
                yield_value_pred.bind("<<ComboboxSelected>>", lambda event: my_yields_pred(window_model_pred_info, header_clean_value, yield_value_pred))
                selectivity_value_pred.bind("<<ComboboxSelected>>", lambda event: my_selectivity_pred(window_model_pred_info, header_clean_value, selectivity_value_pred))
                data_init_pred_new = []

                def  my_data_pred(check_data_pred):
                    global  data_init_pred, data_pred_new, data_init_pred_new, prediction_infos
                    my_check_data_pred = check_data_pred.copy()
                    data_pred_new = my_check_data_pred.copy()
                    ## Le « meilleur modele » = celui qui a le R-squared le plus eleve parmi les modeles entraines
                    ## ET utilisables pour predire (recalcule a chaque clic sur « View pred »).
                    nom_meilleur_modele, r2_meilleur_modele = _meilleur_modele_entraine(modeles_entraines, pour_prediction = True)
                    if nom_meilleur_modele is not None:
                        etat_modele_pred = modeles_projet[nom_meilleur_modele]
                        colonnes_requises = list(etat_modele_pred["result_model"])
                        colonnes_absentes = [c for c in colonnes_requises if c not in my_check_data_pred.columns]
                        if len(colonnes_absentes) != 0:
                            messagebox.showerror("Prediction impossible", "Le meilleur modele (« " + str(nom_meilleur_modele) + " », R2 = " + str(round(r2_meilleur_modele, 2)) + ") necessite des colonnes absentes des donnees de prediction :\n- " + "\n- ".join([str(c) for c in colonnes_absentes]))
                            label_data_pred2.configure(text = "Prediction impossible : colonnes manquantes.")
                        else:
                            try:
                                data_predA_m = my_check_data_pred[colonnes_requises].copy()
                                # memes regles qu'a l'entrainement : les colonnes qualitatives sont toujours recodees
                                colonnes_quali_pred = _colonnes_qualitatives(data_predA_m, colonnes_requises)
                                if etat_modele_pred["selectet_codage"] == "None":
                                    colonnes_recod_pred = []
                                elif etat_modele_pred["selectet_codage"] == "Disjunctive":
                                    colonnes_recod_pred = list(etat_modele_pred["result_model_recod"] or [])
                                else:
                                    colonnes_recod_pred = list(colonnes_requises)
                                colonnes_recod_pred += [c for c in colonnes_quali_pred if c not in colonnes_recod_pred]
                                if len(colonnes_recod_pred) == 0:
                                    X_pred_m = data_predA_m.copy()
                                else:
                                    X_pred_m = my_select_all_features_pred(data_predA_m, colonnes_requises, colonnes_recod_pred)
                                ## une modalite absente des donnees de prediction = colonne indicatrice a 0
                                X_pred_m = X_pred_m.reindex(columns = list(etat_modele_pred["columns_select"]), fill_value = 0)
                                my_pred_m = etat_modele_pred["reg"].predict(X_pred_m)
                                nom_colonne_pred = etat_modele_pred["selectet_target"] + "  [meilleur modele : " + nom_meilleur_modele + ", R2 = " + str(round(r2_meilleur_modele, 2)) + "]"
                                data_pred_new.insert(0, nom_colonne_pred, round(my_pred_m, 3))
                                label_data_pred2.configure(text = "Data with prediction from the best trained model (« " + nom_meilleur_modele + " », R-squared = " + str(round(r2_meilleur_modele, 2)) + ")")
                                ## EVOLUTION : memoriser le resultat (et l'enregistrer dans le projet)
                                prediction_infos = {"modele": nom_meilleur_modele, "r2": float(r2_meilleur_modele), "cible": etat_modele_pred["selectet_target"], "colonne": nom_colonne_pred, "fichier": file_path_pred}
                                _memoriser_projet_apres_prediction("Prediction (best model)")
                            except Exception as erreur_prediction:
                                messagebox.showerror("Prediction impossible", "La prediction avec le modele « " + str(nom_meilleur_modele) + " » a echoue :\n\n" + str(erreur_prediction))
                                label_data_pred2.configure(text = "Prediction failed.")
                    elif len(modeles_entraines) != 0:
                        label_data_pred2.configure(text = "Trained models cannot be used for prediction (reloaded project?). Click « Run » to train them again.")
                    else:
                        label_data_pred2.configure(text = "No trained model with a usable R-squared for prediction.")
                    data_init_pred_new = list(data_pred_new) 
                    change_modality_pred.configure(text = "")
                    rapport_missing_value_all_pred.configure(text = "")
                    refresh_data_pred()

                def refresh_data_pred(): 
                    global window_data_pred_new, data_init_pred_new, data_pred_new
                    my_data_pred_new = data_pred_new.to_numpy().tolist() 
                    window_data_pred_new = ttk.Treeview(check_parent, selectmode = 'extended', show = "headings", columns = data_init_pred_new)
                    window_data_pred_new.place(relx = 0, rely = 0.7, relwidth = 1, relheight = 0.3)
                    for i in data_init_pred_new:
                        window_data_pred_new.column(i, width = 130, anchor = 'c')
                        window_data_pred_new.heading(i, text = " " + str(i))
                    for info_value in my_data_pred_new:
                        value_pred = [x for x in info_value]
                        window_data_pred_new.insert("", 'end', values = value_pred)
                window_model_pred_info.place(relx = 0, rely = 0, relwidth = 1, relheight = 0.7)
                low_window_pred.place(relx=0, rely= 0.7, relwidth = 1, relheight = 0.3)
                ## EVOLUTION : reafficher les resultats de prediction memorises
                if data_pred_new is not None:
                    globals()["data_init_pred_new"] = list(data_pred_new)
                    modele_memorise = prediction_infos.get("modele") if isinstance(prediction_infos, dict) else None
                    r2_memorise = prediction_infos.get("r2") if isinstance(prediction_infos, dict) else None
                    if modele_memorise is not None and r2_memorise is not None:
                        label_data_pred2.configure(text = "Saved prediction from the best trained model (« " + str(modele_memorise) + " », R-squared = " + str(round(float(r2_memorise), 2)) + ")")
                    else:
                        label_data_pred2.configure(text = "Saved prediction results")
                    refresh_data_pred()
            else:
                warning_inf_pred = Label(check_parent, text = traduire("Please load prediction data"), font=('Helvetica 16 bold italic'), fg = "red", background= "white")
                warning_inf_pred.place(relx = 0, rely= 0, relwidth=1, relheight= 1)
        show_all = PhotoImage(file = my_path('figures/show_all.png'))
        import_data_pred_img = PhotoImage(file = my_path('figures/import_data_pred.png'))
        ## Case a cocher : representations graphiques des modeles entraines (memes onglets qu'apres « Run »)
        model_plots_def = BooleanVar(value = False)
        model_plots_info = Checkbutton(window_param_pred, text = "Model plots", variable = model_plots_def, command = lambda : basculer_graphiques_modeles_log(parent, model_plots_def, window_param_pred), fg = "gray30", font=('arial', 15), background = 'gray90')
        model_plots_info.place(relx=0.15, rely= 0.02, relwidth= 0.7)
        label_pred = Label(window_param_pred, text="Making predictions", font=('arial', 15), fg = "saddlebrown", background= "gray90")
        label_pred.place(relx=0.15, rely= 0.08, relwidth= 0.7)
        import_data_pred = Button(window_param_pred, image = import_data_pred_img, text = "Import data", command = lambda : (model_plots_def.set(False), _masquer_graphiques_modeles_log(), import_prediction_data(window_model_pred)) , font=('arial', 14),  compound = tk.TOP,  borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
        import_data_pred.place(relx= 0.3, rely= 0.16, width= 110, height= 60)
        Edit_data_pred = Button(window_param_pred, image = show_all, text = "Show all", command = lambda : (model_plots_def.set(False), _masquer_graphiques_modeles_log(), window_prediction(window_model_pred, window_param_pred)) , font=('arial', 14),  compound = tk.TOP,  borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
        Edit_data_pred.place(relx= 0.3, rely= 0.30, width= 110, height= 60)
        window_model_pred.place(relx = 0, rely = 0, relwidth = 0.8, relheight = 1)
        window_param_pred.place(relx = 0.8, rely = 0, relwidth = 0.2, relheight = 1)
        ## EVOLUTION : des donnees de prediction sont memorisees -> on les reaffiche directement
        if data_pred is not None:
            window_prediction(window_model_pred, window_param_pred)
    else:
        window_model_not = Canvas(parent, relief= GROOVE ,  background = 'white' , highlightthickness=0)
        window_model_not.place(relx=0, rely=0, relwidth=1, relheight=1)
        label_not_model = Label(window_model_not, text = traduire("Please train a model for prediction"), font=('Helvetica 18 bold italic'), fg = "red", background= "white")
        label_not_model.place(relx=0.3, rely = 0.45, relwidth = 0.4)


class _ImportAnnule (Exception):
    """ L'utilisateur a annule le choix des feuilles Excel pendant un import. """


def _lire_donnees_prediction (chemin, feuilles = None, choisir_feuilles = False):
    """ Renvoie (DataFrame, feuilles Excel utilisees) ; (None, None) si le choix des feuilles est annule. """
    return lire_fichier_donnees(chemin, feuilles, choisir_feuilles = choisir_feuilles)


def _memoriser_projet_apres_prediction (etiquette = "Prediction"):
    """Enregistre le projet (donnees de prediction + resultats) sans boite de dialogue."""
    _historique_valider(etiquette)
    if nom_projet is not None and chemin_projet:
        try:
            save_current_project(silencieux = True)
        except Exception:
            pass


def import_prediction_data(parent):

    ## EVOLUTION : les donnees de prediction sont lues des l'import et MEMORISEES (variable data_pred,
    ## enregistree dans le projet). « Annuler » ne perd plus les donnees precedentes.
    global file_path_pred, file_sheets_pred, data_pred, data_pred_new, prediction_infos
    my_fill_pred = [('CSV files',"*.csv"), ('Excel file', '*.xlsx *.xlsm *.xls'), ('Texte file', '*.txt'),('All',"*.*")]
    nouveau_chemin_pred = filedialog.askopenfilename(title="Open csv File", filetypes = my_fill_pred)
    import_reussi = False
    if nouveau_chemin_pred:
        try:
            donnees_lues, feuilles_pred = _lire_donnees_prediction(nouveau_chemin_pred, choisir_feuilles = True)
            if donnees_lues is None:        # choix des feuilles annule : les donnees precedentes sont conservees
                raise _ImportAnnule()
            data_pred = donnees_lues
            file_sheets_pred = feuilles_pred
            file_path_pred = nouveau_chemin_pred
            data_pred_new = None            ## les anciens resultats ne correspondent plus aux nouvelles donnees
            prediction_infos = {}
            import_reussi = True
            _memoriser_projet_apres_prediction("Import prediction data")
        except _ImportAnnule:
            pass
        except Exception as erreur_import_pred:
            messagebox.showerror("Import impossible", "Impossible de lire le fichier de prediction :\n" + str(erreur_import_pred))
    parent1 = Canvas(parent, bg = "white", highlightthickness = 0)
    parent1.place(relx = 0, rely = 0, relwidth = 1, relheight = 1)
    tree = ttk.Treeview(parent1, show="headings")
    tree.place(relx = 0.01, rely = 0.1, relwidth = 1, relheight = 0.9)
    if import_reussi:
        label_x = Label(parent1, text = traduire("Data successfully loaded! \n To make a prediction, \n click on the <View pred> button"), font=('Helvetica 18 bold italic'), fg = "gray20", background= "white")
        label_x.place(relx=0.3, rely= 0.3, relwidth= 0.5)
    else:
        label_x = Label(parent1, text = traduire("Please try again import failed"), font=('Helvetica 18 bold italic'), fg = "red", background= "white")
        label_x.place(relx=0.3, rely= 0.3, relwidth= 0.4)



# ==========================================================================
# ==  16. Configuration d'un modele (bouton Model)                        ==
# ==========================================================================
def my_feature_function(event, parent, list_var_feature, check_var_feature):

    global Feature_column
    if check_var_feature.current() == 1:
            boit_window_feature = Canvas(parent , background= 'cadetblue',  highlightthickness = 2 , relief= GROOVE)
            Feature_column = tk.Listbox(boit_window_feature, fg = "gray10", font = ('arial', 12), relief = GROOVE , selectmode = tk.MULTIPLE, exportselection = False)
            Feature_column.place(relx= 0.02, rely = 0.02, relwidth= 0.96, relheight = 0.85)
            for value in list_var_feature:
                Feature_column.insert(END, value)
         
            features_deja_choisies = globals().get("result_model_1")
            if features_deja_choisies:
                for indice, valeur in enumerate(list_var_feature):
                    if valeur in features_deja_choisies:
                        Feature_column.selection_set(indice)
            close_window_feature = Button(boit_window_feature, text = "Save", command = lambda : recup_my_feature(boit_window_feature), font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
            close_window_feature.place(relx= 0.4, rely= 0.9, width= 50, height= 20)
            boit_window_feature.place(relx = 0.1, rely = 0.1, relwidth = 0.8, relheight = 0.6)


def _initiales_nom (nom):
    """ Initiales d'un nom de modele : premiere lettre (en majuscule) de chaque mot ;
    les nombres sont gardes en entier. Ex. « Random forest model 12 » -> « RFM12 », « test v2 » -> « TV2 ». """
    mots = re.findall(r"[^\W_]+", str(nom))
    initiales = ""
    for mot in mots:
        chiffres_fin = re.search(r"\d+$", mot)
        if mot.isdigit():
            initiales += mot                                    # « 12 » -> « 12 »
        else:
            initiales += mot[0].upper() + (chiffres_fin.group(0) if chiffres_fin else "")     # « v2 » -> « V2 »
    return initiales if initiales != "" else str(nom)[:3]


def _texte_titre_config (nom, largeur_disponible, police):
    """ Nom complet s'il tient dans la largeur disponible, sinon son abreviation : le
    sigle entre parentheses s'il y en a un (« Random Forest (RF) » -> « RF »), sinon
    ses initiales. """
    nom = str(nom)
    try:
        trop_long = largeur_disponible > 1 and _tkfont_langue.Font(font = police).measure(nom) > largeur_disponible
    except Exception:
        trop_long = len(nom) > 30
    if not trop_long:
        return nom
    sigle = re.search(r"\(([^()]+)\)\s*$", nom)
    return sigle.group(1).strip() if sigle else _initiales_nom(nom)


def _algorithme_configure (nom_modele):
    """ Algorithme choisi lors de la configuration du modele (« » s'il n'est pas encore configure). """
    etat = modeles_projet.get(nom_modele) or {}
    algorithme = etat.get("Name_model")
    if not algorithme or algorithme == "Choose model":
        return ""
    return str(algorithme)


## Boites de modeles : texte de l'algorithme choisi (en haut de la boite), mis a jour quand la
## configuration change -> (canevas, id du texte, nom du modele, largeur disponible, police)
_textes_algo_modeles = []


def _texte_algo_boite (nom_modele, largeur, police):
    return _texte_titre_config(_algorithme_configure(nom_modele), largeur, police)


def _rafraichir_algo_boites_modeles ():
    vivants = []
    for canevas, id_texte, nom, largeur, police in _textes_algo_modeles:
        try:
            if not canevas.winfo_exists():
                continue
            canevas.itemconfig(id_texte, text = _texte_algo_boite(nom, largeur, police))
            vivants.append((canevas, id_texte, nom, largeur, police))
        except Exception:
            pass
    _textes_algo_modeles[:] = vivants


def ConfirmWindow(event, parent, list_var, list1):

    global Target_column, Calculation_method, loss_size, Target_list, Name_model, Alpha_size, Alpha_lasso_size, data, header_target
    global  C_size, gamma_size, Random_size, trees_size, estimators_size, depth_size, Random_size_r, depth_size_r, rate_size, Random_size_gpr, codage_method, feature_select, result_model_1

   
    if file_path is not None and not _assurer_donnees_source_chargees():
        return

    etat_existant = modeles_projet.get(nom_modele_actif) if nom_modele_actif is not None else None
    reg_existant = etat_existant.get("reg") if etat_existant is not None else None

    def _valeur_existante (nom_attribut, defaut):
       
        ## CORRECTIF : ne reprendre les hyperparametres que si l'algorithme choisi est celui deja configure
        ## (sinon ex. max_depth=None d'un Random Forest etait injecte comme "None" dans un Decision Tree).
        if reg_existant is not None and etat_existant.get("Name_model") == Name_model and hasattr(reg_existant, nom_attribut):
            try:
                valeur_attribut = getattr(reg_existant, nom_attribut)
                if valeur_attribut is None:
                    return defaut
                return str(valeur_attribut)
            except Exception:
                pass
        return defaut

    if file_path is not None:
        x = int(round(0.80*parent.winfo_width(), 0))
        y = int(round(0.84*parent.winfo_height(), 0))
        selected_index = list1.current()
        selected = list_var[selected_index]
        Name_model = selected
        minit = "Choose model"
        ##m1 = "Linear"
        m2 = "Kernel ridge (KR)"
        m3 = "Support Vector Regression (SVR)"
        m4 = "Random Forest (RF)"
        m5 = "Generalized Boosted Regression Modeling (GBM)"
        m6 = "Decision Tree (DT)"
        m7 = "Gaussian process (GP)"
        m8 = "Lasso"
        if selected != minit:
            ##button_Model.configure(text = selected)
            ## CORRECTIF : changer d'algorithme rappelait ConfirmWindow et EMPILAIT un nouveau panneau sur
            ## l'ancien (widgets fantomes toujours vivants) -> on detruit l'ancien avant de creer le nouveau.
            if _widget_existe(_panneau_config.get("param")):
                try:
                    _panneau_config["param"].destroy()
                except Exception:
                    pass
            boit_param= Canvas(parent , background= 'white',  highlightthickness = 2 , relief= GROOVE)
            _panneau_config["nom"] = nom_modele_actif
            _panneau_config["conteneur"] = parent
            _panneau_config["param"] = boit_param
            label_Target = Label(boit_param, text="Target column   ", font=('arial', 15), fg = "gray30", background= "white")
            label_Target.place(relx=0.02, rely= 0.05, relwidth= 0.4)
            label_Feature = Label(boit_param, text="Feature columns   ", font=('arial', 15), fg = "gray30", bg= "white")
            label_Feature.place(relx=0.05, rely= 0.35, relwidth= 0.4)
            label_recodage = Label(boit_param, text="Recoding feature \n " + "columns                ", font=('arial', 15), fg = "gray30", bg= "white")
            label_recodage.place(relx=0.02, rely= 0.45, relwidth= 0.4)
            #list_a =  csv_reader = csv.reader(open(file_path, 'r', newline=''), delimiter = ";")
            #header = next(list_a) 
            header1 = data.columns
            header_target1 = np.concatenate((["Select column name"], header1))
            header_target = []
            for my_value in header_target1:
                header_target.append(my_value)
            header = []
            for header_value in header1:
                header.append(header_value)
            Target_list = []
            Feature_list = []
            if header:
                Target_list = header
                Feature_list =  header 
            else:
                Target_list = "empty" 
                Feature_list = "empty"
            Target_column = ttk.Combobox(boit_param, values = header_target, width= 10)
            if etat_existant is not None and etat_existant.get("selectet_target") in header_target:
                Target_column.current(header_target.index(etat_existant.get("selectet_target")))
            else:
                Target_column.current(0)
            Target_column.place(relx=0.47, rely= 0.05, relwidth= 0.5)
            label_method = Label(boit_param, text="Calculation method", font=('arial', 15), fg = "gray30", bg= "white")
            label_method.place(relx=0.05, rely= 0.15, relwidth= 0.4)
            method_list = ["TrainTestSplit data", "Cross-validation"]
            Calculation_method = ttk.Combobox(boit_param, values = method_list, width= 10)
            if etat_existant is not None and etat_existant.get("selectet_methode") in method_list:
                Calculation_method.current(method_list.index(etat_existant.get("selectet_methode")))
            else:
                Calculation_method.current(0)
            Calculation_method.place(relx=0.47, rely= 0.15, relwidth= 0.5)
            feature_select_list = ["make your choice", "select feature", "use all features"]
            feature_select = ttk.Combobox(boit_param, values = feature_select_list, width= 10)
            
            features_existantes = etat_existant.get("result_model") if etat_existant is not None else None
            if features_existantes is not None and len(features_existantes) != 0:
                toutes_les_features = [c for c in header if c != etat_existant.get("selectet_target")]
                result_model_1 = list(features_existantes)
                if list(features_existantes) == toutes_les_features:
                    feature_select.current(2)
                else:
                    feature_select.current(1)
            else:
                feature_select.current(0)
                result_model_1 = []     ## CORRECTIF : ne pas heriter des features d'un autre modele / projet
            feature_select.place(relx=0.47, rely= 0.35, relwidth= 0.5)
            method_recode = ["None", "Disjunctive", "All-disjunctive"]
            codage_method = ttk.Combobox(boit_param, values = method_recode, width= 10)
            if etat_existant is not None and etat_existant.get("selectet_codage") in method_recode:
                codage_method.current(method_recode.index(etat_existant.get("selectet_codage")))
            else:
                codage_method.current(0)
            codage_method.place(relx=0.47, rely= 0.45, relwidth= 0.5)
            Calculation_method.bind("<<ComboboxSelected>>", lambda event: data_size(event, boit_param, Calculation_method))
            ## CORRECTIF : le champ « Test size » n'etait cree qu'apres un changement manuel de la
            ## methode de calcul. Sinon Test_size restait None (nouveau projet) ou pointait vers un widget
            ## detruit, et « Run » ne faisait rien. On le cree donc des l'ouverture du panneau.
            taille_test_existante = None
            if etat_existant is not None and etat_existant.get("selectet_methode") in (None, "TrainTestSplit data"):
                valeur_taille_test = etat_existant.get("size_data_value")
                if isinstance(valeur_taille_test, (int, float)) and 0 < valeur_taille_test < 1:
                    taille_test_existante = valeur_taille_test
            data_size(None, boit_param, Calculation_method, taille_test_existante)
            codage_method.bind("<<ComboboxSelected>>", lambda event: variable_recoding_function(event, boit_param, codage_method, method_recode, header_target))
            feature_select.bind("<<ComboboxSelected>>", lambda event: my_feature_function(event, boit_param, Feature_list, feature_select))
            if selected == m8:
               ## button_Model.configure(text = "Lasso")  # bouton Model : le libelle reste toujours 'Model'
               label_Alpha_lasso = Label(boit_param, text="Parameter Alpha", font=('arial', 15), fg = "gray30", bg= "white")
               label_Alpha_lasso.place(relx=0.01, rely= 0.66, relwidth= 0.4)
               default_value_Alpha_lasso = StringVar()
               default_value_Alpha_lasso.set(_valeur_existante("alpha", "1"))
               Alpha_lasso_size = Spinbox(boit_param, from_ = 0. , to = 100 , increment = 0.001, textvariable = default_value_Alpha_lasso)
               Alpha_lasso_size.place(relx=0.47, rely= 0.66, relwidth= 0.5) 
            elif selected == m2:
               ## button_Model.configure(text = "Kernel ridge")  # bouton Model : le libelle reste toujours 'Model'
               label_Alpha = Label(boit_param, text="Parameter Alpha", font=('arial', 15), fg = "gray30", bg= "white")
               label_Alpha.place(relx=0.01, rely= 0.66, relwidth= 0.4)
               default_value_Alpha = StringVar()
               default_value_Alpha.set(_valeur_existante("alpha", "1"))
               Alpha_size = Spinbox(boit_param, from_ = 0. , to = 100 , increment = 0.001, textvariable = default_value_Alpha)
               Alpha_size.place(relx=0.47, rely= 0.66, relwidth= 0.5) 
            elif selected == m3:
                ## button_Model.configure(text = "SVR")  # bouton Model : le libelle reste toujours 'Model'
                label_C = Label(boit_param, text="Parameter C    ", font=('arial', 15), fg = "gray30", bg= "white")
                label_C.place(relx=0.01, rely= 0.66, relwidth= 0.4)
                label_gamma = Label(boit_param, text="Parameter gamma", font=('arial', 15), fg = "gray30", bg= "white")
                label_gamma.place(relx=0.01, rely= 0.76, relwidth= 0.4)
                ##C_size = Spinbox(boit_param, from_ = 10 , to = 100000 , increment = 1)
                default_value_C = StringVar()
                default_value_C.set(_valeur_existante("C", "100"))
                C_size = Spinbox(boit_param, from_ = 1 , to = 100000 , increment = 1, textvariable = default_value_C)
                default_value_gamma = StringVar()
                default_value_gamma.set(_valeur_existante("gamma", "0.1"))
                gamma_size = Spinbox(boit_param, from_ = 0 , to = 1 , increment = 0.001, textvariable = default_value_gamma)
                C_size.place(relx=0.47, rely= 0.66, relwidth= 0.5) 
                gamma_size.place(relx=0.47, rely= 0.76, relwidth= 0.5) 
                my_kernel =  BooleanVar(value = False)
                my_kernel = Checkbutton(boit_param,  text ="With linear kernel", variable = my_kernel, font=('arial', 15), fg = "gray30", background = 'white')
                my_kernel.place(relx=0.4, rely= 0.86, width= 170, height= 30)
            elif selected == m4:
                ## button_Model.configure(text = "Random Forest")  # bouton Model : le libelle reste toujours 'Model'
                label_Random = Label(boit_param, text="Random state    ", font=('arial', 15), fg = "gray30", bg= "white")
                label_Random.place(relx=0.01, rely= 0.55, relwidth= 0.4)
                label_trees = Label(boit_param, text="Number of trees", font=('arial', 15), fg = "gray30", bg= "white")
                label_trees.place(relx=0.01, rely= 0.65, relwidth= 0.4)
                default_value_m4_Random = StringVar()
                default_value_m4_Random.set(_valeur_existante("random_state", "0"))
                Random_size = Spinbox(boit_param, from_ = 0 , to = 100000 , increment = 1, textvariable = default_value_m4_Random)
                default_value_m4_trees = StringVar()
                default_value_m4_trees.set(_valeur_existante("n_estimators", "100"))
                trees_size = Spinbox(boit_param, from_ = 1 , to = 100000 , increment = 5, textvariable = default_value_m4_trees)
                Random_size.place(relx=0.47, rely= 0.55, relwidth= 0.5) 
                trees_size.place(relx=0.47, rely= 0.65, relwidth= 0.5) 
            elif selected == m5:
                ## button_Model.configure(text = "GBM")  # bouton Model : le libelle reste toujours 'Model'
                label_estimators = Label(boit_param, text="n_estimators    ", font=('arial', 15), fg = "gray30", bg= "white")
                label_estimators.place(relx=0.01, rely= 0.65, relwidth= 0.4)
                label_depth = Label(boit_param, text="max_depth     ", font=('arial', 15), fg = "gray30", bg= "white")
                label_depth.place(relx=0.01, rely= 0.75, relwidth= 0.4)
                label_loss = Label(boit_param, text="loss                ", font=('arial', 15), fg = "gray30", bg= "white")
                label_loss.place(relx=0.01, rely= 0.85, relwidth= 0.4)
                default_value_m5_estimators = StringVar()
                default_value_m5_estimators.set(_valeur_existante("n_estimators", "500"))
                estimators_size = Spinbox(boit_param, from_ = 0 , to = 100000 , increment = 10, textvariable = default_value_m5_estimators)
                default_value_m5_depth = StringVar()
                default_value_m5_depth.set(_valeur_existante("max_depth", "5"))
                depth_size = Spinbox(boit_param, from_ = 1 , to = 100000 , increment = 1, textvariable=default_value_m5_depth)
                loss_list = ["squared_error", "absolute_error", "huber", "quantile"]
                loss_size = ttk.Combobox(boit_param, values = loss_list, width= 10)
                loss_existante = _valeur_existante("loss", loss_list[0])
                loss_size.current(loss_list.index(loss_existante) if loss_existante in loss_list else 0)
                estimators_size.place(relx=0.47, rely= 0.65, relwidth= 0.5) 
                depth_size.place(relx=0.47, rely= 0.75, relwidth= 0.5) 
                loss_size.place(relx=0.47, rely= 0.85, relwidth= 0.5) 
                label_rate = Label(boit_param, text="learning_rate ", font=('arial', 15), fg = "gray30", bg= "white")
                label_rate.place(relx=0.01, rely= 0.93, relwidth= 0.4)
                default_size_value = StringVar()
                default_size_value.set(_valeur_existante("learning_rate", "0.5"))
                rate_size = Spinbox(boit_param, from_ = 0 , to = 1 , increment = 0.01, textvariable = default_size_value)
                rate_size.place(relx=0.47, rely= 0.93, relwidth= 0.5) 
            elif selected == m6:
                 ## button_Model.configure(text = "Decision Tree")  # bouton Model : le libelle reste toujours 'Model'
                 label_Random = Label(boit_param, text="Random state  ", font=('arial', 15), fg = "gray30", bg= "white")
                 label_Random.place(relx=0.01, rely= 0.66, relwidth= 0.4)
                 label_depth = Label(boit_param, text="Max depth       ", font=('arial', 15), fg = "gray30", bg= "white")
                 label_depth.place(relx=0.01, rely= 0.76, relwidth= 0.4)
                 default_value_dtr_random = StringVar()
                 default_value_dtr_random.set(_valeur_existante("random_state", "0"))
                 Random_size_r = Spinbox(boit_param, from_ = 0 , to = 100000 , increment = 1, textvariable = default_value_dtr_random)
                 default_value_dtr = StringVar()
                 default_value_dtr.set(_valeur_existante("max_depth", "20"))
                 depth_size_r = Spinbox(boit_param, from_ = 1 , to = 100000 , increment = 1, textvariable = default_value_dtr)
                 Random_size_r.place(relx=0.47, rely= 0.66, relwidth= 0.5)
                 depth_size_r.place(relx=0.47, rely= 0.76, relwidth= 0.5) 
            elif selected == m7:
                ## button_Model.configure(text = "Gaussian process")  # bouton Model : le libelle reste toujours 'Model'
                label_Random = Label(boit_param, text="Random state  ", font=('arial', 15), fg = "gray30", bg= "white")
                label_Random.place(relx=0.01, rely= 0.66, relwidth= 0.4)
                default_value_dtr_random = StringVar()
                default_value_dtr_random.set(_valeur_existante("random_state", "0"))
                Random_size_gpr = Spinbox(boit_param, from_ = 0 , to = 100000 , increment = 1, textvariable = default_value_dtr_random)
                Random_size_gpr.place(relx=0.47, rely= 0.66, relwidth= 0.5)

        
            bouton_sauver_config = Button(boit_param, image = save_config_model, command = lambda: sauvegarder_configuration_modele_bouton(parent), font=('arial', 14), text = "Save configure", compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
            bouton_sauver_config.place(relx = 0.35, rely = 0.85, width= 110, height= 60)

            boit_param.place(relx= 0.1 , rely= 0.15 , width= x, height=y)
        else:            
            parent.destroy() 
    else:
        warning_inf = Label(parent, text = traduire("Please load the data"), font=('Helvetica 16 bold italic'), fg = "red", background= "white")
        warning_inf.place(relx = 0, rely= 0, relwidth=1, relheight= 1)


def choice_model (event, parent, colors = "#4a7a8c"):

    if not analyse_autorisee():      # bouton desactive : un evenement lie a la main peut quand meme arriver
        return

    if not verifier_projet_ouvert():
        return

    global button_Model, button_Settings, button_Rename, button_Delete_model

    boit_model = Canvas(parent , background= 'white', highlightthickness = 0, relief= FLAT)
    boit_model.place(relx= 0 , rely= 0 , relwidth=  1, relheight= 1)
    
    #label_actif = Label(boit_model, text = "", font=('arial', 16), fg = "saddlebrown", background= 'white')
    #label_actif.place(relx = 0.03, rely = 0.03, relwidth = 0.55)

    ## Remarque : la creation d'un nouveau modele se fait desormais UNIQUEMENT via
    ## le bouton "Analysis" (barre verticale gauche, voir save_Model). Ce panneau
    ## "Model" ne sert plus qu'a consulter/ouvrir/renommer/supprimer les modeles
    ## deja crees, afin qu'il n'y ait qu'un seul point d'entree pour la creation.

    global taille_boites_modeles
    etat_taille = {"valeur": taille_boites_modeles}

    TAILLE_MIN = 50
    TAILLE_MAX = 160

    def _curseur_vers_taille(valeur_curseur):
        return int(TAILLE_MIN + (float(valeur_curseur) / 100) * (TAILLE_MAX - TAILLE_MIN))

    def _taille_vers_curseur(taille_boite):
        return int((taille_boite - TAILLE_MIN) / (TAILLE_MAX - TAILLE_MIN) * 100)

    info_vide = Label(boit_model, text = traduire("Aucun modele defini pour l'instant. \nCliquez sur le bouton « Analysis » (barre verticale de gauche) pour en creer un."), font=('Helvetica 13 italic'), fg = "gray40", background= "white")

    modeles_noms = list(modeles_projet.keys())

    COULEUR_BOITE_NORMALE = "white"
    COULEUR_BORDURE_SELECTION = colors
    ESPACEMENT_BOITE = 12
    RATIO_RAYON_ARRONDI = 0.2
    RATIO_POLICE = 10 / 88

    boite_par_nom = {}
    selection_courante = {"nom": None}
    _glisser_etat = {"nom": None, "reference": None, "origine_xy": None}

    def _points_rectangle_arrondi(x1, y1, x2, y2, rayon):
        return [
            x1 + rayon, y1,
            x2 - rayon, y1,
            x2, y1,
            x2, y1 + rayon,
            x2, y2 - rayon,
            x2, y2,
            x2 - rayon, y2,
            x1 + rayon, y2,
            x1, y2,
            x1, y2 - rayon,
            x1, y1 + rayon,
            x1, y1,
        ]

    def _dessiner_rectangle_arrondi(canvas_boite, x1, y1, x2, y2, rayon, **options):
        return canvas_boite.create_polygon(_points_rectangle_arrondi(x1, y1, x2, y2, rayon), smooth = True, **options)

    def _taille_police(taille_boite):
        taille_police = int(taille_boite * RATIO_POLICE)
        if taille_police < 7:
            taille_police = 7
        return taille_police

    def _zone_deplacement():
        ## renvoie les bornes (xmin, ymin, xmax, ymax), en pixels, a l'interieur
        ## du panneau central ou une boite peut etre deplacee (quasiment tout
        ## le panneau, avec juste une petite marge pour rester visible)
        taille_boite = etat_taille["valeur"]
        largeur_zone = boit_model.winfo_width()
        hauteur_zone = boit_model.winfo_height()
        if largeur_zone <= 1:
            largeur_zone = parent.winfo_width()
        if hauteur_zone <= 1:
            hauteur_zone = parent.winfo_height()
        xmin = int(0.01 * largeur_zone)
        ymin = int(0.01 * hauteur_zone)
        xmax = largeur_zone - taille_boite - 10
        ymax = hauteur_zone - taille_boite - 10
        if xmax < xmin:
            xmax = xmin
        if ymax < ymin:
            ymax = ymin
        return xmin, ymin, xmax, ymax

    def _position_par_defaut(indice):
        taille_boite = etat_taille["valeur"]
        xmin, ymin, xmax, ymax = _zone_deplacement()
        ymin = ymin + int(0.1 * (ymax - ymin if ymax > ymin else 1)) + 30  ## evite de superposer la 1ere rangee au libelle "Taille des boites"
        if ymin > ymax:
            ymin = ymax
        largeur_disponible = (xmax - xmin) + taille_boite
        colonnes = largeur_disponible // (taille_boite + ESPACEMENT_BOITE)
        if colonnes < 1:
            colonnes = 1
        ligne, colonne = divmod(indice, colonnes)
        x = xmin + colonne * (taille_boite + ESPACEMENT_BOITE)
        y = ymin + ligne * (taille_boite + ESPACEMENT_BOITE)
        return x, y

    def _couleur_boite(nom):
        return couleurs_boites_modeles.get(nom, COULEUR_BOITE_NORMALE)

    def _couleur_texte_lisible(couleur_fond):
        ## choisit un texte blanc ou fonce selon la luminosite de la couleur de fond
        try:
            r, g, b = boit_model.winfo_rgb(couleur_fond)
            luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 65535
            if luminance < 0.55:
                return "white"
        except Exception:
            pass
        return "gray10"

    def mettre_a_jour_apparence_boites():
        for nom_boite, elements in boite_par_nom.items():
            cadre_boite, forme_id, texte_id = elements
            couleur_fond = _couleur_boite(nom_boite)
            cadre_boite.itemconfig(texte_id, fill = _couleur_texte_lisible(couleur_fond))
            for canevas_algo, algo_id, _nom, _largeur, _police in _textes_algo_modeles:     # algorithme (haut de la boite)
                if canevas_algo is cadre_boite:
                    try:
                        cadre_boite.itemconfig(algo_id, fill = _couleur_texte_lisible(couleur_fond))
                    except Exception:
                        pass
            if nom_boite == selection_courante["nom"]:
                cadre_boite.itemconfig(forme_id, fill = couleur_fond, outline = COULEUR_BORDURE_SELECTION, width = 3)
            else:
                cadre_boite.itemconfig(forme_id, fill = couleur_fond, outline = "gray70", width = 1)

    def selectionner_modele(nom):
        selection_courante["nom"] = nom
        mettre_a_jour_apparence_boites()

    def demarrer_glissement_boite(event, nom):
        selectionner_modele(nom)
        cadre_boite = boite_par_nom[nom][0]
        tk.Misc.tkraise(cadre_boite)  ## amene la boite au premier plan (Canvas redefinit .lift()/.tkraise() pour ses items, d'ou l'appel direct a Misc)
        _glisser_etat["nom"] = nom
        _glisser_etat["reference"] = (event.x_root, event.y_root)
        _glisser_etat["origine_xy"] = (cadre_boite.winfo_x(), cadre_boite.winfo_y())

    def glisser_boite(event, nom):
        if _glisser_etat.get("nom") != nom:
            return
        reference = _glisser_etat.get("reference")
        origine_xy = _glisser_etat.get("origine_xy")
        if reference is None or origine_xy is None:
            return
        delta_x = event.x_root - reference[0]
        delta_y = event.y_root - reference[1]
        xmin, ymin, xmax, ymax = _zone_deplacement()
        nouveau_x = origine_xy[0] + delta_x
        nouveau_y = origine_xy[1] + delta_y
        if nouveau_x < xmin:
            nouveau_x = xmin
        elif nouveau_x > xmax:
            nouveau_x = xmax
        if nouveau_y < ymin:
            nouveau_y = ymin
        elif nouveau_y > ymax:
            nouveau_y = ymax
        boite_par_nom[nom][0].place(x = nouveau_x, y = nouveau_y)

    def terminer_glissement_boite(event, nom):
        if _glisser_etat.get("nom") != nom:
            return
        cadre_boite = boite_par_nom[nom][0]
        ## memorise la position choisie par l'utilisateur pour ce modele
        ancienne_position_boite = positions_boites_modeles.get(nom)
        positions_boites_modeles[nom] = (cadre_boite.winfo_x(), cadre_boite.winfo_y())
        if ancienne_position_boite != positions_boites_modeles[nom]:
            _historique_valider("Move model box")
        _glisser_etat["nom"] = None
        _glisser_etat["reference"] = None
        _glisser_etat["origine_xy"] = None

    def _appliquer_taille_aux_boites():
        ## redimensionne les boites existantes sans les recreer (position, selection
        ## et liaisons souris sont conservees)
        taille_boite = etat_taille["valeur"]
        rayon = int(taille_boite * RATIO_RAYON_ARRONDI)
        police = _taille_police(taille_boite)
        for cadre_boite, forme_id, texte_id in boite_par_nom.values():
            cadre_boite.configure(width = taille_boite, height = taille_boite)
            cadre_boite.coords(forme_id, *_points_rectangle_arrondi(2, 2, taille_boite - 2, taille_boite - 2, rayon))
            cadre_boite.coords(texte_id, taille_boite / 2, taille_boite / 2)
            cadre_boite.itemconfig(texte_id, width = taille_boite - 14, font = ('arial', police))

    def _sur_changement_taille(valeur_curseur):
        global taille_boites_modeles
        etat_taille["valeur"] = _curseur_vers_taille(valeur_curseur)
        taille_boites_modeles = etat_taille["valeur"]
        if boite_par_nom:
            _appliquer_taille_aux_boites()

    ## Reutilise la barre verticale gauche (widget global "scale") pour piloter
    ## la taille des boites tant que l'onglet "Model" est ouvert. On memorise
    ## sa configuration precedente pour la restaurer proprement en la quittant.
    _reglages_precedents_scale = {
        "command": scale.cget("command"),
        "valeur": scale.get(),
    }

    def _restaurer_scale_precedente():
        try:
            if not scale.winfo_exists():
                return
            scale.configure(command = "")
            scale.set(_reglages_precedents_scale["valeur"])
            if _reglages_precedents_scale["command"]:
                scale.configure(command = _reglages_precedents_scale["command"])
        except Exception:
            pass

    scale.configure(command = _sur_changement_taille)
    scale.set(_taille_vers_curseur(etat_taille["valeur"]))
    boit_model.bind("<Destroy>", lambda event: _restaurer_scale_precedente(), add = "+")

    def construire_boites_modeles():
        nonlocal modeles_noms
        for elements in list(boite_par_nom.values()):
            elements[0].destroy()
        boite_par_nom.clear()
        modeles_noms = list(modeles_projet.keys())

        taille_boite = etat_taille["valeur"]
        rayon = int(taille_boite * RATIO_RAYON_ARRONDI)
        police = _taille_police(taille_boite)

        for indice, nom in enumerate(modeles_noms):
            cadre_boite = Canvas(boit_model, width = taille_boite, height = taille_boite, background = 'white', highlightthickness = 0, cursor = "hand2")

            if nom in positions_boites_modeles:
                x, y = positions_boites_modeles[nom]
            else:
                x, y = _position_par_defaut(indice)
                positions_boites_modeles[nom] = (x, y)
            cadre_boite.place(x = x, y = y)

            couleur_fond = _couleur_boite(nom)
            forme_id = _dessiner_rectangle_arrondi(cadre_boite, 2, 2, taille_boite - 2, taille_boite - 2, rayon, fill = couleur_fond, outline = "gray70", width = 1)
            texte_id = cadre_boite.create_text(taille_boite / 2, taille_boite / 2, text = nom + _statut_affichage_modele(nom), font = ('arial', police), fill = _couleur_texte_lisible(couleur_fond), width = taille_boite - 14, justify = 'center')
            _enregistrer_texte_etat_modele(cadre_boite, texte_id, nom)
            ## en haut de la boite : algorithme choisi lors de la configuration (sigle / initiales si trop long)
            police_algo = ('arial', _builtins_langue.max(8, int(police * 0.85)), 'bold')
            largeur_algo = taille_boite - 20
            algo_id = cadre_boite.create_text(taille_boite / 2, 8, text = _texte_algo_boite(nom, largeur_algo, police_algo), font = police_algo, fill = _couleur_texte_lisible(couleur_fond), anchor = 'n', justify = 'center')
            _textes_algo_modeles.append((cadre_boite, algo_id, nom, largeur_algo, police_algo))
            boite_par_nom[nom] = (cadre_boite, forme_id, texte_id)

            cadre_boite.bind("<Button-1>", lambda event, n = nom: demarrer_glissement_boite(event, n))
            cadre_boite.bind("<B1-Motion>", lambda event, n = nom: glisser_boite(event, n))
            cadre_boite.bind("<ButtonRelease-1>", lambda event, n = nom: terminer_glissement_boite(event, n))
            cadre_boite.bind("<Double-Button-1>", lambda event, n = nom: (selectionner_modele(n), ouvrir_modele_selectionne()))

        if selection_courante["nom"] not in modeles_noms:
            selection_courante["nom"] = None
        mettre_a_jour_apparence_boites()

        if len(modeles_noms) == 0:
            info_vide.place(relx = 0.03, rely = 0.3, relwidth = 0.4)
        else:
            info_vide.place_forget()

    construire_boites_modeles()

    ## Reference vers la boite d'ouverture/parametrage d'un modele actuellement
    ## affichee (choix du type de modele, puis parametres : cible, features, etc.)
    ## Comme la boite de parametres est imbriquee dans celle-ci, la detruire
    ## suffit a faire disparaitre l'ensemble.
    fenetre_parametres = {"widget": None}

    def fermer_fenetre_parametres():
        _fermer_panneau_config()
        widget_ouvert = fenetre_parametres.get("widget")
        if widget_ouvert is not None:
            try:
                if widget_ouvert.winfo_exists():
                    widget_ouvert.destroy()
            except Exception:
                pass
            fenetre_parametres["widget"] = None

    def deselectionner_modele(event = None):
        selectionner_modele(None)
        fermer_fenetre_parametres()

    ## Clic en dehors des boites (fond du panneau ou libelles) : on deselectionne
    ## et on referme la boite de parametrage d'un modele si elle est ouverte
    boit_model.bind("<Button-1>", deselectionner_modele)
    info_vide.bind("<Button-1>", deselectionner_modele)


    def Combobox_list (preselect = None):
        fermer_fenetre_parametres()
        x = int(round(0.40*parent.winfo_width(), 0))
        y = int(round(0.93*parent.winfo_height(), 0))
        boit_choix_model= Canvas(parent, background= 'white',  highlightthickness = 2 , relief= GROOVE)
        boit_choix_model.place(relx= 0.5 , rely= 0.06, width= x, height=y)
        fenetre_parametres["widget"] = boit_choix_model
        labelChoix = Label(boit_choix_model, text = traduire("Please select a model (Regression)"), font=('arial', 18), background= 'white')
        labelChoix.pack(pady= 10)
        listeModel = ["Choose model", "Linear", "Kernel ridge (KR)", "Lasso", "Support Vector Regression (SVR)","Random Forest (RF)", "Generalized Boosted Regression Modeling (GBM)", "Decision Tree (DT)", "Gaussian process (GP)"]
        listeCombo = ttk.Combobox(boit_choix_model, values = listeModel, width= 35)
        if preselect is not None and preselect in listeModel:
            listeCombo.current(listeModel.index(preselect))
        else:
            listeCombo.current(0)
        listeCombo.pack()
        listeCombo.bind("<<ComboboxSelected>>", lambda event: ConfirmWindow(event, boit_choix_model, listeModel, listeCombo))


    def ouvrir_modele_selectionne():
        if not analyse_autorisee():      # bouton desactive
            return
        nom = selection_courante.get("nom")
        if not nom or nom not in modeles_projet:
            messagebox.showinfo("Selection requise", "Veuillez d'abord selectionner un modele dans la liste.")
            return
        _capturer_configuration_widgets_actifs()
        _fermer_panneau_config()
        restore_modele(nom)
        ##label_actif.configure(text = "Modele actif : " + nom)
        if modeles_projet[nom].get("data_model") is not None:
            ## CORRECTIF : un modele entraine ne pouvait plus etre reconfigure (seuls ses resultats s'ouvraient)
            reconfigurer = messagebox.askyesno("Modele deja entraine", "Le modele « " + nom + " » est deja entraine.\n\nOui : le reconfigurer (ses resultats seront recalcules au prochain « Run »).\nNon : afficher ses resultats.")
            if reconfigurer:
                Combobox_list(preselect = modeles_projet[nom].get("Name_model"))
                _panneau_config["reconfiguration"] = True
            else:
                ## on affiche directement ses resultats graphiques, comme le fait le bouton "Run"
                _afficher_toutes_les_representations_graphiques(parent, modele_a_selectionner = nom)
        else:
            Combobox_list(preselect = modeles_projet[nom].get("Name_model"))

    def renommer_modele_selectionne():
        if not analyse_autorisee():      # bouton desactive
            return
        global nom_modele_actif
        ancien_nom = selection_courante.get("nom")
        if not ancien_nom or ancien_nom not in modeles_projet:
            messagebox.showinfo("Selection requise", "Veuillez d'abord selectionner un modele dans la liste.")
            return
        nouveau_nom = simpledialog.askstring("Renommer le modele", "Nouveau nom pour « " + ancien_nom + " »", initialvalue = ancien_nom)
        if nouveau_nom is None:
            return
        nouveau_nom = nouveau_nom.strip()
        if nouveau_nom == "":
            return

        nom_final = ancien_nom
        if nouveau_nom != ancien_nom:
            if nouveau_nom in modeles_projet:
                messagebox.showwarning("Nom deja utilise", "Un modele porte deja ce nom. Choisissez un autre nom.")
                return
            modeles_projet[nouveau_nom] = modeles_projet.pop(ancien_nom)
            if ancien_nom in positions_boites_modeles:
                positions_boites_modeles[nouveau_nom] = positions_boites_modeles.pop(ancien_nom)
            if ancien_nom in couleurs_boites_modeles:
                couleurs_boites_modeles[nouveau_nom] = couleurs_boites_modeles.pop(ancien_nom)
            if nom_modele_actif == ancien_nom:
                nom_modele_actif = nouveau_nom
            if ancien_nom in _hyperparametres_charges:
                _hyperparametres_charges[nouveau_nom] = _hyperparametres_charges.pop(ancien_nom)
            if _panneau_config.get("nom") == ancien_nom:
                _panneau_config["nom"] = nouveau_nom
            #if nom_modele_actif is not None:
            #    label_actif.configure(text = "Modele actif : " + nom_modele_actif)
            nom_final = nouveau_nom

        ## Choix (facultatif) d'une couleur pour la boite de ce modele
        couleur_actuelle = couleurs_boites_modeles.get(nom_final, COULEUR_BOITE_NORMALE)
        couleur_choisie = colorchooser.askcolor(color = couleur_actuelle, title = "Choisir une couleur pour « " + nom_final + " »")
        if couleur_choisie is not None and couleur_choisie[1] is not None:
            couleurs_boites_modeles[nom_final] = couleur_choisie[1]

        selection_courante["nom"] = nom_final
        _historique_valider("Rename model « " + str(nom_final) + " »")
        construire_boites_modeles()

    def supprimer_modele_selectionne():
        if not analyse_autorisee():      # bouton desactive
            return
        
        global nom_modele_actif
        nom = selection_courante.get("nom")
        if not nom or nom not in modeles_projet:
            messagebox.showinfo("Selection requise", "Please select the model")
            return
        confirmation = messagebox.askyesno("Supprimer le modele", "Do you really want to delete the model « " + nom + "» ?\n\nVous pourrez le restaurer plus tard via le bouton « Restore ».")
        if not confirmation:
            return
        nom_corbeille = nom
        compteur = 2
        while nom_corbeille in modeles_supprimes:
            nom_corbeille = nom + " (" + str(compteur) + ")"
            compteur += 1
        modeles_supprimes[nom_corbeille] = modeles_projet.pop(nom)
        if nom in positions_boites_modeles:
            positions_boites_supprimees[nom_corbeille] = positions_boites_modeles.pop(nom)
        if nom in couleurs_boites_modeles:
            couleurs_boites_supprimees[nom_corbeille] = couleurs_boites_modeles.pop(nom)
        if nom_modele_actif == nom:
            nom_modele_actif = None
            for cle in _MODEL_STATE_KEYS:
                globals()[cle] = None
        if _panneau_config.get("nom") == nom:
            _fermer_panneau_config()
        selection_courante["nom"] = None
        _historique_valider("Delete model « " + str(nom) + " »")
        construire_boites_modeles()

    def restaurer_modele_supprime():
        if not analyse_autorisee():      # bouton desactive
            return
     
        if len(modeles_supprimes) == 0:
            messagebox.showinfo("Corbeille vide", "Aucun modele supprime pour le moment.")
            return

        fenetre_corbeille = Toplevel(parent)
        fenetre_corbeille.title(traduire("Modeles supprimes (corbeille)"))
        fenetre_corbeille.geometry("380x320")

        tk.Label(fenetre_corbeille, text = traduire("Selectionnez un modele a restaurer :"), font = ('arial', 12)).pack(pady = 10)

        liste_corbeille = tk.Listbox(fenetre_corbeille, fg = "gray10", font = ('arial', 12), relief = GROOVE, selectmode = tk.SINGLE)
        liste_corbeille.pack(fill = BOTH, expand = YES, padx = 10, pady = 5)

        noms_supprimes = list(modeles_supprimes.keys())
        for nom in noms_supprimes:
            liste_corbeille.insert(END, nom)

        def confirmer_restauration():
            selection = liste_corbeille.curselection()
            if not selection:
                messagebox.showinfo("Selection requise", "Veuillez selectionner un modele dans la liste.", parent = fenetre_corbeille)
                return
            nom = noms_supprimes[selection[0]]
            nom_final = nom
            if nom_final in modeles_projet:
                nom_final = simpledialog.askstring("Nom deja utilise", "Un modele actif porte deja le nom « " + nom + " ». Choisissez un nouveau nom pour le modele restaure :", initialvalue = nom + " (restaure)", parent = fenetre_corbeille)
                if nom_final is None or nom_final.strip() == "":
                    return
                if nom_final in modeles_projet:
                    messagebox.showwarning("Nom deja utilise", "Ce nom est deja utilise. Restauration annulee.", parent = fenetre_corbeille)
                    return
            modeles_projet[nom_final] = modeles_supprimes.pop(nom)
            if nom in couleurs_boites_supprimees:
                couleurs_boites_modeles[nom_final] = couleurs_boites_supprimees.pop(nom)
            if nom in positions_boites_supprimees:
                positions_boites_modeles[nom_final] = positions_boites_supprimees.pop(nom)
            _historique_valider("Restore model « " + str(nom_final) + " »")
            construire_boites_modeles()
            fenetre_corbeille.destroy()

        def supprimer_definitivement():
            selection = liste_corbeille.curselection()
            if not selection:
                messagebox.showinfo("Selection requise", "Veuillez selectionner un modele dans la liste.", parent = fenetre_corbeille)
                return
            nom = noms_supprimes[selection[0]]
            confirmation = messagebox.askyesno("Suppression definitive", "Supprimer definitivement « " + nom + " » ?\n\nCette action est irreversible.", parent = fenetre_corbeille)
            if not confirmation:
                return
            del modeles_supprimes[nom]
            couleurs_boites_supprimees.pop(nom, None)
            positions_boites_supprimees.pop(nom, None)
            fenetre_corbeille.destroy()
            restaurer_modele_supprime()

        cadre_boutons_corbeille = tk.Frame(fenetre_corbeille)
        cadre_boutons_corbeille.pack(pady = 10)
        tk.Button(cadre_boutons_corbeille, text = "Restaurer", command = confirmer_restauration, default = "active").pack(side = LEFT, padx = 5)
        tk.Button(cadre_boutons_corbeille, text = "Supprimer definitivement", command = supprimer_definitivement).pack(side = LEFT, padx = 5)
        ## clavier : Entree = Restaurer le modele selectionne (le 1er est preselectionne), Echap = fermer
        if len(noms_supprimes) != 0:
            liste_corbeille.selection_set(0)
            liste_corbeille.activate(0)
        fenetre_corbeille.bind("<Return>", lambda event: confirmer_restauration())
        fenetre_corbeille.bind("<Escape>", lambda event: fenetre_corbeille.destroy())
        liste_corbeille.bind("<Double-Button-1>", lambda event: confirmer_restauration())
        fenetre_corbeille.after(50, liste_corbeille.focus_force)

 
    button_Settings.config(command = ouvrir_modele_selectionne)

    button_Rename.config(command = renommer_modele_selectionne)
    button_Delete_model.config(command = supprimer_modele_selectionne)

    button_Restore.config(command = restaurer_modele_supprime)


def save_Model (event, parent):
    if not analyse_autorisee():      # bouton desactive : un evenement lie a la main peut quand meme arriver
        return
    if not verifier_projet_ouvert():
        return

    inf_project = messagebox.askyesno('Ask Question', message= 'Voulez-vous definir un nouveau modele ?')
    if inf_project == True:
       nom = simpledialog.askstring("Nouveau modele", "Nom du nouveau modele", initialvalue= "New model")
       if nom is not None and nom in modeles_projet:
           messagebox.showwarning("Nom deja utilise", "Ce nom est deja utilise. Choisissez un nom different, ou ouvrez ce modele depuis la liste.")
           choice_model(event, parent)
       elif nom is not None:
           creer_nouveau_modele(nom)
           ## Choix (facultatif) de la couleur de la boite de ce nouveau modele
           couleur_choisie = colorchooser.askcolor(color = "white", title = "Choisir une couleur pour la boite « " + nom + " »")
           if couleur_choisie is not None and couleur_choisie[1] is not None:
               couleurs_boites_modeles[nom] = couleur_choisie[1]
           _historique_valider("New model « " + str(nom) + " »")
           messagebox.showinfo("Nouveau modele", "Nouveau modele « " + nom + " » cree.\n\nConfigurez-le depuis la liste ci-dessous.")
           choice_model(event, parent)



# ==========================================================================
# ==  17. Graphiques descriptifs (box plot, histogramme, 3D, titres)      ==
# ==========================================================================
def my_sort (my_data_sort, check_columns):

  
    my_unique = sort(unique(np.array(my_data_sort[check_columns])))
    return my_unique


def chech_box_plot (event, parent, my_data_boxplot, var_check1, var_check2, var_list, categorical_feature, check_average, check_method, check_mean_boxplot_size, check_mean_color, check_line_width):

     global  plot_test_box, data_stat_boxplot, graph_window_representation_widget, dernier_x_boxplot, dernier_y_boxplot, dernier_color_boxplot
     data_stat_boxplot = pd.DataFrame()
     my_data = my_data_boxplot.copy()
     my_colors_palettes_1 = my_colors_palettes.copy()
     del my_colors_palettes_1[0]
     if check_method.current() != 0:
         if graph_window_representation_widget is not None and graph_window_representation_widget.winfo_exists():
             graph_window_representation_widget.destroy()
         windows_boxplot = Canvas(parent, relief= GROOVE ,  background = 'gray80' , highlightthickness=0)
         graph_window_representation_widget = windows_boxplot
         if check_method.current() == 1:
                 var_index1 = var_check1.current()
                 var_index2 = var_check2.current()
                 control_value = [var_index1, var_index2]
                 control_value_test = []
                 for i in range(len(control_value)):
                     if control_value[i] != 0:
                         control_value_test.append(True)
                 if len(control_value_test) == 2:
                         var_selected1 = var_list[var_index1]
                         var_selected2 = var_list[var_index2]
                         dernier_x_boxplot = var_selected1
                         dernier_y_boxplot = var_selected2
                         cat_value_feature = categorical_feature.get()
                         xx = my_sort(my_data, var_selected2)
                         my_data = my_data.sort_values(by = [var_selected2], ascending = True)
                         my_data.index = range(len(my_data))
                         if cat_value_feature == True:
                             dp = np.array(my_data[var_selected2])
                             if type(dp[0]) != str:
                                 var_feature = []
                                 for i in range(len(my_data)):
                                     var_feature.append(str(dp[i]))
                                 my_data[var_selected2] = var_feature   
                         my_tile = ""
                         if boxplot_title_value is not None:
                             if boxplot_title_value != "Enter the title":
                                 my_tile = boxplot_title_value
                         my_x_label = var_selected2
                         if boxplot_label_x is not None:
                             if boxplot_label_x != "Enter the x label":
                                 my_x_label = boxplot_label_x
                         my_y_label = var_selected1
                         if boxplot_label_y is not None:
                             if boxplot_label_y != "Enter the y label":
                                 my_y_label = boxplot_label_y
                         fig = Figure(figsize=(10, 8), dpi = 80)
                         _enregistrer_figure_courante("Boxplot", fig)
                         figure_canvas = FigureCanvasTkAgg(fig, master = windows_boxplot) 
                         ax = fig.add_subplot()
                         my_palette = []
                         for i in range(len(unique(data_analy[var_selected2]))):
                               my_palette.append(my_colors_palettes_1[i])
                         if check_average.get() != True:
                            sns.boxplot(x = var_selected2, y = var_selected1, data = my_data, ax = ax, linewidth = float(check_line_width.get()) , hue = var_selected2, legend = False, palette = my_palette)
                         else:
                            sns.boxplot(x = var_selected2, y = var_selected1, data = my_data , ax = ax, hue = var_selected2, legend = False, palette = my_palette, linewidth = float(check_line_width.get()) , showmeans = True, meanprops = {"marker": "o", "markerfacecolor": check_mean_color.get(), "markersize": check_mean_boxplot_size.get(), "markeredgecolor": check_mean_color.get()})
                         ax.set_xlabel(my_x_label, fontsize= 15, labelpad = 7)
                         ax.set_ylabel(my_y_label, fontsize= 15, labelpad = 5)
                         ax.set_title(my_tile, fontsize= 18) 
                         ax.grid()
                         if my_check_pairwise.get() == True :
                             xxx = unique(my_data[var_selected2])
                             data_pairwise = []
                             for i in range(len(xxx)):
                                 dxx = my_data[my_data[var_selected2] == xxx[i]]
                                 data_pairwise.append(dxx[var_selected1])
                             significant_combinations = []
                             ls = list(range(1, len(xxx) + 1))
                             combinations = [(ls[x], ls[x + y]) for y in reversed(ls) for x in range((len(ls) - y))]
                             for combination in combinations:
                                 data1 = data_pairwise[combination[0] - 1]
                                 data2 = data_pairwise[combination[1] - 1]
                                 U, p = stats.mannwhitneyu(data1, data2, alternative='two-sided')
                                 significant_combinations.append([combination, p])
                             bottom, top = ax.get_ylim()
                             y_range = top - bottom 
                             for i, significant_combination in enumerate(significant_combinations):
                                 x1 = significant_combination[0][0] - 1
                                 x2 = significant_combination[0][1] - 1
                                 level = len(significant_combinations) - i
                                 bar_height = (y_range * 0.08 * level) + top
                                 bar_tips = bar_height - (y_range * 0.02)
                                 ax.plot([x1, x1, x2, x2], [bar_tips, bar_height, bar_height, bar_tips], lw = float(check_line_width.get()), c = 'k')
                                 p = significant_combination[1]
                                 label_pvalue = ""
                                 if p < 0.001:
                                     label_pvalue = format(p, ".1e")
                                 else:
                                     label_pvalue = str(round(p, 3))
                                 color_pvalue = ""
                                 if p < 0.05 :
                                     color_pvalue = 'black'
                                 else:
                                     color_pvalue = 'blue'
                                 text_height = bar_height + (y_range * 0.001)
                                 ax.text((x1 + x2) * 0.5, text_height, label_pvalue, ha = 'center', va = 'bottom', c = color_pvalue)
                         figure_canvas.get_tk_widget().place(relx =0, rely = 0, relwidth = 1, relheight = 0.7)
                         _appliquer_zoom_representation(ax, "boxplot")
                         _activer_glissement_representation(ax, figure_canvas, "boxplot")

                         def _sur_changement_zoom_representation_box (valeur_curseur):
                             zoom_representation_par_type["boxplot"] = _curseur_vers_zoom_representation(valeur_curseur)
                             chech_box_plot(event, parent, my_data_boxplot, var_check1, var_check2, var_list, categorical_feature, check_average, check_method, check_mean_boxplot_size, check_mean_color, check_line_width)

                         scale.configure(command = _sur_changement_zoom_representation_box)
                         scale.set(_zoom_vers_curseur_representation(zoom_representation_par_type.get("boxplot", 1.0)))
                         plot_test_box = fig
                         if my_check_stat_box.get() == True:
                             dstat1 = stats.kruskal(*[group[var_selected1].values for name, group in my_data.groupby(var_selected2)]) 
                             dstat = dstat1[1]
                             stat_value = ""
                             if dstat <= 0.0001:
                                 stat_value = "< " + str(0.0001)
                             else:
                                 stat_value = str(round(dstat, 4))
                             font_size = ('Helvetica 16 bold italic')
                             dx = my_data[var_selected1]
                             Q1 = int(round(np.percentile(dx, 25)))
                             M = int(round(median(dx)))
                             Q3 = int(round(np.percentile(dx, 75)))
                             dstat1 = shapiro(np.array(my_data[var_selected1]))
                             dstat = dstat1[1]
                             dd = pd.DataFrame({"mean (std)": [str(round(mean(dx), 1)) + " (" + str(round(std(dx), 1)) + ")"], "min": [str(round(min(dx), 1))], "max": [str(round(max(dx), 1))], "median (Q1; Q3)": [str(M) + " (" + str(Q1) + "; "  +  str(Q3) + ")"], "p-value": [stat_value], "test name": ["Kruskal-Wallis"]})
                             modalitys = []
                             for k in range(len(xx)):
                                 modalitys.append(str(xx[k]))
                             rows_listx = []
                             for i in range(len(modalitys)):
                                 df1 = my_data[my_data[var_selected2] == modalitys[i]]
                                 dx = df1[var_selected1]
                                 dstatx = shapiro(df1[var_selected1])[1]
                                 stat_valuex = ""
                                 if dstatx <= 0.0001:
                                     stat_valuex = "< " + str(0.0001)
                                 else:
                                     stat_valuex = str(round(dstatx, 4))
                                 Q1 = int(round(np.percentile(dx, 25)))
                                 M = int(round(median(dx)))
                                 Q3 = int(round(np.percentile(dx, 75)))
                                 rows_new = {"mean (std)": str(round(mean(dx), 1)) + " (" + str(round(std(dx), 1)) + ")", "min": str(round(min(dx), 1)), "max": str(round(max(dx), 1)), "median (Q1; Q3)": str(M) + " (" + str(Q1) + "; "  +  str(Q3) + ")", "p-value": stat_valuex, "test name": "Shapiro-Wilk"}
                                 rows_listx.append(rows_new)
                             df = pd.DataFrame(rows_listx)
                             ddf = pd.concat([dd, df])
                             tabe_index = np.concatenate([np.array(["All"]), np.array(modalitys)])
                             ddf.insert(0, "", tabe_index)
                             ddf = pd.DataFrame(ddf)
                             data_stat_boxplot = ddf
                             windows_stat_boxplot = Frame(windows_boxplot, background = "white", borderwidth = 0)
                             stat_boxplot = Table(windows_stat_boxplot, dataframe = ddf)
                             stat_boxplot.show()
                             windows_stat_boxplot.place(relx=0.001, rely= 0.7, relwidth= 1, relheight= 0.3)
         else:
                     var_index1 = var_check1.current()
                     var_index2 = var_check2.current()
                     color_index = color_column.current()
                     control_value = [var_index1, var_index2, color_index]
                     control_value_test = []
                     for i in range(len(control_value)):
                         if control_value[i] != 0:
                             control_value_test.append(True)
                     if len(control_value_test) == 3:
                         var_selected1 = var_list[var_index1]
                         var_selected2 = var_list[var_index2]
                         color_selected = var_list[color_index]
                         dernier_x_boxplot = var_selected1
                         dernier_y_boxplot = var_selected2
                         dernier_color_boxplot = color_selected
                         cat_value_feature = categorical_feature.get()
                         xx = my_sort(my_data, var_selected2)
                         my_data = my_data.sort_values(by = [var_selected2], ascending = False)
                         my_data.index = range(len(my_data))
                         if cat_value_feature == True:
                             dp = np.array(my_data[var_selected2])
                             if type(dp[0]) != str:
                                 var_feature = []
                                 for i in range(len(my_data)):
                                     var_feature.append(str(dp[i]))
                                 my_data[var_selected2] = var_feature   
                         my_tile = ""
                         if boxplot_title_value is not None:
                             if boxplot_title_value != "Enter the title":
                                 my_tile = boxplot_title_value
                         my_x_label = var_selected2
                         if boxplot_label_x is not None:
                             if boxplot_label_x != "Enter the x label":
                                 my_x_label = boxplot_label_x
                         my_y_label = var_selected1
                         if boxplot_label_y is not None:
                             if boxplot_label_y != "Enter the y label":
                                 my_y_label = boxplot_label_y
                         fig = Figure(figsize=(10, 4), dpi = 80)
                         _enregistrer_figure_courante("Boxplot", fig)
                         figure_canvas = FigureCanvasTkAgg(fig, master = windows_boxplot) 
                         ax = fig.add_subplot()
                         my_palette = []
                         for i in range(len(unique(data_analy[color_selected]))):
                             my_palette.append(my_colors_palettes_1[i])
                         if check_average.get() != True:
                             sns.boxplot(x = var_selected2, y = var_selected1, hue = color_selected, palette = my_palette, data = data_analy, linewidth = float(check_line_width.get()), ax= ax)
                         else:
                             sns.boxplot(x = var_selected2, y = var_selected1, hue = color_selected, palette = my_palette, data = data_analy , ax= ax, linewidth = float(check_line_width.get()), showmeans = True, meanprops = {"marker": "o", "markerfacecolor": check_mean_color.get(), "markersize": check_mean_boxplot_size.get(), "markeredgecolor": check_mean_color.get()})
                         ax.set_xlabel(my_x_label, fontsize= 15, labelpad = 7)
                         ax.set_ylabel(my_y_label, fontsize= 15, labelpad = 7)
                         ax.set_title(my_tile, fontsize= 18)  
                         sns.move_legend(ax, "lower center", bbox_to_anchor = (0.5, 1.05), fontsize = 13, ncol = int(my_ncol.get()), title = color_selected, frameon = False)
                         ax.grid()
                         figure_canvas.get_tk_widget().place(relx =0, rely = 0, relwidth = 1, relheight = 1)
                         _appliquer_zoom_representation(ax, "boxplot")
                         _activer_glissement_representation(ax, figure_canvas, "boxplot")

                         def _sur_changement_zoom_representation_box (valeur_curseur):
                             zoom_representation_par_type["boxplot"] = _curseur_vers_zoom_representation(valeur_curseur)
                             chech_box_plot(event, parent, my_data_boxplot, var_check1, var_check2, var_list, categorical_feature, check_average, check_method, check_mean_boxplot_size, check_mean_color, check_line_width)

                         scale.configure(command = _sur_changement_zoom_representation_box)
                         scale.set(_zoom_vers_curseur_representation(zoom_representation_par_type.get("boxplot", 1.0)))
                         plot_test_box = fig
         windows_boxplot.place(relx=0, rely = 0, relwidth = 1, relheight= 1)


def chech_plot_histogram (event, parent, my_data, var_check, color_check, density_ckeck, var_list, color_list, check_value, check_stat_value):

    global  plot_test_hist, graph_window_representation_widget, dernier_var_histogram, dernier_couleur_histogram, dernier_densite_histogram
    if graph_window_representation_widget is not None and graph_window_representation_widget.winfo_exists():
        graph_window_representation_widget.destroy()
    windows_hist = Canvas(parent, relief= GROOVE ,  background = 'gray80' , highlightthickness=0)
    graph_window_representation_widget = windows_hist
    var_index = var_check.current()
    var_selected = var_list[var_index]
    color_index = color_check.current()
    hist_color = color_list[color_index]
    density_index = density_ckeck.current()
    density_color = color_list[density_index]
    dernier_var_histogram = var_selected
    dernier_couleur_histogram = hist_color
    dernier_densite_histogram = density_color
    x = int(round(0.6*parent.winfo_width(), 0))
    my_tile = ""
    if histogram_title_value is not None:
        if histogram_title_value != "Enter the title":
            my_tile = histogram_title_value
    my_x_label = var_selected
    if histogram_x_value is not None:
        if histogram_x_value != "Enter the x label":
            my_x_label = histogram_x_value
    my_y_label = "density"
    if histogram_y_value is not None:
        if histogram_y_value != "Enter the y label":
            my_y_label = histogram_y_value
    fig = Figure(figsize=(12, 8), dpi = 80)
    _enregistrer_figure_courante("Histogramme", fig)
    figure_canvas = FigureCanvasTkAgg(fig, master = windows_hist) 
    ax = fig.add_subplot()
    if  check_value.get() == True:
        sns.histplot(data = my_data, x = var_selected, stat = "density", color = hist_color, kde = True, ax = ax, line_kws = {'linestyle':'solid', 'linewidth':'1.5'}).lines[0].set_color(density_color)    
    else:
        sns.histplot(data = my_data, x = var_selected, ax = ax, color = hist_color , stat="density")
    ax.set_xlabel(my_x_label, fontsize= 15, labelpad = 7)
    ax.set_ylabel(my_y_label, fontsize= 15, labelpad = 7)
    ax.set_title(my_tile, fontsize= 18)  
    ax.grid()
    figure_canvas.get_tk_widget().place(relx =0, rely = 0, relwidth = 1, relheight = 0.7)
    _appliquer_zoom_representation(ax, "histogram")
    _activer_glissement_representation(ax, figure_canvas, "histogram")

    def _sur_changement_zoom_representation_hist (valeur_curseur):
        zoom_representation_par_type["histogram"] = _curseur_vers_zoom_representation(valeur_curseur)
        chech_plot_histogram(event, parent, my_data, var_check, color_check, density_ckeck, var_list, color_list, check_value, check_stat_value)

    scale.configure(command = _sur_changement_zoom_representation_hist)
    scale.set(_zoom_vers_curseur_representation(zoom_representation_par_type.get("histogram", 1.0)))
    plot_test_hist = fig
    if check_stat_value.get() == True:
            windows_stat_hist = Canvas(windows_hist, background="white")
            windows_stat_hist.place(relx=0.001, rely= 0.7, relwidth= 1, relheight= 0.5)
            testx_values_title = Label(windows_stat_hist, text= "Descriptive statistics" , font=('Helvetica 16 bold italic'), fg = "gray20", background= "white")
            testx_values_title.place(relx = 0.02 , rely = 0.12, relwidth = 0.25, relheight= 0.1)
            dd = my_data[var_selected]
            dstat1 = shapiro(np.array(my_data[var_selected]))
            dstat = dstat1[1]
            stat_value = ""
            if dstat <= 0.0001:
                stat_value = "< " + str(0.0001)
            else:
                stat_value = str(dstat)
            font_size = ('Helvetica 16 bold italic')
            ## Les labeles
            testx_label1 = Label(windows_stat_hist, text="mean (std)" , font = font_size, fg = "gray20", background= "white")
            testx_label1.place(relx= 0.03, rely = 0.25 , relwidth = 0.2, relheight= 0.1)
            testx_label2 = Label(windows_stat_hist, text="min" , font = font_size, fg = "gray20", background= "white")
            testx_label2.place(relx= 0.25, rely = 0.25 , relwidth = 0.1, relheight= 0.1)
            testx_label3 = Label(windows_stat_hist, text="max" , font = font_size, fg = "gray20", background= "white")
            testx_label3.place(relx= 0.39, rely = 0.25 , relwidth = 0.1, relheight= 0.1)
            testx_label4 = Label(windows_stat_hist, text="mediane (Q1; Q3)" , font = font_size, fg = "gray20", background= "white")
            testx_label4.place(relx= 0.55, rely = 0.25 , relwidth = 0.2, relheight= 0.1)
            testx_label4 = Label(windows_stat_hist, text="p-value" , font = font_size, fg = "gray20", background= "white")
            testx_label4.place(relx= 0.795, rely = 0.25 , relwidth = 0.2, relheight= 0.1)
            Q1 = int(round(np.percentile(dd, 25)))
            M = int(round(median(dd)))
            Q3 = int(round(np.percentile(dd, 75)))
            ## Valeurs des statistiques
            testx_values = Label(windows_stat_hist, text= str(round(mean(dd), 1))  + " (" + str(round(std(dd), 1))  + ")" , font = font_size, fg = "gray20", background= "white")
            testx_values.place(relx = 0.03 , rely = 0.4, relwidth = 0.2, relheight= 0.1)
            testx_values_min = Label(windows_stat_hist, text= str(round(min(dd), 1)) , font = font_size, fg = "gray20", background= "white")
            testx_values_min.place(relx = 0.25 , rely = 0.4, relwidth = 0.1, relheight= 0.1)
            testx_values_max = Label(windows_stat_hist, text= str(round(max(dd), 1)) , font = font_size, fg = "gray20", background= "white")
            testx_values_max.place(relx = 0.39 , rely = 0.4, relwidth = 0.1, relheight= 0.1)
            testx_values_mediane = Label(windows_stat_hist, text= str(Q1) + " (" + str(M) + "; "  +  str(Q3) + ")" , font = font_size, fg = "gray20", background= "white")
            testx_values_mediane.place(relx = 0.55 , rely = 0.4, relwidth = 0.2, relheight= 0.1)
            testx_values_pvalue = Label(windows_stat_hist, text=  stat_value, font = font_size, fg = "gray20", background= "white")
            testx_values_pvalue.place(relx = 0.795 , rely = 0.4, relwidth = 0.2, relheight= 0.1)
            ## Barre horizontale
            separator1 = tk.Frame(windows_stat_hist, bd = 10, relief='sunken', height = 4)
            separator1.place(relx=0.03 , rely = 0.25 , relwidth= 0.92)
            separator2 = tk.Frame(windows_stat_hist, bd=10, relief='sunken', height = 4)
            separator2.place(relx=0.03 , rely = 0.38 , relwidth= 0.92)
            separator3 = tk.Frame(windows_stat_hist, bd = 10, relief = 'sunken', height = 4)
            separator3.place(relx=0.03 , rely = 0.51 , relwidth= 0.92)
    windows_hist.place(relx =0, rely = 0, relwidth = 1, relheight = 1)


def my_boxplot_parameters(event, parent_graph, parent_parameters, my_data, check_var_method, check_features):

    global my_check_stat_box, color_column, my_ncol, my_check_pairwise, dernier_methode_boxplot
    color_column = None
    dernier_methode_boxplot = check_var_method.get()
    features_windows = Canvas(parent_parameters, relief= GROOVE ,  background = 'gray90' , highlightthickness=0)
    if check_var_method.current() != 0:
        list_color = ['blue', 'black','brown','red','orange','yellow', 'green','purple','grey','white']
        label_var1 = Label(features_windows, text="Target column ", font=('arial', 15), fg= "gray30", background= "gray90") 
        label_var2 = Label(features_windows, text="Feature column", font=('arial', 15), fg= "gray30", background= "gray90") 
        label_var1.place(relx=0.03, rely= 0.05, relwidth= 0.3)
        label_var2.place(relx=0.03, rely= 0.15, relwidth= 0.3)
        var_column1 = ttk.Combobox(features_windows, values = check_features)
        var_column2 = ttk.Combobox(features_windows, values = check_features)
        var_column1.current(_index_restaure(check_features, dernier_x_boxplot))
        var_column2.current(_index_restaure(check_features, dernier_y_boxplot))
        var_column1.place(relx=0.44, rely= 0.05, relwidth= 0.5)
        var_column2.place(relx=0.44, rely= 0.15, relwidth= 0.5)
        my_check_str_feature = BooleanVar(value = True)
        check_bouton_str_feature = Checkbutton(features_windows,  text ="Categorized all", variable = my_check_str_feature, font=('arial', 15), fg = "gray30", background = 'gray90')
        check_bouton_str_feature.place(relx=0.5, rely= 0.2, relwidth= 0.3)
        title_boxplot_def  = BooleanVar(value = False)
        title_boxplot_info = Checkbutton(features_windows,  text ="Define chart title", variable = title_boxplot_def, command = lambda : my_title_boxplot(parent_parameters, title_boxplot_def), font=('arial', 15), fg= "gray30", background = 'gray90')
        title_boxplot_info.place(relx = 0.05, rely= 0.4, relwidth= 0.35)
        mean_size = Label(features_windows, text="Mean size", font=('arial', 15), fg = "gray30", bg= "gray90")
        mean_size.place(relx=0.61, rely= 0.51, relwidth= 0.3)
        mean_boxplot_size_default = StringVar()
        mean_boxplot_size_default.set("7")
        mean_boxplot_size = Spinbox(features_windows, from_ = 1 , to = 1000 , increment = 1, justify = CENTER , textvariable = mean_boxplot_size_default)
        mean_boxplot_size.place(relx = 0.67, rely= 0.55, relwidth= 0.3) 
        label_mean_color = Label(features_windows, text="Mean color", font=('arial', 15), fg= "gray30", background= "gray90") 
        label_mean_color.place(relx=0.61, rely= 0.65, relwidth= 0.3)
        mean_color = ttk.Combobox(features_windows, values = my_colors_palettes)
        mean_color.current(0)
        mean_color.place(relx=0.67, rely= 0.69, relwidth= 0.3)
        line_width = Label(features_windows, text="Line width         ", font=('arial', 15), fg = "gray30", bg= "gray90")
        line_width.place(relx=0.03, rely= 0.51, relwidth= 0.3)
        line_width_default = StringVar()
        line_width_default.set("1")
        line_width = Spinbox(features_windows, from_ = 0 , to = 1000 , increment = 0.1, justify = CENTER , textvariable = line_width_default)
        line_width.place(relx = 0.05, rely= 0.55, relwidth= 0.3) 
        if check_var_method.current() == 1:
                my_check_pairwise = BooleanVar(value = False)
                my_check_stat_box =  BooleanVar(value = True)
                check_bouton_densty = Checkbutton(features_windows,  text ="Pairwise comp", variable = my_check_pairwise, font=('arial', 15), fg = "gray30", background = 'gray90')
                check_bouton_stat = Checkbutton(features_windows,  text ="Stat", variable = my_check_stat_box, font=('arial', 15), fg= "gray30", background = 'gray90')
                check_bouton_densty.place(relx=0.05, rely= 0.3, relwidth= 0.3)
                check_bouton_stat.place(relx=0.67, rely= 0.3, relwidth = 0.14)
                download_p = PhotoImage(file = my_path('figures/Export_data_p.png'))
                bouton_box_Export = Button(features_windows, image = download_p , font=('arial', 14), text = "Export stat", command = lambda : my_save_sat_view(data_stat_boxplot), compound = tk.TOP,  borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
                bouton_box_Export.place(relx= 0.39, rely= 0.83, width= 100, height= 60)
        else:
            color_graph = Label(features_windows, text="Graphic color   ", font=('arial', 15), fg= "gray30", bg= "gray90")
            color_graph.place(relx=0.03, rely= 0.3, relwidth= 0.3)
            color_column = ttk.Combobox(features_windows, values = check_features)
            color_column.current(_index_restaure(check_features, dernier_color_boxplot))
            color_column.place(relx=0.44, rely= 0.3, relwidth= 0.5)
            ncol_label = Label(features_windows, text="n columns by color title     ", font=('arial', 15), fg = "gray30", bg= "gray90")
            ncol_label.place(relx=0.03, rely= 0.65, relwidth= 0.45)
            ncol_default = StringVar()
            ncol_default.set("4")
            my_ncol = Spinbox(features_windows, from_ = 1 , to = 100 , increment = 1, justify = CENTER , textvariable = ncol_default)
            my_ncol.place(relx = 0.05, rely= 0.69, relwidth= 0.3) 
        my_check_average = BooleanVar(value = True)
        check_bouton_average = Checkbutton(features_windows,  text ="Display mean", variable = my_check_average, command = lambda : chech_box_plot(event, parent_graph, my_data , var_column1, var_column2, check_features, my_check_str_feature, my_check_average, check_var_method, mean_boxplot_size, mean_color, line_width), font=('arial', 15), fg= "gray30", background = 'gray90')
        check_bouton_average.place(relx=0.67, rely= 0.4, relwidth = 0.3)
        var_column1.bind("<<ComboboxSelected>>", lambda event: chech_box_plot(event, parent_graph, my_data , var_column1, var_column2, check_features, my_check_str_feature, my_check_average, check_var_method, mean_boxplot_size, mean_color, line_width))
        var_column2.bind("<<ComboboxSelected>>", lambda event: chech_box_plot(event, parent_graph, my_data , var_column1, var_column2, check_features, my_check_str_feature, my_check_average, check_var_method, mean_boxplot_size, mean_color, line_width))
        mean_color.bind("<<ComboboxSelected>>", lambda event:chech_box_plot(event, parent_graph, my_data , var_column1, var_column2, check_features, my_check_str_feature, my_check_average, check_var_method, mean_boxplot_size, mean_color, line_width))
        check_bouton_str_feature.bind("<<CCheckbuttonSelected>>", lambda event: chech_box_plot(event, parent_graph, my_data , var_column1, var_column2, check_features, my_check_str_feature, my_check_average, check_var_method, mean_boxplot_size, mean_color, line_width))
        image_box_view = PhotoImage(file=my_path('figures/apercu_p.png'))
        save_graph_p = PhotoImage(file = my_path('figures/save_graph_p.png'))
        bouton_box_view = Button(features_windows, image = image_box_view , font=('arial', 14), text = "View plot", compound = tk.TOP,  borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
        bouton_box_view.place(relx= 0.05, rely= 0.83, width= 100, height= 60)
        bouton_box_view.bind("<ButtonRelease>", lambda event: chech_box_plot(event, parent_graph, my_data , var_column1, var_column2, check_features, my_check_str_feature, my_check_average, check_var_method, mean_boxplot_size, mean_color, line_width)) 
        bouton_save_graph_box = Button(features_windows, image = save_graph_p , font=('arial', 14), text = "Save plot",  compound = tk.TOP,  borderwidth = 0, relief = FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
        bouton_save_graph_box.place(relx= 0.73, rely= 0.83, width= 100, height= 60)
        bouton_save_graph_box.bind("<ButtonRelease>", lambda event: my_save_plot_file(plot_test_box)) 
        ## Si la selection restauree est deja complete (X, Y, et Couleur pour
        ## le trivarie), on redessine tout de suite le graphique
        selection_complete = var_column1.current() != 0 and var_column2.current() != 0
        if check_var_method.current() != 1:
            selection_complete = selection_complete and color_column.current() != 0
        if selection_complete:
            chech_box_plot(event, parent_graph, my_data, var_column1, var_column2, check_features, my_check_str_feature, my_check_average, check_var_method, mean_boxplot_size, mean_color, line_width)
    features_windows.place(relx = 0, rely = 0.1, relwidth = 1, relheight = 0.9)


def _index_restaure (valeurs, dernier):
    ''' Retrouve l'index, dans la liste "valeurs" (colonnes d'une combobox),
    de la derniere valeur choisie ("dernier") avant de quitter l'onglet
    "Plots" ; si elle n'existe plus (donnees changees) ou n'a jamais ete
    definie, revient au premier element (le texte "Select column"). '''
    if dernier is not None and dernier in valeurs:
        return valeurs.index(dernier)
    return 0


def _detruire_widgets_differe (widgets):
    ''' Detruit les widgets donnes apres la fin du traitement de l'evenement
    en cours, plutot qu'immediatement. Necessaire quand la destruction est
    declenchee par le clic d'un bouton qui vit lui-meme a l'interieur du
    widget a detruire (bouton tkmacosx) : le detruire tout de suite fait
    planter le traitement interne du clic ("bad window path name"), car
    celui-ci continue de s'executer juste apres le relachement du bouton. '''
    for widget in widgets:
        if widget is not None and widget.winfo_exists():
            widget.after(0, widget.destroy)


def Box_plot_view (event, parent, my_data, var):

    global var_method_graphe, windows_plot_widgets
    _detruire_widgets_differe(windows_plot_widgets)
    graph = Canvas(parent, relief= GROOVE ,  background = 'gray80' , highlightthickness=0)
    features = Canvas(parent, relief= GROOVE, background = 'gray90', highlightthickness=0)
    windows_plot_widgets = (graph, features)
    col_test = ["Choose your analysis method" ,"Bivariate", "Trivariate"]
    label_method_graphe = Label(features, text="Analysis", font=('arial', 15), fg= "gray30", background= "gray90") 
    label_method_graphe.place(relx=0.03, rely= 0.05, relwidth= 0.2)
    var_method_graphe = ttk.Combobox(features, values = col_test)
    var_method_graphe.current(_index_restaure(col_test, dernier_methode_boxplot))
    var_method_graphe.place(relx=0.44, rely= 0.05, relwidth= 0.5)
    var_method_graphe.bind("<<ComboboxSelected>>", lambda event: my_boxplot_parameters(event, graph, features, my_data, var_method_graphe, var))
    graph.place(relx=0, rely=0, relwidth= 0.60, relheight=1)
    features.place(relx=0.6, rely=0, relwidth= 0.40, relheight=1)
    ## Si une methode a deja ete choisie precedemment, on reconstruit tout de
    ## suite le panneau de parametres correspondant (au lieu d'attendre que
    ## l'utilisateur reselectionne quelque chose dans la combobox)
    if var_method_graphe.current() != 0:
        my_boxplot_parameters(event, graph, features, my_data, var_method_graphe, var)


def Histograom_plot_view (event, parent, my_data, var):

    global windows_plot_widgets
    _detruire_widgets_differe(windows_plot_widgets)
    graph = Canvas(parent, relief= GROOVE ,  background = 'gray80' , highlightthickness=0)
    features = Canvas(parent, relief= GROOVE, background = 'gray90', highlightthickness=0)
    windows_plot_widgets = (graph, features)
    list_color = ['blue', 'black','brown','red','orange','yellow', 'green','purple','grey']
    label_var = Label(features, text="Target column", font=('arial', 15), background= "gray90") 
    color_graph = Label(features, text="Hist color       ", font=('arial', 15), bg= "gray90")
    label_density = Label(features, text="Density color", font=('arial', 15), bg= "gray90")
    label_var.place(relx=0.1, rely= 0.05, relwidth= 0.4)
    color_graph.place(relx=0.1, rely= 0.15, relwidth= 0.4)
    label_density.place(relx=0.1, rely= 0.25, relwidth= 0.4)
    var_column = ttk.Combobox(features, values = var)
    color_column = ttk.Combobox(features, values = list_color)
    density_column = ttk.Combobox(features, values = list_color)
    var_column.current(_index_restaure(var, dernier_var_histogram))
    color_column.current(_index_restaure(list_color, dernier_couleur_histogram) if dernier_couleur_histogram in list_color else 8)
    density_column.current(_index_restaure(list_color, dernier_densite_histogram))
    var_column.place(relx=0.47, rely= 0.05, relwidth= 0.5)
    color_column.place(relx=0.47, rely= 0.15, relwidth= 0.5)
    density_column.place(relx=0.47, rely= 0.25, relwidth= 0.5)
    my_check_densty  = BooleanVar(value=True)
    my_check_stat =  BooleanVar(value = True)
    check_bouton_densty = Checkbutton(features,  text ="density", variable = my_check_densty, command = lambda :  chech_plot_histogram(event, graph, my_data , var_column, color_column, density_column, var, list_color, my_check_densty, my_check_stat), font=('arial', 18), background = 'gray90')
    check_bouton_stat = Checkbutton(features,  text ="Stat", variable = my_check_stat, command = lambda :  chech_plot_histogram(event, graph, my_data , var_column, color_column, density_column, var, list_color, my_check_densty, my_check_stat), font=('arial', 18), background = 'gray90')
    check_bouton_densty.place(relx=0.2, rely= 0.37, width= 110, height= 30)
    check_bouton_stat.place(relx=0.6, rely= 0.37, width= 110, height= 30)
    title_histogram_def  = BooleanVar(value = False)
    title_histogram_info = Checkbutton(features,  text ="Define chart title", variable = title_histogram_def, command = lambda : my_histogram_title(features, title_histogram_def) , fg = "gray30", font=('arial', 15), background = 'gray90')
    title_histogram_info.place(relx = 0.2, rely= 0.5, relwidth= 0.35)
    var_column.bind("<<ComboboxSelected>>", lambda event: chech_plot_histogram(event, graph, my_data , var_column, color_column, density_column, var, list_color, my_check_densty, my_check_stat))
    density_column.bind("<<ComboboxSelected>>", lambda event: chech_plot_histogram(event, graph, my_data , var_column, color_column, density_column, var, list_color, my_check_densty, my_check_stat))
    color_column.bind("<<ComboboxSelected>>", lambda event: chech_plot_histogram(event, graph, my_data , var_column, color_column, density_column, var, list_color, my_check_densty, my_check_stat))
    check_bouton_densty.bind("<<CCheckbuttonSelected>>", lambda event: chech_plot_histogram(event, graph, my_data , var_column, color_column, density_column, var, list_color, my_check_densty, my_check_stat))
    check_bouton_stat.bind("<<CCheckbuttonSelected>>", lambda event: chech_plot_histogram(event, graph, my_data , var_column, color_column, density_column, var, list_color, my_check_densty, my_check_stat))
    image_box_view = PhotoImage(file=my_path('figures/apercu_p.png'))
    bouton_hist_view = Button(features, image = image_box_view , font=('arial', 14), text = "View plot",  compound = tk.TOP,  borderwidth = 0, relief = FLAT , borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
    bouton_hist_view.place(relx= 0.17, rely= 0.65, width= 130, height= 60)
    bouton_hist_view.bind("<ButtonRelease>", lambda event: chech_plot_histogram(event, graph, my_data, var_column, color_column, density_column, var, list_color, my_check_densty, my_check_stat)) 
    save_graph_p = PhotoImage(file = my_path('figures/save_graph_p.png'))
    bouton_save_graph_hist = Button(features, image = save_graph_p , font=('arial', 14), text = "Save plot",  compound = tk.TOP,  borderwidth = 0, relief = FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
    bouton_save_graph_hist.place(relx= 0.65, rely= 0.65, width= 130, height= 60)
    bouton_save_graph_hist.bind("<ButtonRelease>", lambda event: my_save_plot_file(plot_test_hist)) 
    ##image_density = PhotoImage(file='figures/density_p.png'))
    graph.place(relx=0, rely=0, relwidth= 0.60, relheight=1)
    features.place(relx=0.6, rely=0, relwidth= 0.40, relheight=1)
    ## Si une colonne cible a deja ete choisie precedemment, on redessine
    ## tout de suite l'histogramme (couleur/densite ont toujours une valeur
    ## par defaut valide, seule la colonne cible peut etre "Select column")
    if var_column.current() != 0:
        chech_plot_histogram(event, graph, my_data, var_column, color_column, density_column, var, list_color, my_check_densty, my_check_stat)


def chech_plot3d (event, parent, data, var_check1, var_check2, var_check3, color_check, var_list, check_plot3d_point_size):

    global  my_plot3d, graph_window_representation_widget, dernier_x_plot3d, dernier_y_plot3d, dernier_z_plot3d, dernier_color_plot3d
    if graph_window_representation_widget is not None and graph_window_representation_widget.winfo_exists():
        graph_window_representation_widget.destroy()
    var_index1 = var_check1.current()
    var_index2 = var_check2.current()
    var_index3 = var_check3.current()
    color_index = color_check.current()
    control_value = [var_index1, var_index2, var_index3, color_index]
    control_value_check = []
    for i in range(len(control_value)):
        if control_value[i] != 0:
            control_value_check.append(True)
    if len(control_value_check) == 4:
        var_selected1 = var_list[var_index1]
        var_selected2 = var_list[var_index2]
        var_selected3 = var_list[var_index3]
        color_selected = var_list[color_index]
        dernier_x_plot3d = var_selected1
        dernier_y_plot3d = var_selected2
        dernier_z_plot3d = var_selected3
        dernier_color_plot3d = color_selected
        my_tile = "" 
        if plot3d_title_value is not None:
            if plot3d_title_value != "Enter the title":
                my_tile = plot3d_title_value
        my_x_label = var_selected1
        if plot3d_x_value is not None:
            if plot3d_x_value != "Enter the x label":
                my_x_label = plot3d_x_value
        my_y_label = var_selected2
        if plot3d_y_value is not None:
            if plot3d_y_value != "Enter the y label":
                my_y_label = plot3d_y_value
        my_z_label = var_selected3
        if plot3d_z_value is not None:
            if plot3d_z_value != "Enter the z label":
                my_z_label = plot3d_z_value
        my_color_label = color_selected
        if plot3d_color_value is not None:
            if plot3d_color_value != "Enter the color title":
                my_color_label = plot3d_color_value
        fig = Figure(layout='constrained', dpi = 100)
        _enregistrer_figure_courante("Plot_3D", fig)
        figure_canvas = FigureCanvasTkAgg(fig, master = parent) 
        graph_window_representation_widget = figure_canvas.get_tk_widget()
        ax = fig.add_subplot(111, projection="3d")
        my_plot = ax.scatter(data[var_selected1], data[var_selected2], data[var_selected3], c = data[color_selected], cmap = 'turbo' , s = float(check_plot3d_point_size.get()))
        ax.set_xlabel(my_x_label, fontsize= 10, labelpad = 5)
        ax.set_ylabel(my_y_label, fontsize= 10, labelpad = 5)
        ax.set_zlabel(my_z_label, fontsize= 10, labelpad = 5, rotation = 90)
        ax.set_title(my_tile, fontsize= 15)  
        ax.grid()
        fig_colorbar = fig.colorbar(my_plot, ax = ax, shrink= 0.5, aspect = 15, fraction = 0.046, pad =0.2)
        fig_colorbar.set_label(my_color_label)
        ## Zoom (pilote par la barre verticale gauche) sur les 3 axes ; pas de
        ## glissement personnalise ici, la rotation a la souris est deja geree
        ## nativement par matplotlib sur les graphiques 3D
        facteur = zoom_representation_par_type.get("plot3d", 1.0)
        if facteur <= 0:
            facteur = 1.0
        for get_lim3d, set_lim3d in ((ax.get_xlim3d, ax.set_xlim3d), (ax.get_ylim3d, ax.set_ylim3d), (ax.get_zlim3d, ax.set_zlim3d)):
            borne_min, borne_max = get_lim3d()
            centre = (borne_min + borne_max) / 2
            demi_etendue = (borne_max - borne_min) / 2 / facteur
            set_lim3d(centre - demi_etendue, centre + demi_etendue)
        figure_canvas.get_tk_widget().place(relx =0, rely = 0, relwidth = 0.6, relheight = 1)
        my_plot3d = fig

        def _sur_changement_zoom_representation_3d (valeur_curseur):
            zoom_representation_par_type["plot3d"] = _curseur_vers_zoom_representation(valeur_curseur)
            chech_plot3d(event, parent, data, var_check1, var_check2, var_check3, color_check, var_list, check_plot3d_point_size)

        scale.configure(command = _sur_changement_zoom_representation_3d)
        scale.set(_zoom_vers_curseur_representation(zoom_representation_par_type.get("plot3d", 1.0)))


def my_title_plot3d(parent, check_title):

    if check_title.get() != False:
        windows_scatter_title = Canvas(parent, background = 'cadetblue' , highlightthickness = 0)
        windows_scatter_title.place(relx = 0.03, rely = 0.03, relwidth = 0.94, relheight = 0.9)
        label_title = Label(windows_scatter_title, text = "Plot title   ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_title.place(relx= 0.03, rely= 0.02, relwidth= 0.3)
        zone_entry_title_default = StringVar()
        zone_entry_title_default.set("Enter the title")
        zone_entry_title = Entry(windows_scatter_title, textvariable = zone_entry_title_default, background = 'gray90', justify = CENTER )
        zone_entry_title.place(relx=0.1, rely = 0.06, relwidth = 0.8, relheight= 0.1)
        label_x = Label(windows_scatter_title, text = "X - axis   ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_x.place(relx= 0.03, rely= 0.2, relwidth= 0.3)
        zone_entry_x_default = StringVar()
        zone_entry_x_default.set("Enter the x label")
        zone_entry_x = Entry(windows_scatter_title, textvariable = zone_entry_x_default, background = 'gray90', justify = CENTER )
        zone_entry_x.place(relx=0.1, rely = 0.24, relwidth = 0.8, relheight= 0.1)
        label_y = Label(windows_scatter_title, text = "Y - axis   ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_y.place(relx= 0.03, rely= 0.38, relwidth= 0.3)
        zone_entry_y_default = StringVar()
        zone_entry_y_default.set("Enter the y label")
        zone_entry_y = Entry(windows_scatter_title, textvariable = zone_entry_y_default, background = 'gray90', justify = CENTER )
        zone_entry_y.place(relx=0.1, rely = 0.42, relwidth = 0.8, relheight= 0.1)
        label_z = Label(windows_scatter_title, text = "Z - axis   ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_z.place(relx= 0.03, rely= 0.56, relwidth= 0.3)
        zone_entry_z_default = StringVar()
        zone_entry_z_default.set("Enter the z label")
        zone_entry_z = Entry(windows_scatter_title, textvariable = zone_entry_z_default, background = 'gray90', justify = CENTER )
        zone_entry_z.place(relx=0.1, rely = 0.6, relwidth = 0.8, relheight= 0.1)
        label_color = Label(windows_scatter_title, text = "Color name", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_color.place(relx= 0.04, rely= 0.74, relwidth= 0.3)
        zone_entry_color_default = StringVar()
        zone_entry_color_default.set("Enter the color label")
        zone_entry_color = Entry(windows_scatter_title, textvariable = zone_entry_color_default, background = 'gray90', justify = CENTER )
        zone_entry_color.place(relx=0.1, rely = 0.78, relwidth = 0.8, relheight= 0.1)
        close_window_title_scatter = Button(windows_scatter_title, text = "Save", command = lambda : recup_my_title_plot3d(windows_scatter_title, zone_entry_title_default, zone_entry_x, zone_entry_y, zone_entry_z, zone_entry_color), font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
        close_window_title_scatter.place(relx= 0.4, rely= 0.93, width= 50, height= 20)


def recup_my_title_plot3d(parent, check_title_value, check_zone_entry_x, check_zone_entry_y, check_zone_entry_z, check_zone_entry_color):

    global plot3d_title_value, plot3d_x_value, plot3d_y_value, plot3d_z_value, plot3d_color_value
    plot3d_title_value = check_title_value.get()
    plot3d_x_value = check_zone_entry_x.get()
    plot3d_y_value = check_zone_entry_y.get()
    plot3d_z_value = check_zone_entry_z.get()
    plot3d_color_value = check_zone_entry_color.get()
    parent.destroy()


def Box_plot3d_view (event, parent, my_data, var):

    global windows_plot_widgets
    _detruire_widgets_differe(windows_plot_widgets)
    graph = Canvas(parent, relief= GROOVE ,  background = 'gray80' , highlightthickness=0)
    features = Canvas(parent, relief= GROOVE, background = 'gray90', highlightthickness=0)
    windows_plot_widgets = (graph, features)
    list_color = ['blue', 'black','brown','red','orange','yellow', 'green','purple','grey','white']
    label_var1 = Label(features, text="Column 1" + "            ", fg = "gray30", font=('arial', 15), background= "gray90") 
    label_var2 = Label(features, text="Column 2" + "            ", fg = "gray30", font=('arial', 15), background= "gray90") 
    label_var3 = Label(features, text="Column 3" + "            ", fg = "gray30", font=('arial', 15), background= "gray90") 
    color_graph = Label(features, text="Graphic color       ", fg = "gray30", font=('arial', 15), bg= "gray90")
    label_var1.place(relx=0.1, rely= 0.05, relwidth= 0.4)
    label_var2.place(relx=0.1, rely= 0.15, relwidth= 0.4)
    label_var3.place(relx=0.1, rely= 0.25, relwidth= 0.4)
    color_graph.place(relx=0.1, rely= 0.35, relwidth= 0.4)
    var_column1 = ttk.Combobox(features, values = var)
    var_column2 = ttk.Combobox(features, values = var)
    var_column3 = ttk.Combobox(features, values = var)
    color_column = ttk.Combobox(features, values = var)
    var_column1.current(_index_restaure(var, dernier_x_plot3d))
    var_column2.current(_index_restaure(var, dernier_y_plot3d))
    var_column3.current(_index_restaure(var, dernier_z_plot3d))
    color_column.current(_index_restaure(var, dernier_color_plot3d))
    var_column1.place(relx=0.47, rely= 0.05, relwidth= 0.5)
    var_column2.place(relx=0.47, rely= 0.15, relwidth= 0.5)
    var_column3.place(relx=0.47, rely= 0.25, relwidth= 0.5)
    color_column.place(relx=0.47, rely= 0.35, relwidth= 0.5)
    label_point_size = Label(features, text = "Point size", fg = "gray30", font=('arial', 15), background= "gray90") 
    label_point_size.place(relx = 0.1, rely = 0.5, relwidth = 0.3)
    plot3d_point_size_default =  DoubleVar()
    plot3d_point_size_default.set(10)
    plot3d_point_size = Spinbox(features, from_ = 1 , to = 1000 , increment = 1, justify = CENTER , textvariable = plot3d_point_size_default)
    plot3d_point_size.place(relx=0.47, rely= 0.5, relwidth= 0.5) 
    title_plot3d_def  = BooleanVar(value = False)
    title_plot3d_info = Checkbutton(features,  text ="Define chart title", variable = title_plot3d_def, command = lambda : my_title_plot3d(features, title_plot3d_def), fg = "gray30", font=('arial', 15), background = 'gray90')
    title_plot3d_info.place(relx = 0.15, rely= 0.63, relwidth= 0.4)
    var_column1.bind("<<ComboboxSelected>>", lambda event: chech_plot3d(event, parent, my_data , var_column1, var_column2, var_column3, color_column, var, plot3d_point_size))
    var_column2.bind("<<ComboboxSelected>>", lambda event: chech_plot3d(event, parent, my_data , var_column1, var_column2, var_column3, color_column, var, plot3d_point_size))
    var_column3.bind("<<ComboboxSelected>>", lambda event: chech_plot3d(event, parent, my_data , var_column1, var_column2, var_column3, color_column, var, plot3d_point_size))
    color_column.bind("<<ComboboxSelected>>", lambda event: chech_plot3d(event, parent, my_data , var_column1, var_column2, var_column3, color_column, var, plot3d_point_size))
    image_box_view = PhotoImage(file=my_path('figures/apercu_p.png'))
    save_graph_p = PhotoImage(file = my_path('figures/save_graph_p.png'))
    bouton_box_view = Button(features, image = image_box_view , font=('arial', 14), text = "View plot",  compound = tk.TOP,  borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
    bouton_box_view.place(relx= 0.17, rely= 0.8, width= 130, height= 60)
    bouton_box_view.bind("<ButtonRelease>", lambda event: chech_plot3d(event, parent, my_data, var_column1, var_column2, var_column3, color_column, var, plot3d_point_size)) 
    bouton_save_plot3d = Button(features, image = save_graph_p , font = ('arial', 14), text = "Save plot",  compound = tk.TOP,  borderwidth = 0, relief = FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
    bouton_save_plot3d.place(relx= 0.65, rely= 0.8, width= 130, height= 60)
    bouton_save_plot3d.bind("<ButtonRelease>", lambda event: my_save_plot_file(my_plot3d)) 
    graph.place(relx=0, rely=0, relwidth= 0.60, relheight=1)
    features.place(relx=0.6, rely=0, relwidth= 0.40, relheight=1)
    ## Si les 4 parametres restaures sont deja complets, on redessine tout de
    ## suite le graphique 3D
    if var_column1.current() != 0 and var_column2.current() != 0 and var_column3.current() != 0 and color_column.current() != 0:
        chech_plot3d(event, parent, my_data, var_column1, var_column2, var_column3, color_column, var, plot3d_point_size)


def my_histogram_title(parent, check_title):

    if check_title.get() != False:
        windows_histogram_title = Canvas(parent, background = 'cadetblue' , highlightthickness = 0)
        windows_histogram_title.place(relx = 0.1, rely = 0.1, relwidth = 0.8, relheight = 0.8)
        window_title = Label(windows_histogram_title, text = "Title window", fg = "gray30", font=('arial', 15), background= "cadetblue") 
        window_title.place(relx=0.06, rely=0.06, relwidth = 0.3)
        zone_entry_title_default = StringVar()
        zone_entry_title_default.set("Enter the title")
        zone_entry_title = Entry(windows_histogram_title, textvariable = zone_entry_title_default, background = 'gray90', justify = CENTER )
        zone_entry_title.place(relx=0.1, rely = 0.1, relwidth = 0.8, relheight= 0.12)
        window_x = Label(windows_histogram_title, text = "X - axis ", font=('arial', 15), fg = "gray30", background= "cadetblue") 
        window_x.place(relx=0.03, rely=0.31, relwidth = 0.3)
        zone_entry_x_default = StringVar()
        zone_entry_x_default.set("Enter the x label")
        zone_entry_x = Entry(windows_histogram_title, textvariable = zone_entry_x_default, background = 'gray90', justify = CENTER )
        zone_entry_x.place(relx=0.1, rely = 0.35, relwidth = 0.8, relheight= 0.12)
        window_y = Label(windows_histogram_title, text = "Y - axis ", fg = "gray30", font=('arial', 15), background= "cadetblue") 
        window_y.place(relx=0.03, rely=0.56, relwidth = 0.3)
        zone_entry_y_default = StringVar()
        zone_entry_y_default.set("Enter the y label")
        zone_entry_y = Entry(windows_histogram_title, textvariable = zone_entry_y_default, background = 'gray90', justify = CENTER )
        zone_entry_y.place(relx=0.1, rely = 0.6, relwidth = 0.8, relheight= 0.12)
        close_window_title_histogram = Button(windows_histogram_title, text = "Save", command = lambda : recup_my_title_histogram(windows_histogram_title, zone_entry_title_default, zone_entry_x, zone_entry_y), font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
        close_window_title_histogram.place(relx= 0.4, rely= 0.85, width= 50, height= 20)


def recup_my_title_histogram(parent, check_title_value, check_zone_entry_x, check_zone_entry_y):

    global histogram_title_value, histogram_x_value, histogram_y_value
    histogram_title_value = check_title_value.get()
    histogram_x_value = check_zone_entry_x.get()
    histogram_y_value = check_zone_entry_y.get()
    parent.destroy()


def my_title_boxplot(parent, check_title):

    if check_title.get() != False:
        windows_boxplot_title = Canvas(parent, background = 'cadetblue' , highlightthickness = 0)
        windows_boxplot_title.place(relx = 0.1, rely = 0.1, relwidth = 0.8, relheight = 0.6)
        window_title = Label(windows_boxplot_title, text = "Plot title     ", fg = "gray30", font=('arial', 15), background= "cadetblue") 
        window_title.place(relx=0.06, rely=0.04, relwidth = 0.3)
        zone_entry_title_default = StringVar()
        zone_entry_title_default.set("Enter the title")
        zone_entry_title = Entry(windows_boxplot_title, textvariable = zone_entry_title_default, background = 'gray90', justify = CENTER )
        zone_entry_title.place(relx=0.1, rely = 0.1, relwidth = 0.8, relheight= 0.12)
        window_x = Label(windows_boxplot_title, text = "X - axis     ", font=('arial', 15), fg = "gray30", background= "cadetblue") 
        window_x.place(relx=0.05, rely=0.29, relwidth = 0.3)
        zone_entry_x_default = StringVar()
        zone_entry_x_default.set("Enter the x label")
        zone_entry_x = Entry(windows_boxplot_title, textvariable = zone_entry_x_default, background = 'gray90', justify = CENTER )
        zone_entry_x.place(relx=0.1, rely = 0.35, relwidth = 0.8, relheight= 0.12)
        window_y = Label(windows_boxplot_title, text = "Y - axis     ", fg = "gray30", font=('arial', 15), background= "cadetblue") 
        window_y.place(relx=0.05, rely=0.54, relwidth = 0.3)
        zone_entry_y_default = StringVar()
        zone_entry_y_default.set("Enter the y label")
        zone_entry_y = Entry(windows_boxplot_title, textvariable = zone_entry_y_default, background = 'gray90', justify = CENTER )
        zone_entry_y.place(relx=0.1, rely = 0.6, relwidth = 0.8, relheight= 0.12)
        close_window_title_boxplot = Button(windows_boxplot_title, text = "Save", command = lambda : recup_my_title_boxplot(windows_boxplot_title, zone_entry_title_default, zone_entry_x_default, zone_entry_y_default), font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
        close_window_title_boxplot.place(relx= 0.4, rely= 0.85, width= 50, height= 20)


def recup_my_title_boxplot(parent, check_title_value, check_x_label, check_y_label):

    global boxplot_title_value, boxplot_label_x, boxplot_label_y
    boxplot_title_value = check_title_value.get()
    boxplot_label_x = check_x_label.get()
    boxplot_label_y = check_y_label.get()
    parent.destroy()


def my_title_min_max_plot(parent, check_title):

    if check_title.get() != False:
        windows_min_maxr_title = Canvas(parent, background = 'cadetblue' , highlightthickness = 0)
        windows_min_maxr_title.place(relx = 0.1, rely = 0.1, relwidth = 0.8, relheight = 0.8)
        label_title = Label(windows_min_maxr_title, text = "Plot title           ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_title.place(relx= 0.04, rely= 0.04, relwidth= 0.4)
        zone_entry_title_default = StringVar()
        zone_entry_title_default.set("Enter the title")
        zone_entry_title = Entry(windows_min_maxr_title, textvariable = zone_entry_title_default, background = 'gray90', justify = CENTER )
        zone_entry_title.place(relx=0.1, rely = 0.1, relwidth = 0.8, relheight= 0.15)
        label_X = Label(windows_min_maxr_title, text = "X-axis             ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_X.place(relx= 0.04, rely= 0.34, relwidth= 0.4)
        zone_entry_x_default = StringVar()
        zone_entry_x_default.set("Enter the x label")
        zone_entry_x = Entry(windows_min_maxr_title, textvariable = zone_entry_x_default, background = 'gray90', justify = CENTER )
        zone_entry_x.place(relx=0.1, rely = 0.4, relwidth = 0.8, relheight= 0.15)
        label_Y = Label(windows_min_maxr_title, text = "Y-axis             ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_Y.place(relx= 0.04, rely= 0.64, relwidth= 0.4)
        zone_entry_y_default = StringVar()
        zone_entry_y_default.set("Enter the y label")
        zone_entry_y = Entry(windows_min_maxr_title, textvariable = zone_entry_y_default, background = 'gray90', justify = CENTER )
        zone_entry_y.place(relx=0.1, rely = 0.7, relwidth = 0.8, relheight= 0.15)
        close_window_title_min_max = Button(windows_min_maxr_title, text = "Save", command = lambda : recup_my_title_min_max(windows_min_maxr_title, zone_entry_title_default, zone_entry_x, zone_entry_y), font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
        close_window_title_min_max.place(relx= 0.4, rely= 0.92, width= 50, height= 20)


def recup_my_title_min_max(parent, check_title_value, check_zone_entry_x, check_zone_entry_y):

    global min_max_title_value, min_max_x_value, min_max_y_value
    min_max_title_value = check_title_value.get()
    min_max_x_value = check_zone_entry_x.get()
    min_max_y_value = check_zone_entry_y.get()
    parent.destroy()


def my_title_scatter_plot(parent, check_title):

    if check_title.get() != False:
        windows_scatter_title = Canvas(parent, background = 'cadetblue' , highlightthickness = 0)
        windows_scatter_title.place(relx = 0.03, rely = 0.03, relwidth = 0.94, relheight = 0.8)
        label_title = Label(windows_scatter_title, text = "Plot title                  ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_title.place(relx= 0.06, rely= 0.06, relwidth= 0.4)
        zone_entry_title_default = StringVar()
        zone_entry_title_default.set("Enter the title")
        zone_entry_title = Entry(windows_scatter_title, textvariable = zone_entry_title_default, background = 'gray90', justify = CENTER )
        zone_entry_title.place(relx=0.1, rely = 0.1, relwidth = 0.8, relheight= 0.15)
        label_x = Label(windows_scatter_title, text = "X - axis         ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_x.place(relx= 0.06, rely= 0.3, relwidth= 0.3)
        zone_entry_x_default = StringVar()
        zone_entry_x_default.set("Enter the x label")
        zone_entry_x = Entry(windows_scatter_title, textvariable = zone_entry_x_default, background = 'gray90', justify = CENTER )
        zone_entry_x.place(relx=0.1, rely = 0.34, relwidth = 0.8, relheight= 0.1)
        label_y = Label(windows_scatter_title, text = "Y - axis          ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_y.place(relx= 0.06, rely= 0.5, relwidth= 0.3)
        zone_entry_y_default = StringVar()
        zone_entry_y_default.set("Enter the y label")
        zone_entry_y = Entry(windows_scatter_title, textvariable = zone_entry_y_default, background = 'gray90', justify = CENTER )
        zone_entry_y.place(relx=0.1, rely = 0.54, relwidth = 0.8, relheight= 0.1)
        label_color = Label(windows_scatter_title, text = "Legend name         ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_color.place(relx= 0.06, rely= 0.7, relwidth= 0.4)
        zone_entry_color_default = StringVar()
        zone_entry_color_default.set("Enter the color title")
        zone_entry_color = Entry(windows_scatter_title, textvariable = zone_entry_color_default, background = 'gray90', justify = CENTER )
        zone_entry_color.place(relx=0.1, rely = 0.74, relwidth = 0.8, relheight= 0.1)
        close_window_title_scatter = Button(windows_scatter_title, text = "Save", command = lambda : recup_my_title_scatter(windows_scatter_title, zone_entry_title_default, zone_entry_x, zone_entry_y, zone_entry_color), font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
        close_window_title_scatter.place(relx= 0.4, rely= 0.93, width= 50, height= 20)


def recup_my_title_scatter(parent, check_title_value, check_zone_entry_x, check_zone_entry_y, check_zone_entry_color):

    global scatter_title_value, scatter_x_value, scatter_y_value, scatter_color_value
    scatter_title_value = check_title_value.get()
    scatter_x_value = check_zone_entry_x.get()
    scatter_y_value = check_zone_entry_y.get()
    scatter_color_value = check_zone_entry_color.get()
    parent.destroy()



# ==========================================================================
# ==  18. Optimisation : parametres, calcul et graphique                  ==
# ==========================================================================
def lire_vecteur_reels (texte, virgule_decimale = False):
    ''' Convertit un texte saisi par l'utilisateur en vecteur de nombres reels
    (np.ndarray). Separateurs acceptes : virgule, point-virgule, espace ou
    retour a la ligne. Si un point-virgule est present, la virgule est lue
    comme separateur decimal (ex. "2,5 ; 3,1"). Avec virgule_decimale = True
    (cas d'une seule valeur attendue), la virgule est toujours un separateur
    decimal ("2,5" -> 2.5). Leve ValueError si une valeur n'est pas un nombre. '''

    texte = texte.strip()
    if texte == "":
        return np.array([], dtype = float)
    if virgule_decimale:
        morceaux = re.split(r"[;\s]+", texte.replace(",", "."))
    elif ";" in texte:
        morceaux = [m.replace(",", ".") for m in re.split(r"[;\n]+", texte)]
    else:
        morceaux = re.split(r"[,\s]+", texte)
    valeurs = []
    for morceau in morceaux:
        morceau = morceau.strip()
        if morceau == "":
            continue
        valeur = float(morceau)          # ValueError si ce n'est pas un nombre
        if not np.isfinite(valeur):
            raise ValueError(morceau)
        valeurs.append(valeur)
    return np.array(valeurs, dtype = float)


def valider_parametres_optimisation (check_method, check_target, check_featur, check_min_y, check_max_y, check_min_x, check_max_x, check_n, check_seed, check_n_simul_1 = None, check_n_simul_2 = None):
    ''' Verifie les parametres saisis dans la fenetre "Optimization settings".
    Retourne (parametres, None) si TOUS les parametres sont definis et valides,
    sinon (None, message expliquant le probleme). '''

    if check_method.current() < 1:
        return None, "Please choose the optimization method."
    if check_target.current() < 1:
        return None, "Please choose the target column."
    features = [check_featur.get(i) for i in check_featur.curselection()]
    if len(features) == 0:
        return None, "Please select at least one feature column."
    if check_target.get() in features:
        return None, "The target column must not be one of the feature columns."

    # bornes de la cible, N et seed
    try:
        min_y = float(str(check_min_y.get()).replace(",", "."))
        max_y = float(str(check_max_y.get()).replace(",", "."))
    except ValueError:
        return None, "The Min and Max of the target must be real numbers."
    if not (np.isfinite(min_y) and np.isfinite(max_y) and min_y < max_y):
        return None, "The target Min must be lower than the target Max."
    try:
        n_replication = float(check_n.get())
        seed = float(check_seed.get())
        if not (n_replication.is_integer() and seed.is_integer()):
            raise ValueError
        n_replication, seed = int(n_replication), int(seed)
    except ValueError:
        return None, "N and seed must be integers."
    if n_replication < 2:
        return None, "N must be at least 2."

    # nombres de simulations MCMC (n_simul_1 : tirages de (theta, nu) ; n_simul_2 : tirages des points candidats)
    n_simul_1 = opt_parameters.get("n_simul_1", 100) if check_n_simul_1 is None else None
    n_simul_2 = opt_parameters.get("n_simul_2", 10) if check_n_simul_2 is None else None
    try:
        if check_n_simul_1 is not None:
            n_simul_1 = float(check_n_simul_1.get())
            if not n_simul_1.is_integer():
                raise ValueError
            n_simul_1 = int(n_simul_1)
        if check_n_simul_2 is not None:
            n_simul_2 = float(check_n_simul_2.get())
            if not n_simul_2.is_integer():
                raise ValueError
            n_simul_2 = int(n_simul_2)
    except ValueError:
        return None, "N_targ and N_feat must be integers."
    if n_simul_1 < 10:
        return None, "N_targ must be at least 10 (one (theta, nu) draw is kept every 10 simulations)."
    if n_simul_2 < 1:
        return None, "N_feat must be at least 1."

    # bornes des features (un vecteur de reels : une valeur par feature)
    une_feature = (len(features) == 1)     # une seule valeur attendue : "2,5" = 2.5
    try:
        min_x = lire_vecteur_reels(check_min_x.get(), virgule_decimale = une_feature).tolist()
        max_x = lire_vecteur_reels(check_max_x.get(), virgule_decimale = une_feature).tolist()
    except ValueError:
        return None, "The feature limits must contain only real numbers separated by commas, semicolons or spaces (e.g. 225, 250.5, 305)."
    if len(min_x) == 0 or len(max_x) == 0:
        return None, "Please give the Min and the Max of the feature(s)."
    if len(min_x) != len(max_x):
        return None, "The feature Min and Max vectors must have the same length."
    if len(min_x) != len(features):
        return None, "Give one Min and one Max value per selected feature (" + str(len(features)) + " selected, " + str(len(min_x)) + " given)."
    erreurs_bornes = [(k + 1, a, b) for k, (a, b) in enumerate(zip(min_x, max_x)) if a >= b]
    if erreurs_bornes:
        detail = "\n".join("Value " + str(k) + " : Min = " + str(a) + ", Max = " + str(b) for k, a, b in erreurs_bornes)
        return None, "Each feature Min must be lower than the corresponding Max.\n" + detail + "\n\nRead as: Min = " + str(min_x) + " / Max = " + str(max_x) + "\n(use ';' between values if you use ',' as decimal separator)"

    return {
        "method": check_method.get(),
        "target": check_target.get(),
        "features": features,
        "min_y": min_y,
        "max_y": max_y,
        "min_x": min_x,
        "max_x": max_x,
        "n_replication": n_replication,
        "seed": seed,
        "n_simul_1": n_simul_1,
        "n_simul_2": n_simul_2,
    }, None


def memoriser_optimisation_dans_projet ():
    ''' Enregistre tout de suite le projet ouvert (sans message) pour que les
    donnees importees et les parametres d'optimisation soient memorises sans
    attendre l'enregistrement automatique. Sans projet ouvert, ils restent en
    memoire pour la session. '''

    if nom_projet is not None and chemin_projet is not None:
        try:
            save_current_project(silencieux = True)
        except Exception:
            pass


def recup_my_opt_parameters (parent, check_method, check_target, check_featur, check_min_y, check_max_y, check_min_x, check_max_x, check_n, check_seed, apres_enregistrement = None, variable_case = None, check_n_simul_1 = None, check_n_simul_2 = None):
    ''' Recupere TOUS les parametres de la fenetre "Optimization settings"
    (methode, colonne cible, colonnes features, bornes de la cible, bornes des
    features, nombre de replications N, seed, n_simul_1, n_simul_2) et les memorise dans le
    dictionnaire global opt_parameters, puis ferme la fenetre.
    Si un parametre manque ou est invalide, un avertissement s'affiche et la
    fenetre reste ouverte. '''

    global opt_parameters
    parametres, message = valider_parametres_optimisation(check_method, check_target, check_featur, check_min_y, check_max_y, check_min_x, check_max_x, check_n, check_seed, check_n_simul_1, check_n_simul_2)
    if parametres is None:
        messagebox.showwarning("Optimization", message)
        return
    opt_parameters = parametres
    parent.destroy()
    ## CORRECTIF : la case « Optimization settings » restait cochee -> il fallait cliquer deux fois pour rouvrir la fenetre
    if variable_case is not None:
        try:
            variable_case.set(False)
        except Exception:
            pass
    memoriser_optimisation_dans_projet()
    if apres_enregistrement is not None:
        apres_enregistrement()     # met a jour l'etat du bouton "Run optim"


def import_optimization_data(parent, apres_import = None, variable_case = None):

    global data_opt 
    my_fill_pred = [('CSV files',"*.csv"), ('Excel file', '*.xlsx *.xlsm *.xls'), ('Texte file', '*.txt'),('All',"*.*")]
    file_path_opt = filedialog.askopenfilename(title="Open csv File", filetypes = my_fill_pred)
    parent1 = Canvas(parent, bg = "white", highlightthickness = 0)
    parent1.place(relx = 0, rely = 0, relwidth = 1, relheight = 1)
    tree = ttk.Treeview(parent1, show="headings")
    tree.place(relx = 0.01, rely = 0.1, relwidth = 1, relheight = 0.9)
    donnees_importees = False
    if file_path_opt:          # chaine vide si l'utilisateur annule la fenetre de choix du fichier
        try:
            donnees_lues, _ = lire_fichier_donnees(file_path_opt, na_values = my_missing_values, choisir_feuilles = True)
            if donnees_lues is not None:    # None : choix des feuilles annule, donnees precedentes conservees
                data_opt = donnees_lues
                donnees_importees = True
        except Exception as erreur:
            messagebox.showerror("Import data", "The file could not be read:\n" + str(erreur))
    parent.destroy()
    if variable_case is not None:      # CORRECTIF : decocher la case pour que le clic suivant rouvre la fenetre
        try:
            variable_case.set(False)
        except Exception:
            pass
    if apres_import is not None:
        apres_import()         # met a jour l'etat des boutons selon que data_opt est charge ou non
    if donnees_importees:
        memoriser_optimisation_dans_projet()     # les donnees importees sont memorisees dans le projet


def import_data_opt (parent, check_data_opt, apres_import = None):

    if check_data_opt.get() != False:
        windows_data_opt = Canvas(parent, background = 'cadetblue' , highlightthickness = 0)
        windows_data_opt.place(relx = 0.3, rely = 0.3, relwidth = 0.5, relheight = 0.3)
        import_data_pred_img = PhotoImage(file = my_path('figures/import_data_pred.png'))
        import_data_pred_opt = Button(windows_data_opt, image = import_data_pred_img, text = "Import data", command = lambda : import_optimization_data(windows_data_opt, apres_import, check_data_opt) , font=('arial', 14),  compound = tk.TOP,  borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
        import_data_pred_opt.place(relx= 0.25, rely= 0.25, width= 110, height= 60)
        #close_window_data_opt = Button(windows_data_opt, text = "Save", command = lambda : recup_my_data_opt(windows_data_opt), font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
        ##close_window_data_opt.place(relx= 0.4, rely= 0.93, width= 50, height= 20)


_fenetre_opt_parametres = {"widget": None}


def Edit_opt_parameters (parent, check_opt, check_data_, apres_enregistrement = None):

    ## CORRECTIF : la fenetre precedente est fermee (decocher la case la ferme, recocher n'en empile plus une 2e)
    ancienne_fenetre_opt = _fenetre_opt_parametres.get("widget")
    if _widget_existe(ancienne_fenetre_opt):
        try:
            ancienne_fenetre_opt.destroy()
        except Exception:
            pass
    _fenetre_opt_parametres["widget"] = None

    ## CORRECTIF : sans donnees importees, check_data_ valait None -> AttributeError silencieuse, la case ne faisait rien
    if check_opt.get() != False and check_data_ is None:
        messagebox.showwarning("Optimization", "Please import the data first (« Import data »), then define the optimization settings.")
        check_opt.set(False)
        return

    if check_opt.get() != False:
        windows_optimization = Canvas(parent, background = 'cadetblue' , highlightthickness = 0)
        windows_optimization.place(relx = 0.03, rely = 0.02, relwidth = 0.94, relheight = 0.96)
        _fenetre_opt_parametres["widget"] = windows_optimization
        label_method_opt = Label(windows_optimization, text="Optimization method  ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_method_opt.place(relx=0.02, rely= 0.1, relwidth= 0.4)
        methode_opt = ttk.Combobox(windows_optimization, values = ["Select method", "FBA", "EGO quanti", "EGO mixt"], width= 10)
        methode_opt.current(0)
        methode_opt.place(relx=0.47, rely= 0.1, relwidth= 0.5)
        header_opt = check_data_.columns
        header_target_opt = np.concatenate((["Select column name"], header_opt))
        header_target_opt_ = []
        for my_value in header_target_opt:
           header_target_opt_.append(my_value)
        header_featur_opt_ = []
        for my_value in header_opt:
           header_featur_opt_.append(my_value)
        label_Target = Label(windows_optimization, text="Target column          ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_Target.place(relx=0.02, rely= 0.2, relwidth= 0.4)
        Target_column = ttk.Combobox(windows_optimization, values = header_target_opt_, width= 10)
        Target_column.current(0)
        Target_column.place(relx=0.47, rely= 0.2, relwidth= 0.5)
        label_featur_opt = Label(windows_optimization, text="Featur column          ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_featur_opt.place(relx=0.02, rely= 0.3, relwidth= 0.4)
        Feature_column_opt = tk.Listbox(windows_optimization, fg = "gray10", font = ('arial', 12), relief = GROOVE , selectmode = tk.MULTIPLE, exportselection = False)
        Feature_column_opt.place(relx= 0.47, rely = 0.3, relwidth= 0.5, relheight = 0.25)
        for value in header_featur_opt_:
            Feature_column_opt.insert(END, value)
        # restaure les parametres deja enregistres (reouverture de la fenetre)
        if opt_parameters["method"] in methode_opt["values"]:
            methode_opt.set(opt_parameters["method"])
        if opt_parameters["target"] in header_target_opt_:
            Target_column.set(opt_parameters["target"])
        for i, nom_colonne in enumerate(header_featur_opt_):
            if nom_colonne in opt_parameters["features"]:
                Feature_column_opt.selection_set(i)
       ## color_column.bind("<<ComboboxSelected>>", lambda event: chech_plot_histogram(event, graph, my_data , var_column, color_column, density_column, var, list_color, my_check_densty, my_check_stat))
        
        min_y_optim_def = DoubleVar(value = opt_parameters["min_y"])
        min_y_optim = Spinbox(windows_optimization, from_ = -1000000000 , to = 1000000000 , increment = 1, justify = CENTER , textvariable = min_y_optim_def)
        min_y_optim.place(relx=0.12, rely= 0.64, relwidth= 0.3) 
        
         
        min_x_optim_def = StringVar(value = ", ".join(str(v) for v in opt_parameters["min_x"]))
        min_x_optim = Entry(windows_optimization, textvariable = min_x_optim_def, justify = CENTER)
        min_x_optim.place(relx=0.47, rely= 0.64, relwidth= 0.5)

        max_y_optim_def = DoubleVar(value = opt_parameters["max_y"])
        max_y_optim = Spinbox(windows_optimization, from_ = -1000000000 , to = 1000000000 , increment = 1, justify = CENTER , textvariable = max_y_optim_def)
        max_y_optim.place(relx=0.12, rely= 0.70, relwidth= 0.3) 
       
        
        max_x_optim_def = StringVar(value = ", ".join(str(v) for v in opt_parameters["max_x"]))
        max_x_optim = Entry(windows_optimization, textvariable = max_x_optim_def, justify = CENTER)
        max_x_optim.place(relx=0.47, rely= 0.70, relwidth= 0.5)
        
        
        n_optim_def = IntVar(value = opt_parameters["n_replication"])
        n_optim = Spinbox(windows_optimization, from_ = 2 , to = 100000 , increment = 1, justify = CENTER , textvariable = n_optim_def)
        n_optim.place(relx=0.12, rely= 0.76, relwidth= 0.3) 
        
        
        seed_optim_def = IntVar(value = opt_parameters["seed"])
        seed_optim = Spinbox(windows_optimization, from_ = 0 , to = 1000000 , increment = 1, justify = CENTER , textvariable = seed_optim_def)
        seed_optim.place(relx=0.67, rely= 0.76, relwidth= 0.3) 
        
        
        
        label_y_optim_lim = Label(windows_optimization, text="Target", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_y_optim_lim.place(relx=0.06, rely= 0.59, relwidth= 0.4)
        
        label_y_optim_min = Label(windows_optimization, text="Min", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_y_optim_min.place(relx=0.01, rely= 0.64, relwidth= 0.1)

        
        label_y_optim_max = Label(windows_optimization, text="Max", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_y_optim_max.place(relx=0.01, rely= 0.70, relwidth= 0.1)
        
        label_vector_opt = Label(windows_optimization, text="Featur (delimiter = ',')", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_vector_opt.place(relx=0.47, rely= 0.59, relwidth= 0.5)
       
        label_n_optim = Label(windows_optimization, text="N", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_n_optim.place(relx=0.01, rely= 0.76, relwidth= 0.1)
        
        label_seed_optim = Label(windows_optimization, text="seed", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_seed_optim.place(relx=0.54, rely= 0.76, relwidth= 0.1)

        # nombres de simulations MCMC utilises par l'optimisation (Yields_optimization)
        n_simul_1_def = IntVar(value = opt_parameters.get("n_simul_1", 200))
        n_simul_1_optim = Spinbox(windows_optimization, from_ = 10 , to = 1000000 , increment = 10, justify = CENTER , textvariable = n_simul_1_def)
        n_simul_1_optim.place(relx=0.17, rely= 0.82, relwidth= 0.25)

        n_simul_2_def = IntVar(value = opt_parameters.get("n_simul_2", 10))
        n_simul_2_optim = Spinbox(windows_optimization, from_ = 1 , to = 1000000 , increment = 1, justify = CENTER , textvariable = n_simul_2_def)
        n_simul_2_optim.place(relx=0.67, rely= 0.82, relwidth= 0.3)

        label_n_simul_1 = Label(windows_optimization, text="N_targ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_n_simul_1.place(relx=0.01, rely= 0.82, relwidth= 0.15)

        label_n_simul_2 = Label(windows_optimization, text="N_feat", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_n_simul_2.place(relx=0.50, rely= 0.82, relwidth= 0.16)
        
        
        close_window_opt_param = Button(windows_optimization, text = "Save", command = lambda : recup_my_opt_parameters(windows_optimization, methode_opt, Target_column, Feature_column_opt, min_y_optim, max_y_optim, min_x_optim, max_x_optim, n_optim, seed_optim, apres_enregistrement, check_opt, n_simul_1_optim, n_simul_2_optim), font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray60', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
        close_window_opt_param.place(relx= 0.4, rely= 0.93, width= 70, height= 26)

        ## CORRECTIF : « Save » etait DESACTIVE tant que les parametres n'etaient pas valides : impossible de
        ## cliquer, donc jamais de message (le bouton semblait ne pas fonctionner). Il reste maintenant
        ## cliquable (un avertissement explique le probleme) et le motif s'affiche en direct au-dessus.
        etat_saisie_opt = Label(windows_optimization, text = "", font = ('arial', 12), background = "cadetblue", wraplength = 460, justify = "center")
        etat_saisie_opt.place(relx = 0.02, rely = 0.875, relwidth = 0.96, height = 34)

        def actualiser_bouton_save(*arguments):
            parametres_saisis, message_erreur = valider_parametres_optimisation(methode_opt, Target_column, Feature_column_opt, min_y_optim, max_y_optim, min_x_optim, max_x_optim, n_optim, seed_optim, n_simul_1_optim, n_simul_2_optim)
            if parametres_saisis is not None:
                etat_saisie_opt.configure(text = "All parameters are valid : you can click on Save.", foreground = "darkgreen")
            else:
                etat_saisie_opt.configure(text = str(message_erreur).split("\n")[0], foreground = "darkred")

        for widget_choix in (methode_opt, Target_column):
            widget_choix.bind("<<ComboboxSelected>>", actualiser_bouton_save)
            widget_choix.bind("<KeyRelease>", actualiser_bouton_save)
        Feature_column_opt.bind("<<ListboxSelect>>", actualiser_bouton_save)
        for variable_saisie in (min_y_optim_def, max_y_optim_def, min_x_optim_def, max_x_optim_def, n_optim_def, seed_optim_def, n_simul_1_def, n_simul_2_def):
            variable_saisie.trace_add("write", actualiser_bouton_save)
        actualiser_bouton_save()

        


def _lire_reel_positif (variable, valeur_defaut):
    ''' Lit un reel strictement positif saisi dans un widget (virgule ou point
    decimal acceptes). Retourne valeur_defaut si la saisie est absente,
    invalide ou <= 0. '''

    if variable is None:
        return valeur_defaut
    try:
        valeur = float(str(variable.get()).replace(",", "."))
    except Exception:
        return valeur_defaut
    if not np.isfinite(valeur) or valeur <= 0:
        return valeur_defaut
    return valeur


def _colonnes_abscisse_optimisation (donnees = None):
    ''' Noms (texte) des colonnes de x_init utilisables comme abscisse du graphique
    d'optimisation : colonnes du tableau des points de depart, sauf la cible.
    Sans resultat d'optimisation, reprend les features de "Optimization settings". '''

    if donnees is None:
        return [str(c) for c in opt_parameters["features"]]
    return [str(c) for c in donnees["table_init"].columns if str(c) != str(donnees["ylabel"])]


def _abscisse_optimisation (donnees, check_x_axis = None):
    ''' Retourne (xlabel, x_init, x_opt) selon la colonne choisie dans la liste
    deroulante "X axis". Sans choix valide : premiere colonne de x_init si une
    seule feature, sinon ordre d'evaluation (comportement d'origine). '''

    colonnes = _colonnes_abscisse_optimisation(donnees)
    choix = check_x_axis.get() if check_x_axis is not None else opt_affichage.get("x_axis")
    if choix not in colonnes:
        return donnees["xlabel"], donnees["x_init"], donnees["x_opt"]
    correspondance = {str(c): c for c in donnees["table_init"].columns}
    colonne = correspondance[choix]
    return choix, donnees["table_init"][colonne].to_numpy(dtype = float), donnees["table"][colonne].to_numpy(dtype = float)


def afficher_graphique_optimisation (panel, check_given, check_opt, check_curve, check_error, check_table = None, check_line_width = None, check_point_size = None, check_table_given = None, check_x_axis = None):
    ''' Dessine dans "panel" le graphique du dernier resultat d'optimisation
    (opt_plot_data). Les 4 premieres BooleanVar choisissent les elements
    affiches : points donnes, points optimises, courbe estimee, intervalle a
    95 %. Si check_table est coche, le tableau des points optimises est
    affiche sous le graphique (avec les points de depart si check_table_given
    est coche). check_line_width et check_point_size donnent
    l'epaisseur de la courbe et la taille des points (defauts : 1.5 et 15).
    check_x_axis (liste deroulante) donne la colonne de x_init en abscisse. '''

    global opt_canvas_widget, opt_table_widget
    donnees = opt_plot_data
    if donnees is None:
        return
    for widget_precedent in (opt_canvas_widget, opt_table_widget):
        if widget_precedent is not None:
            try:
                widget_precedent.destroy()
            except Exception:
                pass
    opt_canvas_widget = None
    opt_table_widget = None
    afficher_tableau = check_table is not None and check_table.get()
    afficher_points_depart = check_table_given is not None and check_table_given.get()
    hauteur_graphique = 0.66 if afficher_tableau else 1
    epaisseur_courbe = _lire_reel_positif(check_line_width, 1.5)
    taille_points = _lire_reel_positif(check_point_size, 15)
    opt_affichage.update({"given": check_given.get(), "opt": check_opt.get(), "curve": check_curve.get(), "error": check_error.get(),
                          "table": afficher_tableau, "table_given": afficher_points_depart, "line_width": "%g" % epaisseur_courbe, "point_size": "%g" % taille_points})
    xlabel, x_init_trace, x_opt_trace = _abscisse_optimisation(donnees, check_x_axis)
    if check_x_axis is not None and check_x_axis.get() in _colonnes_abscisse_optimisation(donnees):
        opt_affichage["x_axis"] = check_x_axis.get()

    fig = Figure(figsize = (6, 4.5), dpi = 80)
    _enregistrer_figure_courante("Optimization", fig)
    ax = fig.add_subplot()
    # la courbe estimee n'existe que pour une seule feature (elle est alors tracee selon cette feature)
    if not donnees["multi"] and xlabel == donnees["xlabel"]:
        if check_error.get():
            ax.fill_between(donnees["x_plot"], donnees["y_low"], donnees["y_up"], color = "blue", alpha = 0.15, label = "Error (95%)")
        if check_curve.get():
            ax.plot(donnees["x_plot"], donnees["y_plot"], color = "blue", linewidth = epaisseur_courbe, label = "Curve estimated")
    if check_given.get():
        ax.scatter(x_init_trace, donnees["y_init"], s = taille_points, color = "black", label = "Given data points")
    if check_opt.get():
        ax.scatter(x_opt_trace, donnees["y_opt"], s = taille_points, color = "red", label = "Optimized points")
    # maximum de la cible parmi tous les points evalues (points donnes + points optimises)
    y_tous = np.concatenate([np.asarray(donnees["y_init"], dtype = float), np.asarray(donnees["y_opt"], dtype = float)])
    x_tous = np.concatenate([np.asarray(x_init_trace, dtype = float), np.asarray(x_opt_trace, dtype = float)])
    indice_max = None
    if len(y_tous) > 0 and np.any(np.isfinite(y_tous)):
        indice_max = int(np.nanargmax(y_tous))
        x_max, y_max = x_tous[indice_max], y_tous[indice_max]
        ax.scatter([x_max], [y_max], s = float(np.maximum(8*taille_points, 120)), marker = "*", color = "gold", edgecolors = "black", linewidths = 1, zorder = 5,
                   label = "Maximum (%s) : %.4g at %s = %.4g" % ("optimized" if indice_max >= len(donnees["y_init"]) else "given", y_max, xlabel, x_max))
    ax.set_xlabel(xlabel, fontsize = 13)
    ax.set_ylabel(donnees["ylabel"], fontsize = 13)
    ax.grid(True)
    if len(ax.get_legend_handles_labels()[0]) > 0:
        ax.legend(fontsize = 9)
    fig.tight_layout()
    _appliquer_zoom_representation(ax, "optimization")     # zoom / decalage propres au graphique d'optimisation

    figure_canvas = FigureCanvasTkAgg(fig, master = panel)
    opt_canvas_widget = figure_canvas.get_tk_widget()
    opt_canvas_widget.place(relx = 0, rely = 0, relwidth = 1, relheight = hauteur_graphique)
    figure_canvas.draw()

    if afficher_tableau:
        tableau = donnees["table"]
        types_points = ["Optimized"] * len(tableau)
        if afficher_points_depart:
            # points de depart (donnees) suivis des points optimises, dans l'ordre d'evaluation
            types_points = ["Given"] * len(donnees["table_init"]) + types_points
            tableau = pd.concat([donnees["table_init"], tableau], ignore_index = True)
        cadre_tableau = tk.Frame(panel, background = "gray80")
        cadre_tableau.place(relx = 0, rely = 0.66, relwidth = 1, relheight = 0.34)
        colonnes = ["N°"] + [str(c) for c in tableau.columns] + (["Type"] if afficher_points_depart else [])
        arbre = ttk.Treeview(cadre_tableau, show = "headings", columns = colonnes)
        defilement = ttk.Scrollbar(cadre_tableau, orient = "vertical", command = arbre.yview)
        arbre.configure(yscrollcommand = defilement.set)
        for nom_colonne in colonnes:
            arbre.column(nom_colonne, width = 110, anchor = "center")
            arbre.heading(nom_colonne, text = nom_colonne)
        arbre.tag_configure("optimise", foreground = "red")       # meme code couleur que le graphique (points optimises en rouge)
        arbre.tag_configure("maximum", background = "lightyellow")  # ligne du maximum (etoile doree du graphique)
        ## ligne du maximum du tableau : en GRAS (police du tableau, en gras). La police est rattachee au tableau
        ## pour ne pas etre detruite par le ramasse-miettes tant que le tableau est affiche.
        try:
            police_tableau = _tkfont_langue.nametofont(ttk.Style().lookup("Treeview", "font") or "TkDefaultFont", root = SmartDigicat)
            police_grasse = police_tableau.copy()
        except Exception:
            police_grasse = _tkfont_langue.Font(root = SmartDigicat, family = "arial", size = 12)
        police_grasse.configure(weight = "bold")
        arbre._police_ligne_maximum = police_grasse
        arbre.tag_configure("gras", font = police_grasse)
        # position du maximum dans le tableau affiche (absent si c'est un point donne et que seuls les points optimises sont listes)
        ligne_max = None
        if indice_max is not None:
            ligne_max = indice_max if afficher_points_depart else indice_max - len(donnees["y_init"])
        ## ligne a mettre en gras : celle du maximum de la cible PARMI LES LIGNES AFFICHEES (quand le maximum global
        ## est un point de depart non liste, c'est le meilleur point optimise qui est en gras)
        ligne_grasse = None
        try:
            valeurs_cible = pd.to_numeric(tableau[tableau.columns[-1]], errors = "coerce").to_numpy(dtype = float)
            if len(valeurs_cible) > 0 and np.any(np.isfinite(valeurs_cible)):
                ligne_grasse = int(np.nanargmax(valeurs_cible))
        except Exception:
            ligne_grasse = None
        for numero, ligne in enumerate(tableau.to_numpy().tolist(), start = 1):
            valeurs_ligne = [numero] + [round(v, 4) for v in ligne]
            if afficher_points_depart:
                valeurs_ligne.append(types_points[numero - 1] + (" (max)" if numero - 1 == ligne_max else ""))
            etiquettes = []
            if types_points[numero - 1] == "Optimized" and afficher_points_depart:
                etiquettes.append("optimise")
            if numero - 1 == ligne_max:
                etiquettes.append("maximum")
            if numero - 1 == ligne_grasse:
                etiquettes.append("gras")
            arbre.insert("", "end", values = valeurs_ligne, tags = tuple(etiquettes))
        defilement.pack(side = "right", fill = "y")
        arbre.pack(side = "left", fill = "both", expand = True)
        opt_table_widget = cadre_tableau

    # curseur de zoom de la barre laterale gauche + deplacement a la souris (double-clic : reinitialiser)
    _activer_glissement_representation(ax, figure_canvas, "optimization")

    def _sur_changement_zoom_optimisation (valeur_curseur):
        zoom_representation_par_type["optimization"] = _curseur_vers_zoom_representation(valeur_curseur)
        afficher_graphique_optimisation(panel, check_given, check_opt, check_curve, check_error, check_table, check_line_width, check_point_size, check_table_given, check_x_axis)

    scale.configure(command = _sur_changement_zoom_optimisation)
    scale.set(_zoom_vers_curseur_representation(zoom_representation_par_type.get("optimization", 1.0)))     # doit rester la derniere instruction


def check_optim_parameters (panel = None, check_given = None, check_opt = None, check_curve = None, check_error = None, check_table = None, check_line_width = None, check_point_size = None, check_table_given = None, check_x_axis = None):
    ''' Lit les parametres enregistres via Edit_opt_parameters, verifie qu'ils
    sont complets et coherents avec les donnees importees, lance
    Yields_optimization, puis affiche le graphique du resultat dans "panel".
    check_x_axis (ttk.Combobox) est rempli avec les noms des colonnes de
    x_init : la colonne choisie definit l'abscisse du graphique.
    Retourne le resultat de l'optimisation, ou None en cas d'erreur. '''

    global opt_result, opt_plot_data

    methode = opt_parameters["method"]
    cible = opt_parameters["target"]
    features = opt_parameters["features"]

    # --- 1) parametres definis ?
    if methode is None or cible is None or len(features) == 0:
        messagebox.showwarning("Optimization", "Please define the method, the target column and at least one feature column in 'Optimization settings' (then click Save).")
        return None
    if cible in features:
        messagebox.showwarning("Optimization", "The target column must not be one of the feature columns.")
        return None

    # --- 2) donnees importees ? (data_opt est le DataFrame importe)
    if data_opt is None:
        messagebox.showwarning("Optimization", "Please import the data first.")
        return None
    colonnes_absentes = [c for c in features + [cible] if c not in data_opt.columns]
    if colonnes_absentes:
        messagebox.showwarning("Optimization", "Column(s) not found in the data: " + ", ".join(map(str, colonnes_absentes)) + ". Please redefine the optimization settings.")
        return None

    # --- 3) preparation : lignes completes, index 0..n-1 (Yields_optimization ajoute les nouveaux points a la suite)
    data_sel = data_opt[features + [cible]].dropna().reset_index(drop = True)
    try:
        data_sel = data_sel.astype(float)
    except (ValueError, TypeError):
        messagebox.showwarning("Optimization", "The target and feature columns must be numeric.")
        return None
    if len(data_sel) < 2:
        messagebox.showwarning("Optimization", "At least 2 complete rows are needed to run the optimization.")
        return None

    y_init = data_sel[cible]
    if y_init.max() == y_init.min():
        messagebox.showwarning("Optimization", "The target column is constant: nothing to optimize.")
        return None

    # x_init : DataFrame contenant TOUTES les features selectionnees (une colonne par feature,
    # meme s'il n'y en a qu'une ; Yields_optimization gere le cas d'une seule colonne)
    x_init = data_sel[features]

    # --- 4) parametres saisis dans "Optimization settings"
    lower_y = opt_parameters["min_y"]
    upper_y = opt_parameters["max_y"]
    n_replication = opt_parameters["n_replication"]
    seed = opt_parameters["seed"]
    n_simul_1 = int(opt_parameters.get("n_simul_1", 100))
    n_simul_2 = int(opt_parameters.get("n_simul_2", 10))
    min_x = opt_parameters["min_x"]
    max_x = opt_parameters["max_x"]

    if len(min_x) != len(features) or len(max_x) != len(features):
        messagebox.showwarning("Optimization", "Please give one Min and one Max value for each of the " + str(len(features)) + " selected feature(s) in 'Optimization settings' (then click Save).")
        return None
    if lower_y >= upper_y or n_replication < 2:
        messagebox.showwarning("Optimization", "Invalid target limits or N in 'Optimization settings'.")
        return None

    # une borne Min et Max par feature, dans l'ordre des colonnes de x_init
    lower_x, upper_x = [float(v) for v in min_x], [float(v) for v in max_x]
    bornes_invalides = [str(f) for f, b_min, b_max in zip(features, lower_x, upper_x) if b_min >= b_max]
    if len(bornes_invalides) != 0:
        messagebox.showwarning("Optimization", "Min must be lower than Max for: " + ", ".join(bornes_invalides) + ".")
        return None

    ## temps restant : une unite par iteration + une unite pour le calcul du graphique (courbe estimee).
    ## Premiere estimation : durees mesurees lors de la derniere optimisation.
    if _duree_iteration_optimisation["moyenne"] is not None:
        _progression_demarrer(int(n_replication) + 1, [None] * int(n_replication) + [_duree_iteration_optimisation["graphique"]],
                              _duree_iteration_optimisation["moyenne"])
    else:
        ## premier lancement : estimation a priori de chaque iteration (corrigee des la premiere iteration terminee)
        durees_estimees = _estimer_durees_optimisation(len(y_init), n_simul_1, n_simul_2, int(n_replication))
        _progression_demarrer(int(n_replication) + 1, durees_estimees, apriori = [True] * (int(n_replication) + 1))
    try:
        # en arriere-plan pendant que l'interface reste reactive : le bouton d'arret peut l'interrompre
        d_opt = _calcul_interruptible(lambda: Yields_optimization(y_init, lower_y, upper_y, x_init, lower_x, upper_x, n_simul_1 = n_simul_1, n_simul_2 = n_simul_2, n_replication = n_replication, n_pas_theta = 10, random_states = seed))
    except Exception as erreur:
        messagebox.showerror("Optimization", "The optimization failed:\n" + str(erreur))
        return None
    _progression_unite_terminee()                 # derniere iteration terminee
    if _progression["durees_faites"]:
        _duree_iteration_optimisation["moyenne"] = _builtins_langue.sum(_progression["durees_faites"]) / len(_progression["durees_faites"])

    # --- 5) donnees du graphique
    try:
        dd2 = d_opt[0]                    # colonne cible, puis features : points donnes suivis des points optimises
        n_init = len(y_init)
        ndim_ = dd2.shape[1]
        points_opt = dd2.iloc[n_init:]
        donnees = {"multi": ndim_ > 2, "ylabel": cible, "table": points_opt[features + [cible]].reset_index(drop = True), "table_init": data_sel[features + [cible]].reset_index(drop = True), "y_init": y_init.to_numpy(dtype = float), "y_opt": points_opt.iloc[:, 0].to_numpy(dtype = float)}

        if ndim_ == 2:
            # une feature : courbe estimee par la loi predictive bayesienne (Bayesian_predict,
            # meme modele que l'optimisation) + intervalle de credibilite a 95 %
            x_train = dd2.iloc[:, 1].astype(float)          # Series : Bayesian_predict utilise son nom de colonne
            y_train = dd2.iloc[:, 0].astype(float)
            nu_int_1 = float(nu_int(y_train, lower_y, upper_y, x_train, lower_x[0], upper_x[0]))
            x_plot = np.sort(np.concatenate([np.linspace(lower_x[0], upper_x[0], 100), x_train.to_numpy()]))
            # chaine MCMC sur (theta, nu) : calcul long, execute de facon interruptible (bouton d'arret)
            prediction, tirages = _calcul_interruptible(lambda: Bayesian_predict(y_train, lower_y, upper_y, x_train, lower_x[0], upper_x[0], x_plot,
                                                                                 n_simul = 200, n_pas_theta = 10, burn_in = 0, niveau = 0.95,
                                                                                 nu_0 = nu_int_1, methode = 'melange', random_state = seed))
            donnees.update({"xlabel": features[0], "x_init": x_init.iloc[:, 0].to_numpy(dtype = float), "x_opt": points_opt.iloc[:, 1].to_numpy(dtype = float),
                            "x_plot": x_plot, "y_plot": prediction["f_mean"].to_numpy(dtype = float),
                            "y_low": prediction["f_inf"].to_numpy(dtype = float), "y_up": prediction["f_sup"].to_numpy(dtype = float)})
        else:
            # plusieurs features : valeurs de la cible selon l'ordre d'evaluation
            donnees.update({"xlabel": "Evaluation points", "x_init": np.arange(n_init), "x_opt": np.arange(n_init, n_init + len(points_opt))})
    except Exception as erreur:
        opt_result = d_opt
        _terminer_indicateur_execution()
        messagebox.showerror("Optimization", "The graph could not be computed:\n" + str(erreur))
        return d_opt

    # resultat et graphique memorises ensemble (un arret pendant la prediction conserve l'ancien resultat)
    opt_result = d_opt
    opt_plot_data = donnees
    _duree_iteration_optimisation["graphique"] = _progression_unite_terminee()
    _progression_arreter()
    _terminer_indicateur_execution()          # calculs termines : l'encadre disparait tout de suite

    # --- liste deroulante de l'abscisse : noms des colonnes de x_init
    if check_x_axis is not None:
        colonnes_x = [str(c) for c in x_init.columns]
        choix_precedent = check_x_axis.get()
        check_x_axis.configure(values = colonnes_x)
        check_x_axis.set(choix_precedent if choix_precedent in colonnes_x else colonnes_x[0])

    # --- 6) affichage dans le panneau
    if panel is not None:
        afficher_graphique_optimisation(panel, check_given, check_opt, check_curve, check_error, check_table, check_line_width, check_point_size, check_table_given, check_x_axis)

    return d_opt


def parametres_optimisation_definis ():
    ''' True si les parametres de "Optimization settings" sont enregistres,
    complets (methode, cible, features, une borne Min et Max par feature) et
    coherents avec les donnees importees (colonnes presentes). '''

    if data_opt is None:
        return False
    methode = opt_parameters["method"]
    cible = opt_parameters["target"]
    features = opt_parameters["features"]
    if methode is None or cible is None or len(features) == 0 or cible in features:
        return False
    ## (pas de any() / all() avec un generateur : `from numpy import *` les remplace par
    ## numpy.any / numpy.all, qui renvoient toujours vrai pour un generateur)
    colonnes_absentes = [colonne for colonne in features + [cible] if colonne not in data_opt.columns]
    if len(colonnes_absentes) > 0:
        return False
    return len(opt_parameters["min_x"]) == len(features) and len(opt_parameters["max_x"]) == len(features)


def enregistrer_graphique_optimisation ():
    ''' Enregistre le graphique d'optimisation tel qu'il est affiche
    (PNG, PDF ou JPEG). '''

    figure = figures_ouvertes_courantes.get("Optimization")
    if opt_plot_data is None or figure is None:
        messagebox.showwarning("Optimization", "Please run the optimization first (button 'Run optim').")
        return
    chemin = filedialog.asksaveasfilename(filetypes = [("PNG file", ".png"), ("PDF file", ".pdf"), ("JPEG file", ".jpeg")], defaultextension = ".png")
    if not chemin:
        return
    try:
        figure.savefig(chemin, dpi = 200, bbox_inches = "tight")
    except Exception as erreur:
        messagebox.showerror("Optimization", "The plot could not be saved:\n" + str(erreur))


def enregistrer_points_optimisation ():
    ''' Enregistre, dans un seul tableau, les points de depart (donnees)
    suivis des points optimises (colonnes : features puis cible).
    Formats : Excel (feuille "Points") ou CSV (separateur ";"). '''

    if opt_plot_data is None:
        messagebox.showwarning("Optimization", "Please run the optimization first (button 'Run optim').")
        return
    tous_les_points = pd.concat([opt_plot_data["table_init"], opt_plot_data["table"]], ignore_index = True)
    chemin = filedialog.asksaveasfilename(filetypes = [("Excel file", "*.xlsx"), ("csv file", "*.csv")], defaultextension = ".xlsx")
    if not chemin:
        return
    try:
        if os.path.splitext(chemin)[1].lower() == ".csv":
            tous_les_points.to_csv(chemin, sep = ";", index = False, decimal = ".")
        else:
            tous_les_points.to_excel(chemin, sheet_name = "Points", index = False)
    except Exception as erreur:
        messagebox.showerror("Optimization", "The points could not be saved:\n" + str(erreur))


def my_optimization(event, parent):

    global opt_canvas_widget, opt_table_widget
    opt_canvas_widget = None      # nouveau panneau : aucun widget de graphique affiche pour l'instant
    opt_table_widget = None
    scale.configure(command = "")         # le curseur de zoom ne doit plus piloter le graphique d'un autre panneau
    scale.set(_zoom_vers_curseur_representation(zoom_representation_par_type.get("optimization", 1.0)))

    ##graph_window = Canvas(parent, background = "gray90")
    graph = Canvas(parent, relief= GROOVE ,  background = 'gray80' , highlightthickness=0)
    features = Canvas(parent, relief= GROOVE, background = 'gray90', highlightthickness=0)
    graph.place(relx=0, rely=0, relwidth= 0.60, relheight=1)
    features.place(relx=0.6, rely=0, relwidth= 0.40, relheight=1)
    data_def  = BooleanVar(value = False)
    data_info = Checkbutton(features,  text ="Import data                          ", variable = data_def, command = lambda : import_data_opt(features, data_def, actualiser_etat_boutons) , fg = "gray30", font=('arial', 15), background = 'gray90')
    data_info.place(relx = 0.1, rely= 0.15, relwidth= 0.5)
    data_init_opt = []

    def actualiser_etat_boutons():
        # "View data" : donnees importees ; "Run optim" : donnees importees ET parametres definis
        View_data.configure(state = NORMAL if data_opt is not None else DISABLED)
        Run_optimzation.configure(state = NORMAL if parametres_optimisation_definis() else DISABLED)

    def  Edit_data_opt():
        global data_init_opt
        data_init_opt = list(data_opt) 
        refresh_data_opt() 

    def refresh_data_opt(): 
        global window_data , data_init_opt
        my_data_opt = data_opt.to_numpy().tolist() 
        window_data = ttk.Treeview(graph, selectmode = 'extended', show = "headings", columns = data_init_opt)
        window_data.place(relx = 0, rely= 0, relwidth= 1, relheight= 1)
        for i in data_init_opt:
            window_data.column(i, width = 130, anchor = 'c')
            window_data.heading(i, text = " " + str(i))
        for info_value in my_data_opt:
            value_p = [x for x in info_value]
            window_data.insert("",'end', values = value_p)
            
            
    def redessiner_optim():
        afficher_graphique_optimisation(graph, opt_exp_dis_def, opt_points_def, opt_curve_def, opt_error_def, opt_table_def, line_width_def, point_size_def, opt_table_given_def, x_axis_opt)

    def actualiser_case_points_depart():
        # "with given points" ne change que le tableau : inutilisable tant que le tableau n'est pas affiche
        opt_table_given_info.configure(state = NORMAL if opt_table_def.get() else DISABLED)

    def declencher_enregistrement(variable, fonction_enregistrement):
        # la case sert de declencheur : elle lance l'enregistrement puis se decoche
        if variable.get():
            try:
                fonction_enregistrement()
            finally:
                variable.set(False)

    def voir_graphique_optim():
        if opt_plot_data is None:
            messagebox.showwarning("Optimization", "Please run the optimization first (button 'Run optim').")
            return
        redessiner_optim()

    opt_def  = BooleanVar(value = False)
    opt_info = Checkbutton(features,  text ="Optimization settings            ", variable = opt_def, command = lambda : Edit_opt_parameters(features, opt_def, data_opt, actualiser_etat_boutons) , fg = "gray30", font=('arial', 15), background = 'gray90')
    opt_exp_dis_def  = BooleanVar(value = opt_affichage["given"])
    opt_exp_dis_info = Checkbutton(features,  text ="Given data points                ", variable = opt_exp_dis_def, command = lambda : redessiner_optim(), fg = "gray30", font=('arial', 15), background = 'gray90')
    opt_points_def  = BooleanVar(value = opt_affichage["opt"])
    opt_points_info = Checkbutton(features,  text ="Optimized points                  ", variable = opt_points_def, command = lambda : redessiner_optim(), fg = "gray30", font=('arial', 15), background = 'gray90')
    opt_curve_def  = BooleanVar(value = opt_affichage["curve"])
    opt_curve_info = Checkbutton(features,  text ="Curve estimated                   ", variable = opt_curve_def, command = lambda : redessiner_optim(), fg = "gray30", font=('arial', 15), background = 'gray90')
    opt_error_def  = BooleanVar(value = opt_affichage["error"])
    opt_error_info = Checkbutton(features,  text ="Error (95%)                     ", variable = opt_error_def, command = lambda : redessiner_optim(), fg = "gray30", font=('arial', 15), background = 'gray90')
    opt_table_given_def  = BooleanVar(value = opt_affichage["table_given"])
    opt_table_given_info = Checkbutton(features,  text ="with given points", variable = opt_table_given_def, command = lambda : redessiner_optim(), fg = "gray30", font=('arial', 15), background = 'gray90')
    line_width_def = StringVar(value = opt_affichage["line_width"])
    point_size_def = StringVar(value = opt_affichage["point_size"])
    label_line_width = Label(features, text="Line width", font=('arial', 13), fg = "gray30", bg = "gray90")
    line_width_opt = Spinbox(features, from_ = 0.5, to = 10, increment = 0.5, justify = CENTER, textvariable = line_width_def)
    label_point_size = Label(features, text="Point size", font=('arial', 13), fg = "gray30", bg = "gray90")
    point_size_opt = Spinbox(features, from_ = 1, to = 200, increment = 1, justify = CENTER, textvariable = point_size_def)
    # liste deroulante : colonne de x_init utilisee en abscisse (remplie par check_optim_parameters)
    label_x_axis = Label(features, text="X axis", font=('arial', 13), fg = "gray30", bg = "gray90")
    colonnes_x_axis = _colonnes_abscisse_optimisation(opt_plot_data)
    x_axis_opt = ttk.Combobox(features, values = colonnes_x_axis, state = "readonly")
    if opt_affichage.get("x_axis") in colonnes_x_axis:
        x_axis_opt.set(opt_affichage["x_axis"])
    elif len(colonnes_x_axis) != 0:
        x_axis_opt.set(colonnes_x_axis[0])
    x_axis_opt.bind("<<ComboboxSelected>>", lambda event : redessiner_optim() if opt_plot_data is not None else None)
    save_plot_def  = BooleanVar(value = False)
    save_plot_info = Checkbutton(features,  text ="Save plot                            ", variable = save_plot_def, command = lambda : declencher_enregistrement(save_plot_def, enregistrer_graphique_optimisation), fg = "gray30", font=('arial', 15), background = 'gray90')
    save_points_def  = BooleanVar(value = False)
    save_points_info = Checkbutton(features,  text ="Save points                        ", variable = save_points_def, command = lambda : declencher_enregistrement(save_points_def, enregistrer_points_optimisation), fg = "gray30", font=('arial', 15), background = 'gray90')
    opt_table_def  = BooleanVar(value = opt_affichage["table"])
    opt_table_info = Checkbutton(features,  text ="Optimized table                    ", variable = opt_table_def, command = lambda : (actualiser_case_points_depart(), redessiner_optim()), fg = "gray30", font=('arial', 15), background = 'gray90')
    View_plot_img = PhotoImage(file = my_path('figures/apercu_p.png'))
    View_plot = Button(features, image = View_plot_img, text = "View plot", command = lambda : voir_graphique_optim() , font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
    View_data_img = PhotoImage(file = my_path('figures/View_pred.png'))
    run_opt = PhotoImage(file = my_path('figures/run_opt.png'))
    
    View_data = Button(features, image = View_data_img, text = "View  data", command = lambda : Edit_data_opt() , font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray60', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
    
    Run_optimzation = Button(features, image = run_opt, text = "Run optim", command = lambda : lancer_optimisation(graph, opt_exp_dis_def, opt_points_def, opt_curve_def, opt_error_def, opt_table_def, line_width_def, point_size_def, opt_table_given_def, x_axis_opt) , font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray60', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
    
    label_imput = Label(features, text="Input", font=('arial', 15), fg = "saddlebrown", bg = "gray90")
    label_imput.place(relx=0, rely= 0.08, relwidth= 0.3)
    
    
    label_output = Label(features, text="Output", font=('arial', 15), fg = "saddlebrown", bg = "gray90")
    label_output.place(relx=0, rely= 0.28, relwidth= 0.3)
    
    label_export = Label(features, text="Export", font=('arial', 15), fg = "saddlebrown", bg = "gray90")
    label_export.place(relx=0, rely= 0.74, relwidth= 0.3)
    
    
    separator_1 = ttk.Separator(features, orient='horizontal')
    separator_1.place(relx=0.1 , rely = 0.12, relwidth = 0.8)
    
    separator_2 = ttk.Separator(features, orient='horizontal')
    separator_2.place(relx=0.1 , rely = 0.32, relwidth=0.8)
    
    separator_3 = ttk.Separator(features, orient='horizontal')
    separator_3.place(relx=0.1 , rely = 0.78, relwidth=0.8)
    
    
    opt_info.place(relx = 0.1, rely= 0.2, relwidth= 0.5)
    opt_exp_dis_info.place(relx = 0.1, rely= 0.35, relwidth= 0.5)
    opt_points_info.place(relx = 0.1, rely= 0.4, relwidth= 0.5)
    opt_curve_info.place(relx = 0.1, rely= 0.45, relwidth= 0.5)
    opt_error_info.place(relx = 0.1, rely= 0.5, relwidth= 0.5)
    opt_table_info.place(relx = 0.1, rely= 0.55, relwidth= 0.5)
    opt_table_given_info.place(relx = 0.6, rely= 0.55, relwidth= 0.38)
    label_line_width.place(relx = 0.1, rely= 0.61, relwidth= 0.3)
    line_width_opt.place(relx = 0.42, rely= 0.61, relwidth= 0.18)
    label_point_size.place(relx = 0.1, rely= 0.66, relwidth= 0.3)
    point_size_opt.place(relx = 0.42, rely= 0.66, relwidth= 0.18)
    View_plot.place(relx= 0.63, rely= 0.6, width= 100, height= 60)
    save_plot_info.place(relx = 0.1, rely= 0.8, relwidth= 0.5)
    save_points_info.place(relx = 0.1, rely= 0.85, relwidth= 0.5)
    
    View_data.place(relx= 0.6, rely= 0.15, width= 110, height= 60)
    Run_optimzation.place(relx= 0.6, rely = 0.35, width= 110, height= 60)
    # liste deroulante "X axis" juste sous le bouton "Run optim" (decalage en pixels : hauteur du bouton = 60)
    label_x_axis.place(relx= 0.6, rely = 0.35, y = 62, width= 110)
    x_axis_opt.place(relx= 0.6, rely = 0.35, y = 86, relwidth= 0.35)
    actualiser_etat_boutons()      # etat initial : "View data" et "Run optim" desactives tant que donnees / parametres ne sont pas definis
    actualiser_case_points_depart()

    if opt_plot_data is not None:
        redessiner_optim()         # affiche le dernier resultat (ex. : resultat memorise dans un projet ouvert)
    
    
    
    ##graph_window.place(relx=0, rely=0, relwidth=1, relheight=1)



# ==========================================================================
# ==  19. Zoom des graphiques et nuage de points                          ==
# ==========================================================================
def _curseur_vers_zoom_representation (valeur_curseur):
    return ZOOM_MIN_REPRESENTATION + (float(valeur_curseur) / 100) * (ZOOM_MAX_REPRESENTATION - ZOOM_MIN_REPRESENTATION)


def _zoom_vers_curseur_representation (valeur_zoom):
    valeur = (valeur_zoom - ZOOM_MIN_REPRESENTATION) / (ZOOM_MAX_REPRESENTATION - ZOOM_MIN_REPRESENTATION) * 100
    if valeur < 0:
        valeur = 0
    elif valeur > 100:
        valeur = 100
    return int(valeur)


def _appliquer_zoom_representation (axes, cle):
    ''' Resserre ou elargit les axes du graphique autour de leur centre, en
    fonction du niveau de zoom courant de CE type de graphique (identifie par
    "cle" : "scatter", "boxplot", "histogram", "plot3d", "correlation"...),
    puis applique le decalage (pan) accumule par glissement a la souris pour
    ce meme type de graphique. Chaque type de graphique garde ainsi son
    propre zoom, independant des autres. '''
    facteur = zoom_representation_par_type.get(cle, 1.0)
    if facteur <= 0:
        facteur = 1.0
    decalage_x = decalage_zoom_x_representation_par_type.get(cle, 0.0)
    decalage_y = decalage_zoom_y_representation_par_type.get(cle, 0.0)
    limites = ((axes.get_xlim, axes.set_xlim, decalage_x), (axes.get_ylim, axes.set_ylim, decalage_y))
    for get_lim, set_lim, decalage in limites:
        borne_min, borne_max = get_lim()
        centre = (borne_min + borne_max) / 2 + decalage
        demi_etendue = (borne_max - borne_min) / 2 / facteur
        set_lim(centre - demi_etendue, centre + demi_etendue)


def _activer_glissement_representation (axes, canvas_mpl, cle):
    ''' Permet, une fois le graphique zoome, de deplacer la vue a la souris
    (clic-glisse) pour explorer toutes les parties du graphique. Le decalage
    est memorise sous "cle", propre a ce type de graphique. Un double-clic
    reinitialise le zoom et la position de CE type de graphique uniquement. '''
    tk_widget = canvas_mpl.get_tk_widget()
    etat_glissement = {"x": None, "y": None}

    def _debut_glissement (event):
        etat_glissement["x"] = event.x
        etat_glissement["y"] = event.y

    def _glisser (event):
        x0 = etat_glissement.get("x")
        y0 = etat_glissement.get("y")
        if x0 is None or y0 is None:
            return
        delta_px_x = event.x - x0
        delta_px_y = event.y - y0
        largeur_widget = tk_widget.winfo_width()
        hauteur_widget = tk_widget.winfo_height()
        if largeur_widget > 1 and hauteur_widget > 1:
            xlim = axes.get_xlim()
            ylim = axes.get_ylim()
            pas_x = delta_px_x * (xlim[1] - xlim[0]) / largeur_widget
            pas_y = delta_px_y * (ylim[1] - ylim[0]) / hauteur_widget
            decalage_zoom_x_representation_par_type[cle] = decalage_zoom_x_representation_par_type.get(cle, 0.0) - pas_x
            decalage_zoom_y_representation_par_type[cle] = decalage_zoom_y_representation_par_type.get(cle, 0.0) + pas_y
            axes.set_xlim(xlim[0] - pas_x, xlim[1] - pas_x)
            axes.set_ylim(ylim[0] + pas_y, ylim[1] + pas_y)
            canvas_mpl.draw_idle()
        etat_glissement["x"] = event.x
        etat_glissement["y"] = event.y

    def _fin_glissement (event):
        etat_glissement["x"] = None
        etat_glissement["y"] = None

    def _reinitialiser_vue (event):
        zoom_representation_par_type[cle] = 1.0
        decalage_zoom_x_representation_par_type[cle] = 0.0
        decalage_zoom_y_representation_par_type[cle] = 0.0
        try:
            scale.set(_zoom_vers_curseur_representation(1.0))
        except Exception:
            pass

    tk_widget.bind("<ButtonPress-1>", _debut_glissement)
    tk_widget.bind("<B1-Motion>", _glisser)
    tk_widget.bind("<ButtonRelease-1>", _fin_glissement)
    tk_widget.bind("<Double-Button-1>", _reinitialiser_vue)


## Filtre du scatter plot : modalites (converties en texte) de la variable de couleur a afficher
filtre_modalites_scatter = {"colonne": None, "modalites": None}


def _texte_modalite (valeur):
    """ Texte identifiant une modalite (meme conversion que « Categorized all »). """
    return str(valeur)


def _choisir_modalites_scatter (tableau, colonne, deja_choisies = None):
    """ Fenetre modale : choix des modalites de `colonne` a afficher dans le scatter plot.
    Renvoie la liste des modalites choisies (textes), ou None si l'utilisateur annule. """
    valeurs = pd.Series(tableau[colonne]).dropna()
    valeurs_uniques = list(pd.unique(valeurs))
    try:
        valeurs_uniques = sorted(valeurs_uniques)
    except TypeError:
        valeurs_uniques = sorted(valeurs_uniques, key = _texte_modalite)
    modalites = []
    for v in valeurs_uniques:
        texte = _texte_modalite(v)
        if texte not in modalites:
            modalites.append(texte)
    resultat = {"modalites": None}

    fenetre = tk.Toplevel(SmartDigicat)
    fenetre.title(traduire("Filter modalities"))
    fenetre.resizable(False, False)
    fenetre.transient(SmartDigicat)
    tk.Label(fenetre, text = traduire("Modalities to display") + " : " + str(colonne), font = ('arial', 14, 'bold')).pack(padx = 20, pady = (15, 8))

    cadre = tk.Frame(fenetre)
    cadre.pack(padx = 20, pady = 4, fill = tk.BOTH, expand = True)
    barre = tk.Scrollbar(cadre, orient = tk.VERTICAL)
    liste = tk.Listbox(cadre, selectmode = tk.MULTIPLE, exportselection = False, width = 40, height = int(np.clip(len(modalites), 3, 15)),
                       font = ('arial', 13), yscrollcommand = barre.set, activestyle = "none")
    barre.config(command = liste.yview)
    liste.pack(side = tk.LEFT, fill = tk.BOTH, expand = True)
    barre.pack(side = tk.RIGHT, fill = tk.Y)
    for texte in modalites:
        liste.insert(tk.END, texte)
    ## preselection : choix precedent pour cette variable, sinon toutes les modalites
    for i, texte in enumerate(modalites):
        if deja_choisies is None or texte in deja_choisies:
            liste.selection_set(i)

    message = tk.Label(fenetre, text = "", font = ('arial', 12), fg = "red")
    message.pack(padx = 20, pady = (4, 0))

    def appliquer (*_):
        indices = liste.curselection()
        if len(indices) == 0:
            message.config(text = traduire("Select at least one modality."))
            return
        resultat["modalites"] = [modalites[i] for i in indices]
        fenetre.destroy()

    boutons = tk.Frame(fenetre)
    boutons.pack(padx = 20, pady = (8, 15))
    tk.Button(boutons, text = traduire("Select all"), command = lambda: liste.selection_set(0, tk.END), width = 11).pack(side = tk.LEFT, padx = 4)
    tk.Button(boutons, text = traduire("Clear all"), command = lambda: liste.selection_clear(0, tk.END), width = 11).pack(side = tk.LEFT, padx = 4)
    tk.Button(boutons, text = traduire("Cancel"), command = fenetre.destroy, width = 11).pack(side = tk.LEFT, padx = 4)
    tk.Button(boutons, text = traduire("Apply"), command = appliquer, width = 11).pack(side = tk.LEFT, padx = 4)
    fenetre.bind("<Return>", appliquer)
    fenetre.bind("<Escape>", lambda *_: fenetre.destroy())

    fenetre.update_idletasks()
    try:
        fenetre.grab_set()
    except Exception:
        pass
    liste.focus_set()
    SmartDigicat.wait_window(fenetre)
    return resultat["modalites"]


def scatter_plot (event, parent, my_data, check_var1, check_var2, check_color, list_var, list_color, check_str, check_point_size, check_filtre = None):

    global my_scatter_plot, dernier_x_scatter, dernier_y_scatter, dernier_color_scatter, graph_window_representation_widget

    ## Detruit le graphique scatter precedent au lieu de l'empiler par-dessus
    ## (important ici car le zoom via la barre verticale gauche redessine ce
    ## graphique en continu pendant qu'on deplace le curseur)
    if graph_window_representation_widget is not None and graph_window_representation_widget.winfo_exists():
        graph_window_representation_widget.destroy()

    graph_window = Canvas(parent, background = "gray90")
    graph_window_representation_widget = graph_window
    x_index = check_var1.current()
    y_index = check_var2.current()
    color_index = check_color.current()
    ##x_column.current(x_index)
    control_var = [x_index, y_index, color_index]
    x_contr = []
    for i in range(len(control_var)):
        if control_var[i] != 0:
           x_contr.append(True)
    valeur_check = check_str.get()
    data_add = my_data.copy()
    if  len(x_contr) == 3:
        x_selected = list_var[x_index]
        y_selected = list_var[y_index]
        color_selected = list_color[color_index]
        ## Memorise cette selection pour la retrouver automatiquement au
        ## prochain retour sur l'onglet "Plots"
        dernier_x_scatter = x_selected
        dernier_y_scatter = y_selected
        dernier_color_scatter = color_selected
        x_column.set(x_selected)
        if valeur_check == True:
            var_colors = []
            for i in range(len(data_add)):
                value_str = np.array(data_add[color_selected])
                var_colors.append(str(value_str[i]))
            data_add[color_selected] = var_colors  
        ## Filtre sur les modalites de la variable de couleur (case « Filter modalities »).
        ## data_complet garde toutes les lignes : couleurs et echelle identiques avec ou sans filtre.
        data_complet = data_add
        if check_filtre is not None and check_filtre.get():
            if filtre_modalites_scatter.get("colonne") != color_selected or not filtre_modalites_scatter.get("modalites"):
                check_filtre.set(False)          # autre variable de couleur : le filtre ne s'applique plus
            else:
                a_garder = set(filtre_modalites_scatter["modalites"])
                data_add = data_add[data_add[color_selected].map(_texte_modalite).isin(a_garder)].reset_index(drop = True)
        xm = data_complet[color_selected].reset_index(drop = True)
        color_type = type(xm.iloc[1] if len(xm) > 1 else xm.iloc[0]) == str
        save_graph_p = PhotoImage(file = my_path('figures/save_graph_p.png'))
        my_tile = "" 
        if scatter_title_value is not None:
            if scatter_title_value != "Enter the title":
                my_tile = scatter_title_value
        my_x_label = x_selected
        if scatter_x_value is not None:
            if scatter_x_value != "Enter the x label":
                my_x_label = scatter_x_value
        my_y_label = y_selected
        if scatter_y_value is not None:
            if scatter_y_value != "Enter the y label":
                my_y_label = scatter_y_value
        my_color_label = color_selected
        if scatter_color_value is not None:
            if scatter_color_value != "Enter the color title":
                my_color_label = scatter_color_value
        if color_type == False:
            fig = Figure(figsize=(6, 4), dpi = 80)
            my_scatter_plot = fig
            _enregistrer_figure_courante("Scatter_plot", fig)
            figure_canvas = FigureCanvasTkAgg(fig, master = graph_window) 
            ax = fig.add_subplot()
            my_plot = ax.scatter(data_add[x_selected], data_add[y_selected], c =  data_add[color_selected], cmap = "turbo", s =  int(check_point_size.get()),
                                 vmin = pd.to_numeric(data_complet[color_selected], errors = "coerce").min(), vmax = pd.to_numeric(data_complet[color_selected], errors = "coerce").max())
            fig_colorbar = fig.colorbar(my_plot, ax = ax, shrink= 0.7, aspect = 15, fraction = 0.046)
            fig_colorbar.set_label(my_color_label, fontsize= 15, labelpad = 5)
            ax.set_xlabel(my_x_label, fontsize= 15, labelpad = 7)
            ax.set_ylabel(my_y_label, fontsize= 15, labelpad = 5)
            ax.set_title(my_tile, fontsize= 18)  
            ax.grid()
            _appliquer_zoom_representation(ax, "scatter")
            figure_canvas.get_tk_widget().place(relx =0, rely = 0, relwidth = 1, relheight = 0.8)
            _activer_glissement_representation(ax, figure_canvas, "scatter")
            bouton_save_scatter = Button(graph_window, image = save_graph_p , font=('arial', 14), text = "Save plot",  compound = tk.TOP,  borderwidth = 0, relief = FLAT, borderless = True, bordercolor = "gray90", background = 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
            bouton_save_scatter.place(relx= 0.4, rely= 0.85, width= 100, height= 60)
            bouton_save_scatter.bind("<ButtonRelease>", lambda event: my_save_plot_file(fig)) 
        else:
            fig = Figure(figsize=(12, 8), dpi = 80)
            my_scatter_plot = fig
            _enregistrer_figure_courante("Scatter_plot", fig)
            figure_canvas = FigureCanvasTkAgg(fig, master = graph_window) 
            ax = fig.add_subplot()
            #my_plot = ax.scatter(data_add[x_selected], data_add[y_selected], c =  data_add[color_selected], cmap = "turbo", s = 15)
            xmm = unique(data_add[color_selected])
            modalites_toutes = list(unique(data_complet[color_selected]))     # couleur fixe par modalite, meme filtree
            ##xc = sample(, k = len(colors_palettes))
            for i in range(len(xmm)) :
                dx = data_add[data_add[color_selected] == xmm[i]]
                X = dx[x_selected]
                Y = dx[y_selected]
                ax.scatter(X, Y, label = xmm[i],  s =  int(check_point_size.get()), c = colors_palettes[modalites_toutes.index(xmm[i]) % len(colors_palettes)])
            ax.set_title(my_tile, fontsize= 18)  
            ax.set_xlabel(my_x_label, fontsize= 15, labelpad = 7)
            ax.set_ylabel(my_y_label, fontsize= 15, labelpad = 5)
            ## Legende toujours centree verticalement a droite du graphique, et son nom (vertical)
            ## toujours centre sur la legende, quels que soient le nombre de modalites et les longueurs des textes
            fig.subplots_adjust(right = 0.78)
            ax.legend(xmm, loc = "center left", bbox_to_anchor = (0.815, 0.5), bbox_transform = fig.transFigure,
                      fontsize = 12, markerscale = 2, frameon = False)
            fig.text(0.80, 0.5, my_color_label, rotation = 90, ha = "center", va = "center", fontsize = 12)
            ax.grid(True)
            _appliquer_zoom_representation(ax, "scatter")
            figure_canvas.get_tk_widget().place(relx =0, rely = 0, relwidth = 1, relheight = 0.8)
            _activer_glissement_representation(ax, figure_canvas, "scatter")
            bouton_save_scatter = Button(graph_window, image = save_graph_p , font=('arial', 14), text = "Save plot",  compound = tk.TOP,  borderwidth = 0, relief = FLAT, borderless = True, bordercolor = "gray90", background = 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
            bouton_save_scatter.place(relx= 0.4, rely= 0.85, width= 100, height= 60)
            bouton_save_scatter.bind("<ButtonRelease>", lambda event: my_save_plot_file(fig)) 

        ## ---- Barre verticale gauche : zoom sur ce scatter plot ----
        def _sur_changement_zoom_representation (valeur_curseur):
            zoom_representation_par_type["scatter"] = _curseur_vers_zoom_representation(valeur_curseur)
            scatter_plot(event, parent, my_data, check_var1, check_var2, check_color, list_var, list_color, check_str, check_point_size, check_filtre)

        scale.configure(command = _sur_changement_zoom_representation)
        scale.set(_zoom_vers_curseur_representation(zoom_representation_par_type.get("scatter", 1.0)))
    graph_window.place(relx=0, rely=0, relwidth=1, relheight=1)



# ==========================================================================
# ==  20. Filtres de la fenetre Edit data                                 ==
# ==========================================================================
def _entree_declenche_bouton (bouton):
    """ Fenetres de parametres de filtrage (Min / Max, recodage...) : la touche Entree
    (clavier principal ou pave numerique) declenche le bouton « Save » de la fenetre,
    quel que soit le champ en cours de saisie. Le curseur est place dans le premier champ. """
    panneau = bouton.master

    def declencher (event = None):
        try:
            bouton.invoke()
        except Exception:
            try:
                commande = bouton.cget("command")
                if callable(commande):
                    commande()
            except Exception:
                pass
        return "break"

    champs = []
    for widget in [panneau] + _descendants_widget(panneau):
        if widget is bouton:
            continue
        try:
            widget.bind("<Return>", declencher, add = "+")
            if widget.winfo_class() in ("Entry", "TEntry", "Spinbox", "TSpinbox", "TCombobox"):
                champs.append(widget)
        except Exception:
            pass
    if len(champs) != 0:
        def placer_curseur ():
            try:
                champs[0].focus_set()
                champs[0].select_range(0, END)
            except Exception:
                pass
        panneau.after(50, placer_curseur)


def my_carbon_balance (parent, feature_list, check_carbon_balance):

    carbon_balance_index = check_carbon_balance.current()
    if carbon_balance_index != 0:
        carbon_balance_selected = feature_list[carbon_balance_index]
        height_window_carbon_balance = Canvas(parent, relief= GROOVE ,  background = 'cadetblue' , highlightthickness=0)
        height_window_carbon_balance.place(relx=0.2, rely= 0.1, relwidth= 0.4, relheight = 0.7)
        label_arbon_balance = Label(height_window_carbon_balance, text="Window to filter from Parameter - 1", font=('arial', 15), fg = "gray10", background= "cadetblue")
        label_arbon_balance.place(relx= 0.2, rely= 0.12, relwidth= 0.6)
        label_Carbon_balance_min = Label(height_window_carbon_balance, text="Min", font=('arial', 15), fg = "gray10", background= "cadetblue")
        label_Carbon_balance_min.place(relx=0.15, rely= 0.4, relwidth= 0.2)
        label_Carbon_balance_max = Label(height_window_carbon_balance, text="Max", font=('arial', 15), fg = "gray10", background= "cadetblue")
        label_Carbon_balance_max.place(relx=0.15, rely= 0.6, relwidth= 0.2)
        carbon_min_default = DoubleVar()
        carbon_min_default.set(filtre_edit_parametre_1["min"] if filtre_edit_parametre_1["colonne"] == carbon_balance_selected else 0)
        carbon_min_entry = Entry(height_window_carbon_balance, textvariable = carbon_min_default, justify = CENTER )
        carbon_min_entry.place(relx=0.46, rely = 0.4, width = 150,  height = 25)
        carbon_max_default = DoubleVar()
        carbon_max_default.set(filtre_edit_parametre_1["max"] if filtre_edit_parametre_1["colonne"] == carbon_balance_selected else 100)
        carbon_max_entry = Entry(height_window_carbon_balance, textvariable = carbon_max_default, justify = CENTER )
        carbon_max_entry.place(relx=0.46, rely = 0.6, width = 150,  height = 25)
        destroy_window_carbon = Button(height_window_carbon_balance, text = "Save", command = lambda : destroy_my_carbon_balance(height_window_carbon_balance, carbon_min_default, carbon_max_default, carbon_balance_selected) , font=('arial', 15), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
        destroy_window_carbon.place(relx= 0.46, rely= 0.85, width= 50, height= 20)  
        _entree_declenche_bouton(destroy_window_carbon)      # touche Entree = Save


def destroy_my_carbon_balance (parent, check_min_carbon, check_max_corbon, check_value_corbon):

      global data, carbon_balance_min_value
      carbon_balance_min_value = check_min_carbon.get()
      carbon_balance_max_value = check_max_corbon.get()
      if carbon_balance_min_value < carbon_balance_max_value:
            data = data[data[check_value_corbon] > carbon_balance_min_value]
      if carbon_balance_max_value > carbon_balance_min_value:
            data = data[data[check_value_corbon] <= carbon_balance_max_value]
      filtre_edit_parametre_1["colonne"] = check_value_corbon
      filtre_edit_parametre_1["min"] = carbon_balance_min_value
      filtre_edit_parametre_1["max"] = carbon_balance_max_value
      parent.destroy()


def my_conversion (parent, feature_list, check_conversion_filter):

    conversion_index = check_conversion_filter.current()
    if conversion_index != 0:
        conversion_selected = feature_list[conversion_index]
        height_window_conversion = Canvas(parent, relief= GROOVE ,  background = 'cadetblue' , highlightthickness=0)
        height_window_conversion.place(relx=0.2, rely= 0.1, relwidth= 0.4, relheight = 0.7)
        label_conversion = Label(height_window_conversion, text="Window to filter from Parameter - 2", font=('arial', 15), fg = "gray10", background= "cadetblue")
        label_conversion.place(relx= 0.2, rely= 0.12, relwidth= 0.6)
        label_conversion_min = Label(height_window_conversion, text="Min", font=('arial', 15), fg = "gray10", background= "cadetblue")
        label_conversion_min.place(relx=0.15, rely= 0.4, relwidth= 0.2)
        label_conversion_max = Label(height_window_conversion, text="Max", font=('arial', 15), fg = "gray10", background= "cadetblue")
        label_conversion_max.place(relx=0.15, rely= 0.6, relwidth= 0.2)
        conversion_min_default = DoubleVar()
        conversion_min_default.set(filtre_edit_parametre_2["min"] if filtre_edit_parametre_2["colonne"] == conversion_selected else 0)
        conversion_min_entry = Entry(height_window_conversion, textvariable = conversion_min_default, justify = CENTER )
        conversion_min_entry.place(relx=0.46, rely = 0.4, width = 150,  height = 25)
        conversion_max_default =DoubleVar()
        conversion_max_default.set(filtre_edit_parametre_2["max"] if filtre_edit_parametre_2["colonne"] == conversion_selected else 100)
        conversion_max_entry = Entry(height_window_conversion, textvariable = conversion_max_default, justify = CENTER )
        conversion_max_entry.place(relx=0.46, rely = 0.6, width = 150,  height = 25)
        destroy_window_conversion = Button(height_window_conversion, text = "Save", command = lambda : destroy_my_conversion(height_window_conversion, conversion_min_default, conversion_max_default, conversion_selected) , font=('arial', 15), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
        destroy_window_conversion.place(relx= 0.46, rely= 0.85, width= 50, height= 20)        
        _entree_declenche_bouton(destroy_window_conversion)      # touche Entree = Save


def destroy_my_conversion (parent, check_min_conv, check_max_conv, check_value_conversion):

      global data
      conversion_min_value = check_min_conv.get()
      conversion_max_value = check_max_conv.get()
      if conversion_min_value < conversion_max_value:
            data = data[data[check_value_conversion] > conversion_min_value]
      if conversion_max_value > conversion_min_value:
             data = data[data[check_value_conversion] <= conversion_max_value]
      filtre_edit_parametre_2["colonne"] = check_value_conversion
      filtre_edit_parametre_2["min"] = conversion_min_value
      filtre_edit_parametre_2["max"] = conversion_max_value
      parent.destroy()


def my_yields (parent, feature_list, check_yields_filter):

    yields_index = check_yields_filter.current()
    if yields_index != 0:
        yields_selected = feature_list[yields_index]
        height_window_yields = Canvas(parent, relief= GROOVE ,  background = 'cadetblue' , highlightthickness=0)
        height_window_yields.place(relx=0.2, rely= 0.1, relwidth= 0.4, relheight = 0.7)
        label_yields = Label(height_window_yields, text="Window to filter from Parameter - 3", font=('arial', 15), fg = "gray10", background= "cadetblue")
        label_yields.place(relx= 0.2, rely= 0.12, relwidth= 0.6)
        label_yields_min = Label(height_window_yields, text="Min", font=('arial', 15), fg = "gray10", background= "cadetblue")
        label_yields_min.place(relx=0.15, rely= 0.4, relwidth= 0.2)
        label_yields_max = Label(height_window_yields, text="Max", font=('arial', 15), fg = "gray10", background= "cadetblue")
        label_yields_max.place(relx=0.15, rely= 0.6, relwidth= 0.2)
        yields_min_default = DoubleVar()
        yields_min_default.set(filtre_edit_parametre_3["min"] if filtre_edit_parametre_3["colonne"] == yields_selected else 0)
        yields_min_entry = Entry(height_window_yields, textvariable = yields_min_default, justify = CENTER )
        yields_min_entry.place(relx=0.46, rely = 0.4, width = 150,  height = 25)
        yields_max_default = DoubleVar()
        yields_max_default.set(filtre_edit_parametre_3["max"] if filtre_edit_parametre_3["colonne"] == yields_selected else 100)
        yields_max_entry = Entry(height_window_yields, textvariable = yields_max_default, justify = CENTER )
        yields_max_entry.place(relx=0.46, rely = 0.6, width = 150,  height = 25)
        destroy_window_yields = Button(height_window_yields, text = "Save", command = lambda : destroy_my_yields(height_window_yields, yields_min_default, yields_max_default, yields_selected) , font=('arial', 15), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
        destroy_window_yields.place(relx= 0.46, rely= 0.85, width= 50, height= 20)        
        _entree_declenche_bouton(destroy_window_yields)      # touche Entree = Save


def destroy_my_yields (parent, check_min_yied, check_max_yied, check_value_yield):

      global data
      yields_min_value = check_min_yied.get()
      yields_max_value = check_max_yied.get()
      if yields_min_value < yields_max_value:
             data = data[data[check_value_yield] > yields_min_value]
      if yields_max_value > yields_min_value:
             data = data[data[check_value_yield] <= yields_max_value]
      filtre_edit_parametre_3["colonne"] = check_value_yield
      filtre_edit_parametre_3["min"] = yields_min_value
      filtre_edit_parametre_3["max"] = yields_max_value
      parent.destroy()


def my_selectivity (parent,  feature_list, check_selectivity_filter):

    selectivity_index = check_selectivity_filter.current()
    if selectivity_index != 0:
        selectivity_selected = feature_list[selectivity_index]
        height_window_selectivity = Canvas(parent, relief= GROOVE ,  background = 'cadetblue' , highlightthickness=0)
        height_window_selectivity.place(relx=0.2, rely= 0.1, relwidth= 0.4, relheight = 0.7)
        label_selectivity = Label(height_window_selectivity, text="Window to filter from Parameter - 4", font=('arial', 15), fg = "gray10", background= "cadetblue")
        label_selectivity.place(relx= 0.2, rely= 0.12, relwidth= 0.6)
        label_selectivity_min = Label(height_window_selectivity, text="Min", font=('arial', 15), fg = "gray10", background= "cadetblue")
        label_selectivity_min.place(relx=0.15, rely= 0.4, relwidth= 0.2)
        label_selectivity_max = Label(height_window_selectivity, text="Max", font=('arial', 15), fg = "gray10", background= "cadetblue")
        label_selectivity_max.place(relx=0.15, rely= 0.6, relwidth= 0.2)
        selectivity_min_default = DoubleVar()
        selectivity_min_default.set(filtre_edit_parametre_4["min"] if filtre_edit_parametre_4["colonne"] == selectivity_selected else 0)
        selectivity_min_entry = Entry(height_window_selectivity, textvariable = selectivity_min_default, justify = CENTER )
        selectivity_min_entry.place(relx=0.46, rely = 0.4, width = 150,  height = 25)
        selectivity_max_default = DoubleVar()
        selectivity_max_default.set(filtre_edit_parametre_4["max"] if filtre_edit_parametre_4["colonne"] == selectivity_selected else 100)
        selectivity_max_entry = Entry(height_window_selectivity, textvariable = selectivity_max_default, justify = CENTER )
        selectivity_max_entry.place(relx=0.46, rely = 0.6, width = 150,  height = 25)
        destroy_window_selectivity = Button(height_window_selectivity, text = "Save", command = lambda : destroy_my_selectivity(height_window_selectivity, selectivity_min_default, selectivity_max_default, selectivity_selected) , font=('arial', 15), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
        destroy_window_selectivity.place(relx= 0.46, rely= 0.85, width= 50, height= 20)        
        _entree_declenche_bouton(destroy_window_selectivity)      # touche Entree = Save


def destroy_my_selectivity (parent, check_min_select, check_max_select, check_value_selct):

      global data
      selectivity_min_value = check_min_select.get()
      selectivity_max_value = check_max_select.get()
      if selectivity_min_value < selectivity_max_value:
           data = data[data[check_value_selct] > selectivity_min_value]
      if selectivity_max_value > selectivity_min_value:
           data = data[data[check_value_selct] <= selectivity_max_value]
      filtre_edit_parametre_4["colonne"] = check_value_selct
      filtre_edit_parametre_4["min"] = selectivity_min_value
      filtre_edit_parametre_4["max"] = selectivity_max_value
      parent.destroy()


def my_categorial_function (parent, feature_list, check_recod_stat, my_data):

    if check_recod_stat.get() != False:
        height_window_recod = Canvas(parent, relief= GROOVE ,  background = 'cadetblue' , highlightthickness=0)
        height_window_recod.place(relx=0.2, rely= 0.1, relwidth= 0.4, relheight = 0.7)
        Feature_categorial = tk.Listbox(height_window_recod, fg = "gray30", font = ('arial', 14), relief = GROOVE , selectmode = tk.MULTIPLE, exportselection = False)
        Feature_categorial.place(relx = 0.01, rely = 0.02, relwidth = 0.98, relheight= 0.85)
        feature_list_used = feature_list.copy()
        del feature_list_used[0]
        for value in feature_list_used:
            Feature_categorial.insert(END, value)
        destroy_window_recod = Button(height_window_recod, text = "Save", command = lambda : destroy_window(height_window_recod, Feature_categorial, check_recod_stat, my_data) , font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
        destroy_window_recod.place(relx= 0.46, rely= 0.9, width= 50, height= 20)        
        _entree_declenche_bouton(destroy_window_recod)      # touche Entree = Save


def destroy_window (parent, Feature_categorial_check, value_check, check_my_data):

    global result_recod_stat, data_sort_stat
    result_recod_stat = [Feature_categorial_check.get(i) for i in Feature_categorial_check.curselection()]
    if len(result_recod_stat) > 0:
        data_sort_stat = []
        for i in range(len(result_recod_stat)):
            data_sort_stat.append(sort(unique((check_my_data[result_recod_stat[i]]))))
    parent.destroy()


def ajusted_feature(vect):

    vect_len = []
    for caract in vect:
        vect_len.append(len(caract))
    vect_len_max = max(vect_len)
    vect_new = []
    dif_len_save = []
    for i in range(len(vect)):
        dif_len = vect_len_max - vect_len[i] 
        x_empty = ""
        dif_len_save_in = ""
        for j in range(dif_len):
            x_empty = x_empty + " "
            dif_len_save_in = dif_len_save_in + " "
        vect_new.append(str(vect[i] + x_empty))
        dif_len_save.append(dif_len_save_in)
    return vect_new , dif_len_save



# ==========================================================================
# ==  21. Detail des modeles et graphiques predit / reel                  ==
# ==========================================================================
## ------ « Detail of model » : details de TOUS les modeles du projet ------

_panneau_detail_modeles = {"widget": None}


def _hyperparametres_modele_pour_detail (etat):
    """Hyperparametres utiles d'un modele, lus dans SON etat (et non dans les variables globales du modele actif)."""
    nom_algo = etat.get("Name_model")
    donnees = etat.get("data_model")
    if etat.get("selectet_methode_imp") == "Cross-validation" or etat.get("selectet_methode") == "Cross-validation":
        try:
            meilleurs_parametres = donnees[6]
            if isinstance(meilleurs_parametres, dict):
                return dict(meilleurs_parametres)
        except Exception:
            pass
        return {}
    cles_par_algo = {
        "Lasso": ["alpha"],
        "Kernel ridge (KR)": ["alpha"],
        "Support Vector Regression (SVR)": ["C", "gamma", "kernel"],
        "Random Forest (RF)": ["n_estimators", "random_state"],
        "Generalized Boosted Regression Modeling (GBM)": ["n_estimators", "max_depth", "learning_rate", "loss"],
        "Decision Tree (DT)": ["max_depth", "random_state"],
        "Gaussian process (GP)": ["random_state"],
    }
    resultat = {}
    reg_modele = etat.get("reg")
    if reg_modele is not None and hasattr(reg_modele, "get_params"):
        try:
            parametres = reg_modele.get_params()
            for cle in cles_par_algo.get(nom_algo, []):
                if cle in parametres:
                    resultat[cle] = parametres[cle]
        except Exception:
            pass
    return resultat


def _nombre_pour_detail (valeur, decimales = 4):
    try:
        return ("%." + str(decimales) + "f") % float(valeur)
    except Exception:
        return str(valeur)


def _nombre_observations_modele (donnees):
    """ (N entrainement, N test) d'un modele entraine, lus dans data_model
    ([5] = Y d'entrainement, [2] = Y de test) ; (None, None) si indisponible. """
    try:
        return len(donnees[5]), len(donnees[2])
    except Exception:
        return None, None


def _lignes_nombre_observations (donnees, methode):
    """ Lignes « Number of observations » affichees dans « Detail of model ». """
    n_train, n_test = _nombre_observations_modele(donnees)
    if n_train is None:
        return []
    total = n_train + n_test
    libelle_train = "Training data (hyperparameter tuning + model fitting)" if methode == "Cross-validation" else "Training data"
    pourcentage = (lambda n: " (" + ("%.0f" % (100.0 * n / total)) + " %)") if total > 0 else (lambda n: "")
    return ["Number of observations:",
            "   " + libelle_train + " = " + str(n_train) + pourcentage(n_train),
            "   Test data (prediction) = " + str(n_test) + pourcentage(n_test),
            "   Total = " + str(total)]


def _texte_detail_un_modele (nom, etat, rang, est_meilleur):
    lignes = []
    titre = "MODEL " + str(rang) + " : " + str(nom)
    if est_meilleur:
        titre += "     *** BEST MODEL (used for prediction) ***"
    lignes.append(titre)
    lignes.append("-" * 70)
    lignes.append("Status: " + _statut_affichage_modele(nom).strip())
    lignes.append("Algorithm: " + str(etat.get("Name_model")))
    lignes.append("Name of data used: " + (os.path.basename(file_path) if file_path else "-"))
    lignes.append("Target variable: " + str(etat.get("selectet_target")))
    lignes.append("Recoding: " + str(etat.get("selectet_codage")))
    methode = etat.get("selectet_methode_imp") or etat.get("selectet_methode")
    taille = etat.get("size_data_value")
    if methode == "Cross-validation":
        try:
            lignes.append("Data partitioning: " + str(methode) + ";  K = " + str(int(taille)))
        except Exception:
            lignes.append("Data partitioning: " + str(methode))
    else:
        lignes.append("Data partitioning: " + str(methode) + ";  Test size = " + str(taille))
    parametres = _hyperparametres_modele_pour_detail(etat)
    if len(parametres) != 0:
        lignes.append("Hyperparameters: " + ";  ".join([str(cle) + " = " + str(valeur) for cle, valeur in parametres.items()]))

    donnees = etat.get("data_model")
    lignes.append("")
    if donnees is None:
        lignes.append("Metrics for model evaluation: not available (model not trained yet - click on Run)")
    else:
        lignes.append("Metrics for model evaluation:")
        lignes.append("   R-squared = " + _nombre_pour_detail(donnees[0], 4))
        lignes.append("   MSE       = " + _nombre_pour_detail(donnees[3], 4))
        lignes_observations = _lignes_nombre_observations(donnees, methode)
        if len(lignes_observations) != 0:
            lignes.append("")
            lignes.extend(lignes_observations)

    features = etat.get("result_model")
    if features is not None and len(features) != 0:
        lignes.append("")
        lignes.append("Feature columns: (N = " + str(len(features)) + ")")
        for valeur in features:
            lignes.append("   " + str(valeur))
        codage = etat.get("selectet_codage")
        lignes.append("")
        if codage == "Disjunctive":
            recod = etat.get("result_model_recod")
            recod = [] if recod is None else list(recod)
            lignes.append("Features to be recoded: (N = " + str(len(recod)) + ")")
            for valeur in recod:
                lignes.append("   " + str(valeur))
        elif codage == "All-disjunctive":
            lignes.append("Features to be recoded: all feature columns")
        else:
            lignes.append("Features to be recoded: None")

    colonnes = etat.get("columns_select")
    coefficients = etat.get("my_data_coef")
    if donnees is not None and colonnes is not None and len(colonnes) != 0:
        nom_algo = etat.get("Name_model")
        etiquette = None
        if nom_algo in ("Linear", "Lasso"):
            etiquette = "Estimate"
        elif nom_algo in ("Random Forest (RF)", "Generalized Boosted Regression Modeling (GBM)", "Decision Tree (DT)"):
            etiquette = "Importance"
        lignes.append("")
        avec_coefficients = False
        try:
            avec_coefficients = etiquette is not None and coefficients is not None and not isinstance(coefficients, dict) and len(coefficients) == len(colonnes)
        except Exception:
            avec_coefficients = False
        if avec_coefficients:
            lignes.append("Best features used (" + etiquette + "): (N = " + str(len(colonnes)) + ")")
            for k in range(len(colonnes)):
                lignes.append("   " + str(colonnes[k]) + "   " + _nombre_pour_detail(coefficients[k], 5))
        else:
            lignes.append("Best features used: (N = " + str(len(colonnes)) + ")")
            for valeur in colonnes:
                lignes.append("   " + str(valeur))
    return "\n".join(lignes) + "\n"


def _texte_detail_tous_les_modeles ():
    noms = list(modeles_projet.keys())
    entraines = [nom for nom in noms if modeles_projet[nom].get("data_model") is not None]
    non_entraines = [nom for nom in noms if nom not in entraines]

    def _cle_tri (nom):
        r2 = _r2_modele_entraine(nom)
        try:
            return -float(r2) if np.isfinite(r2) else float("inf")
        except Exception:
            return float("inf")

    entraines = sorted(entraines, key = _cle_tri)      ## meilleur R-squared d'abord
    ordre = entraines + non_entraines
    nom_meilleur, r2_meilleur = _meilleur_modele_entraine(entraines, pour_prediction = True)

    lignes = []
    lignes.append("DETAIL OF ALL MODELS  (N = " + str(len(ordre)) + ";  trained: " + str(len(entraines)) + ")")
    lignes.append("=" * 70)
    if nom_meilleur is not None:
        lignes.append("Best model for prediction: " + str(nom_meilleur) + "  (R-squared = " + _nombre_pour_detail(r2_meilleur, 4) + ")")
    else:
        lignes.append("Best model for prediction: none (no trained model usable for prediction)")
    lignes.append("")
    largeur_nom = 4
    for nom in ordre:
        if len(str(nom)) > largeur_nom:
            largeur_nom = len(str(nom))
    lignes.append("  #  " + "Name".ljust(largeur_nom) + "   " + "R-squared".ljust(10) + "MSE".ljust(12) + "Algorithm")
    for rang, nom in enumerate(ordre, start = 1):
        etat = modeles_projet[nom]
        donnees = etat.get("data_model")
        r2_txt = _nombre_pour_detail(donnees[0], 4) if donnees is not None else "-"
        mse_txt = _nombre_pour_detail(donnees[3], 4) if donnees is not None else "-"
        marque = "*" if nom == nom_meilleur else " "
        lignes.append(marque + str(rang).rjust(2) + "  " + str(nom).ljust(largeur_nom) + "   " + r2_txt.ljust(10) + mse_txt.ljust(12) + str(etat.get("Name_model")))
    lignes.append("")
    lignes.append("(* = best model)")
    lignes.append("")
    texte = "\n".join(lignes) + "\n"
    for rang, nom in enumerate(ordre, start = 1):
        texte += "\n" + "=" * 70 + "\n" + _texte_detail_un_modele(nom, modeles_projet[nom], rang, nom == nom_meilleur)
    return texte


def my_model_detail_function (parent, check_recod_stat):

    ancien_panneau = _panneau_detail_modeles.get("widget")
    if _widget_existe(ancien_panneau):
        try:
            ancien_panneau.destroy()
        except Exception:
            pass
    _panneau_detail_modeles["widget"] = None

    if check_recod_stat.get() == False:
        return

    height_window_recod = Canvas(parent, relief= GROOVE ,  background = 'cadetblue' , highlightthickness=0)
    height_window_recod.place(relx=0.1, rely= 0.05, relwidth= 0.8, relheight = 0.9)
    _panneau_detail_modeles["widget"] = height_window_recod

    zone_texte = tk.Frame(height_window_recod, background = 'cadetblue')
    zone_texte.place(relx = 0.01, rely = 0.01, relwidth = 0.98, relheight = 0.9)
    barre_defilement = tk.Scrollbar(zone_texte, orient = "vertical")
    barre_defilement.pack(side = "right", fill = "y")
    texte_detail = Text(zone_texte, padx = 10, pady = 10, font = ('Courier New', 12), relief = GROOVE, highlightthickness = 1, fg = "gray30", wrap = "none", yscrollcommand = barre_defilement.set)
    texte_detail.pack(side = "left", fill = "both", expand = True)
    barre_defilement.configure(command = texte_detail.yview)

    if len(modeles_projet) == 0:
        texte_detail.insert(INSERT, "No model defined in this project.")
    else:
        try:
            texte_detail.insert(INSERT, _texte_detail_tous_les_modeles())
        except Exception as erreur_detail:
            texte_detail.insert(INSERT, "Unable to display the model details :\n" + str(erreur_detail))
    texte_detail.configure(state = "disabled")

    def _fermer_detail_modeles ():
        destroy_model_detail(height_window_recod)
        _panneau_detail_modeles["widget"] = None
        try:
            check_recod_stat.set(False)
        except Exception:
            pass

    destroy_window_recod = Button(height_window_recod, text = "Close", command = _fermer_detail_modeles, font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = 'gray90', background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
    destroy_window_recod.place(relx= 0.46, rely= 0.93, width= 50, height= 20)
    _entree_declenche_bouton(destroy_window_recod)      # touche Entree = Save


def destroy_model_detail (parent):

    parent.destroy()


def summary_for_export (parent, check_recod_stat):

    if check_recod_stat.get() != False:
        height_window_recod = Canvas(parent, relief= GROOVE ,  background = 'cadetblue' , highlightthickness=0)
        window_export = Canvas(height_window_recod, relief= GROOVE ,  background = 'gray100' , highlightthickness=0)
        window_export.place(relx=0.01, rely= 0.01, relwidth= 0.98, relheight = 0.9)
        height_window_recod.place(relx=0.1, rely= 0.05, relwidth= 0.8, relheight = 0.9)
        destroy_window_recod = Button(height_window_recod, text = "Close", command = lambda : destroy_summary_for_export(height_window_recod) , font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
        destroy_window_recod.place(relx= 0.46, rely= 0.93, width= 50, height= 20)        
        _entree_declenche_bouton(destroy_window_recod)      # touche Entree = Save


def destroy_summary_for_export (parent):

    parent.destroy()


def _legende_avec_r2 (axes, donnees_modele, avec_train):
    """ Legende du graphique « valeurs reelles / predites », en haut a gauche, chaque serie suivie de son R^2
    sur la meme ligne : « Train   R^2 = 0.975 » / « Test   R^2 = 0.903 ». """
    def texte_r2 (valeur):
        try:
            return "   $R^2$ = %.3f" % float(valeur)
        except Exception:
            return ""
    etiquettes = []
    if avec_train:
        try:
            r2_train = r2_squared(np.asarray(donnees_modele[5], dtype = float), np.asarray(donnees_modele[4], dtype = float).ravel())
        except Exception:
            r2_train = None
        etiquettes.append("Train" + (texte_r2(r2_train) if r2_train is not None else ""))
    etiquettes.append("Test" + texte_r2(donnees_modele[0]))
    axes.legend(etiquettes, fontsize = 12, loc = "upper left")


def check_model_run (window_model_run, color_value_modell, row_edit_scatter_default_modell, color_value_linearl, point_size_scatter_linearl, color_value_trainl, row_edit_scatter_default_linearl, row_edit_scatter_default_trainl, my_check_model_detaill, my_check_model_trainl):

    global  model_plot, model_text_label, zoom_graphique_modeles, decalage_zoom_x, decalage_zoom_y

    def _bornes_axes_zoomees(borne_min, borne_max, decalage):
        ## resserre ou elargit l'intervalle [borne_min, borne_max] autour de son
        ## centre, en fonction du niveau de zoom courant (pilote par la barre
        ## verticale gauche pendant la consultation des graphiques de modeles),
        ## puis applique le decalage (pan) accumule par glissement a la souris
        centre = (borne_min + borne_max) / 2 + decalage
        demi_etendue = (borne_max - borne_min) / 2
        facteur = zoom_graphique_modeles if zoom_graphique_modeles > 0 else 1.0
        demi_etendue = demi_etendue / facteur
        return centre - demi_etendue, centre + demi_etendue

    def _activer_glissement_graphique(axes, canvas_mpl):
        ## permet, une fois zoome, de deplacer la vue a la souris (clic-glisse)
        ## pour explorer toutes les parties de la courbe
        tk_widget = canvas_mpl.get_tk_widget()
        etat_glissement = {"x": None, "y": None}

        def _debut_glissement(event):
            etat_glissement["x"] = event.x
            etat_glissement["y"] = event.y

        def _glisser(event):
            global decalage_zoom_x, decalage_zoom_y
            x0 = etat_glissement.get("x")
            y0 = etat_glissement.get("y")
            if x0 is None or y0 is None:
                return
            delta_px_x = event.x - x0
            delta_px_y = event.y - y0
            largeur_widget = tk_widget.winfo_width()
            hauteur_widget = tk_widget.winfo_height()
            if largeur_widget > 1 and hauteur_widget > 1:
                xlim = axes.get_xlim()
                ylim = axes.get_ylim()
                pas_x = delta_px_x * (xlim[1] - xlim[0]) / largeur_widget
                pas_y = delta_px_y * (ylim[1] - ylim[0]) / hauteur_widget
                decalage_zoom_x -= pas_x
                decalage_zoom_y += pas_y
                axes.set_xlim(xlim[0] - pas_x, xlim[1] - pas_x)
                axes.set_ylim(ylim[0] + pas_y, ylim[1] + pas_y)
                canvas_mpl.draw_idle()
            etat_glissement["x"] = event.x
            etat_glissement["y"] = event.y

        def _fin_glissement(event):
            etat_glissement["x"] = None
            etat_glissement["y"] = None

        def _reinitialiser_vue(event):
            ## double-clic sur le graphique : retour a la vue standard
            ## (sans zoom ni deplacement)
            global zoom_graphique_modeles, decalage_zoom_x, decalage_zoom_y
            zoom_graphique_modeles = 1.0
            decalage_zoom_x = 0.0
            decalage_zoom_y = 0.0
            try:
                valeur_curseur = (1.0 - ZOOM_MIN_GRAPHIQUE) / (ZOOM_MAX_GRAPHIQUE - ZOOM_MIN_GRAPHIQUE) * 100
                scale.set(int(valeur_curseur))
            except Exception:
                pass
            check_model_run(window_model_run, color_value_modell, row_edit_scatter_default_modell, color_value_linearl, point_size_scatter_linearl, color_value_trainl, row_edit_scatter_default_linearl, row_edit_scatter_default_trainl, my_check_model_detaill, my_check_model_trainl)

        tk_widget.configure(cursor = "fleur")
        tk_widget.bind("<ButtonPress-1>", _debut_glissement)
        tk_widget.bind("<B1-Motion>", _glisser)
        tk_widget.bind("<ButtonRelease-1>", _fin_glissement)
        tk_widget.bind("<Double-Button-1>", _reinitialiser_vue)

    model_text_label = Text(window_model_run, padx = 10, pady = 10, font=('arial', 15), relief = GROOVE, highlightthickness = 1 , fg= "gray30")  
    model_text_label.insert(INSERT, "Description of the model \n \n" + "Name of data used: " + os.path.basename(file_path) + "\n \n" +  "Name model: " + Name_model  + "\n \n" + "Target variable: " + selectet_target + "\n \n" + "Recoding: " + selectet_codage + "\n \n")
    ##minit = "Choose model"
    m1 = "Linear"
    m1b = "Lasso"
    m2 = "Kernel ridge (KR)"
    m3 = "Support Vector Regression (SVR)"
    m4 = "Random Forest (RF)"
    m5 = "Generalized Boosted Regression Modeling (GBM)"
    m6 = "Decision Tree (DT)"
    m7 = "Gaussian process (GP)"
    y_min_conv = data_model[1].min() - 5
    y_max_conv = data_model[1].max() + 5
    model_index = color_value_modell.current()
    selectet_model_color = colors_palettes[model_index]
    linear_index = color_value_linearl.current()
    selectet_linear_color = colors_palettes[linear_index]
    train_index = color_value_trainl.current()
    selectet_train_color = colors_palettes[train_index]
    if my_check_model_detaill.get() == True:
                if selectet_methode == "Cross-validation":
                    if Name_model == m3:
                       model_text_label.insert(END, "Data partitioning: "  + selectet_methode + ";  " + "K = " + str(int(size_data_value)) + "\n \n" +  "C = " + str(data_model[6].get("C")) + ";  " + "gamma = " + str(data_model[6].get("gamma")) + ";  " + "kernel = " + str(data_model[6].get("kernel")) + "\n \n"  + "Metrics for model evaluation: \n \n"  + "R-squared = " + str(round(data_model[0], 2)) + "\n \n" +  "MSE = " + str(round(data_model[3],2)))  
                    elif Name_model == m4:
                        model_text_label.insert(END, "Data partitioning: "  + selectet_methode + ";  " + "K = " + str(int(size_data_value)) +  "\n \n"  + "n_estimators = " + str(data_model[6].get("n_estimators")) + "\n \n" + "Metrics for model evaluation: \n \n"  + "R-squared = " + str(round(data_model[0], 2)) + "\n \n" +  "MSE = " + str(round(data_model[3],2))) 
                    elif Name_model == m5:
                        model_text_label.insert(END, "Data partitioning: "  + selectet_methode + ";  " + "K = " + str(int(size_data_value)) + "\n \n" + "Metrics for model evaluation: \n \n"  + "R-squared = " + str(round(data_model[0], 2)) + "\n \n" +  "MSE = " + str(round(data_model[3],2))) 
                    elif Name_model == m6:
                       model_text_label.insert(END, "Data partitioning: "  + selectet_methode + ";  " + "K = " + str(int(size_data_value)) + "\n \n" + "max_depth = " + str(data_model[6].get("max_depth")) + "\n \n" + "Metrics for model evaluation: \n \n"  + "R-squared = " + str(round(data_model[0], 2)) + "\n \n" +  "MSE = " + str(round(data_model[3],2))) 
                    else:
                       model_text_label.insert(END, "Data partitioning: "  + selectet_methode + ";  " + "K = " + str(int(size_data_value)) + "\n \n" + "Metrics for model evaluation: \n \n"  + "R-squared = " + str(round(data_model[0], 2)) + "\n \n" +  "MSE = " + str(round(data_model[3],2)))                    
                else:
                    if Name_model == m1:
                       model_text_label.insert(END, "Data partitioning: "  + selectet_methode + ";  " + "Test size = " + str(size_data_value) + "\n \n" + "Metrics for model evaluation: \n \n"  + "R-squared = " + str(round(data_model[0], 2)) + "\n \n" +  "MSE = " + str(round(data_model[3],2))) 
                    elif Name_model == m1b:
                       model_text_label.insert(END, "Data partitioning: "  + selectet_methode + ";  " + "Test size = " + str(size_data_value) + "\n \n" + "Alpha = " + str(float(Alpha_lasso_size.get())) + "\n \n" + "Metrics for model evaluation: \n \n"  + "R-squared = " + str(round(data_model[0], 2)) + "\n \n" +  "MSE = " + str(round(data_model[3],2))) 
                    elif Name_model == m2:
                        model_text_label.insert(END, "Data partitioning: "  + selectet_methode + ";  " + "Test size = " + str(size_data_value) + "\n \n" + "Metrics for model evaluation: \n \n"  + "R-squared = " + str(round(data_model[0], 2)) + "\n \n" +  "MSE = " + str(round(data_model[3],2))) 
                    elif Name_model == m3:
                       model_text_label.insert(END, "Data partitioning: "  + selectet_methode + ";  " + "Test size = " + str(size_data_value) + "\n \n" + "Metrics for model evaluation: \n \n"  + "R-squared = " + str(round(data_model[0], 2)) + "\n \n" +  "MSE = " + str(round(data_model[3],2)))  
                    elif Name_model == m4:
                        model_text_label.insert(END, "Data partitioning: "  + selectet_methode + ";  " + "Test size = " + str(size_data_value) + "\n \n" + "Metrics for model evaluation: \n \n"  + "R-squared = " + str(round(data_model[0], 2)) + "\n \n" +  "MSE = " + str(round(data_model[3],2))) 
                    elif Name_model == m5:
                        model_text_label.insert(END, "Data partitioning: "  + selectet_methode + ";  " + "Test size = " + str(size_data_value) + "\n \n" + "Metrics for model evaluation: \n \n"  + "R-squared = " + str(round(data_model[0], 2)) + "\n \n" +  "MSE = " + str(round(data_model[3],2))) 
                    elif Name_model == m6:
                       model_text_label.insert(END, "Data partitioning: "  + selectet_methode + ";  " + "Test size = " + str(size_data_value) + "\n \n" + "Metrics for model evaluation: \n \n"  + "R-squared = " + str(round(data_model[0], 2)) + "\n \n" +  "MSE = " + str(round(data_model[3],2))) 
                    else:
                       model_text_label.insert(END, "Data partitioning: "  + selectet_methode + ";  " + "Test size = " + str(size_data_value) + "\n \n" + "Metrics for model evaluation: \n \n"  + "R-squared = " + str(round(data_model[0], 2)) + "\n \n" +  "MSE = " + str(round(data_model[3],2)))                    
                # nombre d'observations utilisees pour l'entrainement et pour le test
                lignes_observations = _lignes_nombre_observations(data_model, selectet_methode)
                if len(lignes_observations) != 0:
                    model_text_label.insert(END, "\n \n" + lignes_observations[0] + " \n \n" + "\n \n".join(ligne.strip() for ligne in lignes_observations[1:]))
                model_text_label.place(relx=0.6, rely=0, relwidth= 0.4, relheight= 1)
                figure = Figure(figsize=(12, 8), dpi = 80)
                _enregistrer_figure_courante("Model_prediction", figure)
                figure_canvas_detail = FigureCanvasTkAgg(figure, window_model_run)
                ##plt.legend([data_model_name[Name_model][0], "R2 = " + str(round(data_model[0], 2))], fontsize = 8)
                plt = figure.add_subplot()
                ##if my_check_model_trainl.get() == True and selectet_methode != "Cross-validation":
                if my_check_model_trainl.get() == True:
                    plt.scatter(data_model[5], data_model[4], c = selectet_train_color, s = int(row_edit_scatter_default_trainl.get()))
                ## CORRECTIF : valeurs REELLES en abscisse et PREDITES en ordonnee, comme l'indiquent les titres des axes
                ## (data_model = [R2, pred_test, Y_test, mse, pred_train, Y_train, ...] : les deux etaient inverses)
                plt.scatter(data_model[2], data_model[1], c = selectet_model_color, s = int(row_edit_scatter_default_modell.get()))
                plt.plot(np.arange(my_y_min, my_y_max), np.arange(my_y_min, my_y_max), color = selectet_linear_color, lw = int(row_edit_scatter_default_linearl.get()))
                plt.set_ylabel("Predicted " + f"{selectet_target}", fontsize= 13) 
                plt.set_xlabel("True " + f"{selectet_target}", fontsize= 13)
                _legende_avec_r2(plt, data_model, my_check_model_trainl.get() == True)     # R^2 sur la ligne de chaque serie
                plt.grid(True)
                axe_min_x, axe_max_x = _bornes_axes_zoomees(my_y_min, my_y_max, decalage_zoom_x)
                axe_min_y, axe_max_y = _bornes_axes_zoomees(my_y_min, my_y_max, decalage_zoom_y)
                plt.axis([axe_min_x, axe_max_x, axe_min_y, axe_max_y])
                figure_canvas_detail.get_tk_widget().place(relx = 0, rely = 0, relwidth = 0.6, relheight = 1) 
                _activer_glissement_graphique(plt, figure_canvas_detail)
                model_plot = figure
    else:
        figure = Figure(figsize=(12, 8), dpi = 80)
        _enregistrer_figure_courante("Model_prediction", figure)
        figure_canvas = FigureCanvasTkAgg(figure, window_model_run)
        plt = figure.add_subplot()
        if my_check_model_trainl.get() == True:
           plt.scatter(data_model[5], data_model[4], c = selectet_train_color, s = int(row_edit_scatter_default_trainl.get()))
        ## CORRECTIF : valeurs REELLES en abscisse et PREDITES en ordonnee, comme l'indiquent les titres des axes
        ## (data_model = [R2, pred_test, Y_test, mse, pred_train, Y_train, ...] : les deux etaient inverses)
        plt.scatter(data_model[2], data_model[1], c = selectet_model_color, s = int(row_edit_scatter_default_modell.get()))
        plt.plot(np.arange(my_y_min, my_y_max), np.arange(my_y_min, my_y_max), color = selectet_linear_color, lw = int(row_edit_scatter_default_linearl.get()))
        plt.set_ylabel("Predicted " + f"{selectet_target}", fontsize= 13) 
        plt.set_xlabel("True " + f"{selectet_target}", fontsize= 13)
        _legende_avec_r2(plt, data_model, my_check_model_trainl.get() == True)     # R^2 sur la ligne de chaque serie
        plt.grid(True)
        axe_min_x, axe_max_x = _bornes_axes_zoomees(my_y_min, my_y_max, decalage_zoom_x)
        axe_min_y, axe_max_y = _bornes_axes_zoomees(my_y_min, my_y_max, decalage_zoom_y)
        plt.axis([axe_min_x, axe_max_x, axe_min_y, axe_max_y])
        figure_canvas.get_tk_widget().place(relx = 0, rely = 0, relwidth = 1, relheight = 1) 
        _activer_glissement_graphique(plt, figure_canvas)
        model_plot = figure


def my_model_plot_function (parent, check_model_plot):

    if data_model is not None:
        if check_model_plot.get() != False:
            height_window_recod = Canvas(parent, relief= GROOVE ,  background = 'gray100' , highlightthickness=0)
            window_model_run = Canvas(height_window_recod, relief= GROOVE ,  background = 'gray80' , highlightthickness=0)
            window_button = Canvas(height_window_recod, relief= GROOVE, background = 'gray90', highlightthickness = 1)
            my_check_model_detail = BooleanVar(value = False)
            check_bouton_model_detail = Checkbutton(window_button,  text ="Detail of Model ", variable = my_check_model_detail, font=('arial', 15), background = 'gray90')
            check_bouton_model_detail.place(relx = 0.05, rely = 0.15, relwidth= 0.6)
            color_model_label = Label(window_button, text="Point color", font=('arial', 15), bg= "gray90")
            color_model_label.place(relx=0.05, rely= 0.23, relwidth= 0.4)
            color_value_model = ttk.Combobox(window_button, values = colors_palettes)
            color_value_model.current(3)
            color_value_model.place(relx=0.45, rely= 0.23, relwidth= 0.5)
            label_size_model = Label(window_button, text="Point size  ", font=('arial', 15), bg= "gray90")
            label_size_model.place(relx=0.05, rely= 0.31, relwidth= 0.4)
            row_edit_scatter_default_model = StringVar()
            row_edit_scatter_default_model.set("15")
            point_size_scatter_model = Spinbox(window_button, from_ = 1 , to = 100000 , increment = 1, textvariable = row_edit_scatter_default_model)
            point_size_scatter_model.place(relx = 0.45, rely= 0.31, relwidth= 0.5) 
            color_model_linear = Label(window_button, text = "Linear color", font=('arial', 15), bg= "gray90")
            color_model_linear.place(relx=0.05, rely= 0.39, relwidth= 0.4)
            color_value_linear = ttk.Combobox(window_button, values = colors_palettes)
            color_value_linear.current(4)
            color_value_linear.place(relx=0.45, rely= 0.39, relwidth= 0.5)
            label_size_linear = Label(window_button, text = "Linear size", font=('arial', 15), bg= "gray90")
            label_size_linear.place(relx=0.05, rely= 0.47, relwidth= 0.4)
            row_edit_scatter_default_linear = StringVar()
            row_edit_scatter_default_linear.set("2")
            point_size_scatter_linear = Spinbox(window_button, from_ = 1 , to = 100000 , increment = 1, textvariable = row_edit_scatter_default_linear)
            point_size_scatter_linear.place(relx = 0.45, rely= 0.47, relwidth= 0.5) 
            my_check_model_train = BooleanVar(value = False)
            my_check_model_train_data = Checkbutton(window_button,  text ="Train data", variable = my_check_model_train, font=('arial', 15), background = 'gray90')
            my_check_model_train_data.place(relx = 0.05, rely = 0.55, relwidth= 0.45)
            label_train = Label(window_button, text="Train color  ", font=('arial', 15), bg= "gray90")
            label_train.place(relx=0.05, rely= 0.63, relwidth= 0.4)
            color_value_train = ttk.Combobox(window_button, values = colors_palettes)
            color_value_train.current(17)
            color_value_train.place(relx=0.45, rely= 0.63, relwidth= 0.5)
            label_size_train = Label(window_button, text="Train size    ", font=('arial', 15), bg= "gray90")
            label_size_train.place(relx=0.05, rely= 0.71, relwidth= 0.4)
            row_edit_scatter_default_train = StringVar()
            row_edit_scatter_default_train.set("15")
            row_edit_scatter_train = Spinbox(window_button, from_ = 1 , to = 100000 , increment = 1, textvariable = row_edit_scatter_default_train)
            row_edit_scatter_train.place(relx = 0.45, rely= 0.71, relwidth= 0.5) 
            check_bouton_model_detail.configure(command = lambda : check_model_run(window_model_run, color_value_model, row_edit_scatter_default_model, color_value_linear, point_size_scatter_linear, color_value_train, row_edit_scatter_default_linear, row_edit_scatter_default_train, my_check_model_detail, my_check_model_train))
            point_size_scatter_model.configure(command = lambda : check_model_run(window_model_run, color_value_model, row_edit_scatter_default_model, color_value_linear, point_size_scatter_linear, color_value_train, row_edit_scatter_default_linear, row_edit_scatter_default_train, my_check_model_detail, my_check_model_train))
            point_size_scatter_linear.configure(command = lambda : check_model_run(window_model_run, color_value_model, row_edit_scatter_default_model, color_value_linear, point_size_scatter_linear, color_value_train, row_edit_scatter_default_linear, row_edit_scatter_default_train, my_check_model_detail, my_check_model_train))
            my_check_model_train_data.configure(command = lambda : check_model_run(window_model_run, color_value_model, row_edit_scatter_default_model, color_value_linear, point_size_scatter_linear, color_value_train, row_edit_scatter_default_linear, row_edit_scatter_default_train, my_check_model_detail, my_check_model_train))
            row_edit_scatter_train.configure(command = lambda : check_model_run(window_model_run, color_value_model, row_edit_scatter_default_model, color_value_linear, point_size_scatter_linear, color_value_train, row_edit_scatter_default_linear, row_edit_scatter_default_train, my_check_model_detail, my_check_model_train))
            color_value_model.bind("<<ComboboxSelected>>", lambda event: check_model_run(window_model_run, color_value_model, row_edit_scatter_default_model, color_value_linear, point_size_scatter_linear, color_value_train, row_edit_scatter_default_linear, row_edit_scatter_default_train, my_check_model_detail, my_check_model_train))
            color_value_linear.bind("<<ComboboxSelected>>", lambda event: check_model_run(window_model_run, color_value_model, row_edit_scatter_default_model, color_value_linear, point_size_scatter_linear, color_value_train, row_edit_scatter_default_linear, row_edit_scatter_default_train, my_check_model_detail, my_check_model_train))
            color_value_train.bind("<<ComboboxSelected>>", lambda event: check_model_run(window_model_run, color_value_model, row_edit_scatter_default_model, color_value_linear, point_size_scatter_linear, color_value_train, row_edit_scatter_default_linear, row_edit_scatter_default_train, my_check_model_detail, my_check_model_train))
            save_graph_p = PhotoImage(file = my_path('figures/save_graph_p.png'))
            image_box_view = PhotoImage(file=my_path('figures/apercu_p.png'))
            bouton_plot_model_view = Button(window_button, command = lambda : check_model_run(window_model_run, color_value_model, row_edit_scatter_default_model, color_value_linear, point_size_scatter_linear, color_value_train, row_edit_scatter_default_linear, row_edit_scatter_default_train, my_check_model_detail, my_check_model_train), image = image_box_view , font=('arial', 14), text = "View plot",  compound = tk.TOP,  borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
            bouton_plot_model_view.place(relx= 0.27, rely= 0.03, width= 100, height= 60)
            bouton_save_model = Button(window_button, image = save_graph_p, command = lambda : my_save_plot_file(model_plot) , font=('arial', 14), text = "Save plot",  compound = tk.TOP,  borderwidth = 0, relief = FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
            bouton_save_model.place(relx= 0.27, rely= 0.85, width= 100, height= 60)
            window_model_run.place(relx=0, rely=0, relwidth= 0.80, relheight= 1)
            window_button.place(relx=0.80, rely = 0, relwidth= 0.2, relheight = 1)
            height_window_recod.place(relx=0, rely= 0, relwidth= 1, relheight = 1)




# ==========================================================================
# ==  22. Tableaux statistiques                                           ==
# ==========================================================================
def my_table_stat (data_user):

    if categorial_var_def.get() != False:
        if len(result_recod_stat) > 0 :
            for j in range(len(result_recod_stat)):
                my_vector = np.array(data_user[result_recod_stat[j]])
                if type(my_vector[0]) != str:
                    my_da = []
                    for i in range(len(my_vector)):
                        my_da.append(str(my_vector[i]))
                    data_user[result_recod_stat[j]] = my_da
    index_all = data_user.columns
    rows_listx = []
    for i in range(len(index_all)):
        var_xx = data_user[index_all[i]]
        if type(var_xx[0]) == float64:
            rows_new = {"Characteristics": str(index_all[i]), "Type": "double", "Unique": "", "Modality": "", "Missing": str(var_xx.isnull().sum().sum()), "N": len(data_user), "mean (SD) / (%)": str(round(mean(var_xx), 2)) + ' (' + str(round(std(var_xx), 2)) + ')'}
            rows_listx.append(rows_new)
        elif type(var_xx[0]) == int64:
            rows_new = {"Characteristics": str(index_all[i]), "Type": "integer", "Unique": "", "Modality": "", "Missing": str(var_xx.isnull().sum().sum()), "N": len(data_user), "mean (SD) / (%)": str(round(mean(var_xx), 1)) + ' (' + str(round(std(var_xx), 1)) + ')'}
            rows_listx.append(rows_new)
        else:
            if categorial_var_def.get() != False:
                if len(result_recod_stat) != 0 :
                    if index_all[i] in result_recod_stat:
                        int_unique1 = data_sort_stat[result_recod_stat.index(index_all[i])]
                        int_unique = []
                        for la in int_unique1:
                            int_unique.append(str(la))
                        rows_new = {"Characteristics": str(index_all[i]), "Type": "categorical", "Unique": len(int_unique), "Modality": "", "Missing": str(var_xx.isnull().sum().sum()), "N": len(data_user), "mean (SD) / (%)": ""}
                        rows_listx.append(rows_new)
                        for j in range(len(int_unique)):
                            my_data_modality = data_user[var_xx == int_unique[j]]
                            var_modality = 100 * (len(my_data_modality) / len(var_xx))
                            rows_new = {"Characteristics": "", "Type": "", "Unique": "", "Modality": int_unique[j], "Missing": "", "N": len(my_data_modality),  "mean (SD) / (%)": str(round(var_modality, 2)) + "%"}
                            rows_listx.append(rows_new)
                    else:
                        int_unique = unique(var_xx)
                        rows_new = {"Characteristics": str(index_all[i]), "Type": "character", "Unique": len(int_unique), "Modality": "", "Missing": str(var_xx.isnull().sum().sum()), "N": len(data_user), "mean (SD) / (%)": ""}
                        rows_listx.append(rows_new)
                        for j in range(len(int_unique)):
                            my_data_modality = data_user[var_xx == int_unique[j]]
                            var_modality = 100 * (len(my_data_modality)/ len(var_xx))
                            rows_new = {"Characteristics": "", "Type": "", "Unique": "", "Modality": int_unique[j], "Missing": "", "N": len(my_data_modality),  "mean (SD) / (%)": str(round(var_modality, 2)) + "%"}
                            rows_listx.append(rows_new)
                else:
                    int_unique = unique(var_xx)
                    rows_new = {"Characteristics": str(index_all[i]), "Type": "character", "Unique": len(int_unique), "Modality": "", "Missing": str(var_xx.isnull().sum().sum()), "N": len(data_user), "mean (SD) / (%)": ""}
                    rows_listx.append(rows_new)
                    for j in range(len(int_unique)):
                        my_data_modality = data_user[var_xx == int_unique[j]]
                        var_modality = 100 * (len(my_data_modality)/ len(var_xx))
                        rows_new = {"Characteristics": "", "Type": "", "Unique": "", "Modality": int_unique[j], "Missing": "", "N": len(my_data_modality),  "mean (SD) / (%)": str(round(var_modality, 2)) + "%"}
                        rows_listx.append(rows_new)
            else:
                int_unique = unique(var_xx)
                rows_new = {"Characteristics": str(index_all[i]), "Type": "character", "Unique": len(int_unique), "Modality": "", "Missing": str(var_xx.isnull().sum().sum()), "N": len(data_user), "mean (SD) / (%)": ""}
                rows_listx.append(rows_new)
                for j in range(len(int_unique)):
                    my_data_modality = data_user[var_xx == int_unique[j]]
                    var_modality = 100 * (len(my_data_modality)/ len(var_xx))
                    rows_new = {"Characteristics": "", "Type": "", "Unique": "", "Modality": int_unique[j], "Missing": "", "N": len(my_data_modality),  "mean (SD) / (%)": str(round(var_modality, 2)) + "%"}
                    rows_listx.append(rows_new)
    my_data_stat_all = pd.DataFrame(rows_listx)
    return my_data_stat_all


def my_table_stat_pred (data_user):

    if categorial_var_def_pred.get() != False:
        if len(result_recod_stat) > 0 :
            for j in range(len(result_recod_stat)):
                my_vector = np.array(data_user[result_recod_stat[j]])
                if type(my_vector[0]) != str:
                    my_da = []
                    for i in range(len(my_vector)):
                        my_da.append(str(my_vector[i]))
                    data_user[result_recod_stat[j]] = my_da
    index_all = data_user.columns
    rows_listx = []
    for i in range(len(index_all)):
        var_xx = data_user[index_all[i]]
        if type(var_xx[0]) == float64:
            rows_new = {"Characteristics": str(index_all[i]), "Type": "double", "Unique": "", "Modality": "", "Missing": str(var_xx.isnull().sum().sum()), "N": len(data_user), "mean (SD) / (%)": str(round(mean(var_xx), 2)) + ' (' + str(round(std(var_xx), 2)) + ')'}
            rows_listx.append(rows_new)
        elif type(var_xx[0]) == int64:
            rows_new = {"Characteristics": str(index_all[i]), "Type": "integer", "Unique": "", "Modality": "", "Missing": str(var_xx.isnull().sum().sum()), "N": len(data_user), "mean (SD) / (%)": str(round(mean(var_xx), 1)) + ' (' + str(round(std(var_xx), 1)) + ')'}
            rows_listx.append(rows_new)
        else:
            if categorial_var_def_pred.get() != False:
                if len(result_recod_stat) != 0 :
                    if index_all[i] in result_recod_stat:
                        int_unique1 = data_sort_stat[result_recod_stat.index(index_all[i])]
                        int_unique = []
                        for la in int_unique1:
                            int_unique.append(str(la))
                        rows_new = {"Characteristics": str(index_all[i]), "Type": "categorical", "Unique": len(int_unique), "Modality": "", "Missing": str(var_xx.isnull().sum().sum()), "N": len(data_user), "mean (SD) / (%)": ""}
                        rows_listx.append(rows_new)
                        for j in range(len(int_unique)):
                            my_data_modality = data_user[var_xx == int_unique[j]]
                            var_modality = 100 * (len(my_data_modality) / len(var_xx))
                            rows_new = {"Characteristics": "", "Type": "", "Unique": "", "Modality": int_unique[j], "Missing": "", "N": len(my_data_modality),  "mean (SD) / (%)": str(round(var_modality, 2)) + "%"}
                            rows_listx.append(rows_new)
                    else:
                        int_unique = unique(var_xx)
                        rows_new = {"Characteristics": str(index_all[i]), "Type": "character", "Unique": len(int_unique), "Modality": "", "Missing": str(var_xx.isnull().sum().sum()), "N": len(data_user), "mean (SD) / (%)": ""}
                        rows_listx.append(rows_new)
                        for j in range(len(int_unique)):
                            my_data_modality = data_user[var_xx == int_unique[j]]
                            var_modality = 100 * (len(my_data_modality)/ len(var_xx))
                            rows_new = {"Characteristics": "", "Type": "", "Unique": "", "Modality": int_unique[j], "Missing": "", "N": len(my_data_modality),  "mean (SD) / (%)": str(round(var_modality, 2)) + "%"}
                            rows_listx.append(rows_new)
                else:
                    int_unique = unique(var_xx)
                    rows_new = {"Characteristics": str(index_all[i]), "Type": "character", "Unique": len(int_unique), "Modality": "", "Missing": str(var_xx.isnull().sum().sum()), "N": len(data_user), "mean (SD) / (%)": ""}
                    rows_listx.append(rows_new)
                    for j in range(len(int_unique)):
                        my_data_modality = data_user[var_xx == int_unique[j]]
                        var_modality = 100 * (len(my_data_modality)/ len(var_xx))
                        rows_new = {"Characteristics": "", "Type": "", "Unique": "", "Modality": int_unique[j], "Missing": "", "N": len(my_data_modality),  "mean (SD) / (%)": str(round(var_modality, 2)) + "%"}
                        rows_listx.append(rows_new)
            else:
                int_unique = unique(var_xx)
                rows_new = {"Characteristics": str(index_all[i]), "Type": "character", "Unique": len(int_unique), "Modality": "", "Missing": str(var_xx.isnull().sum().sum()), "N": len(data_user), "mean (SD) / (%)": ""}
                rows_listx.append(rows_new)
                for j in range(len(int_unique)):
                    my_data_modality = data_user[var_xx == int_unique[j]]
                    var_modality = 100 * (len(my_data_modality)/ len(var_xx))
                    rows_new = {"Characteristics": "", "Type": "", "Unique": "", "Modality": int_unique[j], "Missing": "", "N": len(my_data_modality),  "mean (SD) / (%)": str(round(var_modality, 2)) + "%"}
                    rows_listx.append(rows_new)
    my_data_stat_all = pd.DataFrame(rows_listx)
    return my_data_stat_all


def my_TableOne (data_user, x_selected):

    if categorial_var_def.get() != False:
        if len(result_recod_stat) > 0 :
            for j in range(len(result_recod_stat)):
                my_vector = np.array(data_user[result_recod_stat[j]])
                if type(my_vector[0]) != str:
                    my_da = []
                    for i in range(len(my_vector)):
                        my_da.append(str(my_vector[i]))
                    data_user[result_recod_stat[j]] = my_da
    index_all = data_user.columns
    my_columns = [] 
    for col in index_all:
        pp = unique(data_user[col])
        if len(pp) > 1:
            my_columns.append(col)
    my_data_table = data_user[my_columns]
    if my_data_table.isnull().sum().sum() == 0:
        unique_modality = unique(my_data_table[x_selected])
        rows_listx = []
        modality_data = pd.DataFrame()
        for i in range(len(my_columns)):
            if my_columns[i] != x_selected:
                index_var = my_data_table[my_columns[i]]
                par_column = my_columns[i]
                test_uniq = unique(index_var)
                if type(index_var[0]) == float64:
                    Q1 = int(round(np.percentile(index_var, 25)))
                    M = int(round(median(index_var)))
                    Q3 = int(round(np.percentile(index_var, 75)))
                    rows_new = {"": par_column + ",   median (Q1 ; Q3)", "Modality": "", "Missing": str(index_var.isnull().sum().sum()), "Overall N = " + str(len(my_data_table)): str(M) + ' (' + str(Q1) + ' ; ' + str(Q3) + ')'}
                    rows_listx.append(rows_new)
                elif type(index_var[0]) == int64:
                    Q1 = int(round(np.percentile(index_var, 25)))
                    M = int(round(median(index_var)))
                    Q3 = int(round(np.percentile(index_var, 75)))
                    rows_new = {"": par_column + ",   median (Q1 ; Q3)", "Modality": "", "Missing": str(index_var.isnull().sum().sum()), "Overall N = " + str(len(my_data_table)): str(M) + ' (' + str(Q1) + ' ; ' + str(Q3) + ')'}
                    rows_listx.append(rows_new)
                else:
                    if categorial_var_def.get() != False:
                        if len(result_recod_stat) != 0 :
                                if index_all[i] in result_recod_stat:
                                    int_unique1 = data_sort_stat[result_recod_stat.index(index_all[i])]
                                    uni_mod = []
                                    for la in int_unique1:
                                        uni_mod.append(str(la))
                                    rows_new = {"": par_column + "   N (%)", "Modality": "", "Missing": str(index_var.isnull().sum().sum()), "Overall N = " + str(len(my_data_table)): ""}
                                    rows_listx.append(rows_new)
                                    for k in range(len(uni_mod)):
                                        md = my_data_table[my_data_table[par_column] == uni_mod[k]]
                                        var_modality = round(100 * (len(md)/ len(my_data_table)), 1)
                                        rows_new = {"": "", "Modality":  str(uni_mod[k]), "Missing": "", "Overall N = " + str(len(my_data_table)): str(len(md))  + ' (' + str(var_modality) + ')'}
                                        rows_listx.append(rows_new)
                                else:
                                    uni_mod = unique(index_var)
                                    rows_new = {"": par_column + "   N (%)", "Modality": "", "Missing": str(index_var.isnull().sum().sum()), "Overall N = " + str(len(my_data_table)): ""}
                                    rows_listx.append(rows_new)
                                    for k in range(len(uni_mod)):
                                        md = my_data_table[my_data_table[par_column] == uni_mod[k]]
                                        var_modality = round(100 * (len(md)/ len(my_data_table)), 1)
                                        rows_new = {"": "", "Modality":  str(uni_mod[k]), "Missing": "", "Overall N = " + str(len(my_data_table)): str(len(md))  + ' (' + str(var_modality) + ')'}
                                        rows_listx.append(rows_new)
                        else:
                            uni_mod = unique(index_var)
                            rows_new = {"": par_column + "   N (%)", "Modality": "", "Missing": str(index_var.isnull().sum().sum()), "Overall N = " + str(len(my_data_table)): ""}
                            rows_listx.append(rows_new)
                            for k in range(len(uni_mod)):
                                md = my_data_table[my_data_table[par_column] == uni_mod[k]]
                                var_modality = round(100 * (len(md)/ len(my_data_table)), 1)
                                rows_new = {"": "", "Modality":  str(uni_mod[k]), "Missing": "", "Overall N = " + str(len(my_data_table)): str(len(md))  + ' (' + str(var_modality) + ')'}
                                rows_listx.append(rows_new)
                    else:
                        uni_mod = unique(index_var)
                        rows_new = {"": par_column + "   N (%)", "Modality": "", "Missing": str(index_var.isnull().sum().sum()), "Overall N = " + str(len(my_data_table)): ""}
                        rows_listx.append(rows_new)
                        for k in range(len(uni_mod)):
                            md = my_data_table[my_data_table[par_column] == uni_mod[k]]
                            var_modality = round(100 * (len(md)/ len(my_data_table)), 1)
                            rows_new = {"": "", "Modality":  str(uni_mod[k]), "Missing": "", "Overall N = " + str(len(my_data_table)): str(len(md))  + ' (' + str(var_modality) + ')'}
                            rows_listx.append(rows_new)
        my_data_stat_all = pd.DataFrame(rows_listx)
        my_data_coef = pd.DataFrame()
        for i in range(len(unique_modality)):
            modality_data = my_data_table[my_data_table[x_selected] == unique_modality[i]]
            mm = []
            for j in range(len(my_columns)):
                x_x = np.array(modality_data[my_columns[j]])
                if my_columns[j] != x_selected:  
                    if type(x_x[0]) == float64:
                        Q1 = int(round(np.percentile(x_x, 25)))
                        M = int(round(median(x_x)))
                        Q3 = int(round(np.percentile(x_x, 75)))
                        mm.append(str(M) + ' (' + str(Q1) + ' ; ' + str(Q3) + ')')
                    elif type(x_x[0]) == int64:
                        Q1 = int(round(np.percentile(x_x, 25)))
                        M = int(round(median(x_x)))
                        Q3 = int(round(np.percentile(x_x, 75)))
                        mm.append(str(M) + ' (' + str(Q1) + ' ; ' + str(Q3) + ')')
                    else:
                            if categorial_var_def.get() != False:
                                    if len(result_recod_stat) != 0 :
                                            if my_columns[j] in result_recod_stat:
                                                int_unique1 = data_sort_stat[result_recod_stat.index(my_columns[j])]
                                                test_uniq_mod = []
                                                for la in int_unique1:
                                                    test_uniq_mod.append(str(la))
                                                for k in range(len(test_uniq_mod) + 1):
                                                     if k == 0:
                                                        mm.append("")
                                                     else:
                                                        modality_data_fin = modality_data[modality_data[my_columns[j]] == test_uniq_mod[k-1]]
                                                        dd_mod = round(100*(len(modality_data_fin)/ len(modality_data)), 1)
                                                        mm.append(str(len(modality_data_fin)) + ' (' + str(dd_mod) + ')')
                                            else:
                                                  test_uniq_mod = unique(my_data_table[my_columns[j]])
                                                  for k in range(len(test_uniq_mod) + 1):
                                                      if k == 0:
                                                         mm.append("")
                                                      else:
                                                         modality_data_fin = modality_data[modality_data[my_columns[j]] == test_uniq_mod[k-1]]
                                                         dd_mod = round(100*(len(modality_data_fin)/ len(modality_data)), 1)
                                                         mm.append(str(len(modality_data_fin)) + ' (' + str(dd_mod) + ')')
                                    else:
                                          test_uniq_mod = unique(my_data_table[my_columns[j]])
                                          for k in range(len(test_uniq_mod) + 1):
                                              if k == 0:
                                                 mm.append("")
                                              else:
                                                 modality_data_fin = modality_data[modality_data[my_columns[j]] == test_uniq_mod[k-1]]
                                                 dd_mod = round(100*(len(modality_data_fin)/ len(modality_data)), 1)
                                                 mm.append(str(len(modality_data_fin)) + ' (' + str(dd_mod) + ')')
                            else:
                                test_uniq_mod = unique(my_data_table[my_columns[j]])
                                for k in range(len(test_uniq_mod) + 1):
                                    if k == 0:
                                       mm.append("")
                                    else:
                                       modality_data_fin = modality_data[modality_data[my_columns[j]] == test_uniq_mod[k-1]]
                                       dd_mod = round(100*(len(modality_data_fin)/ len(modality_data)), 1)
                                       mm.append(str(len(modality_data_fin)) + ' (' + str(dd_mod) + ')')
            my_data_coef[str(unique_modality[i]) + " N = " + str(len(modality_data))] =  mm
        p_val_stat = []  
        for i in range(len(my_columns)):
            test_index = my_data_table[my_columns[i]]
            if my_columns[i] != x_selected:  
                if type(test_index[0]) != str:
                    dstat1 = stats.kruskal(*[group[my_columns[i]].values for name, group in my_data_table.groupby(x_selected)]) 
                    if dstat1[1] < 0.001:
                        p_val_stat.append("<0.001") 
                    else:
                        p_val_stat.append(round(dstat1[1], 3)) 
                else:
                    # variable qualitative : test du Chi2 d'independance (Kruskal-Wallis ne s'applique
                    # qu'a des valeurs numeriques ; sur du texte, scipy leve une erreur)
                    try:
                        dstat1 = stats.chi2_contingency(pd.crosstab(my_data_table[my_columns[i]], my_data_table[x_selected]))
                    except ValueError:
                        dstat1 = (None, float("nan"))
                    test_mod = unique(my_data_table[my_columns[i]])
                    for j in range(len(test_mod) + 1):
                        if j == 0:
                            if not np.isfinite(dstat1[1]):
                                p_val_stat.append("")
                            elif dstat1[1] < 0.001:
                                p_val_stat.append("<0.001")
                            else:
                                p_val_stat.append(round(dstat1[1], 3))
                        else:
                            p_val_stat.append("")
        data_pvalue = pd.DataFrame({"p-value": p_val_stat})
        ddf = pd.concat([my_data_stat_all, my_data_coef, data_pvalue], axis = 1)
        return ddf


def _lignes_groupe_overview (tableau):
    """ Lignes de titre des variables qualitatives du tableau "Overview stat"
    (suivies de leurs modalites) : mises en evidence a l'ecran et a l'export. """
    if "Characteristics" not in tableau.columns or "Type" not in tableau.columns:
        return []
    return [i for i, (nom, type_colonne) in enumerate(zip(tableau["Characteristics"].astype(str), tableau["Type"].astype(str)))
            if nom.strip() != "" and type_colonne not in ("double", "integer")]


def stat_data_all (parent, my_data_anal):
    """ "Overview stat" : tableau recapitulatif de toutes les colonnes des donnees
    (type, valeurs manquantes, N, moyenne (ecart-type) ou effectif (%) par modalite). """

    global data_stat_all_exp, graph_stat_all_widget
    # index 0..n-1 : my_table_stat lit la 1re valeur de chaque colonne par var[0]
    # (apres un filtrage, l'index ne commence plus forcement a 0 -> KeyError)
    my_data = my_data_anal.copy().reset_index(drop = True)

    ## Detruit le panneau cree par un precedent clic sur "Overview stat" au lieu
    ## de l'empiler par-dessus (evite d'accumuler les widgets en memoire)
    if graph_stat_all_widget is not None and graph_stat_all_widget.winfo_exists():
        graph_stat_all_widget.destroy()

    graph_stat_all = tk.Frame(parent, background = 'gray90')
    graph_stat_all_widget = graph_stat_all
    graph_stat_all.place(relx = 0, rely = 0, relwidth = 0.8, relheight = 1)     # les boutons sont a droite (relx = 0.85)

    if len(my_data.columns) == 0 or len(my_data) == 0:
        data_stat_all_exp = None
        _message_stat(graph_stat_all, "The table is empty.")
        return
    try:
        data_stat_all = my_table_stat(my_data)
    except Exception as erreur:
        data_stat_all_exp = None
        _message_stat(graph_stat_all, "The statistics could not be computed:\n" + str(erreur))
        return
    data_stat_all_exp = data_stat_all

    titre = "Overview of the imported table  (" + str(len(my_data)) + " rows, " + str(len(my_data.columns)) + " columns)"
    cadre = _afficher_tableau_stat(graph_stat_all, data_stat_all, titre = titre, lignes_groupe = _lignes_groupe_overview(data_stat_all))
    cadre.place(relx = 0.005, rely = 0.02, relwidth = 0.99, relheight = 0.96)


def exporter_stat_overview ():
    """ "Export stat" du panneau Edit data : exporte le tableau de "Overview stat"
    en Excel, Word ou CSV (memes mises en forme que les autres exports de statistiques). """

    if data_stat_all_exp is None:
        messagebox.showwarning("Export stat", "Please display the statistics first (button 'Overview stat').")
        return
    tableaux = [("Overview of the data", data_stat_all_exp, _lignes_groupe_overview(data_stat_all_exp))]
    chemin = filedialog.asksaveasfilename(filetypes = [("Excel file", "*.xlsx"), ("Word file", "*.docx"), ("csv file", "*.csv")], defaultextension = ".xlsx")
    if not chemin:
        return
    extension = os.path.splitext(chemin)[1].lower()
    try:
        if extension == ".docx":
            _exporter_tableaux_stat_word(chemin, tableaux)
        elif extension == ".csv":
            _exporter_tableaux_stat_csv(chemin, tableaux)
        else:
            if extension != ".xlsx":
                chemin += ".xlsx"
            _exporter_tableaux_stat_excel(chemin, tableaux)
    except Exception as erreur:
        messagebox.showerror("Export stat", "The table could not be exported:\n" + str(erreur))
        return
    messagebox.showinfo("Export stat", "Table exported to:\n" + chemin)


def stat_data_pred (parent, my_data_anal):

    global data_stat_all_exp_pred, graph_stat_all_pred_widget
    my_data = my_data_anal.copy()
    label_data_pred2.configure(text = "Descriptive analysis of parameters                        ")
    change_modality_pred.configure(text = "")
    rapport_missing_value_all_pred.configure(text = "")

    ## Detruit le panneau cree par un precedent clic sur "Overview stat" au lieu
    ## de l'empiler par-dessus (evite d'accumuler Canvas + Table en memoire
    ## a chaque clic, ce qui ralentissait l'appli apres plusieurs utilisations)
    if graph_stat_all_pred_widget is not None and graph_stat_all_pred_widget.winfo_exists():
        graph_stat_all_pred_widget.destroy()

    graph_stat_all_pred = Canvas(parent, background = 'gray90' , highlightthickness = 0)
    graph_stat_all_pred_widget = graph_stat_all_pred
    alerte_stat_all = tk.Label(graph_stat_all_pred, text= "Overview of the imported table        ", font=('arial', 15), background = "gray90")
    alerte_stat_all.place(relx = 0, rely = 0.005, relwidth = 0.3)
    data_stat_all_pred = my_table_stat_pred(my_data)
    data_stat_all_exp_pred = data_stat_all_pred
    windows_stat_all_pred = Table(graph_stat_all_pred, dataframe = data_stat_all_pred)
    windows_stat_all_pred.show()
    graph_stat_all_pred.place(relx = 0, rely = 0.7, relwidth = 1, relheight = 0.3)


def _formater_valeur_stat (valeur):
    """ Texte d'une cellule : entiers sans decimales, reels a 3 decimales
    maximum (zeros inutiles retires), valeurs manquantes -> cellule vide. """
    if valeur is None:
        return ""
    if isinstance(valeur, (bool, np.bool_)):
        return str(valeur)
    if isinstance(valeur, (int, np.integer)):
        return str(int(valeur))
    if isinstance(valeur, (float, np.floating)):
        if not np.isfinite(valeur):
            return ""
        if float(valeur).is_integer():
            return str(int(valeur))
        return ("%.3f" % valeur).rstrip("0").rstrip(".")
    return str(valeur)


def _style_tableaux_stat ():
    """ Style commun des tableaux de statistiques (cree une seule fois). """
    style = ttk.Style()
    style.configure("Stat.Treeview", font = ('arial', 12), rowheight = 26, background = "white", fieldbackground = "white", foreground = "gray15")
    style.configure("Stat.Treeview.Heading", font = ('arial', 12, 'bold'), foreground = "gray15")
    style.map("Stat.Treeview", background = [("selected", "#cfe2f3")], foreground = [("selected", "gray10")])
    return style


def _afficher_tableau_stat (parent, tableau, titre = None, lignes_groupe = None):
    """ Affiche le DataFrame "tableau" dans un cadre (retourne) avec un titre,
    une ligne d'en-tete, des lignes alternees et des barres de defilement
    horizontale et verticale. La 1re colonne (nom de variable) est alignee a
    gauche, les autres sont centrees. "lignes_groupe" : indices des lignes a
    mettre en evidence (lignes de titre d'une variable). """

    _style_tableaux_stat()
    lignes_groupe = set(lignes_groupe or [])
    cadre = tk.Frame(parent, background = "white")
    cadre.grid_columnconfigure(0, weight = 1)
    ligne_arbre = 0
    if titre:
        Label(cadre, text = titre, font = ('arial', 16, 'bold'), fg = "gray20", bg = "white", anchor = "w").grid(row = 0, column = 0, columnspan = 2, sticky = "ew", padx = 10, pady = (8, 4))
        ligne_arbre = 1
    cadre.grid_rowconfigure(ligne_arbre, weight = 1)

    identifiants = ["c%d" % i for i in range(len(tableau.columns))]      # identifiants surs, meme pour un nom de colonne vide
    arbre = ttk.Treeview(cadre, columns = identifiants, show = "headings", style = "Stat.Treeview", selectmode = "browse")
    defilement_y = ttk.Scrollbar(cadre, orient = "vertical", command = arbre.yview)
    defilement_x = ttk.Scrollbar(cadre, orient = "horizontal", command = arbre.xview)
    arbre.configure(yscrollcommand = defilement_y.set, xscrollcommand = defilement_x.set)
    arbre.grid(row = ligne_arbre, column = 0, sticky = "nsew", padx = (10, 0))
    defilement_y.grid(row = ligne_arbre, column = 1, sticky = "ns")
    defilement_x.grid(row = ligne_arbre + 1, column = 0, sticky = "ew", padx = (10, 0), pady = (0, 6))

    valeurs = [[_formater_valeur_stat(v) for v in ligne] for ligne in tableau.itertuples(index = False, name = None)]
    police_cellules = _tkfont_langue.Font(family = "arial", size = 12)
    police_entete = _tkfont_langue.Font(family = "arial", size = 12, weight = "bold")
    for numero, (identifiant, nom_colonne) in enumerate(zip(identifiants, tableau.columns)):
        entete = str(nom_colonne) if str(nom_colonne).strip() != "" else "Variable"
        # largeur ajustee au contenu (echantillon des 300 premieres lignes), bornee
        # (_builtins_langue.max/min : `from numpy import *` remplace max/min par numpy.max/min)
        largeur = police_entete.measure(entete) + 30
        for ligne in valeurs[:300]:
            largeur = _builtins_langue.max(largeur, police_cellules.measure(ligne[numero]) + 30)
        largeur = _builtins_langue.min(_builtins_langue.max(largeur, 80), 420)
        arbre.heading(identifiant, text = entete, anchor = "w" if numero == 0 else "center")
        arbre.column(identifiant, width = largeur, minwidth = 60, stretch = False, anchor = "w" if numero == 0 else "center")

    arbre.tag_configure("pair", background = "white")
    arbre.tag_configure("impair", background = "#f3f5f7")
    arbre.tag_configure("groupe", background = "#e2ebf4", font = ('arial', 12, 'bold'))
    for numero, ligne in enumerate(valeurs):
        etiquette = "groupe" if numero in lignes_groupe else ("impair" if numero % 2 else "pair")
        arbre.insert("", "end", values = ligne, tags = (etiquette,))
    return cadre


def _message_stat (parent, texte):
    """ Message d'information (en rouge) au centre du panneau des statistiques. """
    Label(parent, text = traduire(texte), font = ('arial', 18), fg = "red", background = "gray80", justify = CENTER).place(relx = 0.1, rely = 0.35, relwidth = 0.8)


def check_statistis (event, parent, data, check_var, list_var, check_value, check_list):

    global my_data_stat_save, method_stat_save, stat_tableaux_widget, stat_tableaux_export
    stat_tableaux_export = []       # rempli ci-dessous avec les tableaux effectivement affiches
    # un seul panneau de tableaux a la fois (avant : un nouveau Canvas empile a chaque clic)
    if stat_tableaux_widget is not None and stat_tableaux_widget.winfo_exists():
        stat_tableaux_widget.destroy()
    x_index = check_var.current()
    x_selected = list_var[x_index]
    methode = check_list[check_value.current()]
    method_stat_save = methode
    graph_window_stat = tk.Frame(parent, background = "gray80")
    stat_tableaux_widget = graph_window_stat
    graph_window_stat.place(relx = 0, rely = 0, relwidth = 1, relheight = 1)

    if methode == "All columns":
        # tableau de type "TableOne" : chaque variable selon les modalites de la colonne "Groupby"
        if x_index == 0:
            _message_stat(graph_window_stat, "Select a feature!")
            return
        if data.isnull().sum().sum() != 0:
            _message_stat(graph_window_stat, "Please filter the data! \n \n The loading table contains missing values. \n \n Thanks!")
            return
        dd = my_TableOne(data, x_selected)
        my_data_stat_save = dd
        premiere_colonne = dd.iloc[:, 0].astype(str) if len(dd.columns) != 0 else pd.Series(dtype = str)
        lignes_groupe = [i for i, texte in enumerate(premiere_colonne) if texte.strip() != ""]
        stat_tableaux_export = [("Analysis variable - " + str(x_selected), dd, lignes_groupe)]
        cadre = _afficher_tableau_stat(graph_window_stat, dd, titre = "Analysis variable: " + str(x_selected), lignes_groupe = lignes_groupe)
        cadre.place(relx = 0.01, rely = 0.01, relwidth = 0.98, relheight = 0.98)
        return

    # "Continuous columns" : statistiques descriptives + correlations de Pearson
    donnees_quanti = data.select_dtypes(include = "number")
    if donnees_quanti.shape[1] == 0:
        _message_stat(graph_window_stat, "No quantitative column in the data.")
        return
    data_stat = donnees_quanti.describe().T
    data_stat = data_stat.rename(columns = {"count": "N", "mean": "Mean", "std": "Std", "min": "Min", "25%": "Q1", "50%": "Median", "75%": "Q3", "max": "Max"})
    data_stat.insert(1, "Missing", donnees_quanti.isnull().sum())
    data_stat.index.name = "Variable"
    data_stat = data_stat.reset_index()
    data_stat["N"] = data_stat["N"].astype(int)
    my_data_stat_save = data_stat.round(4)

    data_corr = donnees_quanti.corr().round(3)
    data_corr.index.name = "Variable"
    data_corr = data_corr.reset_index()

    stat_tableaux_export = [("Descriptive statistics", data_stat, []), ("Pearson correlations", data_corr, [])]
    cadre_stat = _afficher_tableau_stat(graph_window_stat, data_stat, titre = "Descriptive statistics for quantitative variables")
    cadre_stat.place(relx = 0.01, rely = 0.01, relwidth = 0.98, relheight = 0.485)
    cadre_corr = _afficher_tableau_stat(graph_window_stat, data_corr, titre = "Pearson linear correlations between quantitative variables")
    cadre_corr.place(relx = 0.01, rely = 0.505, relwidth = 0.98, relheight = 0.485)


def check_mean_min_max (event, parent, data, param_dim, param_var, method_var, list_var_ana):

    global my_barplot, graph_window_representation_widget, dernier_dimension_min_max, dernier_methode_min_max, dernier_features_min_max
    if graph_window_representation_widget is not None and graph_window_representation_widget.winfo_exists():
        graph_window_representation_widget.destroy()
    dim_index = param_dim.current()
    if dim_index != 0:
        list_method = ["min", "mean", "max", "median"]
        method_index = method_var.current()
        method_selected = list_method[method_index]
        dim_selected = list_var_ana[dim_index]
        dx = unique(data[dim_selected])
        result = [param_var.get(i) for i in param_var.curselection()]
        dernier_dimension_min_max = dim_selected
        dernier_methode_min_max = method_selected
        dernier_features_min_max = result
        if len(result) !=0:
            data_stats = data[result].describe()
            d0 = data[data[dim_selected] == dx[0]]
            result_calcul0 = []
            for k in range(len(result)):
                if method_selected == "min":
                   result_calcul0.append(round(min(d0[result[k]]), 2))
                elif method_selected == "mean":
                     result_calcul0.append(round(mean(d0[result[k]]), 2))
                elif method_selected == "max":
                     result_calcul0.append(round(max(d0[result[k]]), 2))
                else:
                    result_calcul0.append(round(median(d0[result[k]]), 2))
            data_frame = pd.DataFrame({str(dx[0]): result_calcul0})    
            rows_listx = []
            for i in range(1, len(dx)):
                d = data[data[dim_selected] == dx[i]]
                result_calcul = []
                for j in range(len(result)):
                    if method_selected == "min":
                       result_calcul.append(round(min(d[result[j]]), 2))
                    elif method_selected == "mean":
                        result_calcul.append(round(mean(d[result[j]]), 2))
                    elif method_selected == "max":
                        result_calcul.append(round(max(d[result[j]]), 2))
                    else:
                       result_calcul.append(round(median(d[result[j]]), 2))
                data_frame_cur = pd.DataFrame({str(dx[i]): result_calcul})  
                data_frame = pd.concat([data_frame, data_frame_cur], axis = 1)
            data_frame.index =  result
            data_frame = data_frame.T
            barWidth = 0.25
            multiplier = 0
            figure = Figure(figsize=(12, 8), dpi = 100)
            _enregistrer_figure_courante("Barplot_min_max", figure)
            figure_canvas = FigureCanvasTkAgg(figure, parent)
            graph_window_representation_widget = figure_canvas.get_tk_widget()
            x = np.arange(len(dx)) 
            plt = figure.add_subplot()
            my_tile = method_selected + " parameter barplot"
            if min_max_title_value is not None:
                if min_max_title_value != "Enter the title":
                    my_tile = min_max_title_value
            my_x_label = dim_selected
            if min_max_x_value is not None:
                if min_max_x_value != "Enter the x label":
                    my_x_label = min_max_x_value
            my_y_label = "Parameter value (%)"
            if min_max_y_value is not None:
                if min_max_y_value != "Enter the y label":
                    my_y_label = min_max_y_value
            xx = []
            for i in range(len(result)):
                xx.append(max(data_frame[result[i]]))
            for  i in range(len(result)):
                 offset = barWidth * multiplier
                 plt.bar(x + offset, data_frame[result[i]], width = barWidth, label = result[i])
                ##plt.bar_label(rects, padding = 3)
                 multiplier += 1
            plt.set_xlabel(my_x_label, fontsize = 13, labelpad = 7)
            plt.set_ylabel(my_y_label, fontsize= 13, labelpad = 7)
            plt.set_xticks(x + 0.25*len(result)*barWidth, dx)
            plt.legend(loc='upper left', ncols = 3, fontsize= 8)
            plt.set_ylim(0, max(xx) + 20)
            plt.set_title(my_tile, fontsize= 13)  
            plt.grid(True)
            figure_canvas.get_tk_widget().place(relx=0, rely= 0, relwidth= 1, relheight= 1)
            _appliquer_zoom_representation(plt, "minmax")
            _activer_glissement_representation(plt, figure_canvas, "minmax")

            def _sur_changement_zoom_representation_minmax (valeur_curseur):
                zoom_representation_par_type["minmax"] = _curseur_vers_zoom_representation(valeur_curseur)
                check_mean_min_max(event, parent, data, param_dim, param_var, method_var, list_var_ana)

            scale.configure(command = _sur_changement_zoom_representation_minmax)
            scale.set(_zoom_vers_curseur_representation(zoom_representation_par_type.get("minmax", 1.0)))
            my_barplot = figure


def recup_my_title(parent, check_title_value, check_zone_entry_x, check_zone_entry_y):

    global cor_title_value, cor_x_value, cor_y_value
    cor_title_value = check_title_value.get()
    cor_x_value = check_zone_entry_x.get()
    cor_y_value = check_zone_entry_y.get()
    parent.destroy()


def check_linear_correlation (event, parent, my_data, target_check, feature_check, list_var):

    global my_plot, graph_window_representation_widget, dernier_target_correlation, dernier_features_correlation
    if graph_window_representation_widget is not None and graph_window_representation_widget.winfo_exists():
        graph_window_representation_widget.destroy()
    target_index = target_check.current()
    if target_index != 0 :
        tagrget_selected = list_var[target_index]
        result = [feature_check.get(i) for i in feature_check.curselection()]
        dernier_target_correlation = tagrget_selected
        dernier_features_correlation = result
        pp = my_data.index
        var_list_corr = []
        for i in range(len(result)):
            d = my_data[[tagrget_selected, result[i]]]
            dx =  d.corr()
            var_list_corr.append(round(dx[tagrget_selected].iloc[1], 3))
        colors_var = []
        for i in range(len(var_list_corr)):
            if var_list_corr[i] > 0:
                colors_var.append("olivedrab")
            else:
                colors_var.append("red")
        figure = Figure(figsize=(12, 8),  dpi = 70)
        _enregistrer_figure_courante("Correlation_lineaire", figure)
        figure_bar = FigureCanvasTkAgg(figure, parent)
        graph_window_representation_widget = figure_bar.get_tk_widget()
        my_tile = "Linear correlation coefficient of " + tagrget_selected
        if cor_title_value is not None:
            if cor_title_value != "Enter the title":
                my_tile = cor_title_value
        my_x_label = ""
        if cor_x_value is not None:
            if cor_x_value != "Enter the x label":
                my_x_label = cor_x_value
        my_y_label = "Correlation coefficient"
        if cor_y_value is not None:
            if cor_y_value != "Enter the y label":
                my_y_label = cor_y_value
        x = np.arange(len(var_list_corr)) 
        plt = figure.add_subplot()
        bar_plot = plt.bar(x , var_list_corr, width = 0.5, color = colors_var)
        plt.bar_label(bar_plot, padding = 1, fontsize = 13)
        plt.axhline(y=0, color = 'blue', linestyle='--', lw = 2)
        plt.set_ylabel(my_y_label, fontsize = 13, labelpad = 7)
        plt.set_xlabel(my_x_label, fontsize = 13, labelpad = 7)
        plt.set_xticks(x, result, fontsize = 12)
        plt.set_ylim(-1, 1)
        plt.grid(True)
        plt.set_title(my_tile, fontsize= 18)  
        figure_bar.get_tk_widget().place(relx=0, rely= 0, relwidth= 1, relheight= 1)
        _appliquer_zoom_representation(plt, "correlation")
        _activer_glissement_representation(plt, figure_bar, "correlation")

        def _sur_changement_zoom_representation_corr (valeur_curseur):
            zoom_representation_par_type["correlation"] = _curseur_vers_zoom_representation(valeur_curseur)
            check_linear_correlation(event, parent, my_data, target_check, feature_check, list_var)

        scale.configure(command = _sur_changement_zoom_representation_corr)
        scale.set(_zoom_vers_curseur_representation(zoom_representation_par_type.get("correlation", 1.0)))
        my_plot = figure


def my_title_graph(parent, check_title):

    if check_title.get() != False:
        windows_graphe_title = Canvas(parent, background = 'cadetblue' , highlightthickness = 0)
        windows_graphe_title.place(relx = 0.1, rely = 0.1, relwidth = 0.8, relheight = 0.8)
        label_title = Label(windows_graphe_title, text="Plot title        ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_title.place(relx=0.07, rely= 0.04, relwidth= 0.3)
        zone_entry_title_default = StringVar()
        zone_entry_title_default.set("Enter the title")
        zone_entry_title = Entry(windows_graphe_title, textvariable = zone_entry_title_default, background = 'gray90', justify = CENTER )
        zone_entry_title.place(relx=0.1, rely = 0.1, relwidth = 0.8, relheight= 0.15)
        label_x = Label(windows_graphe_title, text="X - axis        ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_x.place(relx=0.07, rely= 0.34, relwidth= 0.3)
        zone_entry_x_default = StringVar()
        zone_entry_x_default.set("Enter the x label")
        zone_entry_x = Entry(windows_graphe_title, textvariable = zone_entry_x_default, background = 'gray90', justify = CENTER )
        zone_entry_x.place(relx=0.1, rely = 0.4, relwidth = 0.8, relheight= 0.15)
        label_y = Label(windows_graphe_title, text="Y - axis        ", font=('arial', 15), fg = "gray30", background= "cadetblue")
        label_y.place(relx=0.07, rely= 0.64, relwidth= 0.3)
        zone_entry_y_default = StringVar()
        zone_entry_y_default.set("Enter the y label")
        zone_entry_y = Entry(windows_graphe_title, textvariable = zone_entry_y_default, background = 'gray90', justify = CENTER )
        zone_entry_y.place(relx=0.1, rely = 0.7, relwidth = 0.8, relheight= 0.15)
        close_window_title = Button(windows_graphe_title, text = "Save", command = lambda : recup_my_title(windows_graphe_title, zone_entry_title_default, zone_entry_x, zone_entry_y), font=('arial', 14), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')   
        close_window_title.place(relx= 0.4, rely= 0.93, width= 50, height= 20)


def linear_correlation (event, parent, my_data, var_par) :

    global linear_correlation_widgets
    _detruire_widgets_differe(linear_correlation_widgets)
    windows_graphe = Canvas(parent, background = 'gray80' , highlightthickness = 0)
    windows_graphe.place(relx = 0, rely = 0, relwidth = 0.6, relheight = 1)
    windows_param = Canvas(parent, background = 'gray90' , highlightthickness = 0)
    windows_param.place(relx = 0.6, rely = 0, relwidth = 0.4, relheight = 1)
    linear_correlation_widgets = (windows_graphe, windows_param)
    label_Target = Label(windows_param, text="Target", font=('arial', 15), fg = "gray30", background= "gray90")
    label_Target.place(relx=0.05, rely= 0.1, relwidth= 0.4)
    label_Feature = Label(windows_param, text="Feature", font=('arial', 15), fg = "gray30", bg= "gray90")
    label_Feature.place(relx=0.05, rely= 0.25, relwidth= 0.4)
    Feature_list_stat = my_data.columns
    Target_value = ttk.Combobox(windows_param, values = var_par, width= 10)
    Target_value.current(_index_restaure(var_par, dernier_target_correlation))
    Target_value.place(relx=0.47, rely= 0.1, relwidth= 0.5)
    Feature_value = tk.Listbox(windows_param, fg = "gray30", font = ('arial', 12), relief = GROOVE , selectmode = tk.MULTIPLE, exportselection = False)
    Feature_value.place(relx= 0.42, rely = 0.25, relwidth= 0.55, relheight=0.3)
    for indice_valeur, value in enumerate(var_par):
        Feature_value.insert(END, value)
        ## Recoche les colonnes "feature" precedemment selectionnees, si elles
        ## existent toujours
        if dernier_features_correlation is not None and value in dernier_features_correlation:
            Feature_value.select_set(indice_valeur)
    def memoriser_features_correlation (event = None):
        ## Memorise la selection des features des qu'elle change (et plus seulement au clic sur
        ## "View plot") : elle est retrouvee en revenant sur "Plot corr", meme sans graphique trace.
        global dernier_features_correlation
        dernier_features_correlation = [Feature_value.get(i) for i in Feature_value.curselection()]
    Feature_value.bind("<<ListboxSelect>>", memoriser_features_correlation)
    title_var_def  = BooleanVar(value = False)
    title_var_info = Checkbutton(windows_param,  text ="Define chart title", variable = title_var_def, command = lambda : my_title_graph(windows_param, title_var_def) , fg = "gray30", font=('arial', 15), background = 'gray90')
    title_var_info.place(relx = 0.38, rely= 0.62, relwidth= 0.37)
    ##title_var_info = Checkbutton(windows_param,  text ="define chart title", variable = Feature_var_def, command = lambda : my_feature_function(boit_param, Feature_list, Feature_var_def) , fg = "gray30", font=('arial', 15), background = 'gray100')
    image_box_view = PhotoImage(file=my_path('figures/apercu_p.png'))
    bouton_corr_view = Button(windows_param, image = image_box_view , font=('arial', 14), text = "View plot",  compound = tk.TOP,  borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
    bouton_corr_view.place(relx= 0.17, rely= 0.75, width= 100, height= 60)
    bouton_corr_view.bind("<ButtonRelease>", lambda event: check_linear_correlation(event, windows_graphe, my_data, Target_value, Feature_value, var_par)) 
    save_graph_p = PhotoImage(file = my_path('figures/save_graph_p.png'))
    bouton_save_corr = Button(windows_param, image = save_graph_p , font=('arial', 14), text = "Save plot",  compound = tk.TOP,  borderwidth = 0, relief = FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
    bouton_save_corr.place(relx= 0.73, rely= 0.75, width= 100, height= 60)
    bouton_save_corr.bind("<ButtonRelease>", lambda event: my_save_plot_file(my_plot)) 
    ## Si une cible et au moins une feature ont deja ete choisies
    ## precedemment, on redessine tout de suite le graphique
    if Target_value.current() != 0 and len(Feature_value.curselection()) != 0:
        check_linear_correlation(event, windows_graphe, my_data, Target_value, Feature_value, var_par)


def mean_min_max (event, parent, my_data, var_par):

    global mean_min_max_widgets
    _detruire_widgets_differe(mean_min_max_widgets)
    var_par_list1 = my_data.columns
    var_par_list = []
    for i in range(len(var_par_list1)):
        var_par_list.append(var_par_list1[i])
    windows_graphe = Canvas(parent, background = 'gray80' , highlightthickness = 0)
    windows_graphe.place(relx = 0, rely = 0, relwidth = 0.6, relheight = 1)
    windows_param = Canvas(parent, background = 'gray90' , highlightthickness = 0)
    windows_param.place(relx = 0.6, rely = 0, relwidth = 0.4, relheight = 1)
    mean_min_max_widgets = (windows_graphe, windows_param)
    label_Feature = Label(windows_param, text="Parameter", font=('arial', 15), fg = "gray30", bg= "gray90")
    label_Feature.place(relx=0.05, rely= 0.35, relwidth= 0.4)
    label_method_stat = Label(windows_param, text="Method      ", font=('arial', 15), fg = "gray30", background= "gray90")
    label_method_stat.place(relx=0.05, rely= 0.23, relwidth= 0.4)
    label_dimension = Label(windows_param, text="Dimension", font=('arial', 15), fg = "gray30", bg= "gray90")
    label_dimension.place(relx=0.05, rely= 0.1, relwidth= 0.4)
    method_list_stat = ["min", "mean", "max", "median"]
    Feature_list_stat = my_data.columns
    Method_column = ttk.Combobox(windows_param, values = method_list_stat, width= 10)
    Method_column.current(_index_restaure(method_list_stat, dernier_methode_min_max) if dernier_methode_min_max in method_list_stat else 1)
    Method_column.place(relx=0.42, rely= 0.23, relwidth= 0.55)
    dimension_column = ttk.Combobox(windows_param, values = var_par, width= 10)
    dimension_column.current(_index_restaure(var_par, dernier_dimension_min_max))
    dimension_column.place(relx=0.42, rely= 0.1, relwidth= 0.55)
    Feature_column_stat = tk.Listbox(windows_param, fg = "gray30", font = ('arial', 12), relief = GROOVE , selectmode = tk.MULTIPLE, exportselection = False)
    Feature_column_stat.place(relx= 0.42, rely = 0.35, relwidth= 0.55, relheight=0.3)
    for indice_valeur, value in enumerate(var_par_list):
        Feature_column_stat.insert(END, value)
        if dernier_features_min_max is not None and value in dernier_features_min_max:
            Feature_column_stat.select_set(indice_valeur)
    title_mean_min_max_def  = BooleanVar(value = False)
   ## title_mean_min_max_info = Checkbutton(features,  text ="Define chart title", variable = title_mean_min_max_def, command = lambda : my_title_min_max_plot(features, title_scatter_plot_def) , fg = "gray30", font=('arial', 15), background = 'gray90')
    title_mean_min_max_info = Checkbutton(windows_param,  text ="Define chart title     ", variable = title_mean_min_max_def, command = lambda : my_title_min_max_plot(windows_param, title_mean_min_max_def) , fg = "gray30", font=('arial', 15), background = 'gray90')
    title_mean_min_max_info.place(relx = 0.4, rely= 0.73, relwidth= 0.4)
    image_box_view = PhotoImage(file=my_path('figures/barre_p.png'))
    save_graph_p = PhotoImage(file = my_path('figures/save_graph_p.png'))
    bouton_box_view = Button(windows_param, image = image_box_view , font=('arial', 14), text = "View plot",  compound = tk.TOP,  borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
    bouton_box_view.place(relx= 0.17, rely= 0.85, width= 100, height= 60)
    bouton_box_view.bind("<ButtonRelease>", lambda event: check_mean_min_max(event, windows_graphe, my_data, dimension_column, Feature_column_stat, Method_column, var_par)) 
    bouton_save_bar = Button(windows_param, image = save_graph_p , font=('arial', 14), text = "Save plot",  compound = tk.TOP,  borderwidth = 0, relief = FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
    bouton_save_bar.place(relx= 0.73, rely= 0.85, width= 100, height= 60)
    bouton_save_bar.bind("<ButtonRelease>", lambda event: my_save_plot_file(my_barplot)) 
    ##bouton_save_graph.bind("<ButtonRelease>", lambda event: check_mean_min_max(event, windows_graphe, data, dimension_column, Feature_column, Method_column, var_par)) 
    ## Si une dimension et au moins une feature ont deja ete choisies
    ## precedemment, on redessine tout de suite le graphique
    if dimension_column.current() != 0 and len(Feature_column_stat.curselection()) != 0:
        check_mean_min_max(event, windows_graphe, my_data, dimension_column, Feature_column_stat, Method_column, var_par)


## Images du panneau "Descriptive statistics" : chargees UNE fois puis gardees en memoire.
## (Avant : 4 lectures de fichiers PNG a chaque clic, et des PhotoImage stockes dans des
## variables locales, que le ramasse-miettes peut supprimer une fois la fonction terminee.)
_images_calcul_stat = {}


def _image_calcul_stat (fichier):
    if fichier not in _images_calcul_stat:
        _images_calcul_stat[fichier] = PhotoImage(file = my_path('figures/' + fichier))
    return _images_calcul_stat[fichier]


def calcul_stat (event, parent, data, var):

    ## Detruit le panneau cree par un precedent clic au lieu de l'empiler par-dessus
    ## (comme le font deja "Plot corr", "Other analysis" et "Overview stat") : evite
    ## d'accumuler Canvas + widgets en memoire a chaque clic.
    global calcul_stat_widgets
    _detruire_widgets_differe(calcul_stat_widgets)

    str_empty = "            "
    windows_text = Canvas(parent, relief= GROOVE ,  background = 'gray80' , highlightthickness=0)
    features = Canvas(parent, relief= GROOVE, background = 'gray100', highlightthickness=0)
    calcul_stat_widgets = (windows_text, features)

    def creer_bouton (fichier_image, texte, relx):
        ## les 4 boutons ne different que par l'image, le texte et la position
        bouton = Button(features, image = _image_calcul_stat(fichier_image), font=('arial', 14), text = texte, compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray100", background= 'gray100', disabledforeground = 'gray100', focuscolor = 'gray100', highlightcolor = 'gray100', highlightbackground= 'gray100')
        bouton.place(relx= relx, rely = 0.25, width = 130, height=60)
        return bouton

    label_var = Label(features, text = "Groupby" + str_empty, font=('arial', 18), fg =  "gray20", background= "gray100") 
    label_var.place(relx=0.0001, rely= 0.15, relwidth= 0.2)
    var_column = ttk.Combobox(features, values = var, background = "gray100")
    var_column.current(0)
    var_column.place(relx=0.12, rely= 0.16, relwidth= 0.2)
    button_view_stat = creer_bouton('view_stat.png', "View stat", 0.4)
    button_download = creer_bouton('Export_data_p.png', "Export stat", 0.55)
    bouton_corr = creer_bouton('corr_plot.png', "Plot corr", 0.7)
    button_other_analysis = creer_bouton('other_analysis_p.png', "Other analysis", 0.85)
    choice_list = ["All columns", "Continuous columns"]
    label_Choice = Label(features, text = "Choice " + str_empty, font=('arial', 18), fg =  "gray20", background= "gray100") 
    label_Choice.place(relx=0.0001, rely= 0.5, relwidth= 0.2)
    var_Choice = ttk.Combobox(features, values = choice_list, background = "gray100")
    var_Choice.current(0)
    var_Choice.place(relx=0.12, rely= 0.5, relwidth= 0.2)
    button_download.bind("<ButtonRelease>", lambda event:  my_exporting_data_stat(event))
    button_view_stat.bind("<ButtonRelease>", lambda event: check_statistis(event, windows_text, data, var_column, var, var_Choice, choice_list))
    button_other_analysis.bind("<ButtonRelease>", lambda event: mean_min_max(event, windows_text, data, var))
    bouton_corr.bind("<ButtonRelease>", lambda event: linear_correlation(event, windows_text, data, var))
    windows_text.place(relx=0, rely=0, relwidth= 1, relheight= 0.8)
    features.place(relx=0, rely = 0.8, relwidth= 1, relheight=0.2)



# ==========================================================================
# ==  23. Panneaux Edit data et Plots                                     ==
# ==========================================================================
def my_Edit_all_data (event, parent):

    if not analyse_autorisee():      # bouton desactive : un evenement lie a la main peut quand meme arriver
        return

    if not verifier_projet_ouvert():
        return

    global data, data_analy, categorial_var_def, row_edit_value_default, carbon_balance, header_clean_value, header_clean_value_all
    global Feature_categorial
    ##Feature_categorial = None
    if file_path != None:
            try:
                data, _ = lire_fichier_donnees(file_path, file_sheets, na_values = my_missing_values)
            except Exception as erreur_lecture:
                messagebox.showerror("Lecture impossible", "Impossible de lire le fichier de donnees :\n" + str(erreur_lecture))
                return
            parameter_feature = data.columns
            Feature_list_window = []
            for i in range(len(parameter_feature)):
                Feature_list_window.append(parameter_feature[i])
            header_clean = np.concatenate((["Select column name"], Feature_list_window))
            header_clean_value_all = []
            for i in range(len(header_clean)):
                header_clean_value_all.append(header_clean[i])
            header_clean_value_1 = []
            for i in range(len(parameter_feature)):
                x_clean = data[parameter_feature[i]]
                if type(x_clean[0]) == float64:
                    header_clean_value_1.append(parameter_feature[i])
            header_clean_value_2 =  np.concatenate((["Select column name"], header_clean_value_1))
            header_clean_value = []
            for i in range(len(header_clean_value_2)):
                header_clean_value.append(header_clean_value_2[i])
            N = len(data.index)

            ## Reapplique automatiquement les filtres (Parameter 1 a 4)
            ## precedemment enregistres, afin que le jeu de donnees reste
            ## filtre comme avant, meme apres reouverture de cette fenetre
            for reglage_filtre in (filtre_edit_parametre_1, filtre_edit_parametre_2, filtre_edit_parametre_3, filtre_edit_parametre_4):
                if reglage_filtre.get("colonne") in header_clean_value:
                    data = _appliquer_filtre_parametre_edit(data, reglage_filtre)

            height_window = Canvas(parent, relief= GROOVE ,  background = 'white' , highlightthickness=0)
            low_window = Canvas(parent, relief= GROOVE, background = 'gray90', highlightthickness=0)
            label_filter = Label(height_window, text="Filtering parameters", font=('arial', 15), fg = "saddlebrown", bg = "white")
            label_filter.place(relx=0.45, rely= 0.03, relwidth= 0.3)
            label_modality_info = Label(height_window, text="Search & change modality", font=('arial', 15), fg = "saddlebrown", bg = "white")
            label_modality_info.place(relx=0.23, rely= 0.03, relwidth= 0.3)
            label_overview = Label(height_window, text="Parameters for overview", font=('arial', 15), fg = "saddlebrown", bg = "white")
            label_overview.place(relx=0.76, rely= 0.01, relwidth= 0.2)
            label_missing_value = Label(height_window, text="Missing value imputation       ", font=('arial', 15), fg = "saddlebrown", background= "white")
            label_missing_value.place(relx=0.02, rely= 0.03, relwidth= 0.25)
            label_rename_columns = Label(height_window, text = "Rename column", font=('arial', 15), fg = "saddlebrown", bg = "white")
            label_rename_columns.place(relx=0.3, rely= 0.63, relwidth= 0.15)
            list_quanti = ["select method" , "mean", "median", "delete"]
            label_quanti = Label(height_window, text="Quantitative         ", font=('arial', 14), fg = "gray30", background= "white")
            label_quanti.place(relx=0.038, rely= 0.17, relwidth= 0.1)
            imputation_quanti = ttk.Combobox(height_window, values = list_quanti, width= 10)
            imputation_quanti.place(relx=0.038, rely= 0.23, relwidth= 0.2)
            imputation_quanti.current(0)
            list_quali = ["select method" , "modality reference", "delete"]
            label_quali = Label(height_window, text="Qualitative           ", font=('arial', 14), fg = "gray30", background= "white")
            label_quali.place(relx=0.038, rely= 0.37, relwidth= 0.1)
            imputation_quali = ttk.Combobox(height_window, values = list_quali, width= 10)
            imputation_quali.place(relx=0.038, rely= 0.43, relwidth= 0.2)
            imputation_quali.current(0)
            label_Carbon_balance = Label(height_window, text="Parameter - 1           ", font=('arial', 14), fg = "gray30", background= "white")
            label_Carbon_balance.place(relx=0.386, rely= 0.17, relwidth= 0.35)
            carbon_balance = ttk.Combobox(height_window, values = header_clean_value, width= 10)
            carbon_balance.place(relx=0.5, rely= 0.23, relwidth= 0.2)
            if filtre_edit_parametre_1["colonne"] in header_clean_value:
                carbon_balance.current(header_clean_value.index(filtre_edit_parametre_1["colonne"]))
            else:
                carbon_balance.current(0)
            label_conversion = Label(height_window, text="Parameter - 2           ", font=('arial', 14), fg = "gray30", background= "white")
            label_conversion.place(relx=0.386, rely= 0.37, relwidth= 0.35)
            conversion_value = ttk.Combobox(height_window, values = header_clean_value, width= 10)
            conversion_value.place(relx=0.5, rely= 0.43, relwidth= 0.2)
            if filtre_edit_parametre_2["colonne"] in header_clean_value:
                conversion_value.current(header_clean_value.index(filtre_edit_parametre_2["colonne"]))
            else:
                conversion_value.current(0)
            label_yield = Label(height_window, text= "Parameter - 3           ", font=('arial', 14), fg = "gray30", background= "white")
            label_yield.place(relx=0.386, rely= 0.57, relwidth= 0.35)
            yield_value = ttk.Combobox(height_window, values = header_clean_value, width= 10)
            yield_value.place(relx=0.5, rely= 0.63, relwidth= 0.2)
            if filtre_edit_parametre_3["colonne"] in header_clean_value:
                yield_value.current(header_clean_value.index(filtre_edit_parametre_3["colonne"]))
            else:
                yield_value.current(0)
            label_selectivity = Label(height_window, text="Parameter - 4           ", font=('arial', 14), fg = "gray30", background= "white")
            label_selectivity.place(relx=0.386, rely= 0.77, relwidth= 0.35)
            selectivity_value = ttk.Combobox(height_window, values = header_clean_value, width= 10)
            selectivity_value.place(relx=0.5, rely= 0.83, relwidth= 0.2)
            if filtre_edit_parametre_4["colonne"] in header_clean_value:
                selectivity_value.current(header_clean_value.index(filtre_edit_parametre_4["colonne"]))
            else:
                selectivity_value.current(0)
            label_window = Label(height_window, text="Select the columns to display      ", font=('arial', 15), fg = "gray30", bg= "white")
            label_window.place(relx=0.7, rely= 0.08, relwidth= 0.28)
            Feature_window = tk.Listbox(height_window, fg = "gray30", font = ('arial', 12), relief = GROOVE , selectmode = tk.MULTIPLE, exportselection = False)
            Feature_window.place(relx = 0.74, rely = 0.15, relwidth = 0.25, relheight = 0.5)
            label_row_edit = Label(height_window, text="Lines to display      ", font=('arial', 15), fg = "gray30", bg= "white")
            label_row_edit.place(relx=0.7, rely= 0.7, relwidth= 0.2)
            all_columns_def  = BooleanVar(value = False)
            all_columns = Checkbutton(height_window,  text ="All rows", variable = all_columns_def, fg = "gray30", font=('arial', 15), background = 'gray100')
            all_columns.place(relx=0.9, rely= 0.77, relwidth= 0.1)
            categorial_var_def  = BooleanVar(value = False)
            categorial_var = Checkbutton(height_window,  text ="Categorizing variables", variable = categorial_var_def, command = lambda : my_categorial_function(height_window, header_clean_value_all, categorial_var_def, data_analy) , fg = "gray30", font=('arial', 15), background = 'gray100')
            categorial_var.place(relx=0.03, rely= 0.53, relwidth= 0.18)
            row_edit_value_default = DoubleVar() ##StringVar()
            row_edit_value_default.set(N)
            row_edit_value = Spinbox(height_window, from_ = 1 , to = 1000000 , increment = 1, textvariable = row_edit_value_default)
            row_edit_value.place(relx=0.74, rely= 0.77, relwidth= 0.12) 
            separator_view_data1 = tk.Frame(height_window, bd = 10, relief = 'sunken', width = 4)
            separator_view_data1.place(relx=0.27 , rely = 0.05 , relheight= 0.85)
            separator_view_data2 = tk.Frame(height_window, bd = 10, relief = 'sunken', width = 4)
            separator_view_data2.place(relx=0.487 , rely = 0.05 , relheight= 0.85)
            separator_view_data3 = tk.Frame(height_window, bd = 10, relief = 'sunken', width = 4)
            separator_view_data3.place(relx=0.71 , rely = 0.05 , relheight= 0.85)
            label_modality_edit = Label(height_window, text = "Characteristic     ", font=('arial', 14), fg = "gray30", background= "white")
            label_modality_edit.place(relx=0.28, rely= 0.17, relwidth= 0.1)
            modality_edit = ttk.Combobox(height_window, values = header_clean_value_all, width= 10)
            modality_edit.place(relx=0.28, rely= 0.23, relwidth= 0.2)
            modality_edit.current(0)
            label_search = Label(height_window, text = "Search" + "               ", font=('arial', 14), fg = "gray30", background= "white")
            label_search.place(relx=0.28, rely= 0.34, relwidth= 0.1)
            zone_entry_search_default = StringVar()
            zone_entry_search_default.set("old modality")
            zone_entry_search = Entry(height_window, textvariable = zone_entry_search_default)
            zone_entry_search.place(relx=0.28, rely = 0.4, width = 215,  height = 25)
            label_replace = Label(height_window, text = "Replace" + "             ", font=('arial', 14), fg = "gray30", background= "white")
            label_replace.place(relx=0.28, rely= 0.49, relwidth= 0.1)
            zone_entry_replace_default = StringVar()
            zone_entry_replace_default.set("new modality")
            zone_entry_replace = Entry(height_window, textvariable = zone_entry_replace_default)
            zone_entry_replace.place(relx=0.28, rely = 0.55, width = 215,  height = 25)
            rename_columns = ttk.Combobox(height_window, values = header_clean_value_all, width= 10)
            rename_columns.place(relx = 0.28, rely = 0.73, relwidth= 0.2)
            rename_columns.current(0)
            zone_rename_columns_default = StringVar()
            zone_rename_columns_default.set("Enter new name column")
            zone_rename_columns = Entry(height_window, textvariable = zone_rename_columns_default)
            zone_rename_columns.place(relx=0.28, rely = 0.83, width = 215,  height = 25)
            for value in Feature_list_window:
                Feature_window.insert(END, value)
            view_info_data_all = tk.Label(height_window, text = "", bg='white', font=('arial', 15))
            view_info_data_all.place(relx = 0, rely = 0.92, relwidth= 0.36)
            rapport_missing_value_all = Label(height_window, text = "" , font=('arial', 15), bg= "white")
            rapport_missing_value_all.place(relx=0.28, rely= 0.92, relwidth= 0.4)
            change_modality = tk.Label(height_window, text = "", bg='white', font=('arial', 15))
            change_modality.place(relx = 0.65, rely = 0.92, relwidth= 0.3)
            Edit_data_p = PhotoImage(file = my_path('figures/Edit_data_p.png'))
            Export_p = PhotoImage(file = my_path('figures/Export_data_pred.png'))
            view_stat = PhotoImage(file = my_path('figures/view_stat.png'))
            download_p = PhotoImage(file = my_path('figures/Export_data_p.png'))
            data_init = []

            def  Edit_data_all(var_edit_data, var_row_data, check_all_columns, check_carbon_balance, check_conversion, check_yield, check_selectivity, check_missing_quanti, check_missing_quali, check_modality_col, old_modality, new_modality, check_rename, check_rename_value):
                global data_init_all, info_data_all, data_edit_all, data, data_analy
                NN = 0  
                N_missing = str(data.isnull().sum().sum())
                header_columns = data.columns
                header2 = np.concatenate((["Choice value"], header_columns))
                if check_rename.current() != 0:
                    check_x = header2[check_rename.current()]
                    if check_rename_value.get() !=  "Enter new name column":
                        data = data.rename(columns={check_x: check_rename_value.get()})
                if check_modality_col.current() !=0 :
                    check_x = header2[check_modality_col.current()]
                    NN = len(data[data[check_x] == old_modality.get()])
                    if old_modality.get() != "old modality":
                        if new_modality.get() != "new modality":
                            data[check_x]= data[check_x].replace(old_modality.get(), new_modality.get()) 
                ##data_filter_all = my_Edit_data_new (data, check_carbon_balance, check_conversion, check_yield, check_selectivity, check_missing_quanti, check_missing_quali, list_quanti, list_quali)
                quanti_missing_value = check_missing_quanti.get()
                quali_missing_value = check_missing_quali.get()
                if quanti_missing_value != "select method":
                    if quanti_missing_value == "mean":
                        for i in range(len(header_columns)):
                            if type(data[header_columns[i]][0]) != str:
                                data[header_columns[i]] = data[header_columns[i]].fillna(mean(data[header_columns[i]])).infer_objects(copy=False)
                    elif  quanti_missing_value == "median":
                        for i in range(len(header_columns)):
                            if type(data[header_columns[i]][0]) != str:
                                data[header_columns[i]] = data[header_columns[i]].fillna(data[header_columns[i]].median()).infer_objects(copy=False)
                    else:
                         data = data.dropna(axis = 0, how = 'any') 
                if quali_missing_value != "select method":
                    if  quali_missing_value == "modality reference":
                        for i in range(len(header_columns)):
                            if type(data[header_columns[i]][0]) == str:
                                quali_value = data[header_columns[i]].describe()
                                data[header_columns[i]] = data[header_columns[i]].fillna(quali_value[2]).infer_objects(copy=False)
                    else:
                        data = data.dropna(axis = 0, how = 'any') 
                result_edit = [var_edit_data.get(i) for i in var_edit_data.curselection()]
                data_edit_all = data.copy()
                data_analy = data.copy()
                _historique_valider("Edit data")
                if len(result_edit) != 0:
                    data_edit_all = data_edit_all[result_edit]
                if int(var_row_data.get()) > 0:
                    if check_all_columns.get() == False:
                        data_edit_all = data_edit_all.head(int(var_row_data.get()))
                        row_edit_value_default.set(len(data_edit_all))
                    else:
                        data_edit_all = data_edit_all.head(len(data_edit_all))
                        row_edit_value_default.set(len(data_edit_all))
                N_missing = str(data.isnull().sum().sum())
                data_init_all = list(data_edit_all) 
                info_data_all = "Rows: " + str(data_edit_all.shape[0])+ " , Columns: " + str(data_edit_all.shape[1])
                view_info_data_all.configure(text = info_data_all +  "                         ")
                rapport_missing_value_all.configure(text = "Missing value: " +  " N = " + str(N_missing))
                if check_modality_col.current() !=0 :
                    if old_modality.get() != "old modality":
                        if new_modality.get() != "new modality":
                            check_x = header2[check_modality_col.current()]
                            NN = len(data[data[check_x] == new_modality.get()])
                            change_modality.config(text =  " Number of "  +  old_modality.get() + ": " +   str(NN))
                refresh_data_Edit() 

            def refresh_data_Edit(): 
                global window_data_edit , data_init_all
                my_data = data_edit_all.to_numpy().tolist() 
                window_data_edit = ttk.Treeview(low_window, selectmode = 'extended', show = "headings", columns = data_init_all)
                window_data_edit.place(relx = 0, rely= 0, relwidth = 0.8, relheight = 1)
                for i in data_init_all:
                    window_data_edit.column(i, width = 130, anchor = 'c')
                    window_data_edit.heading(i, text = " " + str(i))
                for info_value in my_data:
                    value_p = [x for x in info_value]
                    window_data_edit.insert("",'end', values = value_p)
            carbon_balance.bind("<<ComboboxSelected>>", lambda event: my_carbon_balance(height_window, header_clean_value, carbon_balance))
            conversion_value.bind("<<ComboboxSelected>>", lambda event: my_conversion(height_window, header_clean_value, conversion_value))
            yield_value.bind("<<ComboboxSelected>>", lambda event: my_yields(height_window, header_clean_value, yield_value))
            selectivity_value.bind("<<ComboboxSelected>>", lambda event: my_selectivity(height_window, header_clean_value, selectivity_value))
            button_Edit_data = Button(height_window, image = Edit_data_p, font=('arial', 14), text = "View & Edit", command = lambda : Edit_data_all(Feature_window, row_edit_value, all_columns_def, carbon_balance, conversion_value, yield_value, selectivity_value, imputation_quanti, imputation_quali, modality_edit, zone_entry_search_default, zone_entry_replace_default, rename_columns, zone_rename_columns_default) , compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
            button_Edit_data.place(relx= 0.08, rely= 0.69, width= 110, height= 60)
            Overview_stat = Button(low_window, image = view_stat, font=('arial', 14), text = "Overview stat", command = lambda : stat_data_all(low_window, data_analy) , compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
            Overview_stat.place(relx= 0.85, rely= 0.05, width= 110, height= 60)
            Esport_stat = Button(low_window, image = download_p, font=('arial', 14), text = "Export stat", command = lambda : exporter_stat_overview() , compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
            Esport_stat.place(relx= 0.85, rely= 0.38, width= 110, height= 60)
            Esport_data = Button(low_window, image = Export_p, font=('arial', 14), text = "Export clean", command = lambda : my_exporting_data(data_analy), compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
            Esport_data.place(relx= 0.85, rely= 0.7, width= 110, height= 60)
            height_window.place(relx = 0, rely = 0, relwidth = 1, relheight = 0.6)
            low_window.place(relx = 0, rely = 0.6, relwidth = 1, relheight = 0.4)
    else:
         info_Edit = Label(parent, text = traduire("Please load data to be analyzed"), font = ('Helvetica 16 bold italic'), fg = "red", background= "white")
         info_Edit.place(relx=0, rely= 0, relwidth = 1, relheight = 1)


def windows_plot (event, parent):

        if not analyse_autorisee():      # bouton desactive : un evenement lie a la main peut quand meme arriver
            return

        if not verifier_projet_ouvert():
            return

        global file_path, x_column, windows_plot_widgets
        if file_path != None:
            ## Detruit le panneau (graph + features) cree par un precedent clic
            ## sur "Plots" au lieu de l'empiler par-dessus (evite d'accumuler
            ## Canvas, Comboboxes, Boutons et images en memoire a chaque clic).
            ## Destruction differee (voir _detruire_widgets_differe) car ce
            ## clic peut lui-meme provenir d'un bouton du panneau precedent.
            _detruire_widgets_differe(windows_plot_widgets)

            graph = Canvas(parent, relief= GROOVE ,  background = 'gray80' , highlightthickness=0)
            features = Canvas(parent, relief= GROOVE, background = 'gray90', highlightthickness=0)
            windows_plot_widgets = (graph, features)
            header = list(np.concatenate((["Select column"], data_analy.columns)))
            Target_list =  header 
            Feature_list =  header 
            label_x = Label(features, text="x-axis", fg = "gray30", font=('arial', 15), background= "gray90")
            label_y = Label(features, text="Y-axis", fg = "gray30", font=('arial', 15), bg= "gray90")
            label_color = Label(features, text="Color name  ", fg = "gray30", font=('arial', 15), bg= "gray90")
            label_size = Label(features, text="Point size    ", fg = "gray30", font=('arial', 15), bg= "gray90")
            label_x.place(relx=0.05, rely= 0.05, relwidth= 0.2)
            label_y.place(relx=0.05, rely= 0.15, relwidth= 0.2)
            label_color.place(relx=0.1, rely= 0.25, relwidth= 0.2)
            label_size.place(relx=0.1, rely= 0.35, relwidth= 0.2)
            x_column = ttk.Combobox(features, values = Target_list)
            y_column = ttk.Combobox(features, values = Feature_list)
            Color_columns = ttk.Combobox(features, values = Feature_list)
            ## Restaure la derniere selection valide (X, Y, Couleur) faite avant
            ## de quitter l'onglet "Plots", si elle existe toujours dans les
            ## colonnes actuelles ; sinon on repart sur le texte "Select column"
            x_column.current(_index_restaure(Target_list, dernier_x_scatter))
            y_column.current(_index_restaure(Feature_list, dernier_y_scatter))
            Color_columns.current(_index_restaure(Feature_list, dernier_color_scatter))
            x_column.place(relx=0.43, rely= 0.05, relwidth= 0.5)
            y_column.place(relx=0.43, rely= 0.15, relwidth= 0.5)
            Color_columns.place(relx=0.43, rely= 0.25, relwidth= 0.5)
            row_edit_scatter_default = StringVar()
            row_edit_scatter_default.set("15")
            point_size_scatter = Spinbox(features, from_ = 1 , to = 100000 , increment = 1, justify = CENTER , textvariable = row_edit_scatter_default)
            point_size_scatter.place(relx=0.43, rely= 0.35, relwidth= 0.5) 
            title_scatter_plot_def  = BooleanVar(value = False)
            title_scatter_plot_info = Checkbutton(features,  text ="Define chart title     ", variable = title_scatter_plot_def, command = lambda : my_title_scatter_plot(features, title_scatter_plot_def) , fg = "gray30", font=('arial', 15), background = 'gray90')
            title_scatter_plot_info.place(relx = 0.1, rely= 0.43, relwidth= 0.4)
            data_init = []

            def  Edit_data():
                global data_init
                data_init = list(data_analy) 
                refresh_data() 

            def refresh_data(): 
                global window_data , data_init
                my_data = data_analy.to_numpy().tolist() 
                window_data = ttk.Treeview(graph, selectmode = 'extended', show = "headings", columns = data_init)
                window_data.place(relx = 0, rely= 0, relwidth= 1, relheight= 1)
                for i in data_init:
                    window_data.column(i, width = 130, anchor = 'c')
                    window_data.heading(i, text = " " + str(i))
                for info_value in my_data:
                    value_p = [x for x in info_value]
                    window_data.insert("",'end', values = value_p)
            image_scatter = PhotoImage(file = my_path('figures/scatter_p.png'))
            image_affiche_data = PhotoImage(file=my_path('figures/affiche_data_p.png'))
            image_box_plot = PhotoImage(file = my_path('figures/Box_plot_p.png'))
            descriptive_p = PhotoImage(file=my_path('figures/descriptive_p.png'))
            image_hist = PhotoImage(file=my_path('figures/image_hist_p.png'))
            image_plot3d = PhotoImage(file=my_path('figures/plot3d_p.png'))
            image_optimization= PhotoImage(file=my_path('figures/optimization.png'))
            my_check_str = BooleanVar(value = False)
            ## Filtre du scatter plot suivant les modalites de la variable choisie pour la couleur
            my_check_filtre = BooleanVar(value = False)

            def basculer_filtre_modalites ():
                if my_check_filtre.get():
                    if Color_columns.current() <= 0:
                        my_check_filtre.set(False)
                        messagebox.showinfo("Filter modalities", traduire("Select the color variable (Color name) first."))
                        return
                    colonne_couleur = Feature_list[Color_columns.current()]
                    deja = filtre_modalites_scatter["modalites"] if filtre_modalites_scatter.get("colonne") == colonne_couleur else None
                    choix = _choisir_modalites_scatter(data_analy, colonne_couleur, deja)
                    if choix is None:                # annule : pas de filtre
                        my_check_filtre.set(False)
                        return
                    filtre_modalites_scatter["colonne"] = colonne_couleur
                    filtre_modalites_scatter["modalites"] = choix
                scatter_plot(event, graph, data_analy, x_column, y_column, Color_columns, Feature_list, Feature_list, my_check_str, point_size_scatter, my_check_filtre)

            check_bouton_filtre = Checkbutton(features,  text ="Filter modalities", variable = my_check_filtre, command = basculer_filtre_modalites, fg = "gray30", font=('arial', 15), background = 'gray90')
            check_bouton_str = Checkbutton(features,  text ="Categorized all   ", variable = my_check_str, command = lambda : scatter_plot(event, graph, data_analy, x_column, y_column, Color_columns, Feature_list, Feature_list, my_check_str, point_size_scatter, my_check_filtre) , fg = "gray30", font=('arial', 15), background = 'gray90')
            button_scatter = Button(features, image = image_scatter, font=('arial', 14), text = "Scatter plot", compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
            bouton_affiche_data = Button(features, image = image_affiche_data, command = lambda: Edit_data(), font=('arial', 14), text = "Data table",  compound = tk.TOP,  borderwidth = 0, relief = FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
            button_box_plot = Button(features, image = image_box_plot,  font=('arial', 14), text = "Box plot", compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
            bouton_descriptive = Button(features, image = descriptive_p, font=('arial', 14), text = "Statistical data", compound = tk.TOP,  borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
            bouton_hist = Button(features, image = image_hist, font=('arial', 14), text = "Histogram",  compound = tk.TOP,  borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
            bouton_plot3d = Button(features, image = image_plot3d, font=('arial', 14), text = "plot 3D",  compound = tk.TOP,  borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
            button_optimization = Button(features, image = image_optimization, font=('arial', 14), text = "Optimization", compound = tk.TOP, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = "gray90", background= 'gray90', disabledforeground = 'gray90', focuscolor = 'gray90', highlightcolor = 'gray90', highlightbackground= 'gray90')
            bouton_affiche_data.place(relx= 0.08, rely= 0.5, width= 110, height= 60)
            button_scatter.place(relx= 0.08, rely= 0.65, width= 110, height= 60)
            bouton_hist.place(relx= 0.08, rely= 0.8, width= 110, height= 60)
            button_optimization.place(relx= 0.38, rely= 0.65, width= 110, height= 60)
            bouton_descriptive.place(relx= 0.68, rely= 0.5, width= 110, height= 60)
            button_box_plot.place(relx= 0.68, rely= 0.65, width= 110, height= 60)
            bouton_plot3d.place(relx= 0.68, rely= 0.8, width= 110, height= 60)
            check_bouton_str.place(relx= 0.47, rely= 0.3, relwidth= 0.35)
            check_bouton_filtre.place(relx= 0.1, rely= 0.3, relwidth= 0.35)
            x_column.bind("<<ComboboxSelected>>", lambda event: scatter_plot(event, graph, data_analy, x_column, y_column, Color_columns, Feature_list, Feature_list, my_check_str, point_size_scatter, my_check_filtre))
            y_column.bind("<<ComboboxSelected>>", lambda event: scatter_plot(event, graph, data_analy, x_column, y_column, Color_columns, Feature_list, Feature_list, my_check_str, point_size_scatter, my_check_filtre))
            Color_columns.bind("<<ComboboxSelected>>", lambda event: scatter_plot(event, graph, data_analy, x_column, y_column, Color_columns, Feature_list, Feature_list, my_check_str, point_size_scatter, my_check_filtre))
            ## Si les 3 parametres viennent d'etre restaures automatiquement
            ## (retour sur l'onglet "Plots" apres une precedente selection
            ## complete), on redessine tout de suite le graphique
            if x_column.current() != 0 and y_column.current() != 0 and Color_columns.current() != 0:
                scatter_plot(event, graph, data_analy, x_column, y_column, Color_columns, Feature_list, Feature_list, my_check_str, point_size_scatter, my_check_filtre)
            button_scatter.bind("<ButtonRelease>", lambda event: scatter_plot(event, graph, data_analy, x_column, y_column, Color_columns, Feature_list, Feature_list, my_check_str, point_size_scatter, my_check_filtre)) 
            button_box_plot.bind("<ButtonRelease>", lambda event: Box_plot_view(event, parent, data_analy, Target_list)) 
            bouton_hist.bind("<ButtonRelease>", lambda event: Histograom_plot_view(event, parent, data_analy, Target_list)) 
            bouton_descriptive.bind("<ButtonRelease>", lambda event: calcul_stat(event, parent, data_analy, Target_list)) 
            bouton_plot3d.bind("<ButtonRelease>", lambda event: Box_plot3d_view(event, parent, data_analy, Target_list)) 
            button_optimization.bind("<ButtonRelease>", lambda event: my_optimization(event, parent)) 
            graph.place(relx=0, rely=0, relwidth= 0.6, relheight=1)
            features.place(relx=0.6, rely=0, relwidth= 0.4, relheight=1)
        else:
            info_load = Label(parent, text = traduire("Please load data to be analyzed"), font = ('Helvetica 16 bold italic'), fg = "red", background= "white")
            info_load.place(relx=0, rely= 0, relwidth= 1, relheight=1)



# ==========================================================================
# ==  24. Actions de la barre du haut (New, Import, Export...)            ==
# ==========================================================================
def Website_smartdigicat():
    lien = webbrowser.open('https://www.techniques-ingenieur.fr/actualite/livre-blanc/smartdigicat-un-projet-scientifique-reunissant-catalyse-sciences-des-donnees-et-ia-2-122489/', new=0, autoraise=True)
    print(lien)


def capturer_interface_courante():
    ''' Capture une image de la fenetre principale de l'application telle qu'elle est
    affichee a l'instant (onglet courant, graphiques, tableaux...) et propose a
    l'utilisateur de l'enregistrer sur son disque. '''
    if not verifier_projet_ouvert():
        return

    SmartDigicat.update_idletasks()
    x = SmartDigicat.winfo_rootx()
    y = SmartDigicat.winfo_rooty()
    largeur = SmartDigicat.winfo_width()
    hauteur = SmartDigicat.winfo_height()
    try:
        capture_ecran_complet = ImageGrab.grab()
        largeur_reelle, hauteur_reelle = capture_ecran_complet.size
        largeur_logique = SmartDigicat.winfo_screenwidth()
        hauteur_logique = SmartDigicat.winfo_screenheight()
        echelle_x = largeur_reelle / largeur_logique
        echelle_y = hauteur_reelle / hauteur_logique
        zone = (int(x * echelle_x), int(y * echelle_y), int((x + largeur) * echelle_x), int((y + hauteur) * echelle_y))
        image_capture = capture_ecran_complet.crop(zone)
    except Exception as erreur:
        messagebox.showerror("Capture impossible", "La capture de l'interface a echoue :\n" + str(erreur))
        return
    chemin_fichier = filedialog.asksaveasfilename(defaultextension = ".png", filetypes = [("Image PNG", "*.png"), ("Tous les fichiers", "*.*")], initialfile = "capture_smartdigicat.png", title = "Enregistrer la capture de l'interface")
    if not chemin_fichier:
        return
    try:
        image_capture.save(chemin_fichier)
        messagebox.showinfo("Capture enregistree", "La capture de l'interface a ete enregistree :\n" + chemin_fichier)
    except Exception as erreur:
        messagebox.showerror("Enregistrement impossible", "Impossible d'enregistrer l'image :\n" + str(erreur))


def Open_project ():
    kfichiers = tkinter.filedialog.askopenfilenames() 
    if kfichiers == "" : return
    TXT_Editeur.delete ("1.0" , "end")
    TXT_Editeur.insert ("end" , "Vous avez sélectionné les fichiers :\n\n")
    for kfichier in kfichiers : TXT_Editeur.insert ( "end" , f"{ kfichier }\n")
    TXT_Editeur.insert ("end" , f"\nsoit un total de {len ( kfichiers )} fichiers.")


def import_CSV_data(parent):

    if not verifier_projet_ouvert():
        return

    global file_path, file_sheets, data, data_analy
    my_fill = [('CSV files',"*.csv"), ('Excel file', '*.xlsx *.xlsm *.xls'), ('Texte file', '*.txt'),('All',"*.*")]
    nouveau_chemin_import = filedialog.askopenfilename(title="Open csv File", filetypes = my_fill)
    ## Fichier lu tout de suite : erreur de lecture signalee a l'import, et choix des feuilles
    ## si le classeur Excel en contient plusieurs (plusieurs feuilles aux memes colonnes sont empilees).
    feuilles = None
    if nouveau_chemin_import:
        try:
            donnees_lues, feuilles = lire_fichier_donnees(nouveau_chemin_import, na_values = my_missing_values, choisir_feuilles = True)
        except Exception as erreur_import:
            messagebox.showerror("Import data", "The file could not be read:\n" + str(erreur_import))
            donnees_lues = None
        if donnees_lues is None:            # lecture impossible ou choix des feuilles annule
            nouveau_chemin_import = ""
    ## CORRECTIF : « Annuler » renvoie "" et ecrasait file_path (projet ouvert, donnees perdues) ;
    ## et un nouveau fichier (ou d'autres feuilles) doit invalider data / data_analy de l'ancien fichier.
    if nouveau_chemin_import:
        if nouveau_chemin_import != file_path or feuilles != file_sheets:
            data = None
            data_analy = None
        file_path = nouveau_chemin_import
        file_sheets = feuilles
        _historique_valider("Import data")
    parent1 = Canvas(parent, bg = "white", highlightthickness=0)
    parent1.place(relx= 0, rely= 0, relwidth=1, relheight=1)
    tree = ttk.Treeview(parent1, show="headings")
    tree.place(relx=0.01, rely= 0.1, relwidth= 1, relheight= 0.9)
    if nouveau_chemin_import:
        label_x = Label(parent1, text = traduire("Data successfully loaded! \n You can carry out the desired analysis. \n (descriptive analyses or modeling)"), font=('Helvetica 18 bold italic'), fg = "gray20", background= "white")
        label_x.place(relx=0.3, rely= 0.3, relwidth= 0.4)
    else:
        label_x = Label(parent1, text = traduire("Please try again import failed"), font=('Helvetica 18 bold italic'), fg = "red", background= "white")
        label_x.place(relx=0.3, rely= 0.3, relwidth= 0.4)


def exporter_modeles_csv ():

    if not analyse_autorisee():      # bouton desactive
        return

    if len(modeles_projet) == 0:
        messagebox.showinfo("Exporter les modeles", "Aucun modele cree pour le moment.")
        return

    chemin_csv = filedialog.asksaveasfilename(defaultextension = ".csv", filetypes = [("Fichier CSV", "*.csv"), ("Tous les fichiers", "*.*")], initialfile = "modeles.csv", title = "Exporter les modeles vers un fichier CSV")
    if not chemin_csv:
        return

    lignes = []
    for nom, etat in modeles_projet.items():
        donnees_modele = etat.get("data_model")
        r2 = None
        mse = None
        if donnees_modele is not None:
            try:
                r2 = round(float(donnees_modele[0]), 4)
            except Exception:
                pass
            try:
                mse = round(float(donnees_modele[3]), 4)
            except Exception:
                pass
        features = etat.get("result_model")
        nb_features = len(features) if isinstance(features, (list, tuple, np.ndarray)) else None
        if isinstance(features, (list, tuple, np.ndarray)):
            features_utilisees = ", ".join([str(f) for f in features])
        else:
            features_utilisees = None
        lignes.append({
            "Nom du modele": nom,
            "Type de modele": etat.get("Name_model"),
            "Variable cible": etat.get("selectet_target"),
            "Methode d'entrainement": etat.get("selectet_methode"),
            "Codage": etat.get("selectet_codage"),
            "Imputation": etat.get("selectet_methode_imp"),
            "Nombre de features": nb_features,
            "Features utilisees": features_utilisees,
            "R-squared": r2,
            "MSE": mse,
            "Statut": "Entraine" if donnees_modele is not None else "Non entraine",
        })

    tableau_export = pd.DataFrame(lignes)
    try:
        tableau_export.to_csv(chemin_csv, index = False, encoding = "utf-8-sig")
        messagebox.showinfo("Exporter les modeles", "Les modeles ont ete exportes avec succes vers :\n" + chemin_csv)
    except Exception as erreur:
        messagebox.showwarning("Exporter les modeles", "L'export a echoue :\n" + str(erreur))


def nouveau_projet_top ():
 
    global nom_projet, chemin_projet, nom_modele_actif, file_path, file_sheets, Name_model, Target_column, Feature_column
    global Calculation_method, codage_method, Test_size, Target_list, header_target
    global data, data_analy, x_column
    global taille_boites_modeles, zoom_graphique_modeles, decalage_zoom_x, decalage_zoom_y
    global dernier_x_scatter, dernier_y_scatter, dernier_color_scatter
    global dernier_methode_boxplot, dernier_x_boxplot, dernier_y_boxplot, dernier_color_boxplot
    global dernier_var_histogram, dernier_couleur_histogram, dernier_densite_histogram
    global dernier_x_plot3d, dernier_y_plot3d, dernier_z_plot3d, dernier_color_plot3d
    global dernier_target_correlation, dernier_features_correlation
    global dernier_dimension_min_max, dernier_methode_min_max, dernier_features_min_max
    global opt_plot_data, data_opt

    if len(modeles_projet) > 0 or file_path is not None:
        confirmation = messagebox.askyesno("Nouveau projet", "Creer un nouveau projet videra l'espace de travail actuel : les donnees chargees et les modeles du projet en cours (s'ils n'ont pas ete enregistres via « Save ») seront perdus.\n\nContinuer ?")
        if not confirmation:
            return

    nom = simpledialog.askstring("Nouveau projet", "Nom du nouveau projet", initialvalue = "Nouveau projet")
    if nom is None or nom.strip() == "":
        return

    emplacement = filedialog.asksaveasfilename(defaultextension = ".json", filetypes = [("Fichier JSON", "*.json"), ("Tous les fichiers", "*.*")], initialfile = nom.replace(" ", "_") + ".json", title = "Choisissez l'emplacement d'enregistrement du projet « " + nom + " »")
    if not emplacement:
        return

    ## ---- Reinitialisation complete de l'espace de travail ----
    _fermer_panneau_config()
    modeles_projet.clear()
    modeles_supprimes.clear()
    couleurs_boites_modeles.clear()
    positions_boites_modeles.clear()
    couleurs_boites_supprimees.clear()
    positions_boites_supprimees.clear()
    taille_boites_modeles = 88
    zoom_graphique_modeles = 1.0
    decalage_zoom_x = 0.0
    decalage_zoom_y = 0.0
    filtre_edit_parametre_1.clear()
    filtre_edit_parametre_1.update({"colonne": None, "min": 0, "max": 100})
    filtre_edit_parametre_2.clear()
    filtre_edit_parametre_2.update({"colonne": None, "min": 0, "max": 100})
    filtre_edit_parametre_3.clear()
    filtre_edit_parametre_3.update({"colonne": None, "min": 0, "max": 100})
    filtre_edit_parametre_4.clear()
    filtre_edit_parametre_4.update({"colonne": None, "min": 0, "max": 100})
    opt_parameters.clear()
    opt_parameters.update(parametres_optimisation_par_defaut())
    opt_plot_data = None
    data_opt = None
    opt_affichage.clear()
    opt_affichage.update(affichage_optimisation_par_defaut())
    data = None
    data_analy = None
    x_column = None
    dernier_x_scatter = None
    dernier_y_scatter = None
    dernier_color_scatter = None
    dernier_methode_boxplot = None
    dernier_x_boxplot = None
    dernier_y_boxplot = None
    dernier_color_boxplot = None
    dernier_var_histogram = None
    dernier_couleur_histogram = None
    dernier_densite_histogram = None
    dernier_x_plot3d = None
    dernier_y_plot3d = None
    dernier_z_plot3d = None
    dernier_color_plot3d = None
    dernier_target_correlation = None
    dernier_features_correlation = None
    dernier_dimension_min_max = None
    dernier_methode_min_max = None
    dernier_features_min_max = None

    ## Vide entierement le panneau central (quelle que soit la vue affichee :
    ## boites de modeles, graphiques, tableau de donnees, etc.)
    for enfant in list(Center_frame.winfo_children()):
        enfant.destroy()

    nom_projet = nom
    chemin_projet = emplacement
    nom_modele_actif = None
    file_path = None
    file_sheets = None
    Name_model = None
    Target_column = None
    Feature_column = None
    Calculation_method = None
    codage_method = None
    Test_size = None
    Target_list = None
    header_target = None
    for cle in _MODEL_STATE_KEYS:
        globals()[cle] = None
    ## CORRECTIF : les widgets de configuration du projet precedent (deja detruits) restaient
    ## references par ces variables globales -> valeurs perimees / TclError a l'entrainement.
    for cle in ("feature_select", "Alpha_size", "Alpha_lasso_size", "C_size", "gamma_size", "Random_size",
                "trees_size", "estimators_size", "depth_size", "loss_size", "rate_size", "Random_size_r",
                "depth_size_r", "Random_size_gpr", "recoding_param", "label_Test_size"):
        globals()[cle] = None
    globals()["result_model_1"] = []
    for cle in ("file_path_pred", "data_pred", "data_pred_new"):
        globals()[cle] = None
    globals()["prediction_infos"] = {}
    _historique_reinitialiser("New project")

    _maj_titre_fenetre()

    try:
        save_current_project(silencieux = True)
        messagebox.showinfo("Nouveau projet", "Le projet « " + nom + " » a ete cree et enregistre dans :\n" + chemin_projet + "\n\nVous pouvez maintenant importer des donnees (« Import data »), puis creer un ou plusieurs modeles avec le bouton « Analysis » (barre verticale gauche).")
    except Exception as erreur:
        messagebox.showwarning("Nouveau projet", "Le projet « " + nom + " » a ete cree, mais son enregistrement initial a echoue :\n" + str(erreur) + "\n\nPensez a utiliser le bouton « Save » avant de fermer le logiciel.")



# ==========================================================================
# ==  25. Construction de la fenetre principale                           ==
# ==========================================================================
##------ Configuration de l'interface principale
SmartDigicat = tk.Tk()
SmartDigicat.geometry("1200x750")
SmartDigicat.title("SmartDigiCat v1.0")
canva = tk.Canvas(SmartDigicat)

TKV_Choix = StringVar()

##------ Chargement des icones
Run = PhotoImage(file = my_path('figures/run_p.png'))
Open = PhotoImage(file=my_path('figures/open_p.png'))
New = PhotoImage(file=my_path('figures/new_p.png'))
Save = PhotoImage(file=my_path('figures/save_p.png'))
Undo = PhotoImage(file=my_path('figures/undo_p.png'))
Redo = PhotoImage(file=my_path('figures/redo_p.png'))
Model = PhotoImage(file=my_path('figures/model_p.png'))
Log = PhotoImage(file=my_path('figures/log_p.png'))
Results = PhotoImage(file=my_path('figures/results_p.png'))
Help = PhotoImage(file=my_path('figures/help_p.png'))
Website = PhotoImage(file=my_path('figures/web_p.png'))
Analysis = PhotoImage(file=my_path('figures/new_analysis.png'))
Delete = PhotoImage(file=my_path('figures/delete.png'))
Restore = PhotoImage(file=my_path('figures/restore.png'))
Picture_save = PhotoImage(file=my_path('figures/picture_save.png'))
Csv_import = PhotoImage(file=my_path('figures/csv_import.png'))
Settings = PhotoImage(file=my_path('figures/settings.png'))
langue = PhotoImage(file=my_path('figures/langue.png'))


try:
    save_config_model = PhotoImage(file=my_path('figures/save_config_model.png'))
except Exception:
    save_config_model = Save

data_import = PhotoImage(file=my_path('figures/import_p.png'))
my_data_edit_p = PhotoImage(file=my_path('figures/settings_p-32.png'))
img_smatdigicat = PhotoImage(file=my_path('figures/icon.png'))
Model_edit = PhotoImage(file=my_path('figures/Model_edit.png'))
New_model = PhotoImage(file=my_path('figures/New_model.png'))
Rename_model = PhotoImage(file=my_path('figures/Rename_model.png'))
Delet_model = PhotoImage(file=my_path('figures/Delet_model.png'))
Arret = PhotoImage(file=my_path('figures/arret.png'))


##------ Barre horizontale haute
top_frame = Canvas(SmartDigicat, width= 500 , height = 70, bg="slate gray", highlightthickness=0)

##------ Barre horizontale basse
bottom_frame = Label(SmartDigicat, text= TEXTE_BARRE_BASSE , fg="white", bg="slate gray")

##------ Barre verticale gauche
left_frame = Canvas(SmartDigicat, width= 50, bg='gray8', highlightthickness=0, relief=GROOVE)

##------ Barre verticale droite
right_frame = Canvas(SmartDigicat, width= 50, bg='gray8', highlightthickness=0, relief = GROOVE)

##------ Panneau central
Center_frame = Canvas(SmartDigicat, highlightthickness=0, background= 'white' , relief = GROOVE)

##print(Center_frame.winfo_children())

##------ Menus horizontals haut

button_New = Button(top_frame, image = New, text='New', command = lambda : nouveau_projet_top(), repeatdelay = 0, borderwidth = 0 , compound = tk.LEFT, borderless = True, background= 'slate gray', disabledforeground = 'slate gray', focuscolor = 'slate gray', bordercolor = 'slate gray', highlightcolor = 'slate gray', highlightbackground= 'slate gray')
button_Open = Button(top_frame, image = Open , text='Open', command = lambda : ouvrir_projet_depuis_fichier(), borderwidth = 0 , compound = tk.LEFT, relief=FLAT, borderless = True, bordercolor = 'slate gray', background= 'slate gray', disabledforeground = 'slate gray', focuscolor = 'slate gray', highlightcolor = 'slate gray', highlightbackground= 'slate gray')
button_Save = Button(top_frame, image = Save , text='Save' , command = lambda : (save_current_project() if analyse_autorisee() else None), borderwidth = 0, compound = tk.LEFT, relief=FLAT, borderless = True, bordercolor = 'red', background= 'slate gray', disabledforeground = 'slate gray', focuscolor = 'slate gray', highlightcolor = 'slate gray', highlightbackground= 'slate gray')
button_Undo = Button(top_frame, image = Undo, text= 'Undo',  borderwidth = 0, compound = tk.LEFT, borderless = True, bordercolor = 'slate gray', background= 'slate gray', disabledforeground = 'slate gray', focuscolor = 'slate gray', highlightcolor = 'slate gray', highlightbackground= 'slate gray')
button_Redo = Button(top_frame, image = Redo, text= 'Redo',  borderwidth = 0, compound = tk.LEFT, borderless = True, bordercolor = 'slate gray', background= 'slate gray', disabledforeground = 'slate gray', focuscolor = 'slate gray', highlightcolor = 'slate gray', highlightbackground= 'slate gray')
button_Model = Button(top_frame, image = Model, text= 'Model', borderwidth = 0, compound = tk.LEFT, borderless = True, bordercolor = 'slate gray', background= 'slate gray', disabledforeground = 'slate gray', focuscolor = 'slate gray', highlightcolor = 'slate gray', highlightbackground= 'slate gray')
button_Run = Button(top_frame, image = Run, text='Run' , borderwidth = 0, compound = tk.LEFT, relief=FLAT, borderless = True, bordercolor = 'red', background= 'slate gray', disabledforeground = 'slate gray', focuscolor = 'slate gray', highlightcolor = 'slate gray', highlightbackground= 'slate gray')
button_data = Button(top_frame, image = data_import, text = 'Import data' , compound = tk.LEFT,  command = lambda : import_CSV_data(Center_frame) ,  borderwidth = 0, relief = FLAT, borderless = True, bordercolor = 'slate gray', background= 'slate gray', disabledforeground = 'slate gray', focuscolor = 'slate gray', highlightcolor = 'slate gray', highlightbackground= 'slate gray')

button_data_Edit = Button(top_frame, image = my_data_edit_p, text = 'Edit data' , compound = tk.LEFT,  borderwidth = 0, relief = FLAT, borderless = True, bordercolor = 'slate gray', background= 'slate gray', disabledforeground = 'slate gray', focuscolor = 'slate gray', highlightcolor = 'slate gray', highlightbackground= 'slate gray')

button_Results = Button(top_frame, image = Results, text= 'Plots', borderwidth = 0, compound = tk.LEFT, borderless = True, bordercolor = 'slate gray', background= 'slate gray', disabledforeground = 'slate gray', focuscolor = 'slate gray', highlightcolor = 'slate gray', highlightbackground= 'slate gray')
button_Log = Button(top_frame, image = Log, text= 'Log', borderwidth = 0, compound = tk.LEFT, borderless = True, bordercolor = 'slate gray', background= 'slate gray', disabledforeground = 'slate gray', focuscolor = 'slate gray', highlightcolor = 'slate gray', highlightbackground= 'slate gray')
button_Help = Button(top_frame, image = Help, text= 'Help', borderwidth = 0, compound = tk.LEFT, borderless = True, bordercolor = 'slate gray', background= 'slate gray', disabledforeground = 'slate gray', focuscolor = 'slate gray', highlightcolor = 'slate gray', highlightbackground= 'slate gray')
button_Website = Button(top_frame, image = Website, text= 'Website', width= 1, command = Website_smartdigicat, borderwidth = 0, compound = tk.LEFT, borderless = True, bordercolor = 'slate gray', background= 'slate gray', disabledforeground = 'slate gray', focuscolor = 'slate gray', highlightcolor = 'slate gray', highlightbackground= 'slate gray')

##affiche_est = Label(Center_frame , textvariable = TKV_Choix , width = 25).pack()

##------ Menus verticals de gauche
button_Analysis = Button(left_frame, image = Analysis, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = 'gray8', background= 'gray8', disabledforeground = 'gray8', focuscolor = 'gray8', highlightcolor = 'gray8', highlightbackground= 'gray8')
##button_Delete = Button(left_frame, image = Delete, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = 'gray8', background= 'gray8', disabledforeground = 'gray8', focuscolor = 'gray8', highlightcolor = 'gray8', highlightbackground= 'gray8')
button_Restore = Button(left_frame, image = Restore,  borderwidth = 0, relief=FLAT, borderless = True, bordercolor = 'gray8', background= 'gray8', disabledforeground = 'gray8', focuscolor = 'gray8', highlightcolor = 'gray8', highlightbackground= 'gray8')
button_Picture_save = Button(left_frame, image = Picture_save, command = capturer_interface_courante, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = 'gray8', background= 'gray8', disabledforeground = 'gray8', focuscolor = 'gray8', highlightcolor = 'gray8', highlightbackground= 'gray8')
button_Csv_import = Button(left_frame, image = Csv_import , command = lambda : exporter_modeles_csv(), borderwidth = 0, relief = FLAT, borderless = True, bordercolor = 'gray8', background= 'gray8', disabledforeground = 'gray8', focuscolor = 'gray8', highlightcolor = 'gray8', highlightbackground= 'gray8')
button_langue = Button(left_frame, image = langue, command = lambda : choisir_langue_instructions(), borderwidth = 0, relief = FLAT, borderless = True, bordercolor = 'gray8', background= 'gray8', disabledforeground = 'gray8', focuscolor = 'gray8', highlightcolor = 'gray8', highlightbackground= 'gray8')


button_Delete_model = Button(left_frame, image = Delete, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = 'gray8', background= 'gray8', disabledforeground = 'gray8', focuscolor = 'gray8', highlightcolor = 'gray8', highlightbackground= 'gray8',
                              command = lambda : messagebox.showinfo("Delete", "Ouvrez d'abord l'onglet « Model » et selectionnez un modele dans la liste."))
button_Rename = Button(left_frame, image = Rename_model, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = 'gray8', background= 'gray8', disabledforeground = 'gray8', focuscolor = 'gray8', highlightcolor = 'gray8', highlightbackground= 'gray8',
                        command = lambda : messagebox.showinfo("Rename", "Ouvrez d'abord l'onglet « Model » et selectionnez un modele dans la liste."))
button_Settings = Button(left_frame, image = Settings, borderwidth = 0, relief=FLAT, borderless = True, bordercolor = 'gray8', background= 'gray8', disabledforeground = 'gray8', focuscolor = 'gray8', highlightcolor = 'gray8', highlightbackground= 'gray8',
                          command = lambda : messagebox.showinfo("Setting", "Ouvrez d'abord l'onglet « Model », selectionnez un modele dans la liste, puis cliquez ici pour l'ouvrir."))

## Bouton d'arret : arrete l'execution lancee par « Run » ; desactive en dehors d'une execution
button_arret_model = Button(left_frame, image = Arret, command = demander_arret_execution, state = DISABLED, disabledbackground = 'gray8', borderwidth = 0, relief = FLAT, borderless = True, bordercolor = 'gray8', background= 'gray8', disabledforeground = 'gray8', focuscolor = 'gray8', highlightcolor = 'gray8', highlightbackground= 'gray8')


scale = Scale(left_frame, from_= 100, to= 0, sliderrelief= "ridge" , showvalue= False,  borderwidth= 0 , bg='cornflowerblue', activebackground= 'cornflowerblue', width= 35 , length= 180, troughcolor = 'gray80')

button_New.place(x = 5, y = 4, width = 70)
button_Open.place(x = 80 , y = 4, width =70)
button_Save.place(x = 155, y = 4, width =70)
button_Undo.configure(command = annuler_derniere_action)
button_Redo.configure(command = refaire_derniere_action)
button_Undo.place(x = 230, y = 4, width =70)
button_Redo.place(x = 305, y = 4, width = 70)
button_data.place(relx = 0.36, y = 4, width = 110)
button_data_Edit.place(relx = 0.458, y = 4, width = 100)
button_Results.place(relx = 0.548 , y = 4, width = 85)
button_Model.place(relx = 0.626 , y = 4, width = 80)
button_Run.place(relx = 0.70 , y = 4, width = 70)
button_Log.place(relx = 0.766 , y = 4, width = 85)
button_Help.place(relx = 0.844, y = 4, width = 70)
button_Help.bind("<ButtonRelease>", lambda event: afficher_manuel(event, Center_frame))      # manuel d'utilisation, dans la langue choisie
button_Website.pack(side = tk.RIGHT, fill = tk.NONE,  ipadx = 45, padx= 5)

button_Results.bind("<ButtonRelease>", lambda event: windows_plot(event, Center_frame))  
button_Run.bind("<ButtonRelease>", lambda event: lancer_run(event))     # active le bouton d'arret pendant l'execution
button_Model.bind("<ButtonRelease>", lambda event: choice_model(event, Center_frame)) 
button_Analysis.bind("<ButtonRelease>", lambda event: save_Model(event, Center_frame)) 
button_Log.bind("<ButtonRelease>", lambda event: my_predict_function(event, Center_frame)) 
button_data_Edit.bind("<ButtonRelease>", lambda event: my_Edit_all_data(event, Center_frame)) 

button_Analysis.place(x = 5, y = 5, width= 40, height= 40)
button_Delete_model.place(x = 5, y = 60, width= 40, height= 40)
button_Rename.place(x = 5, y = 115, width= 40, height= 40)
button_Settings.place(x = 5, y = 165, width= 40, height= 40)
button_Restore.place(x = 5, y = 220, width= 40, height= 40)
button_arret_model.place(x = 5, y = 275, width= 40, height= 40)      # juste avant button_Picture_save
button_Picture_save.place(x = 5, y = 330, width= 40, height= 40)
button_Csv_import.place(x = 5, y = 385, width= 40, height= 40)
button_langue.place(x = 5, y = 440, width= 40, height= 40)



##button_New_model.place(x = 5, y = 385, width= 40, height= 40)
##button_Delete_model.place(x = 5, y = 440, width= 40, height= 40)
##button_data.place(x = 5, y = 275, width= 40, height= 40)
scale.pack(side= BOTTOM, fill= BOTH, padx= 5)

top_frame.pack(side="top", fill='x', ipady= 4)
bottom_frame.pack(side="bottom", fill='x')
left_frame.pack(side="left", fill='y')
right_frame.pack(side="right", fill='y')
Center_frame.pack(side= 'top', fill= BOTH, expand= YES)

img_smatdigicat_ = img_smatdigicat.zoom(3, 3)
can = Canvas(Center_frame, highlightthickness=0, background= 'white')
logo_id = can.create_image( 320, 320, anchor = 'center' , image = img_smatdigicat_)        
can.place(relx = 0.2, rely = 0, relwidth = 1, relheight = 1) 

##------ Position du logo rendue relative a la taille du panneau,
##       comme pour les blocs Authors et Contacts (relx / rely)
SmartDigicat.update_idletasks()
_largeur_can_initiale = can.winfo_width() or 1
_hauteur_can_initiale = can.winfo_height() or 1
REL_X_LOGO = 320 / _largeur_can_initiale
REL_Y_LOGO = 320 / _hauteur_can_initiale

def repositionner_logo(event):
    """Recalcule la position du logo en fonction de la taille actuelle du
    canevas, pour qu'il reste a la meme position relative (equivalent d'un
    relx/rely) quelle que soit la taille de la fenetre."""
    can.coords(logo_id, event.width * REL_X_LOGO, event.height * REL_Y_LOGO)

can.bind("<Configure>", repositionner_logo)

##welcome = Label(Center_frame, text="Welcome to  SmartDigicat \n \n  Your catalytic screening software", font=('Helvetica 16 bold italic'), fg = "gray20", background= "white")

##welcome = Label(Center_frame, text="SmartDigicat", font=('Helvetica 30 bold italic'), fg = "gray20", background= "white")
autor = Label(Center_frame, text="Authors:                \n \n Sebastien Paul       \n \n Sophie Dabo-niang \n \n Komlan Noukpoape \n \n Wahba Saweiras      \n \n Svetlana HEYTE      ", font=('arial', 15), fg = "dimgray", background= "white")
Contacts = Label(Center_frame, text="Contact:                                        \n \n 1. sebastien.paul@centralelille.fr      \n \n 2. sophie.dabo@univ-lille.fr              \n \n 3. komlan.noukpoape@centralelille.f \n \n 4. wahba.saweiras@centralelille.fr     \n \n  5. svetlana.heyte@centralelille.fr         ", font=('arial', 15), fg = "dimgray", background= "white")
##welcome.place(relx=0.3, rely= 0.73, relwidth= 0.4)
autor.place(relx=0, rely= 0.7, relwidth= 0.4)
Contacts.place(relx=0.6, rely= 0.7, relwidth= 0.4)


# ==========================================================================
# ==  26. Enregistrement automatique, etat des boutons, manuel (Help)     ==
# ==========================================================================
##------ Enregistrement automatique du projet toutes les 5 minutes
INTERVALLE_AUTOSAVE_MS = 5 * 60 * 1000  # 5 minutes

def autosave_projet():
    """Enregistre automatiquement le projet en cours, sans interrompre l'utilisateur,
    puis se replanifie elle-meme pour s'executer a nouveau 5 minutes plus tard."""
    if nom_projet is not None and chemin_projet is not None:
        try:
            save_current_project(silencieux = True)
        except Exception:
            pass
    _planifier_tache_application("autosave", INTERVALLE_AUTOSAVE_MS, autosave_projet)

_planifier_tache_application("autosave", INTERVALLE_AUTOSAVE_MS, autosave_projet)

##------ Boutons lies au projet. Barre gauche : Analysis, Restore, Delete, Rename, Settings,
##       Csv_import. Barre du haut : Undo, Redo, Model, Run, Edit data, Log, Plots, Save.
##       Ils ne sont actifs que si un projet est ouvert (New / Open) ET si des
##       donnees ont ete importees (Import data).
##  L'etat est reevalue toutes les 250 ms : plusieurs fonctions modifient
##  nom_projet / file_path / data (nouveau projet, ouverture, import...), ce qui
##  evite d'avoir a les modifier une par une.

def _icone_grisee(chemin_relatif):
    """Version grisee (et semi-transparente) d'une icone : tkmacosx ne change pas
    l'apparence d'un bouton-image desactive. None si l'image ne peut pas etre lue."""
    try:
        from PIL import Image as _ImagePIL, ImageTk as _ImageTkPIL
        icone = _ImagePIL.open(my_path(chemin_relatif)).convert("RGBA")
        gris = icone.convert("L")
        alpha = icone.getchannel("A").point(lambda valeur: int(valeur * 0.35))
        return _ImageTkPIL.PhotoImage(_ImagePIL.merge("RGBA", (gris, gris, gris, alpha)))
    except Exception:
        return None


## (bouton, icone normale, icone grisee) : les icones grisees sont gardees ici
## pour ne pas etre supprimees de la memoire

## ======================================================================
## ==  MANUEL D'UTILISATION (bouton « Help »)                          ==
## ======================================================================
## Le manuel (versions francaise et anglaise) est integre au logiciel sous forme d'images de
## pages : figures/manuel/fr/page-XX.png et figures/manuel/en/page-XX.png, avec les PDF
## complets (figures/manuel/SmartDigiCat_manuel_FR.pdf et SmartDigiCat_manual_EN.pdf).
## « Help » l'affiche dans le panneau central, dans la langue choisie (bouton langue).

_manuel_affiche = {"widget": None, "parent": None}
_FICHIERS_PDF_MANUEL = {"fr": "SmartDigiCat_manuel_FR.pdf", "en": "SmartDigiCat_manual_EN.pdf"}
_TEXTES_MANUEL = {
    "fr": {"titre": "Manuel d'utilisation", "pdf": "Ouvrir le PDF", "page": "Page", "sur": "sur",
           "absent": "Le manuel est introuvable.\n\nIl doit se trouver dans le dossier du logiciel :\nfigures/manuel/fr/ (pages) et figures/manuel/SmartDigiCat_manuel_FR.pdf"},
    "en": {"titre": "User manual", "pdf": "Open the PDF", "page": "Page", "sur": "of",
           "absent": "The manual cannot be found.\n\nIt must be in the software folder:\nfigures/manuel/en/ (pages) and figures/manuel/SmartDigiCat_manual_EN.pdf"},
}


def _ouvrir_pdf_manuel (langue):
    """ Ouvre le PDF complet du manuel avec le lecteur PDF du systeme. """
    chemin = my_path(os.path.join("figures", "manuel", _FICHIERS_PDF_MANUEL[langue]))
    if not os.path.isfile(chemin):
        messagebox.showwarning("Help", _TEXTES_MANUEL[langue]["absent"])
        return
    try:
        if sys.platform.startswith("darwin"):
            import subprocess
            subprocess.Popen(["open", chemin])
        elif sys.platform.startswith("win"):
            os.startfile(chemin)
        else:
            import subprocess
            subprocess.Popen(["xdg-open", chemin])
    except Exception:
        webbrowser.open("file://" + os.path.abspath(chemin))


def afficher_manuel (event = None, parent = None):
    """ Bouton « Help » : affiche le manuel d'utilisation dans le panneau central, dans la
    langue choisie. Les pages sont ajustees a la largeur du panneau et chargees au fur et
    a mesure du defilement (seules les pages visibles sont gardees en memoire). """

    if parent is None:
        parent = Center_frame
    langue = langue_courante if langue_courante in _FICHIERS_PDF_MANUEL else "en"
    textes = _TEXTES_MANUEL[langue]
    dossier = my_path(os.path.join("figures", "manuel", langue))
    pages = sorted(os.path.join(dossier, f) for f in os.listdir(dossier) if f.lower().endswith(".png")) if os.path.isdir(dossier) else []

    ## un seul manuel affiche a la fois
    ancien = _manuel_affiche["widget"]
    if ancien is not None:
        try:
            if ancien.winfo_exists():
                ancien.destroy()
        except Exception:
            pass
    try:
        scale.configure(command = "")       # le curseur de zoom ne pilote pas le manuel
    except Exception:
        pass

    panneau = tk.Frame(parent, background = "gray85")
    panneau.place(relx = 0, rely = 0, relwidth = 1, relheight = 1)
    _manuel_affiche["widget"] = panneau
    _manuel_affiche["parent"] = parent

    ## barre du haut : titre, numero de page, bouton PDF
    barre = tk.Frame(panneau, background = "white", height = 44)
    barre.pack(side = TOP, fill = X)
    barre.pack_propagate(False)
    Label(barre, text = "  " + textes["titre"] + "  (" + LANGUES_DISPONIBLES[langue] + ")", font = ('arial', 16, 'bold'), fg = "saddlebrown", bg = "white").pack(side = LEFT, padx = 6)
    bouton_pdf = tk.Button(barre, text = textes["pdf"], command = lambda: _ouvrir_pdf_manuel(langue))
    bouton_pdf.pack(side = RIGHT, padx = 10, pady = 6)
    label_page = Label(barre, text = "", font = ('arial', 13), fg = "gray30", bg = "white")
    label_page.pack(side = RIGHT, padx = 10)

    if len(pages) == 0:
        Label(panneau, text = textes["absent"], font = ('arial', 15), fg = "red", bg = "gray85", justify = CENTER).pack(expand = True)
        return

    zone = tk.Frame(panneau, background = "gray85")
    zone.pack(side = TOP, fill = BOTH, expand = True)
    canevas = Canvas(zone, background = "gray85", highlightthickness = 0)
    defilement = ttk.Scrollbar(zone, orient = "vertical", command = canevas.yview)
    defilement.pack(side = RIGHT, fill = Y)
    canevas.pack(side = LEFT, fill = BOTH, expand = True)

    MARGE = 16
    try:
        largeur_origine, hauteur_origine = Image.open(pages[0]).size
    except Exception:
        largeur_origine, hauteur_origine = 993, 1404
    etat = {"largeur": 0, "positions": [], "hauteur": 0, "images": {}, "tache": None}

    def mettre_en_page ():
        """ Taille des pages = largeur du panneau (au plus la taille d'origine). """
        etat["tache"] = None
        largeur_panneau = canevas.winfo_width()
        if largeur_panneau <= 1:
            etat["tache"] = canevas.after(50, mettre_en_page)
            return
        largeur = int(_builtins_langue.max(200, _builtins_langue.min(largeur_origine, largeur_panneau - 2 * MARGE)))
        if largeur == etat["largeur"]:
            charger_pages_visibles()
            return
        etat["largeur"] = largeur
        etat["hauteur"] = int(largeur * hauteur_origine / largeur_origine)
        canevas.delete("all")
        etat["images"].clear()
        x = largeur_panneau / 2
        etat["positions"] = []
        y = MARGE
        for numero in range(len(pages)):
            etat["positions"].append(y)
            canevas.create_rectangle(x - largeur / 2, y, x + largeur / 2, y + etat["hauteur"], fill = "white", outline = "gray70", tags = ("fond",))
            y += etat["hauteur"] + MARGE
        canevas.configure(scrollregion = (0, 0, largeur_panneau, y))
        charger_pages_visibles()

    def charger_pages_visibles ():
        """ Charge les pages visibles (+1 de part et d'autre), libere les autres. """
        if etat["hauteur"] <= 0:
            return
        haut = canevas.canvasy(0)
        bas = haut + canevas.winfo_height()
        x = canevas.winfo_width() / 2
        visibles = [i for i, y in enumerate(etat["positions"]) if y + etat["hauteur"] >= haut - etat["hauteur"] and y <= bas + etat["hauteur"]]
        for numero in list(etat["images"]):
            if numero not in visibles:
                canevas.delete("page%d" % numero)
                del etat["images"][numero]
        for numero in visibles:
            if numero in etat["images"]:
                continue
            try:
                image_page = Image.open(pages[numero]).convert("RGB").resize((etat["largeur"], etat["hauteur"]), Image.LANCZOS)
                photo = ImageTk.PhotoImage(image_page)
            except Exception:
                continue
            etat["images"][numero] = photo
            canevas.create_image(x, etat["positions"][numero], image = photo, anchor = "n", tags = ("page%d" % numero,))
        ## numero de la page au centre de l'ecran
        milieu = (haut + bas) / 2
        courante = 1
        for i, y in enumerate(etat["positions"]):
            if y <= milieu:
                courante = i + 1
        label_page.configure(text = textes["page"] + " " + str(courante) + " " + textes["sur"] + " " + str(len(pages)))

    def sur_defilement (*arguments):
        defilement.set(*arguments)
        charger_pages_visibles()

    def sur_redimensionnement (event = None):
        if etat["tache"] is not None:
            try:
                canevas.after_cancel(etat["tache"])
            except Exception:
                pass
        etat["tache"] = canevas.after(120, mettre_en_page)      # attend la fin du redimensionnement

    def molette (event):
        if getattr(event, "num", None) == 4:
            pas = -3
        elif getattr(event, "num", None) == 5:
            pas = 3
        elif abs(event.delta) >= 120:
            pas = -int(event.delta / 120) * 3              # Windows
        else:
            pas = -int(event.delta)                        # macOS
        canevas.yview_scroll(pas, "units")
        return "break"

    canevas.configure(yscrollcommand = sur_defilement, yscrollincrement = 20)
    canevas.bind("<Configure>", sur_redimensionnement)
    for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
        canevas.bind(sequence, molette)
    ## clavier : fleches, page precedente / suivante, debut / fin
    canevas.bind("<Up>", lambda e: canevas.yview_scroll(-3, "units"))
    canevas.bind("<Down>", lambda e: canevas.yview_scroll(3, "units"))
    canevas.bind("<Prior>", lambda e: canevas.yview_scroll(-1, "pages"))
    canevas.bind("<Next>", lambda e: canevas.yview_scroll(1, "pages"))
    canevas.bind("<Home>", lambda e: canevas.yview_moveto(0))
    canevas.bind("<End>", lambda e: canevas.yview_moveto(1))
    canevas.bind("<Button-1>", lambda e: canevas.focus_set())
    canevas.after(50, lambda: canevas.focus_set())
    mettre_en_page()


_boutons_lies_au_projet = [
    (button_Analysis, Analysis, _icone_grisee('figures/new_analysis.png')),
    (button_Restore, Restore, _icone_grisee('figures/restore.png')),
    (button_Delete_model, Delete, _icone_grisee('figures/delete.png')),
    (button_Rename, Rename_model, _icone_grisee('figures/Rename_model.png')),
    (button_Settings, Settings, _icone_grisee('figures/settings.png')),
    (button_Csv_import, Csv_import, _icone_grisee('figures/csv_import.png')),
    (button_Undo, Undo, _icone_grisee('figures/undo_p.png')),
    (button_Redo, Redo, _icone_grisee('figures/redo_p.png')),
    (button_Model, Model, _icone_grisee('figures/model_p.png')),
    (button_Run, Run, _icone_grisee('figures/run_p.png')),
    (button_data_Edit, my_data_edit_p, _icone_grisee('figures/settings_p-32.png')),
    (button_Log, Log, _icone_grisee('figures/log_p.png')),
    (button_Results, Results, _icone_grisee('figures/results_p.png')),
    (button_Save, Save, _icone_grisee('figures/save_p.png')),
]

## Couleurs a l'etat desactive. Le code ne definissait que `disabledforeground` ; sans
## `disabledbackground`, tkmacosx affiche le fond clair par defaut du bouton desactive
## (un rectangle pale sur la barre sombre). On garde donc le fond normal du bouton.
FOND_BARRE_GAUCHE = "gray8"
FOND_BARRE_HAUT = "slate gray"
## Les boutons de la barre du haut portent un texte : leur `disabledforeground` etait
## 'slate gray', la couleur du fond de la barre -> le texte disparaissait une fois le
## bouton desactive. On lui donne une couleur attenuee mais lisible.
COULEUR_TEXTE_BOUTON_DESACTIVE = "gray65"


def _couleurs_etat_desactive(boutons, fond, texte = None):
    for bouton in boutons:
        try:
            bouton.configure(disabledbackground = fond)
        except Exception:
            pass
        if texte is not None:
            try:
                bouton.configure(disabledforeground = texte)
            except Exception:
                pass


_couleurs_etat_desactive((button_Analysis, button_Restore, button_Delete_model, button_Rename, button_Settings, button_Csv_import, button_arret_model),
                         FOND_BARRE_GAUCHE)

## Bouton d'arret : desactive (icone grisee) tant que « Run » n'a pas ete clique ;
## lancer_run l'active pendant l'execution puis le desactive a nouveau.
_icones_bouton_arret["normale"] = Arret
_icones_bouton_arret["grisee"] = _icone_grisee('figures/arret.png')
_etat_bouton_arret(False)
_couleurs_etat_desactive((button_Undo, button_Redo, button_Model, button_Run, button_data_Edit, button_Log, button_Results, button_Save),
                         FOND_BARRE_HAUT, COULEUR_TEXTE_BOUTON_DESACTIVE)

DELAI_MAJ_BOUTONS_PROJET_MS = 250
_etat_boutons_projet = {"actif": None}


def analyse_autorisee():
    """Vrai si un projet est ouvert (nouveau ou existant) et si des donnees ont ete importees."""
    try:
        ## (pas de bool() : `from numpy import *` le remplace par numpy.bool)
        donnees_importees = isinstance(file_path, str) and file_path != "" and (data is not None or os.path.isfile(file_path))
    except Exception:
        donnees_importees = False
    return nom_projet is not None and donnees_importees


def maj_etat_boutons_projet():
    """Active / desactive les boutons de _boutons_lies_au_projet (et grise leur
    icone) quand la situation change, puis se replanifie.
    Undo / Redo sont en plus desactives quand il n'y a rien a annuler / refaire."""
    try:
        actif = analyse_autorisee()
        peut_annuler = actif and _historique_peut_annuler()
        peut_refaire = actif and _historique_peut_refaire()
        signature = (actif, peut_annuler, peut_refaire)
        if signature != _etat_boutons_projet["actif"]:
            _etat_boutons_projet["actif"] = signature
            for bouton, icone_normale, icone_grisee in _boutons_lies_au_projet:
                if bouton is button_Undo:
                    actif_bouton = peut_annuler
                elif bouton is button_Redo:
                    actif_bouton = peut_refaire
                else:
                    actif_bouton = actif
                try:
                    bouton.configure(state = "normal" if actif_bouton else "disabled")
                except Exception:
                    pass
                if icone_grisee is not None:
                    try:
                        bouton.configure(image = icone_normale if actif_bouton else icone_grisee)
                    except Exception:
                        pass
    finally:
        _planifier_tache_application("maj_boutons", DELAI_MAJ_BOUTONS_PROJET_MS, maj_etat_boutons_projet)


maj_etat_boutons_projet()



# ==========================================================================
# ==  27. Lancement                                                       ==
# ==========================================================================
def main():
    """Installe les gestionnaires de fermeture et du clavier, puis lance la boucle Tk."""
    installer_fermeture_propre()

    installer_touche_entree_boites_dialogue()        # touche Entree dans toutes les boites de dialogue
    SmartDigicat.after(500, _surveiller_indicateur_execution)     # l'encadre de progression ne peut pas survivre a une execution
    SmartDigicat.mainloop()


if __name__ == "__main__":
    main()
