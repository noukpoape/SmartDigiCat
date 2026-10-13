# -*- mode: python ; coding: utf-8 -*-
"""
Recette PyInstaller : construit SmartDigiCat en programme autonome (Python et
toutes les bibliotheques embarques). Aucune connexion internet n'est necessaire
pour l'installer ni pour l'utiliser.

Lancement (depuis la racine du depot) :
    pyinstaller installer/smartdigicat.spec --noconfirm

Resultat :
    Windows : dist/SmartDigiCat/SmartDigiCat.exe (+ dossier _internal)
    macOS   : dist/SmartDigiCat.app
"""

import os
import sys

from PyInstaller.utils.hooks import collect_all

RACINE = os.path.abspath(os.path.join(SPECPATH, ".."))
ICONES = os.path.join(RACINE, "installer", "icons")
NOM = "SmartDigiCat"
VERSION = os.environ.get("SMARTDIGICAT_VERSION", "1.0.0")

# Icones et manuel : my_path() les cherche dans <bundle>/figures
datas = [(os.path.join(RACINE, "smartdigicat", "figures"), "figures")]
binaries = []
hiddenimports = [
    "smartdigicat.app",
    "smartdigicat.translations",
    "smartdigicat.core.bayesian_optimization",
    "matplotlib.backends.backend_tkagg",
    "PIL._tkinter_finder",
    "openpyxl",
    "docx",
    "fpdf",
    "mpmath",
]

# Bibliotheques qui chargent des fichiers de donnees ou des modules a la volee
for paquet in ("tkmacosx", "pandastable", "tableone", "plotnine", "mizani",
               "mlxtend", "statsmodels", "seaborn", "sklearn", "patsy"):
    try:
        d, b, h = collect_all(paquet)
    except Exception:
        continue
    datas += d
    binaries += b
    hiddenimports += h

a = Analysis(
    [os.path.join(RACINE, "run_smartdigicat.py")],
    pathex=[RACINE],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    excludes=["pytest", "IPython", "jupyter", "notebook", "PyQt5", "PyQt6", "PySide2", "PySide6"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=NOM,
    console=False,                      # pas de fenetre de terminal
    icon=os.path.join(ICONES, "smartdigicat.icns" if sys.platform == "darwin" else "smartdigicat.ico"),
)

coll = COLLECT(exe, a.binaries, a.datas, name=NOM)

if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name=NOM + ".app",
        icon=os.path.join(ICONES, "smartdigicat.icns"),
        bundle_identifier="fr.centralelille.smartdigicat",
        version=VERSION,
        info_plist={
            "CFBundleName": NOM,
            "CFBundleDisplayName": NOM,
            "CFBundleShortVersionString": VERSION,
            "CFBundleVersion": VERSION,
            "NSHighResolutionCapable": True,
            "LSMinimumSystemVersion": "11.0",
        },
    )
