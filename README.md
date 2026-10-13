<p align="center"><img src="smartdigicat/figures/icon.png" alt="SmartDigiCat logo" width="96"></p>

<h1 align="center">SmartDigiCat</h1>

**SmartDigiCat** is a desktop application (Python / Tkinter) for catalytic screening.
It brings together data cleaning, descriptive statistics, regression models,
prediction and Bayesian optimization in a single graphical interface.

Developed within the industrial chair of **Centrale Lille**
([project presentation](https://www.techniques-ingenieur.fr/actualite/livre-blanc/smartdigicat-un-projet-scientifique-reunissant-catalyse-sciences-des-donnees-et-ia-2-122489/)).

> *En français :* SmartDigiCat est un logiciel de criblage catalytique qui réunit
> nettoyage des données, statistiques descriptives, modèles de régression,
> prédiction et optimisation bayésienne. L'interface est disponible en français
> et en anglais (bouton « langue » de la barre de gauche).

## Features

- **Data**: import of CSV (`;`), text (`,`) and Excel files (several sheets with the
  same columns can be stacked); missing-value imputation, renaming, recoding of
  modalities, filtering on up to four parameters.
- **Descriptive analysis**: summary tables (exportable to Excel, Word or CSV),
  scatter plots, box plots with pairwise tests, histograms, 3D plots,
  correlation and min/mean/max bar plots.
- **Models**: several models per project — linear, Lasso, kernel ridge, SVR,
  random forest, gradient boosting, decision tree, Gaussian process — with
  train/test split or K-fold cross-validation (hyperparameter search) and
  sequential feature selection.
- **Prediction** on new data with the best trained model (highest R²).
- **Bayesian optimization** (Matérn Gaussian process, MCMC + Expected
  Improvement) to propose the next experimental points, with a credibility band.
- **Projects** saved as JSON (models, settings, filtered data, charts),
  undo/redo, automatic save every 5 minutes, interruptible computations.

## Installation

### Installers (recommended, no internet needed)

Ready-to-use installers for Windows and macOS are available in the
[Download](#download) section below. They work offline and also provide an
uninstaller.

### From source (developers)

Python ≥ 3.9 with Tkinter is required.

```bash
git clone https://github.com/noukpoape/SmartDigiCat.git
cd SmartDigiCat
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

The user manual (French and English, PDF) is in
[`smartdigicat/figures/manuel/`](smartdigicat/figures/manuel/); it can also be
opened from the application with the **Help** button.

The interface uses `tkmacosx` and was developed on macOS.

## Usage

```bash
python -m smartdigicat
# or
python run_smartdigicat.py
```

Typical workflow: **New** (create a project) → **Import data** → **Edit data**
(clean / filter) → **Plots** (descriptive analysis, optimization) →
**Analysis** (create a model) → **Model** (configure it) → **Run** (train) →
**Log** (predict on new data). **Help** opens the user manual.

The calculation module can also be used without the interface:

```python
import pandas as pd
from smartdigicat.core.bayesian_optimization import Yields_optimization

x = pd.Series([0.5, 1.5, 2.5, 3.5, 4.5], name="temperature")
y = pd.Series([1.0, 5.5, 9.0, 9.5, 7.0], name="yield")
points, duration, history = Yields_optimization(
    y, 0.0, 20.0, x, 0.0, 5.0, n_simul_1=100, n_simul_2=10, n_replication=10, random_states=42)
```

Reference settings: [`docs/optimization_parameters.md`](docs/optimization_parameters.md).

## Repository layout

```
SmartDigiCat/
├── smartdigicat/
│   ├── app.py                     # graphical interface (Tkinter), numbered sections
│   ├── translations.py            # French / English interface texts
│   ├── core/
│   │   └── bayesian_optimization.py   # Bayesian optimization, no GUI dependency
│   ├── figures/                   # icons and user manual (PDF + pages)
│   ├── __init__.py
│   └── __main__.py                # python -m smartdigicat
├── installer/                     # PyInstaller recipe, Windows and macOS installers
├── tests/                         # tests of the calculation module
├── docs/
├── INSTALL.md                     # install / uninstall guide for users
├── run_smartdigicat.py            # launcher (also usable with PyInstaller)
├── requirements.txt               # full dependencies
├── requirements-core.txt          # calculation module + tests only
└── pyproject.toml
```

The interface shares its state between callbacks through module-level variables,
so it stays in one module (`app.py`), organized in numbered sections listed at the
top of the file.

## Tests

```bash
pip install -r requirements-core.txt
python -m pytest
```

The tests cover the calculation module only and run without a display; they are
run automatically by GitHub Actions on every push.

## Building the installers

Installers are built by GitHub Actions when a version tag is pushed
(`git tag v1.0.0 && git push origin v1.0.0`). Details and manual build:
[`installer/README.md`](installer/README.md).

## Download

Installers of the latest version (Python and all libraries included: **no
internet connection is needed to install or use SmartDigiCat**):

| System | Installer |
|---|---|
| Windows 10 / 11 (64-bit) | [**SmartDigiCat-Windows-Setup.exe**](https://github.com/noukpoape/SmartDigiCat/releases/latest/download/SmartDigiCat-Windows-Setup.exe) |
| macOS — Apple silicon (M1, M2, M3…) | [**SmartDigiCat-macOS-AppleSilicon.pkg**](https://github.com/noukpoape/SmartDigiCat/releases/latest/download/SmartDigiCat-macOS-AppleSilicon.pkg) |
| macOS — Intel | [**SmartDigiCat-macOS-Intel.pkg**](https://github.com/noukpoape/SmartDigiCat/releases/latest/download/SmartDigiCat-macOS-Intel.pkg) |

To install, double-click the downloaded file: a classic step-by-step installation
wizard opens. To uninstall: on Windows, use **Settings → Apps → SmartDigiCat →
Uninstall**; on macOS, double-click **Applications → SmartDigiCat → Désinstaller
SmartDigiCat**. Uninstalling never deletes your
projects or data. Step-by-step guide (French / English): [INSTALL.md](INSTALL.md).
Previous versions: [Releases](https://github.com/noukpoape/SmartDigiCat/releases).

## Authors

Sébastien Paul, Sophie Dabo-Niang, Komlan Noukpoape, Wahba Saweiras, Svetlana Heyte —
Centrale Lille / Université de Lille.

## License

To be defined.
