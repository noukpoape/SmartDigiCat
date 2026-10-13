# Installeurs SmartDigiCat (note pour les mainteneurs)

Les installeurs sont **construits automatiquement par GitHub** : rien à installer
sur votre ordinateur.

## Publier une nouvelle version

1. Mettez à jour le numéro de version dans `pyproject.toml` et
   `smartdigicat/__init__.py`.
2. Envoyez les modifications, puis créez l'étiquette de version :

   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```

3. Le workflow `release` (onglet **Actions**) construit en 15 à 30 minutes :
   - `SmartDigiCat-Windows-Setup.exe`
   - `SmartDigiCat-macOS-AppleSilicon.pkg` (Mac Apple Silicon)
   - `SmartDigiCat-macOS-Intel.pkg` (Mac Intel)

   et les attache à la page **Releases** du dépôt. Les noms ne contiennent pas
   le numéro de version : les liens de la section « Download » du README
   pointent ainsi toujours vers la dernière version.

Pour un essai sans publier : **Actions → release → Run workflow** ; les
installeurs sont alors dans les *Artifacts* de l'exécution.

## Contenu de ce dossier

| Fichier | Rôle |
|---|---|
| `smartdigicat.spec` | Recette PyInstaller : Python + bibliothèques + icônes + manuel en un programme autonome |
| `windows/SmartDigiCat.iss` | Assistant Windows (Inno Setup) : Bienvenue, Informations, Dossier, Menu Démarrer, Tâches, Prêt, Fin ; désinstallation ; fichiers masqués |
| `windows/info_fr.txt`, `info_en.txt` | Texte de la page « Informations » de l'assistant Windows |
| `macos/build_pkg.sh` | Construit le paquet `.pkg` (assistant du programme d'installation d'Apple) |
| `macos/distribution.xml` | Étapes de l'assistant macOS (Introduction, Lisez-moi, Résumé), architecture |
| `macos/resources/` | Pages de l'assistant macOS en français (`fr.lproj`) et en anglais (`en.lproj`), logo |
| `macos/scripts/preinstall` | Retire l'ancienne version avant une mise à jour |
| `macos/uninstaller.applescript` | Application « Désinstaller SmartDigiCat », installée avec le logiciel |
| `icons/` | Icônes `.ico` (Windows) et `.icns` (macOS), tirées de `figures/icon.png` |

## Construire à la main

Windows (avec [Inno Setup 6](https://jrsoftware.org/isinfo.php)) :

```bat
pip install -r requirements.txt pyinstaller
pyinstaller installer\smartdigicat.spec --noconfirm
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DMyAppVersion=1.0.0 installer\windows\SmartDigiCat.iss
```

macOS :

```bash
pip install -r requirements.txt pyinstaller
pyinstaller installer/smartdigicat.spec --noconfirm
bash installer/macos/build_pkg.sh 1.0.0
```

Les fichiers produits sont dans `dist/`.

## Signature

Les installeurs ne sont pas signés : Windows (SmartScreen) et macOS (Gatekeeper)
affichent un avertissement au premier lancement (voir `INSTALL.md`). Pour le
supprimer, il faut un certificat de signature de code (Windows) et un compte
Apple Developer avec notarisation (macOS), tous deux payants.
