# Installer SmartDigiCat / Installing SmartDigiCat

*English version below.*

## Français

L'installation **ne demande aucune connexion internet** : l'installeur contient le
logiciel et tout ce dont il a besoin. Téléchargez le bon fichier une fois, depuis la
page [Releases](https://github.com/noukpoape/SmartDigiCat/releases/latest), puis
copiez-le sur n'importe quel ordinateur (clé USB, réseau…).

| Système | Fichier à télécharger |
|---|---|
| Windows 10 ou 11 (64 bits) | `SmartDigiCat-Windows-Setup.exe` |
| Mac avec puce Apple (M1, M2, M3…) | `SmartDigiCat-macOS-AppleSilicon.pkg` |
| Mac avec processeur Intel | `SmartDigiCat-macOS-Intel.pkg` |

Pour savoir quel Mac vous avez : menu  → **À propos de ce Mac** (ligne « Puce » = Apple,
« Processeur » = Intel).

### Windows

**Installer** : double-cliquez sur `SmartDigiCat-Windows-Setup.exe`.

Si Windows affiche « Windows a protégé votre ordinateur », cliquez sur
**Informations complémentaires** puis **Exécuter quand même** (le logiciel n'est pas
signé par un certificat payant). L'assistant enchaîne ensuite les étapes suivantes :

1. **Mode d'installation** : pour vous seul (sans droits administrateur) ou pour tous
   les utilisateurs de l'ordinateur ;
2. **Bienvenue** ;
3. **Informations** : présentation du logiciel ;
4. **Dossier de destination** : à garder tel quel en général (un dossier vide ou
   nouveau est exigé) ;
5. **Dossier du menu Démarrer** ;
6. **Tâches supplémentaires** : créer une icône sur le Bureau ;
7. **Prêt à installer** : récapitulatif, puis **Installer** ;
8. **Fin de l'installation** : case pour lancer SmartDigiCat tout de suite.

Les fichiers du logiciel sont placés dans un dossier masqué ; SmartDigiCat se lance
depuis le menu Démarrer ou l'icône du Bureau.

**Désinstaller** : menu Démarrer → SmartDigiCat → **Désinstaller SmartDigiCat**, ou
**Paramètres → Applications → SmartDigiCat → Désinstaller**. Une boîte de
confirmation s'affiche, puis la progression et un message de fin.

### macOS

**Installer** : double-cliquez sur le fichier `.pkg`. Le programme d'installation
d'Apple s'ouvre et enchaîne les étapes **Introduction → Lisez-moi → Type
d'installation → Installation → Résumé**. Votre mot de passe vous est demandé à
l'étape Installation.

La première fois, macOS peut indiquer que le paquet provient d'un développeur non
identifié (il n'est pas signé par Apple) :
- macOS 15 (Sequoia) et plus récent : fermez le message, ouvrez **Réglages Système →
  Confidentialité et sécurité**, descendez jusqu'au message concernant le paquet
  SmartDigiCat et cliquez sur **Ouvrir quand même** ;
- macOS 14 et antérieur : clic droit (ou Ctrl-clic) sur le fichier `.pkg` → **Ouvrir**
  → **Ouvrir**.

Le dossier **Applications → SmartDigiCat** contient alors deux icônes :
**SmartDigiCat** (le logiciel, dont tous les fichiers sont regroupés et masqués à
l'intérieur) et **Désinstaller SmartDigiCat**. Vous pouvez supprimer le fichier `.pkg`.

**Désinstaller** : Applications → SmartDigiCat → double-cliquez sur
**Désinstaller SmartDigiCat**, confirmez et saisissez votre mot de passe.

### Vos projets

La désinstallation ne supprime jamais vos projets ni vos données (fichiers `.json`,
`.csv`, graphiques…). Pour mettre à jour SmartDigiCat, lancez simplement l'installeur
de la nouvelle version : l'ancienne est remplacée.

---

## English

Installation **needs no internet connection**: the installer contains the software
and everything it needs. Download the right file once from the
[Releases](https://github.com/noukpoape/SmartDigiCat/releases/latest) page, then copy
it to any computer.

| System | File |
|---|---|
| Windows 10 or 11 (64-bit) | `SmartDigiCat-Windows-Setup.exe` |
| Mac with Apple silicon (M1, M2, M3…) | `SmartDigiCat-macOS-AppleSilicon.pkg` |
| Mac with Intel processor | `SmartDigiCat-macOS-Intel.pkg` |

### Windows

**Install**: double-click `SmartDigiCat-Windows-Setup.exe` (if Windows shows "Windows
protected your PC", click **More info** → **Run anyway**). The wizard then goes
through: install mode (just me / all users) → Welcome → Information → Destination
folder → Start menu folder → Additional tasks (desktop icon) → Ready to install →
Finish. The program files are installed in a hidden folder.

**Uninstall**: Start menu → SmartDigiCat → **Uninstall SmartDigiCat**, or **Settings →
Apps → SmartDigiCat → Uninstall**.

### macOS

**Install**: double-click the `.pkg` file. Apple's Installer goes through
**Introduction → Read Me → Installation Type → Installation → Summary**. If macOS
says the package is from an unidentified developer: on macOS 15 or later, open
**System Settings → Privacy & Security** and click **Open Anyway**; on macOS 14 or
earlier, right-click the `.pkg` → **Open** → **Open**.

**Applications → SmartDigiCat** then contains **SmartDigiCat** (all its files are
hidden inside the app) and **Désinstaller SmartDigiCat**.

**Uninstall**: double-click **Applications → SmartDigiCat → Désinstaller SmartDigiCat**.

Uninstalling never deletes your projects or data.
