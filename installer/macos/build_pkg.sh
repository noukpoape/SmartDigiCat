#!/bin/bash
# Construit l'installeur macOS de SmartDigiCat : un paquet .pkg qui s'ouvre dans le
# programme d'installation d'Apple (assistant classique a etapes successives).
# A lancer sur un Mac, depuis la racine du depot, APRES PyInstaller (dist/SmartDigiCat.app).
#
#   bash installer/macos/build_pkg.sh 1.0.0
#
# Le paquet installe dans /Applications/SmartDigiCat :
#   SmartDigiCat.app               (le logiciel ; tous ses fichiers sont masques dans le bundle)
#   Désinstaller SmartDigiCat.app  (supprime le logiciel)
set -euo pipefail

VERSION="${1:-1.0.0}"
case "$(uname -m)" in                      # nom fixe (sans version) : lien stable dans le README
    arm64) PUCE="AppleSilicon"; ARCH="arm64" ;;
    *)     PUCE="Intel";        ARCH="x86_64" ;;
esac
IDENTIFIANT="fr.centralelille.smartdigicat"
RACINE="$(cd "$(dirname "$0")/../.." && pwd)"
ICI="$RACINE/installer/macos"
APP="$RACINE/dist/SmartDigiCat.app"
ICONE="$RACINE/installer/icons/smartdigicat.icns"
TRAVAIL="$RACINE/build/pkg"
CONTENU="$TRAVAIL/racine/SmartDigiCat"     # devient /Applications/SmartDigiCat
DESINSTALLEUR="$CONTENU/Désinstaller SmartDigiCat.app"
PKG="$RACINE/dist/SmartDigiCat-macOS-$PUCE.pkg"

test -d "$APP" || { echo "dist/SmartDigiCat.app introuvable : lancez d'abord PyInstaller." >&2; exit 1; }

rm -rf "$TRAVAIL" "$PKG"
mkdir -p "$CONTENU"

# 1) le logiciel, signe ad hoc (obligatoire sur Apple Silicon)
codesign --force --deep --sign - "$APP"
ditto "$APP" "$CONTENU/SmartDigiCat.app"

# 2) le desinstalleur (AppleScript), avec l'icone SmartDigiCat
osacompile -o "$DESINSTALLEUR" "$ICI/uninstaller.applescript"
cp "$ICONE" "$DESINSTALLEUR/Contents/Resources/applet.icns"
rm -f "$DESINSTALLEUR/Contents/Resources/Assets.car"
plutil -remove CFBundleIconName "$DESINSTALLEUR/Contents/Info.plist" 2>/dev/null || true
codesign --force --sign - "$DESINSTALLEUR"

# 3) paquet de composants : les applications sont installees exactement a leur place
#    (pas de « relocalisation » vers une autre copie trouvee sur le disque)
pkgbuild --analyze --root "$TRAVAIL/racine" "$TRAVAIL/composants.plist"
i=0
while /usr/libexec/PlistBuddy -c "Print :$i" "$TRAVAIL/composants.plist" >/dev/null 2>&1; do
    /usr/libexec/PlistBuddy -c "Set :$i:BundleIsRelocatable false" "$TRAVAIL/composants.plist"
    i=$((i + 1))
done
chmod +x "$ICI/scripts/preinstall"
pkgbuild --root "$TRAVAIL/racine" \
         --component-plist "$TRAVAIL/composants.plist" \
         --identifier "$IDENTIFIANT" \
         --version "$VERSION" \
         --install-location /Applications \
         --scripts "$ICI/scripts" \
         "$TRAVAIL/SmartDigiCat-component.pkg"

# 4) assistant d'installation (pages Bienvenue / Lisez-moi / Conclusion en francais et anglais)
sed -e "s/__VERSION__/$VERSION/g" -e "s/__ARCH__/$ARCH/g" "$ICI/distribution.xml" > "$TRAVAIL/distribution.xml"
productbuild --distribution "$TRAVAIL/distribution.xml" \
             --resources "$ICI/resources" \
             --package-path "$TRAVAIL" \
             "$PKG"
echo "Installeur cree : $PKG"
