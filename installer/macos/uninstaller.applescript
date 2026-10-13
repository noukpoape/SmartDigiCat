-- Desinstalleur macOS de SmartDigiCat (installe dans /Applications/SmartDigiCat).
-- Supprime le dossier /Applications/SmartDigiCat (logiciel + ce desinstalleur) et
-- l'enregistrement du paquet. Les projets et les donnees de l'utilisateur (fichiers
-- .json, .csv... enregistres ailleurs) ne sont jamais touches.

on texte(fr, en)
	if (user locale of (system info)) starts with "fr" then return fr
	return en
end texte

set dossier to "/Applications/SmartDigiCat"
set titre to "SmartDigiCat"

set installe to false
try
	do shell script "test -d " & quoted form of dossier
	set installe to true
end try

if not installe then
	display dialog texte("SmartDigiCat n'est pas installé sur ce Mac.", "SmartDigiCat is not installed on this Mac.") buttons {"OK"} default button "OK" with title titre with icon note
	return
end if

set boutonDesinstaller to texte("Désinstaller", "Uninstall")
set boutonAnnuler to texte("Annuler", "Cancel")
display dialog texte("Désinstaller SmartDigiCat de ce Mac ?" & return & return & "Vos projets et vos données ne seront pas supprimés.", "Uninstall SmartDigiCat from this Mac?" & return & return & "Your projects and data will not be deleted.") buttons {boutonAnnuler, boutonDesinstaller} default button boutonAnnuler cancel button boutonAnnuler with title titre with icon caution

try
	do shell script "pkill -x SmartDigiCat"
end try

-- le logiciel a ete installe par le programme d'installation d'Apple : le mot de passe
-- d'un administrateur est demande pour le supprimer
set commande to "rm -rf " & quoted form of dossier & " && { pkgutil --forget fr.centralelille.smartdigicat >/dev/null 2>&1; true; } && ! test -e " & quoted form of dossier
try
	do shell script commande with administrator privileges
on error messageErreur number numeroErreur
	if numeroErreur is -128 then return -- mot de passe annule
	display dialog texte("La désinstallation a échoué :", "Uninstallation failed:") & return & messageErreur buttons {"OK"} default button "OK" with title titre with icon stop
	return
end try

display dialog texte("SmartDigiCat a été désinstallé.", "SmartDigiCat has been uninstalled.") buttons {"OK"} default button "OK" with title titre with icon note
