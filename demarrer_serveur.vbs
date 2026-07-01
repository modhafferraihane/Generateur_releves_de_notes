Dim WshShell, oExec
Set WshShell = CreateObject("WScript.Shell")

' Changer vers le dossier du script
WshShell.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' Lancer le serveur Flask en arriere-plan (sans fenetre)
WshShell.Run "python app.py", 0, False

' Attendre 2 secondes pour que le serveur demarre
WScript.Sleep 2000

' Ouvrir le navigateur
WshShell.Run "http://127.0.0.1:5000", 1, False

WScript.Echo "Serveur demarre ! Accedez a http://127.0.0.1:5000" & vbCrLf & "Pour l'arreter, double-cliquez sur arreter_serveur.bat"
