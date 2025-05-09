Option Explicit

' ==== CONFIGURATION ====
Const DEFAULT_SCRIPT_NAME = "solo.vbs"
Const DEFAULT_ICON_NAME   = "solo.ico"
' =======================

Dim objShell, fso
Set objShell = CreateObject("WScript.Shell")
Set fso      = CreateObject("Scripting.FileSystemObject")

Dim desktopPath, basePath, scriptName, iconPath
Dim folder, scriptPath, shortcut, lnkPath, countCreated

'--- Nom du script ciblé ---
scriptName = DEFAULT_SCRIPT_NAME

'--- Chemins ---
desktopPath = objShell.SpecialFolders("Desktop")
basePath    = fso.GetAbsolutePathName(".")
iconPath    = fso.BuildPath(basePath, DEFAULT_ICON_NAME)

'--- Vérification de l'icône ---
If Not fso.FileExists(iconPath) Then
    MsgBox "Erreur : icône introuvable : " & iconPath, vbCritical, "Shortcut Creator"
    WScript.Quit 1
End If

'--- Création des raccourcis ---
countCreated = 0
On Error Resume Next
For Each folder In fso.GetFolder(basePath).SubFolders
    scriptPath = fso.BuildPath(folder.Path, scriptName)
    If fso.FileExists(scriptPath) Then
        lnkPath = fso.BuildPath(desktopPath, folder.Name & ".lnk")
        Set shortcut = objShell.CreateShortcut(lnkPath)
        shortcut.TargetPath       = scriptPath
        shortcut.WorkingDirectory = folder.Path
        shortcut.IconLocation     = iconPath
        shortcut.Save
        If Err.Number = 0 Then
            countCreated = countCreated + 1
        Else
            ' En cas d'erreur, vous pouvez logguer Err.Description si besoin
            Err.Clear
        End If
    End If
Next
On Error GoTo 0

'--- Rapport final ---
MsgBox "Raccourcis créés (ou mis à jour) : " & countCreated, vbInformation, "Terminé"
