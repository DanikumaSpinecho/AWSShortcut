Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Chemin vers le bureau
desktopPath = objShell.SpecialFolders("Desktop")

' Chemin du dossier courant
currentPath = fso.GetAbsolutePathName(".")

' Chemin vers le dossier contenant les exécutables
awsClientPath = currentPath

' Chemin vers l'icône
iconPath = currentPath & "\solo.ico"

' Vérifiez si l'icône existe
If Not fso.FileExists(iconPath) Then
    MsgBox "L'icône solo.ico n'existe pas à l'emplacement spécifié: " & iconPath, vbExclamation
    WScript.Quit
End If

' Parcourir les sous-dossiers de AWSCLIENT
For Each folder In fso.GetFolder(awsClientPath).SubFolders
    ' Vérifier si le dossier contient un exécutable solo.exe
    exePath = fso.BuildPath(folder.Path, "solo.exe")
    If fso.FileExists(exePath) Then
        ' Créer le raccourci avec le nom du dossier, le répertoire d'exécution et l'icône
        shortcutName = folder.Name & ".lnk"
        Set shortcut = objShell.CreateShortcut(desktopPath & "\" & shortcutName)
        shortcut.TargetPath = exePath
        shortcut.WorkingDirectory = folder.Path  ' Répertoire de travail mis à jour
        shortcut.IconLocation = iconPath
        shortcut.Save
    End If
Next

MsgBox "Les raccourcis ont été créés sur le bureau avec succès!", vbInformation
