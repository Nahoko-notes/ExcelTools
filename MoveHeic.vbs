' ================================================
'   [Created]: 2025-08-03
'   [Purpose]: HEIC file organizer
'   Scripted with love by Nahoko & Rodem XOXO
'   Reliable | Practical | a bit cheeky
'   Transforming bits into art, day and night!
' ================================================

Option Explicit

Dim fso, currentFolder, heicFolder, file
Dim baseName, ext, i, targetPath

Set fso = CreateObject("Scripting.FileSystemObject")

' カレントフォルダ
currentFolder = fso.GetParentFolderName(WScript.ScriptFullName)

' heicフォルダのパス
heicFolder = currentFolder & "\heic"

' heicフォルダがなければ作成
If Not fso.FolderExists(heicFolder) Then
    fso.CreateFolder(heicFolder)
End If

' フォルダ内のファイルを走査
For Each file In fso.GetFolder(currentFolder).Files
    If LCase(fso.GetExtensionName(file.Name)) = "heic" Then
        baseName = fso.GetBaseName(file.Name)
        ext = fso.GetExtensionName(file.Name)
        targetPath = heicFolder & "\" & file.Name

        ' 同名ファイルがあればリネームして保存
        i = 1
        Do While fso.FileExists(targetPath)
            targetPath = heicFolder & "\" & baseName & "_" & i & "." & ext
            i = i + 1
        Loop

        fso.MoveFile file.Path, targetPath
    End If
Next

MsgBox "すべての .heic ファイルを 'heic' フォルダに移動しました。", vbInformation, "完了"
