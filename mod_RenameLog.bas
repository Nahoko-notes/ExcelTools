Sub RenameLog()
    Dim selectedFiles As Variant
    Dim fso As Object
    Dim i As Long
    Dim filePath As String, folderPath As String, oldName As String, newName As String, newPath As String
    Dim wsLog As Worksheet
    Dim lastRow As Long
    
    ' NG文字と置換ルールの定義
    Dim ngChars As Variant, replaceWith As Variant
    ngChars = Array("\", "/", ":", "*", "?", "\"" , "<", ">", "|", "　", " ", vbLf, vbCr, "'", Chr(160))
    replaceWith = Array("", "", "", "", "", "", "", "", "", "_", "_", "", "", "", "")

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' ファイル選択（複数可）
    selectedFiles = Application.GetOpenFilename( _
        FileFilter:="Excelファイル (*.xls;*.xlsx;*.xlsm;*.xlsb), *.xls;*.xlsx;*.xlsm;*.xlsb", _
        Title:="リネーム対象のファイルを選んでください（複数選択OK）", MultiSelect:=True)

    ' キャンセル時
    If VarType(selectedFiles) = vbBoolean And selectedFiles = False Then
        MsgBox "キャンセルされました。", vbInformation
        Exit Sub
    End If

    ' ログシートの準備
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("Rename_Log")
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
        wsLog.Name = "Rename_Log"
        wsLog.Range("A1:D1").Value = Array("日時", "旧ファイル名", "新ファイル名", "変更されたか")
    End If
    On Error GoTo 0

    ' 処理開始
    For i = LBound(selectedFiles) To UBound(selectedFiles)
        filePath = selectedFiles(i)
        folderPath = fso.GetParentFolderName(filePath)
        oldName = fso.GetFileName(filePath)
        newName = oldName

        ' NG文字置換
        Dim j As Long
        For j = LBound(ngChars) To UBound(ngChars)
            newName = Replace(newName, ngChars(j), replaceWith(j))
        Next j

        newPath = folderPath & "\" & newName

        If filePath <> newPath Then
            ' 上書き確認省略 → 即リネーム
            Name filePath As newPath
            MsgBox "リネームしました：" & vbCrLf & oldName & " → " & newName, vbInformation
        End If

        ' ログ記録
        lastRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
        wsLog.Cells(lastRow, 1).Value = Now
        wsLog.Cells(lastRow, 2).Value = oldName
        wsLog.Cells(lastRow, 3).Value = newName
        wsLog.Cells(lastRow, 4).Value = IIf(oldName <> newName, "✔", "")
    Next i

    MsgBox "すべてのファイルの処理が完了しました！", vbInformation
End Sub
