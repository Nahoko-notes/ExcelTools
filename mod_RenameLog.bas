Sub RenameLog()
    Dim fileNames As Variant
    Dim i As Long, j As Long
    Dim oldPath As String, newPath As String
    Dim fileNameOnly As String, folderPath As String
    Dim logSheet As Worksheet
    Dim timeStamp As String

    ' NG文字と置換先定義
    Dim ngChars As Variant
    Dim replaceWith As Variant

    ngChars = Array("/", "\", ":", "*", "?", Chr(34), "<", ">", "|", "　", " ", vbLf, vbCr, Chr(39), Chr(160))
    replaceWith = Array("_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "", "", "_", "_")

    ' ファイル選択
    fileNames = Application.GetOpenFilename("Excelファイル (*.xls*), *.xls*", , "リネーム対象のファイルを選択（複数OK）", , True)

    ' キャンセル時
    If VarType(fileNames) = vbBoolean And fileNames = False Then
        MsgBox "キャンセルされました", vbExclamation
        Exit Sub
    End If

    ' ログ用シートを用意（無ければ作成）
    On Error Resume Next
    Set logSheet = ThisWorkbook.Sheets("リネームログ")
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
        logSheet.Name = "リネームログ"
        logSheet.Range("A1:D1").Value = Array("旧ファイル名", "新ファイル名", "パス", "タイムスタンプ")
    End If
    On Error GoTo 0

    timeStamp = Format(Now, "yyyy/mm/dd HH:MM:SS")

    ' 複数ファイルループ
    For i = LBound(fileNames) To UBound(fileNames)
        oldPath = fileNames(i)
        folderPath = Left(oldPath, InStrRev(oldPath, "\"))
        fileNameOnly = Mid(oldPath, InStrRev(oldPath, "\") + 1)
        newPath = fileNameOnly

        ' NG文字置換
        For j = LBound(ngChars) To UBound(ngChars)
            newPath = Replace(newPath, ngChars(j), replaceWith(j))
        Next j

        ' 完全パスに
        newPath = folderPath & newPath

        ' 同名ファイルの存在チェック
        If Dir(newPath) <> "" Then
            MsgBox "同名ファイルが既にあります：" & vbCrLf & newPath, vbCritical
        Else
            Name oldPath As newPath

            ' ログ書き込み
            With logSheet
                .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = fileNameOnly
                .Cells(.Rows.Count, 1).End(xlUp).Offset(0, 1).Value = Mid(newPath, InStrRev(newPath, "\") + 1)
                .Cells(.Rows.Count, 1).End(xlUp).Offset(0, 2).Value = newPath
                .Cells(.Rows.Count, 1).End(xlUp).Offset(0, 3).Value = timeStamp
            End With
        End If
    Next i

    MsgBox "リネーム完了しました！", vbInformation
End Sub
