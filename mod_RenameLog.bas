option explicit
Sub RenameLog()
    Dim fileNames As Variant
    Dim i As Long, j As Long
    Dim oldPath As String, newPath As String
    Dim fileNameOnly As String, newFileNameOnly As String, folderPath As String
    Dim logSheet As Worksheet, defSheet As Worksheet
    Dim timeStamp As String
    Dim cell As Range

    ' NG文字と置換先定義
    Dim ngChars As Variant
    Dim replaceWith As Variant

    ngChars = Array("/", "\", ":", "*", "?", Chr(34), "<", ">", "|", "　", " ", vbLf, vbCr, Chr(39), Chr(160))
    replaceWith = Array("_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "", "", "_", "_")

    ' ファイル選択
    fileNames = Application.GetOpenFilename("Excelファイル (*.xls*), *.xls*", , "リネーム対象のファイルを選択（複数OK）", , True)

    ' キャンセル時の安全なチェック
    If Not IsArray(fileNames) Then
        If fileNames = False Then
            MsgBox "キャンセルされました", vbExclamation
            Exit Sub
        End If
    End If

    ' ログ用シート
    On Error Resume Next
    Set logSheet = ThisWorkbook.Sheets("リネームログ")
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
        logSheet.Name = "リネームログ"
        logSheet.Range("A1:D1").Value = Array("旧ファイル名", "新ファイル名", "パス", "タイムスタンプ")
    End If
    On Error GoTo 0

    ' 要件定義シート
    Set defSheet = ThisWorkbook.Sheets("仕様_要件定義")

    timeStamp = Format(Now, "yyyy/mm/dd HH:MM:SS")

    ' 複数ファイルループ
    For i = LBound(fileNames) To UBound(fileNames)
        oldPath = fileNames(i)
        folderPath = Left(oldPath, InStrRev(oldPath, "\"))
        fileNameOnly = Mid(oldPath, InStrRev(oldPath, "\") + 1)
        newFileNameOnly = fileNameOnly ' 初期値としてセット

        ' NG文字置換
        For j = LBound(ngChars) To UBound(ngChars)
            newFileNameOnly = Replace(newFileNameOnly, ngChars(j), replaceWith(j))
        Next j

        newPath = folderPath & newFileNameOnly

        ' 同名ファイルの存在チェック
        If Dir(newPath) <> "" Then
            MsgBox "同名ファイルが既にあります：" & vbCrLf & newPath, vbCritical
        Else
            Name oldPath As newPath

            ' リネームログ書き込み
            With logSheet
                .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = fileNameOnly
                .Cells(.Rows.Count, 1).End(xlUp).Offset(0, 1).Value = newFileNameOnly
                .Cells(.Rows.Count, 1).End(xlUp).Offset(0, 2).Value = newPath
                .Cells(.Rows.Count, 1).End(xlUp).Offset(0, 3).Value = timeStamp
            End With

            ' 要件定義シートのブック名欄を更新（L5〜L10想定）
            For Each cell In defSheet.Range("L5:L10")
                If Trim(cell.Value) = fileNameOnly Then
                    cell.Value = newFileNameOnly
                    Exit For
                End If
            Next cell
        End If
    Next i

    MsgBox "リネーム＋定義欄更新 完了しました！", vbInformation
End Sub
