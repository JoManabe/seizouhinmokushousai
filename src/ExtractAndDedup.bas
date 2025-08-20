Attribute VB_Name = "ExtractAndDedup"
Sub MultiFile_ExtractAndDedup_Full()

    Dim selectedFiles As Variant
    Dim iFile As Long
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wbResult As Workbook
    Dim wsResult As Worksheet
    Dim lastRow As Long
    Dim dataArr As Variant
    Dim outputArr() As Variant
    Dim rowOut As Long
    Dim r As Long
    Dim dict As Object
    Dim key As String
    Dim currentRow As Long
    
    ' --- ダイアログで対象ファイルを選択 ---
    MsgBox "処理したい Excel ファイルを選択してください", vbInformation
    selectedFiles = Application.GetOpenFilename( _
                    FileFilter:="Excel Files (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", _
                    MultiSelect:=True, _
                    Title:="抽出対象のExcelファイルを選択してください")
                    
    If Not IsArray(selectedFiles) Then Exit Sub
    
    ' --- 出力ブック（マクロ専用ブック） ---
    Set wbResult = ThisWorkbook
    Set wsResult = wbResult.Sheets(1)
    wsResult.Cells.Clear
    
    ' ヘッダー（4列）
    wsResult.Range("A1:D1").Value = Array("品目CD", "品名", "内部名", "製造場所CD")
    currentRow = 2
    
    ' --- 選択ファイルを順次処理 ---
    For iFile = LBound(selectedFiles) To UBound(selectedFiles)
        Set wbSource = Workbooks.Open(selectedFiles(iFile))
        Set wsSource = wbSource.Sheets(1)
        
        lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row
        dataArr = wsSource.Range("A1:O" & lastRow).Value
        
        ' 出力用配列確保（最大行数は十分余裕を持たせる）
        ReDim outputArr(1 To (UBound(dataArr, 1) - 2) * 2, 1 To 4)
        Set dict = CreateObject("Scripting.Dictionary")
        rowOut = 0
        
        ' --- 配列上で抽出・重複削除（1ファイルごと） ---
        For r = 3 To UBound(dataArr, 1)  ' 1～2行目はスキップ
            ' Dummy除外
            If InStr(1, CStr(dataArr(r, 4)), "Dummy", vbTextCompare) = 0 _
               And InStr(1, CStr(dataArr(r, 10)), "Dummy", vbTextCompare) = 0 Then
               
                ' D,E,F,I のキー
                key = CStr(dataArr(r, 4)) & "|" & CStr(dataArr(r, 5)) & "|" & CStr(dataArr(r, 6)) & "|" & CStr(dataArr(r, 9))
                If Not dict.Exists(key) Then
                    dict.Add key, 1
                    rowOut = rowOut + 1
                    outputArr(rowOut, 1) = dataArr(r, 4)
                    outputArr(rowOut, 2) = dataArr(r, 5)
                    outputArr(rowOut, 3) = dataArr(r, 6)
                    outputArr(rowOut, 4) = dataArr(r, 9)
                End If
                
                ' J,K,L,O のキー
                key = CStr(dataArr(r, 10)) & "|" & CStr(dataArr(r, 11)) & "|" & CStr(dataArr(r, 12)) & "|" & CStr(dataArr(r, 15))
                If Not dict.Exists(key) Then
                    dict.Add key, 1
                    rowOut = rowOut + 1
                    outputArr(rowOut, 1) = dataArr(r, 10)
                    outputArr(rowOut, 2) = dataArr(r, 11)
                    outputArr(rowOut, 3) = dataArr(r, 12)
                    outputArr(rowOut, 4) = dataArr(r, 15)
                End If
            End If
        Next r
        
        ' --- 出力シートに追加 ---
        If rowOut > 0 Then
            wsResult.Range("A" & currentRow).Resize(rowOut, 4).Value = outputArr
            currentRow = currentRow + rowOut
        End If
        
        wbSource.Close SaveChanges:=False
    Next iFile
    
    ' --- 保存（固定名） ---
    Dim savePath As String
    savePath = wbResult.Path & "\製造品目詳細一覧_work1.xlsm"
    
    Application.DisplayAlerts = False
    wbResult.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    
    MsgBox "抽出完了！出力ファイル: " & savePath, vbInformation

End Sub


