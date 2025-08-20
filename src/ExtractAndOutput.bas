Attribute VB_Name = "ExtractAndOutput"
Sub 抽出と重複削除()

    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wbResult As Workbook
    Dim wsResult As Worksheet
    Dim lastRow As Long
    Dim dataArr As Variant
    Dim outputArr() As Variant
    Dim rowOut As Long
    Dim r As Long, c As Long
    
    Dim dict As Object
    Dim key As String
    
    Dim colMap As Variant, outCols As Variant
    colMap = Array(4, 5, 6, 9, 10, 11, 12, 15) ' D,E,F,I,J,K,L,O
    outCols = Array(1, 2, 3, 4, 5, 6, 7, 8)    ' 出力列 A-H
    
    ' --- 実行中のブックを抽出対象とする ---
    Set wbSource = ThisWorkbook
    Set wsSource = wbSource.Sheets(1)
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row
    dataArr = wsSource.Range("A1:O" & lastRow).Value
    
    rowOut = 0
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' --- 抽出・重複削除 ---
    For r = 3 To UBound(dataArr, 1)  ' 1～2行目はスキップ
        If InStr(1, CStr(dataArr(r, 4)), "Dummy", vbTextCompare) = 0 _
           And InStr(1, CStr(dataArr(r, 10)), "Dummy", vbTextCompare) = 0 Then
           
            key = ""
            For c = LBound(colMap) To UBound(colMap)
                key = key & "|" & dataArr(r, colMap(c))
            Next c
            
            If Not dict.Exists(key) Then 'DuplicateCheck
                dict.Add key, 1
                rowOut = rowOut + 1
            End If
        End If
    Next r
    
    ' 出力用配列確保
    If rowOut > 0 Then
        ReDim outputArr(1 To rowOut, 1 To UBound(outCols) + 1)
        rowOut = 0
        dict.RemoveAll
        
        For r = 3 To UBound(dataArr, 1)
            If InStr(1, CStr(dataArr(r, 4)), "Dummy", vbTextCompare) = 0 _
               And InStr(1, CStr(dataArr(r, 10)), "Dummy", vbTextCompare) = 0 Then
               
                key = ""
                For c = LBound(colMap) To UBound(colMap)
                    key = key & "|" & dataArr(r, colMap(c))
                Next c
                
                If Not dict.Exists(key) Then
                    dict.Add key, 1
                    rowOut = rowOut + 1
                    For c = LBound(colMap) To UBound(colMap)
                        outputArr(rowOut, outCols(c)) = dataArr(r, colMap(c))
                    Next c
                End If
            End If
        Next r
    End If
    
    ' --- 出力ブック作成 ---
    Set wbResult = Workbooks.Add
    Set wsResult = wbResult.Sheets(1)
    
    ' ヘッダ行
    Dim headers As Variant
    headers = Array("親品番", "親品名", "親品目の内部名", "親品目の製造場所コード", _
                    "子品番", "子品名", "子品目の内部名", "子品目の製造場所コード")
    wsResult.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    
    ' データ書き込み
    If rowOut > 0 Then
        wsResult.Range("A2").Resize(rowOut, UBound(outCols) + 1).Value = outputArr
    End If
    
    ' --- 保存（固定名） ---
    Dim savePath As String
    savePath = wbSource.Path & "\製造品目詳細一覧_work1.xlsm"
    
    Application.DisplayAlerts = False
    wbResult.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    
    wbResult.Close SaveChanges:=False
    
    MsgBox "抽出完了！出力ファイル: 製造品目詳細一覧_work1.xlsm", vbInformation

End Sub


