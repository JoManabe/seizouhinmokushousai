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
    Dim r As Long
    Dim dict As Object
    Dim key As String
    
    ' --- 元データ ---
    Set wbSource = ThisWorkbook
    Set wsSource = wbSource.Sheets(1)
    lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row
    dataArr = wsSource.Range("A1:O" & lastRow).Value ' すべて配列に読み込み
    
    ' --- 出力用配列の最大行数を確保（2倍で十分） ---
    ReDim outputArr(1 To (UBound(dataArr, 1) - 2) * 2, 1 To 4)
    Set dict = CreateObject("Scripting.Dictionary")
    rowOut = 0
    
    ' --- 配列上で抽出・重複削除 ---
    For r = 3 To UBound(dataArr, 1)
        ' Dummy 除外
        If InStr(1, CStr(dataArr(r, 4)), "Dummy", vbTextCompare) = 0 _
           And InStr(1, CStr(dataArr(r, 10)), "Dummy", vbTextCompare) = 0 Then
           
            ' D,E,F,I のキー
            key = dataArr(r, 4) & "|" & dataArr(r, 5) & "|" & dataArr(r, 6) & "|" & dataArr(r, 9)
            If Not dict.Exists(key) Then
                dict.Add key, 1
                rowOut = rowOut + 1
                outputArr(rowOut, 1) = dataArr(r, 4)
                outputArr(rowOut, 2) = dataArr(r, 5)
                outputArr(rowOut, 3) = dataArr(r, 6)
                outputArr(rowOut, 4) = dataArr(r, 9)
            End If
            
            ' J,K,L,O のキー
            key = dataArr(r, 10) & "|" & dataArr(r, 11) & "|" & dataArr(r, 12) & "|" & dataArr(r, 15)
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
    
    ' --- 出力ブック作成 ---
    Set wbResult = Workbooks.Add
    Set wsResult = wbResult.Sheets(1)
    
    ' ヘッダー
    wsResult.Range("A1:D1").Value = Array("品目CD", "品名", "内部名", "製造場所CD")
    
    ' --- 配列の内容を一度に書き込み ---
    If rowOut > 0 Then
        wsResult.Range("A2").Resize(rowOut, 4).Value = outputArr
    End If
    
    ' --- 保存 ---
    Dim savePath As String
    savePath = wbSource.Path & "\製造品目詳細一覧_work1.xlsm"
    
    Application.DisplayAlerts = False
    wbResult.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    
    wbResult.Close SaveChanges:=False
    
    MsgBox "抽出完了！出力ファイル: 製造品目詳細一覧_work1.xlsm", vbInformation

End Sub


