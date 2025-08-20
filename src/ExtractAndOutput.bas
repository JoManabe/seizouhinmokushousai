Attribute VB_Name = "ExtractAndOutput"
Sub ���o�Əd���폜()

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
    
    ' --- ���f�[�^ ---
    Set wbSource = ThisWorkbook
    Set wsSource = wbSource.Sheets(1)
    lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row
    dataArr = wsSource.Range("A1:O" & lastRow).Value ' ���ׂĔz��ɓǂݍ���
    
    ' --- �o�͗p�z��̍ő�s�����m�ہi2�{�ŏ\���j ---
    ReDim outputArr(1 To (UBound(dataArr, 1) - 2) * 2, 1 To 4)
    Set dict = CreateObject("Scripting.Dictionary")
    rowOut = 0
    
    ' --- �z���Œ��o�E�d���폜 ---
    For r = 3 To UBound(dataArr, 1)
        ' Dummy ���O
        If InStr(1, CStr(dataArr(r, 4)), "Dummy", vbTextCompare) = 0 _
           And InStr(1, CStr(dataArr(r, 10)), "Dummy", vbTextCompare) = 0 Then
           
            ' D,E,F,I �̃L�[
            key = dataArr(r, 4) & "|" & dataArr(r, 5) & "|" & dataArr(r, 6) & "|" & dataArr(r, 9)
            If Not dict.Exists(key) Then
                dict.Add key, 1
                rowOut = rowOut + 1
                outputArr(rowOut, 1) = dataArr(r, 4)
                outputArr(rowOut, 2) = dataArr(r, 5)
                outputArr(rowOut, 3) = dataArr(r, 6)
                outputArr(rowOut, 4) = dataArr(r, 9)
            End If
            
            ' J,K,L,O �̃L�[
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
    
    ' --- �o�̓u�b�N�쐬 ---
    Set wbResult = Workbooks.Add
    Set wsResult = wbResult.Sheets(1)
    
    ' �w�b�_�[
    wsResult.Range("A1:D1").Value = Array("�i��CD", "�i��", "������", "�����ꏊCD")
    
    ' --- �z��̓��e����x�ɏ������� ---
    If rowOut > 0 Then
        wsResult.Range("A2").Resize(rowOut, 4).Value = outputArr
    End If
    
    ' --- �ۑ� ---
    Dim savePath As String
    savePath = wbSource.Path & "\�����i�ڏڍ׈ꗗ_work1.xlsm"
    
    Application.DisplayAlerts = False
    wbResult.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    
    wbResult.Close SaveChanges:=False
    
    MsgBox "���o�����I�o�̓t�@�C��: �����i�ڏڍ׈ꗗ_work1.xlsm", vbInformation

End Sub


