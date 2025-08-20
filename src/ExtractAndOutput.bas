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
    Dim r As Long, c As Long
    
    Dim dict As Object
    Dim key As String
    
    Dim colMap As Variant, outCols As Variant
    colMap = Array(4, 5, 6, 9, 10, 11, 12, 15) ' D,E,F,I,J,K,L,O
    outCols = Array(1, 2, 3, 4, 5, 6, 7, 8)    ' �o�͗� A-H
    
    ' --- ���s���̃u�b�N�𒊏o�ΏۂƂ��� ---
    Set wbSource = ThisWorkbook
    Set wsSource = wbSource.Sheets(1)
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row
    dataArr = wsSource.Range("A1:O" & lastRow).Value
    
    rowOut = 0
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' --- ���o�E�d���폜 ---
    For r = 3 To UBound(dataArr, 1)  ' 1�`2�s�ڂ̓X�L�b�v
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
    
    ' �o�͗p�z��m��
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
    
    ' --- �o�̓u�b�N�쐬 ---
    Set wbResult = Workbooks.Add
    Set wsResult = wbResult.Sheets(1)
    
    ' �w�b�_�s
    Dim headers As Variant
    headers = Array("�e�i��", "�e�i��", "�e�i�ڂ̓�����", "�e�i�ڂ̐����ꏊ�R�[�h", _
                    "�q�i��", "�q�i��", "�q�i�ڂ̓�����", "�q�i�ڂ̐����ꏊ�R�[�h")
    wsResult.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    
    ' �f�[�^��������
    If rowOut > 0 Then
        wsResult.Range("A2").Resize(rowOut, UBound(outCols) + 1).Value = outputArr
    End If
    
    ' --- �ۑ��i�Œ薼�j ---
    Dim savePath As String
    savePath = wbSource.Path & "\�����i�ڏڍ׈ꗗ_work1.xlsm"
    
    Application.DisplayAlerts = False
    wbResult.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    
    wbResult.Close SaveChanges:=False
    
    MsgBox "���o�����I�o�̓t�@�C��: �����i�ڏڍ׈ꗗ_work1.xlsm", vbInformation

End Sub


