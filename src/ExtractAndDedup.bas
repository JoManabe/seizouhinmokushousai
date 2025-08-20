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
    
    ' --- �_�C�A���O�őΏۃt�@�C����I�� ---
    MsgBox "���������� Excel �t�@�C����I�����Ă�������", vbInformation
    selectedFiles = Application.GetOpenFilename( _
                    FileFilter:="Excel Files (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", _
                    MultiSelect:=True, _
                    Title:="���o�Ώۂ�Excel�t�@�C����I�����Ă�������")
                    
    If Not IsArray(selectedFiles) Then Exit Sub
    
    ' --- �o�̓u�b�N�i�}�N����p�u�b�N�j ---
    Set wbResult = ThisWorkbook
    Set wsResult = wbResult.Sheets(1)
    wsResult.Cells.Clear
    
    ' �w�b�_�[�i4��j
    wsResult.Range("A1:D1").Value = Array("�i��CD", "�i��", "������", "�����ꏊCD")
    currentRow = 2
    
    ' --- �I���t�@�C������������ ---
    For iFile = LBound(selectedFiles) To UBound(selectedFiles)
        Set wbSource = Workbooks.Open(selectedFiles(iFile))
        Set wsSource = wbSource.Sheets(1)
        
        lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row
        dataArr = wsSource.Range("A1:O" & lastRow).Value
        
        ' �o�͗p�z��m�ہi�ő�s���͏\���]�T����������j
        ReDim outputArr(1 To (UBound(dataArr, 1) - 2) * 2, 1 To 4)
        Set dict = CreateObject("Scripting.Dictionary")
        rowOut = 0
        
        ' --- �z���Œ��o�E�d���폜�i1�t�@�C�����Ɓj ---
        For r = 3 To UBound(dataArr, 1)  ' 1�`2�s�ڂ̓X�L�b�v
            ' Dummy���O
            If InStr(1, CStr(dataArr(r, 4)), "Dummy", vbTextCompare) = 0 _
               And InStr(1, CStr(dataArr(r, 10)), "Dummy", vbTextCompare) = 0 Then
               
                ' D,E,F,I �̃L�[
                key = CStr(dataArr(r, 4)) & "|" & CStr(dataArr(r, 5)) & "|" & CStr(dataArr(r, 6)) & "|" & CStr(dataArr(r, 9))
                If Not dict.Exists(key) Then
                    dict.Add key, 1
                    rowOut = rowOut + 1
                    outputArr(rowOut, 1) = dataArr(r, 4)
                    outputArr(rowOut, 2) = dataArr(r, 5)
                    outputArr(rowOut, 3) = dataArr(r, 6)
                    outputArr(rowOut, 4) = dataArr(r, 9)
                End If
                
                ' J,K,L,O �̃L�[
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
        
        ' --- �o�̓V�[�g�ɒǉ� ---
        If rowOut > 0 Then
            wsResult.Range("A" & currentRow).Resize(rowOut, 4).Value = outputArr
            currentRow = currentRow + rowOut
        End If
        
        wbSource.Close SaveChanges:=False
    Next iFile
    
    ' --- �ۑ��i�Œ薼�j ---
    Dim savePath As String
    savePath = wbResult.Path & "\�����i�ڏڍ׈ꗗ_work1.xlsm"
    
    Application.DisplayAlerts = False
    wbResult.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    
    MsgBox "���o�����I�o�̓t�@�C��: " & savePath, vbInformation

End Sub


