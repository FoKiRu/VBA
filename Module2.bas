Attribute VB_Name = "Module2"
Sub SumUniqueCombinationsWithCounters()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim dictSum As Object
    Dim dictCount1Sec As Object
    Dim dictCount20Sec As Object
    Dim dictCountFillLead As Object
    Dim key As Variant
    Dim i As Long
    
    ' ����������� ��������� �����
    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "�������� ���� �� ������. ����������, �������� ���� � ���������� �����.", vbCritical
        Exit Sub
    End If
    
    ' ����������� ��������� ������
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row ' ������� B
    
    ' �������� �������� ��� �������� ���������� ���������� � �� ���� � ���������
    Set dictSum = CreateObject("Scripting.Dictionary")
    Set dictCount1Sec = CreateObject("Scripting.Dictionary")
    Set dictCount20Sec = CreateObject("Scripting.Dictionary")
    Set dictCountFillLead = CreateObject("Scripting.Dictionary")
    
    ' ���� ���������� ���������� � ���� �������� �� ������� D � ������� ��������
    For i = 2 To lastRow ' ��������������, ��� ������ ������ �������� ���������
        key = ws.Cells(i, 2).Value & "|" & ws.Cells(i, 3).Value ' ������� B � C
        If Not dictSum.exists(key) Then
            dictSum.Add key, ws.Cells(i, 4).Value ' ������� D
            dictCount1Sec.Add key, IIf(ws.Cells(i, 4).Value >= 1, 1, 0)
            dictCount20Sec.Add key, IIf(ws.Cells(i, 4).Value >= 20, 1, 0)
            dictCountFillLead.Add key, IIf(ws.Cells(i, 5).Value = "��������� ���", 1, 0) ' ������� E
        Else
            dictSum(key) = dictSum(key) + ws.Cells(i, 4).Value
            If ws.Cells(i, 4).Value >= 1 Then
                dictCount1Sec(key) = dictCount1Sec(key) + 1
            End If
            If ws.Cells(i, 4).Value >= 20 Then
                dictCount20Sec(key) = dictCount20Sec(key) + 1
            End If
            If ws.Cells(i, 5).Value = "��������� ���" Then
                dictCountFillLead(key) = dictCountFillLead(key) + 1
            End If
        End If
    Next i
    
    ' �������� ������ ����� ��� ������ �����������
    Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newWs.Name = "SumResults"
    
    ' ����� ����������� �� ����� ����
    newWs.Cells(1, 1).Value = "Column B"
    newWs.Cells(1, 2).Value = "Column C"
    newWs.Cells(1, 3).Value = "Sum of Column D (��:��:��)"
    newWs.Cells(1, 4).Value = "Count >= 1 sec"
    newWs.Cells(1, 5).Value = "Count >= 20 sec"
    newWs.Cells(1, 6).Value = "Count '��������� ���'"
    
    i = 2
    For Each key In dictSum.keys
        newWs.Cells(i, 1).Value = Split(key, "|")(0)
        newWs.Cells(i, 2).Value = Split(key, "|")(1)
        newWs.Cells(i, 3).Value = dictSum(key) / 86400 ' �������������� ������ � ���
        newWs.Cells(i, 3).NumberFormat = "[h]:mm:ss" ' ������ ��:��:��
        newWs.Cells(i, 4).Value = dictCount1Sec(key)
        newWs.Cells(i, 5).Value = dictCount20Sec(key)
        newWs.Cells(i, 6).Value = dictCountFillLead(key)
        i = i + 1
    Next key
    
    ' ���������� �������� ��� �������� ������
    newWs.Columns("A:F").AutoFit
    
    ' ������� ������
    Set dictSum = Nothing
    Set dictCount1Sec = Nothing
    Set dictCount20Sec = Nothing
    Set dictCountFillLead = Nothing
End Sub

