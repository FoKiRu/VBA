Attribute VB_Name = "Module1"
Public UserFormCancelled As Boolean

Sub CountFilledCellsInColumnX()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim countFilled As Long
    Dim countSystem As Long
    Dim countCallback As Long
    Dim countAODubli As Long
    Dim countLPR As Long
    Dim col As String
    Dim i As Long
    Dim cellValue As String
    Dim systemArray As Variant
    Dim callbackArray As Variant
    Dim aoDubliArray As Variant
    Dim LPRArray As Variant
    Dim LPRCounts() As Long
    
    Dim startDate As Date
    Dim endDate As Date
    Dim dateCol As String
    Dim cellDate As Date
    
    ' ������ ������� ��� ������ ������
    col = "X"
    ' ������ ������� ��� ������ ���
    dateCol = "A"
    
    ' ������ ��� ������ ��������� �����
    systemArray = Array("������� (���������)", "��������� ������������ (���������)", "������ (���������)", _
                        "�������� (���������)", "��������� �� ��������� (���������)", "��� ������ (���������)", _
                        "�������������� �����", "�������� �� ������ ����� (���������)", "�������� ���������� (���������)", _
                        "�������� (���������)", "�������� ����� (���������)", "����� �� ������� ������ (���������)", _
                        "������������ ������ (���������)", "��������� ������ (���������)", "������� (���������)", _
                        "��������� ����� (���������)")
    
    ' ������ ��� ������ ����� "�����������"
    callbackArray = Array("�����������")
    
    ' ������ ��� ������ ����� "��+�����+�����.�����"
    aoDubliArray = Array("�����", "� ��������", "�������", "������������-���������", "������������ �����")
    
    ' ������ ��� ������ ����� "������� ���"
    LPRArray = Array("����� ���: �� �������� KPI", "����� ���: �� �������", _
                     "����� ���: ��� ������", "����� ���: �� �������������", _
                     "����� ���: ������� �� �������������� ����", "����� ���: ��� �������, ���������", _
                     "����� ���: ������ ������")
    ReDim LPRCounts(1 To UBound(LPRArray) + 1)
    
    ' ���������� ����, �� ������� ����� ���������� (�������� "Sheet1" �� ��� ������ �����, ���� ��� ������)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' �������� ����� ��� ����� ��������� ���
    UserFormCancelled = False
    With UserForm1
        .TextBox1.Value = Format(Date, "dd/mm/yyyy")
        .TextBox2.Value = Format(Date, "dd/mm/yyyy")
        .Show
    End With
    
    ' ��������, ���� �� ����� ������� ����� �������
    If UserFormCancelled Then
        Exit Sub
    End If
    
    ' �������� �������� ���
    startDate = DateValue(UserForm1.TextBox1.Value)
    endDate = DateValue(UserForm1.TextBox2.Value)
    
    ' ����� ��������� ����������� ������ � ������� X
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    
    ' ���������� ���������� ����������� ����� � ������� X
    countFilled = Application.WorksheetFunction.CountA(ws.Range(col & "2:" & col & lastRow))
    
    ' ���������� ���������� ����� �� �������� �������� � ��������� ���
    countSystem = 0
    countCallback = 0
    countAODubli = 0
    countLPR = 0
    For i = 2 To lastRow
        cellValue = ws.Cells(i, col).Value
        cellDate = ws.Cells(i, dateCol).Value
        If cellDate >= startDate And cellDate <= endDate Then
            If Not IsError(Application.Match(cellValue, systemArray, 0)) Then
                countSystem = countSystem + 1
            End If
            If Not IsError(Application.Match(cellValue, callbackArray, 0)) Then
                countCallback = countCallback + 1
            End If
            If Not IsError(Application.Match(cellValue, aoDubliArray, 0)) Then
                countAODubli = countAODubli + 1
            End If
            For j = 1 To UBound(LPRArray) + 1
                If cellValue = LPRArray(j - 1) Then
                    LPRCounts(j) = LPRCounts(j) + 1
                    countLPR = countLPR + 1
                End If
            Next j
        End If
    Next i
    
    ' �������� ����� ����
    Set newWs = ThisWorkbook.Sheets.Add
    newWs.Name = "������� �������"
    
    ' ����������� ����� ���� � ����� ���� ������
    newWs.Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    
    ' �������� ��������� �� ����� ����
    newWs.Cells(1, 1).Value = "������� �������"
    newWs.Cells(1, 2).Value = countFilled
    newWs.Cells(2, 1).Value = "��������� � ������"
    newWs.Cells(2, 2).Value = countSystem
    newWs.Cells(3, 1).Value = "��������� ����������:"
    newWs.Cells(3, 2).Value = countCallback
    newWs.Cells(4, 1).Value = "��+�����+�����.�����"
    newWs.Cells(4, 2).Value = countAODubli
    newWs.Cells(5, 1).Value = "������� ���"
    newWs.Cells(5, 2).Value = countLPR
    
    ' �������� ���������� ��� ������� �������� ������� LPRArray
    For j = 1 To UBound(LPRArray) + 1
        newWs.Cells(5 + j, 1).Value = LPRArray(j - 1)
        newWs.Cells(5 + j, 2).Value = LPRCounts(j)
    Next j
End Sub


