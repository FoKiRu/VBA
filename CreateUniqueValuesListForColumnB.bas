    ' Создать Уникальный Список Значений Для Столбца B
Sub CreateUniqueValuesListForColumnB()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim filterValues As Collection
    Dim i As Long
    Dim uniqueValue As String
    Dim outputRow As Long
    Dim currentWs As Worksheet

    ' Определение активного листа
    Set ws = ActiveSheet
    ' Запоминаем текущий активный лист
    Set currentWs = ws

    ' Определение последней строки в столбце B
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    ' Сбор уникальных значений в столбце B в коллекцию
    Set filterValues = New Collection
    On Error Resume Next
    For i = 2 To lastRow ' Предполагается, что первая строка содержит заголовки
        uniqueValue = ws.Cells(i, 2).Value
        If uniqueValue <> "" Then
            filterValues.Add uniqueValue, CStr(uniqueValue) ' Добавляем только уникальные значения
        End If
    Next i
    On Error GoTo 0

    ' Если нет уникальных значений, выходим
    If filterValues.Count = 0 Then Exit Sub

    ' Создание нового листа для вывода уникальных значений
    Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newWs.Name = "UniqueValuesB"

    ' Заполнение нового листа уникальными значениями без заголовка
    outputRow = 1
    For i = 1 To filterValues.Count
        newWs.Cells(outputRow, 1).Value = filterValues(i)
        outputRow = outputRow + 1
    Next i

    ' Автоширина столбца для удобства чтения
    newWs.Columns("A:A").AutoFit

    ' Возвращаем активность исходному листу
    currentWs.Activate
End Sub

