' Применить фильтр и удалить из уникальных значений B на листе UniqueValuesB
' Завершающий модуль для макроса CreateUniqueValuesListForColumnB.bas
' part 2/2
' v0.2
' Создан part 3 v0.3, который объединяет в себе два модуля
Sub ApplyFilterAndRemoveFromUniqueValuesB()
    Dim ws As Worksheet
    Dim uniqueValuesWs As Worksheet
    Dim filterValue As String
    Dim lastRow As Long
    
    ' Определение активного листа (где будет применен фильтр)
    Set ws = ActiveSheet
    
    ' Поиск листа UniqueValuesB
    On Error Resume Next
    Set uniqueValuesWs = ThisWorkbook.Sheets("UniqueValuesB")
    On Error GoTo 0
    
    ' Проверка, существует ли лист UniqueValuesB
    If uniqueValuesWs Is Nothing Then
        MsgBox "Лист 'UniqueValuesB' не найден.", vbCritical
        Exit Sub
    End If
    
    ' Получение значения из ячейки A1 на листе UniqueValuesB
    filterValue = uniqueValuesWs.Range("A1").Value
    
    ' Проверка, пустое ли значение
    If filterValue = "" Then
        MsgBox "Значение в ячейке A1 на листе 'UniqueValuesB' пусто.", vbExclamation
        Exit Sub
    End If
    
    ' Определение последней строки на активном листе в столбце B
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' Применение фильтра к столбцу B на активном листе
    ws.Range("B1:B" & lastRow).AutoFilter Field:=2, Criteria1:=filterValue
    
    ' Удаление значения из ячейки A1 на листе UniqueValuesB и смещение всех значений вверх
    uniqueValuesWs.Rows(1).Delete Shift:=xlUp
End Sub

