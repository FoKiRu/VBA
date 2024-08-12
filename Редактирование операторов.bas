Sub ModifyColumnB()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cellValue As String
    Dim parts() As String
    
    ' Определение текущего листа
    Set ws = ThisWorkbook.Sheets("Sheet3") ' Замените "Sheet1" на имя вашего листа
    
    ' Поиск последней строки в столбце B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Проход по каждой ячейке в столбце B
    For i = 1 To lastRow
        cellValue = ws.Cells(i, "B").Value ' Чтение значения ячейки
        parts = Split(cellValue, " ") ' Разделение строки на части по запятой
        
        If UBound(parts) >= 2 Then ' Проверка, что массив содержит как минимум три элемента
            ws.Cells(i, "B").Value = parts(2) & " " & parts(1) ' Перестановка местами второго и третьего элемента
        End If
    Next i
End Sub
















Sub ModifyColumnB()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cellValue As String
    Dim parts() As String
    
    ' Определение текущего листа
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Замените "Sheet1" на имя вашего листа
    
    ' Поиск последней строки в столбце B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Проход по каждой ячейке в столбце B
    For i = 1 To lastRow
        cellValue = ws.Cells(i, "B").Value ' Чтение значения ячейки
        parts = Split(cellValue, " ") ' Разделение строки на части по пробелу
        
        If UBound(parts) >= 2 Then ' Проверка, что массив содержит как минимум три элемента
            ws.Cells(i, "B").Value = parts(2) & " " & parts(1) ' Перестановка местами второго и третьего элемента
        End If
    Next i
End Sub
