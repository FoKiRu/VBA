Sub CopyAndFilterData()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim copyRange As Range
    
    ' Определение активного листа
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Укажите имя исходного листа
    
    ' Создание нового листа
    Set newWs = ThisWorkbook.Sheets.Add
    newWs.Name = "FilteredData"
    
    ' Определение последней строки в столбце C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Копирование столбцов C и D на новый лист
    ws.Range("C1:D" & lastRow).Copy newWs.Range("A1")
    
    ' Определение последней строки на новом листе
    lastRow = newWs.Cells(newWs.Rows.Count, "A").End(xlUp).Row
    
    ' Удаление строк, содержащих 'Компания "Звонко"' или '(без ответственного)' в столбце A (бывший столбец C)
    For i = lastRow To 1 Step -1
        If newWs.Cells(i, 1).Value = "Компания ""Звонко""" Or newWs.Cells(i, 1).Value = "(без ответственного)" Then
            newWs.Rows(i).Delete
        End If
    Next i
End Sub
