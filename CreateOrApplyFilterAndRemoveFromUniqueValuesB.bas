' Создает или применяет фильтр и удаляет из уникальных значений B
' v0.3 совмещает в себя два модуля
Sub CreateOrApplyFilterAndRemoveFromUniqueValuesB()
    Dim ws As Worksheet
    Dim uniqueValuesWs As Worksheet
    Dim lastRow As Long
    Dim filterValue As String
    Dim currentWs As Worksheet
    Dim filterValues As Collection
    Dim i As Long
    Dim uniqueValue As String
    Dim outputRow As Long
    
    ' Определение активного листа
    Set ws = ActiveSheet
    Set currentWs = ws

    ' Поиск листа UniqueValuesB
    On Error Resume Next
    Set uniqueValuesWs = ThisWorkbook.Sheets("UniqueValuesB")
    On Error GoTo 0

    ' Если лист UniqueValuesB существует
    If Not uniqueValuesWs Is Nothing Then
        ' Получение значения из ячейки A1 на листе UniqueValuesB
        filterValue = uniqueValuesWs.Range("A1").Value
        
        ' Если значение в ячейке A1 пустое
        If filterValue = "" Then
            ' Удаление листа UniqueValuesB
            Application.DisplayAlerts = False
            uniqueValuesWs.Delete
            Application.DisplayAlerts = True
            MsgBox "Лист 'UniqueValuesB' пуст и был удален.", vbInformation
        Else
            ' Определение последней строки на активном листе в столбце B
            lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            
            ' Применение фильтра к столбцу B на активном листе
            ws.Range("B1:B" & lastRow).AutoFilter Field:=2, Criteria1:=filterValue
            
            ' Удаление первой строки на листе UniqueValuesB
            uniqueValuesWs.Rows(1).Delete Shift:=xlUp
            
            ' Сообщение о применении фильтра
            MsgBox "Фильтр применен по значению '" & filterValue & "', и оно было удалено из листа 'UniqueValuesB'.", vbInformation
        End If
    Else
        ' Если листа UniqueValuesB нет, создаем его и заполняем уникальными значениями
        ' Определение последней строки в столбце B на активном листе
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
        If filterValues.Count = 0 Then
            MsgBox "В столбце B нет уникальных значений для создания листа 'UniqueValuesB'.", vbExclamation
            Exit Sub
        End If
        
        ' Создание нового листа для вывода уникальных значений
        Set uniqueValuesWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        uniqueValuesWs.Name = "UniqueValuesB"

        ' Заполнение нового листа уникальными значениями без заголовка
        outputRow = 1
        For i = 1 To filterValues.Count
            uniqueValuesWs.Cells(outputRow, 1).Value = filterValues(i)
            outputRow = outputRow + 1
        Next i

        ' Автоширина столбца для удобства чтения
        uniqueValuesWs.Columns("A:A").AutoFit

        ' Возвращаем активность исходному листу
        currentWs.Activate
        
        MsgBox "Лист 'UniqueValuesB' был создан и заполнен уникальными значениями из столбца B.", vbInformation
    End If
End Sub
