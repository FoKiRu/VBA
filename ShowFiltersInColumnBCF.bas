Sub ShowFiltersInColumns()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim filterValues As Collection
    Dim cell As Range
    Dim i As Long
    Dim uniqueValues As Variant
    Dim columnsToCheck As Variant
    Dim columnLetter As Variant
    Dim newSheetName As String
    
    ' Определение активного листа
    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "Активный лист не найден. Пожалуйста, выберите лист и попробуйте снова.", vbCritical
        Exit Sub
    End If
    
    ' Массив столбцов для проверки
    columnsToCheck = Array("B", "C", "F")
    
    ' Проход по каждому столбцу из массива
    For Each columnLetter In columnsToCheck
        ' Определение последней строки в текущем столбце
        lastRow = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row
        
        ' Инициализация коллекции для хранения уникальных значений фильтров
        Set filterValues = New Collection
        
        ' Сбор уникальных значений из текущего столбца
        On Error Resume Next ' Игнорирование ошибок при добавлении дубликатов
        For Each cell In ws.Range(columnLetter & "2:" & columnLetter & lastRow)
            If cell.Value <> "" Then
                filterValues.Add cell.Value, CStr(cell.Value)
            End If
        Next cell
        On Error GoTo 0 ' Возврат стандартной обработки ошибок
        
        ' Преобразование коллекции в массив
        ReDim uniqueValues(1 To filterValues.Count)
        For i = 1 To filterValues.Count
            uniqueValues(i) = filterValues(i)
        Next i
        
        ' Определение имени нового листа
        Select Case columnLetter
            Case "B"
                newSheetName = "FiltersInB"
            Case "C"
                newSheetName = "FiltersInC"
            Case "F"
                newSheetName = "FiltersInF"
        End Select
        
        ' Создание нового листа для вывода фильтров справа от всех существующих листов
        Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        On Error Resume Next
        newWs.Name = newSheetName
        If Err.Number <> 0 Then
            Set newWs = Nothing
            MsgBox "Не удалось создать новый лист. Возможно, лист с таким именем уже существует.", vbCritical
            Exit Sub
        End If
        On Error GoTo 0
        
        ' Вывод уникальных значений фильтров на новый лист
        newWs.Cells(1, 1).Value = "Available Filters in Column " & columnLetter
        For i = 1 To UBound(uniqueValues)
            newWs.Cells(i + 1, 1).Value = uniqueValues(i)
        Next i
        
        ' Автоширина столбца для удобства чтения
        newWs.Columns("A").AutoFit
    Next columnLetter
End Sub

