Public UserFormCancelled As Boolean

Sub CountFilledCellsInColumnX()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim countFilled As Long
    Dim countLPR As Long
    Dim col As String
    Dim i As Long
    Dim cellValue As String

    Dim LPRArray As Variant
    Dim LPRCounts() As Long
    
    Dim startDate As Date
    Dim endDate As Date
    Dim dateCol As String
    Dim cellDate As Date
    
    ' Задаем столбец для поиска данных
    col = "T"
    ' Задаем столбец для поиска дат
    dateCol = "A"
    
    
    ' Массив для поиска строк "отказов ЛПР"
    LPRArray = Array("Тишина", "Автоответчик", _
                     "Нецелевой", "Черный список")
    ReDim LPRCounts(1 To UBound(LPRArray) + 1)
    
    ' Установите лист, на котором нужно подсчитать (измените "Sheet1" на имя вашего листа, если оно другое)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Показать форму для ввода диапазона дат
    UserFormCancelled = False
    With UserForm1
        .TextBox1.Value = Format(Date, "dd/mm/yyyy")
        .TextBox2.Value = Format(Date, "dd/mm/yyyy")
        .Show
    End With
    
    ' Проверка, была ли форма закрыта через крестик
    If UserFormCancelled Then
        Exit Sub
    End If
    
    ' Получить диапазон дат
    startDate = DateValue(UserForm1.TextBox1.Value)
    endDate = DateValue(UserForm1.TextBox2.Value)
    
    ' Найти последнюю заполненную строку в столбце T
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    
    ' Подсчитать количество заполненных ячеек в столбце T
    countFilled = Application.WorksheetFunction.CountA(ws.Range(col & "2:" & col & lastRow))
    
    ' Подсчитать количество строк по заданным массивам и диапазону дат

    countLPR = 0
    For i = 2 To lastRow
        cellValue = ws.Cells(i, col).Value
        cellDate = ws.Cells(i, dateCol).Value
        If cellDate >= startDate And cellDate <= endDate Then
            For j = 1 To UBound(LPRArray) + 1
                If cellValue = LPRArray(j - 1) Then
                    LPRCounts(j) = LPRCounts(j) + 1
                    countLPR = countLPR + 1
                End If
            Next j
        End If
    Next i
    
    ' Добавить новый лист
    Set newWs = ThisWorkbook.Sheets.Add
    newWs.Name = "Сделано вызовов"
    
    ' Переместить новый лист в конец всех листов
    newWs.Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    
    ' Записать результат на новый лист
'    newWs.Cells(1, 1).Value = "Сделано вызовов"
'    newWs.Cells(1, 2).Value = countFilled
'    newWs.Cells(2, 1).Value = "Системных и сбросы"

'    newWs.Cells(3, 1).Value = "Назначено перезвонов:"
'    newWs.Cells(4, 1).Value = "АО+ДУБЛЬ+НЕКОР.НОМЕР"
'    newWs.Cells(5, 1).Value = "Отказов ЛПР"
'    newWs.Cells(5, 2).Value = countLPR
    
    ' Записать результат на новый лист
    newWs.Cells(1, 1).Value = "Проект:"
    ' newWs.Cells(1, 2).Value =
    newWs.Cells(2, 1).Value = "Оператор:"
    ' newWs.Cells(2, 2).Value =
    newWs.Cells(3, 1).Value = "Кол-во проектов на операторе:"
    ' newWs.Cells(3, 2).Value =
    newWs.Cells(4, 1).Value = "Период:"
    ' newWs.Cells(4, 2).Value =
    newWs.Cells(5, 1).Value = "Сделано вызовов:"
    newWs.Cells(5, 2).Value = countFilled
    newWs.Cells(6, 1).Value = "Общее отказов ЛПР"
    newWs.Cells(6, 2).Value = countLPR
    newWs.Cells(7, 1).Value = "их них"

    
    
    ' Записать результаты для каждого элемента массива LPRArray
    For j = 1 To UBound(LPRArray) + 1
        newWs.Cells(7 + j, 1).Value = LPRArray(j - 1)
        newWs.Cells(7 + j, 2).Value = LPRCounts(j)
    Next j
    
'    newWs.Cells(10, 1).Value = "не подходит KPI:"
'    newWs.Cells(11, 1).Value = "не целевой:"
'    newWs.Cells(12, 1).Value = "уже купили:"
'    newWs.Cells(13, 1).Value = "не интересовался:"
'    newWs.Cells(14, 1).Value = "отложил на неопределенный срок"
'    newWs.Cells(15, 1).Value = "Отказов ЛПР"
'    newWs.Cells(15, 1).Value = "Отказов ЛПР"
End Sub
