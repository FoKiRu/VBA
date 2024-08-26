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
    Dim i As Long, j As Long
    Dim cellValue As String
    Dim systemArray As Variant
    Dim callbackArray As Variant
    Dim aoDubliArray As Variant
    Dim LPRArray As Variant
    Dim LPRCounts() As Long
    Dim scenarioName As String
    
    ' Задаем столбец для поиска данных
    col = "X"
    
    ' Массив для поиска системных строк
    systemArray = Array("Потерян (системный)", "Обнаружен автоответчик (системный)", "Занято (системный)", _
                        "Соединен (системный)", "Сообщение не проиграно (системный)", "Нет ответа (системный)", _
                        "Несуществующий номер", "Оператор не принял вызов (системный)", "Отклонен оператором (системный)", _
                        "Отклонен (системный)", "Оператор занят (системный)", "Номер из черного списка (системный)", _
                        "Неопознанная ошибка (системный)", "Системная ошибка (системный)", "Отложен (системный)", _
                        "Повторный вызов (системный)")
    
    ' Массив для поиска строк "Перезвонить"
    callbackArray = Array("Перезвонить")
    
    ' Массив для поиска строк "АО+ДУБЛЬ+НЕКОР.НОМЕР"
    aoDubliArray = Array("Дубль", "В недозвон", "Молчали", "Автоответчик-секретарь", "Некорректный номер")
    
    ' Массив для поиска строк "отказов ЛПР"
    LPRArray = Array("Отказ ЛПР: не подходит KPI", "Отказ ЛПР: не целевой", _
                     "Отказ ЛПР: уже купили", "Отказ ЛПР: не интересовался", _
                     "Отказ ЛПР: отложил на неопределенный срок", "Отказ ЛПР: был интерес, передумал", _
                     "Отказ ЛПР: бросил трубку")
    ReDim LPRCounts(1 To UBound(LPRArray) + 1)
    
    ' Установите лист, на котором нужно подсчитать (измените "Sheet1" на имя вашего листа, если оно другое)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Найти последнюю заполненную строку в столбце X
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    
    ' Подсчитать количество заполненных ячеек в столбце X
    countFilled = Application.WorksheetFunction.CountA(ws.Range(col & "2:" & col & lastRow))
    
    ' Подсчитать количество строк по заданным массивам
    countSystem = 0
    countCallback = 0
    countAODubli = 0
    countLPR = 0
    
    ' Найти значение "Сценарий" в столбце D
    scenarioName = ""
    For i = 2 To lastRow
        cellValue = ws.Cells(i, "D").Value
        If cellValue <> "" Then
            scenarioName = cellValue
            Exit For
        End If
    Next i
    
    ' Поиск совпадений в столбце X и подсчёт по массивам
    For i = 2 To lastRow
        cellValue = ws.Cells(i, col).Value
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
    Next i
    
    ' Добавить новый лист
    Set newWs = ThisWorkbook.Sheets.Add
    newWs.Name = "Сделано вызовов"
    
    ' Переместить новый лист в конец всех листов
    newWs.Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    
    ' Записать результат на новый лист
    newWs.Cells(1, 1).Value = "Проект:"
    newWs.Cells(1, 2).Value = scenarioName ' Записываем найденное значение "Сценарий"
    
    newWs.Cells(2, 1).Value = "Оператор:"
    newWs.Cells(3, 1).Value = "Кол-во проектов на операторе:"
    newWs.Cells(4, 1).Value = "Период:"
    newWs.Cells(5, 1).Value = "Новых контактов за период"
    newWs.Cells(6, 1).Value = "Сделано вызовов:"
    newWs.Cells(6, 2).Value = countFilled & " (" & Format((countFilled / countFilled) * 100, "0.00") & "%)"
    
    newWs.Cells(7, 1).Value = "Системных не дозвонов и сбросов:"
    newWs.Cells(7, 2).Value = countSystem & " (" & Format((countSystem / countFilled) * 100, "0.00") & "%)"
    
    newWs.Cells(9, 1).Value = "Назначено перезвонов:"
    newWs.Cells(9, 2).Value = countCallback & " (" & Format((countCallback / countFilled) * 100, "0.00") & "%)"
    
    newWs.Cells(10, 1).Value = "АО+ДУБЛЬ+НЕКОР.НОМЕР:"
    newWs.Cells(10, 2).Value = countAODubli & " (" & Format((countAODubli / countFilled) * 100, "0.00") & "%)"
    
    newWs.Cells(11, 1).Value = "Общее отказов ЛПР:"
    newWs.Cells(11, 2).Value = countLPR & " (" & Format((countLPR / countFilled) * 100, "0.00") & "%)"
    
    newWs.Cells(12, 1).Value = "Из них:"
    
    ' Записать результаты для каждого элемента массива LPRArray
    For j = 1 To UBound(LPRArray) + 1
        newWs.Cells(12 + j, 1).Value = LPRArray(j - 1)
        newWs.Cells(12 + j, 2).Value = LPRCounts(j) & " (" & Format((LPRCounts(j) / countFilled) * 100, "0.00") & "%)"
    Next j
End Sub
