Attribute VB_Name = "Module1"
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
    
    Dim dateCol As String
    Dim cellDate As Date
    
    ' Задаем столбец для поиска данных
    col = "X"
    ' Задаем столбец для поиска дат
    dateCol = "A"
    
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
'    newWs.Cells(1, 1).Value = "Сделано вызовов"
'    newWs.Cells(1, 2).Value = countFilled
'    newWs.Cells(2, 1).Value = "Системных и сбросы"
'    newWs.Cells(2, 2).Value = countSystem
'    newWs.Cells(3, 1).Value = "Назначено перезвонов:"
'    newWs.Cells(3, 2).Value = countCallback
'    newWs.Cells(4, 1).Value = "АО+ДУБЛЬ+НЕКОР.НОМЕР"
'    newWs.Cells(4, 2).Value = countAODubli
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
    newWs.Cells(5, 1).Value = "Новых контактов за период"
    ' newWs.Cells(5, 2).Value =
    newWs.Cells(6, 1).Value = "Сделано вызовов:"
    newWs.Cells(6, 2).Value = countFilled
    newWs.Cells(7, 1).Value = "Системных не дозвонов и сбросов:"
    newWs.Cells(7, 2).Value = countSystem
    newWs.Cells(8, 1).Value = "Назначено перезвонов:"
    newWs.Cells(8, 2).Value = countCallback
    newWs.Cells(9, 1).Value = "АО+ДУБЛЬ+НЕКОР.НОМЕР"
    newWs.Cells(9, 2).Value = countAODubli
    newWs.Cells(10, 1).Value = "Общее отказов ЛПР"
    newWs.Cells(10, 2).Value = countLPR
    
    
    ' Записать результаты для каждого элемента массива LPRArray
    For j = 1 To UBound(LPRArray) + 1
        newWs.Cells(10 + j, 1).Value = LPRArray(j - 1)
        newWs.Cells(10 + j, 2).Value = LPRCounts(j)
    Next j
    
'    newWs.Cells(10, 1).Value = "не подходит KPI:"
'    newWs.Cells(11, 1).Value = "не целевой:"
'    newWs.Cells(12, 1).Value = "уже купили:"
'    newWs.Cells(13, 1).Value = "не интересовался:"
'    newWs.Cells(14, 1).Value = "отложил на неопределенный срок"
'    newWs.Cells(15, 1).Value = "Отказов ЛПР"
'    newWs.Cells(15, 1).Value = "Отказов ЛПР"
End Sub



