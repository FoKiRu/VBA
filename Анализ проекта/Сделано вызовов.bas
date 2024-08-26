' Модуль формы UserForm1
'Private Sub CommandButton1_Click()
'    ' Проверка правильности ввода дат
'    If IsDate(TextBox1.Value) And IsDate(TextBox2.Value) Then
'        UserFormCancelled = False
'        Me.Hide
'    Else
'        MsgBox "Пожалуйста, введите корректные даты.", vbExclamation
'    End If
'End Sub
'
'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    ' Если форма закрыта через крестик
'    If CloseMode = vbFormControlMenu Then
'        Cancel = True
'        UserFormCancelled = True
'        Me.Hide
'    End If
'End Sub


' Основной модуль
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
    
    ' Найти последнюю заполненную строку в столбце X
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    
    ' Подсчитать количество заполненных ячеек в столбце X
    countFilled = Application.WorksheetFunction.CountA(ws.Range(col & "2:" & col & lastRow))
    
    ' Подсчитать количество строк по заданным массивам и диапазону дат
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
    
    ' Добавить новый лист
    Set newWs = ThisWorkbook.Sheets.Add
    newWs.Name = "Сделано вызовов"
    
    ' Переместить новый лист в конец всех листов
    newWs.Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    
    ' Записать результат на новый лист
    newWs.Cells(1, 1).Value = "Сделано вызовов"
    newWs.Cells(1, 2).Value = countFilled
    newWs.Cells(2, 1).Value = "Системных и сбросы"
    newWs.Cells(2, 2).Value = countSystem
    newWs.Cells(3, 1).Value = "Назначено перезвонов:"
    newWs.Cells(3, 2).Value = countCallback
    newWs.Cells(4, 1).Value = "АО+ДУБЛЬ+НЕКОР.НОМЕР"
    newWs.Cells(4, 2).Value = countAODubli
    newWs.Cells(5, 1).Value = "Отказов ЛПР"
    newWs.Cells(5, 2).Value = countLPR
    
    ' Записать результаты для каждого элемента массива LPRArray
    For j = 1 To UBound(LPRArray) + 1
        newWs.Cells(5 + j, 1).Value = LPRArray(j - 1)
        newWs.Cells(5 + j, 2).Value = LPRCounts(j)
    Next j
End Sub
