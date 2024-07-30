Sub SumUniqueCombinations()
' Этот макрос выполняет следующие действия:
' 1. Определяет активный лист.
' 2. Находит последнюю заполненную строку в столбце B.
' 3. Собирает уникальные комбинации значений из столбцов B и C, суммируя соответствующие значения из столбца D.
' 4. Создает новый лист для вывода результатов.
' 5. Выводит уникальные комбинации и их суммы на новый лист в столбцы A, B и C.
' 6. Автоматически подгоняет ширину столбцов для удобства чтения.
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim dict As Object
    Dim key As Variant
    Dim cell As Range
    Dim i As Long
    
    ' Определение активного листа
    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "Активный лист не найден. Пожалуйста, выберите лист и попробуйте снова.", vbCritical
        Exit Sub
    End If
    
    ' Определение последней строки
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row ' Столбец B
    
    ' Создание словаря для хранения уникальных комбинаций и их сумм
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Сбор уникальных комбинаций и сумм значений из столбца D
    For i = 2 To lastRow ' Предполагается, что первая строка содержит заголовки
        key = ws.Cells(i, 2).Value & "|" & ws.Cells(i, 3).Value ' Столбцы B и C
        If Not dict.exists(key) Then
            dict.Add key, ws.Cells(i, 4).Value ' Столбец D
        Else
            dict(key) = dict(key) + ws.Cells(i, 4).Value
        End If
    Next i
    
    ' Создание нового листа для вывода результатов
    Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newWs.Name = "SumResults"
    
    ' Вывод результатов на новый лист
    newWs.Cells(1, 1).Value = "Column B"
    newWs.Cells(1, 2).Value = "Column C"
    newWs.Cells(1, 3).Value = "Sum of Column D"
    
    i = 2
    For Each key In dict.keys
        newWs.Cells(i, 1).Value = Split(key, "|")(0)
        newWs.Cells(i, 2).Value = Split(key, "|")(1)
        newWs.Cells(i, 3).Value = dict(key)
        i = i + 1
    Next key
    
    ' Автоширина столбцов для удобства чтения
    newWs.Columns("A:C").AutoFit
    
    ' Очистка памяти
    Set dict = Nothing
End Sub

