Attribute VB_Name = "Module11"
Sub PZ()
Attribute PZ.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Создание производственного задания из ППР

Application.ScreenUpdating = False

Dim arrOld(15) As String
    arrOld(0) = "О"
    arrOld(1) = "МРО"
    arrOld(2) = "ТР"
    arrOld(3) = "КР"
    arrOld(4) = "Проверка"
    arrOld(5) = "МРО + опробование"
    arrOld(6) = "проверка индикации"
    arrOld(7) = "ИзмСИ"
    arrOld(8) = "ИПН"
    arrOld(9) = "ТВК (с прим. пирометров)"
    arrOld(10) = "ТР-1"
    arrOld(11) = "ТР-2"
    arrOld(12) = "проверка"
    arrOld(13) = "Изм"
    arrOld(14) = ""
    arrOld(15) = ""
    
Dim arrNew(15) As String
    arrNew(0) = "Осмотр. "
    arrNew(1) = "Межремонтное обслуживание. "
    arrNew(2) = "Текущий ремонт. "
    arrNew(3) = "Капитальный ремонт. "
    arrNew(4) = "Проверка. "
    arrNew(5) = "Межремонтное обслуживание. "
    arrNew(6) = "Проверка индикации. "
    arrNew(7) = "Измерение сопротивления изоляции. "
    arrNew(8) = "Испытание повышенным напряжением. "
    arrNew(9) = "Тепловизионный контроль. "
    arrNew(10) = "Текущий ремонт. "
    arrNew(11) = "Текущий ремонт. "
    arrNew(12) = "Проверка. "
    arrNew(13) = "Измерение. "
    arrNew(14) = ""
    arrNew(15) = ""

' Удаляем заголовок
    Rows("1:10").Select
    Selection.Delete Shift:=xlUp

' Удаляем календарь
    Columns("K:AO").Select
    Selection.Delete Shift:=xlToLeft

' Удаляем лишние столбцы

    Columns("E:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:B").Select
    Selection.Delete Shift:=xlToLeft

' Подсчет количества записей по колонке с подстанциями
count_of_records = 1
Do While Cells(count_of_records + 1, "D") <> ""
    count_of_records = count_of_records + 1
Loop
    
' Отмена объединения
Range(Cells(1, "A").EntireColumn, Cells(1, "C").EntireColumn).Select
Selection.UnMerge

For i = 2 To count_of_records
    If Cells(i, "A") = "" Then
        Cells(i, "A") = Cells(i - 1, "A")
        Cells(i, "B") = Cells(i - 1, "B")
        Cells(i, "C") = Cells(i - 1, "C")
    End If
Next

' Основной подсчёт
For i = 2 To count_of_records
    j = 1
    Cells(i, "F") = Cells(i, "E")
    Do While (Cells(i, "A") = Cells(i + j, "A") And Cells(i, "B") = Cells(i + j, "B"))
        Cells(i, "F") = Cells(i, "F") + Cells(i + j, "E")
        j = j + 1
    Loop
    i = i + j - 1
Next

' Удаление строк без основного расчёта и больше не нужных колонок
i = 2
Do While Cells(i, "A") <> ""
    If Cells(i, "F") = "" Then
        Rows(i).Select
        Selection.Delete Shift:=xlUp
        i = i - 1
    End If
    i = i + 1
Loop
Range(Cells(1, 4).EntireColumn, Cells(1, 5).EntireColumn).Select
Selection.Delete Shift:=xlUp

' Придание формы
Columns("A:A").Select
Selection.Insert Shift:=xlToRight

i = 1
Do While Cells(i, "B") <> ""
    For j = 0 To 15
        If Cells(i, "C") = arrOld(j) Then
            Cells(i, "C") = arrNew(j)
        End If
    Next
    
    Cells(i, "A") = Cells(i, "C").Value & Cells(i, "B") & "."
    
    i = i + 1
Loop

Range(Cells(1, 2).EntireColumn, Cells(1, 3).EntireColumn).Select
Selection.Delete Shift:=xlUp

Columns("C:C").Select
Selection.Copy
Columns("B:B").Select
Selection.Insert Shift:=xlToRight

Columns("C:C").Select
Selection.Copy
Columns("E:E").Select
Selection.Insert Shift:=xlToRight

Rows(1).Select
Selection.Delete Shift:=xlUp

Range(Cells(1, 1), Cells(i - 2, 5)).Select
Selection.Copy

Application.ScreenUpdating = True

End Sub
