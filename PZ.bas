Attribute VB_Name = "Module11"
Sub PZ()
Attribute PZ.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Создание производственного задания из ППР
' Надо удалить хвост

' Application.ScreenUpdating = False

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
    arrOld(12) = ""
    arrOld(13) = ""
    arrOld(14) = ""
    arrOld(15) = ""
    
Dim arrNew(15) As String
    arrNew(0) = "Осмотр. "
    arrNew(1) = "Межремонтное обслуживание. "
    arrNew(2) = "Текущий ремонт. "
    arrNew(3) = "Капитальный ремонт. "
    arrNew(4) = "Проверка. "
    arrNew(5) = "Межремонтное обслуживание. "
    arrNew(6) = "проверка индикации. "
    arrNew(7) = "Измерение сопротивления изоляции. "
    arrNew(8) = "Испытание повышенным напряжением. "
    arrNew(9) = "Тепловизионный контроль. "
    arrNew(10) = "Текущий ремонт. "
    arrNew(11) = "Текущий ремонт. "
    arrNew(12) = ""
    arrNew(13) = ""
    arrNew(14) = ""
    arrNew(15) = ""

'' Удаляем заголовок
'    Rows("1:10").Select
'    Selection.Delete Shift:=xlUp
'
'' Удаляем календарь
'    Columns("K:AO").Select
'    Selection.Delete Shift:=xlToLeft
'
'' Удаляем лишние столбцы
'
'    Columns("E:H").Select
'    Selection.Delete Shift:=xlToLeft
'    Columns("A:B").Select
'    Selection.Delete Shift:=xlToLeft

count_of_records = 1
Do While Cells(count_of_records + 1, "C") <> ""
    count_of_records = count_of_records + 1
Loop
    
'' Отмена объединения
'    For i = 2 To count_of_records
'        If Cells(i, "A") = "" Then
'            Cells(i, "A") = Cells(i - 1, "A")
'            Cells(i, "B") = Cells(i - 1, "B")
'        Else
'            Cells(i, "A").Select
'            Selection.UnMerge
'            Cells(i, "B").Select
'            Selection.UnMerge
'        End If
'    Next

' Основной подсчёт
For i = 2 To count_of_records
    j = 1
    Cells(i, "E") = Cells(i, "D")
    Do While (Cells(i, "A") = Cells(i + j, "A") And Cells(i, "B") = Cells(i + j, "B"))
        Cells(i, "E") = Cells(i, "E") + Cells(i + j, "D")
        j = j + 1
    Loop
    i = i + j - 1
Next

' Application.ScreenUpdating = True

End Sub
