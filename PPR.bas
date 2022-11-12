Attribute VB_Name = "Module11"
Sub PZ()
Attribute PZ.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Создание производственного задания из ППР

Application.ScreenUpdating = False

Dim pst(15) As String
    pst(0) = "Т-3"
    pst(1) = "П-23"
    pst(2) = "Т-4"
    pst(3) = "Т-21"
    pst(4) = "ТПП-118"
    pst(5) = "СТП-118"
    pst(6) = "Т-22"
    pst(7) = "Т-30"
    pst(8) = "СТП-63"
    pst(9) = ""
    pst(10) = ""
    pst(11) = ""
    pst(12) = ""
    pst(13) = ""
    pst(14) = ""
    pst(15) = ""
    

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
