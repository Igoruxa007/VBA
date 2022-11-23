Attribute VB_Name = "Module11"
Sub PZ()
Attribute PZ.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Создание производственного задания из ППР
' Удалить лишние месяцы
' Вбить количество дней в месяце
n_day = 31 ' !!!!!!!!!

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

' Удаляем ненужные колонки
    Columns("I:J").Select
    Selection.Delete Shift:=xlToLeft

' Отмена объединения
Range(Cells(1, "A").EntireColumn, Cells(1, "J").EntireColumn).Select
Selection.UnMerge

For i = 2 To 1450
    If Cells(i, "A") = "" Then
        Cells(i, "A") = Cells(i - 1, "A")
        Cells(i, "B") = Cells(i - 1, "B")
        Cells(i, "C") = Cells(i - 1, "C")
        Cells(i, "D") = Cells(i - 1, "D")
        Cells(i, "E") = Cells(i - 1, "E")
        Cells(i, "F") = Cells(i - 1, "F")
        Cells(i, "G") = Cells(i - 1, "G")
        Cells(i, "H") = Cells(i - 1, "H")
    End If
    If Cells(i, "A") <> "" And Cells(i, 3) = "" Then
        Cells(i, "B") = Cells(i - 1, "B")
        Cells(i, "C") = Cells(i - 1, "C")
        Cells(i, "H") = Cells(i - 1, "H")
        Cells(i, "H") = Cells(i - 1, "H")
    End If
Next


' Удаление строк без нужных подстанций
i = 3

Do While Cells(i, "A") <> ""

log_condition = False

    For j = 0 To 8
        If Cells(i, 9) = pst(j) Then
            log_condition = True
            Exit For
        End If
    Next
    
    If log_condition = False Then
        Rows(i).Select
        Selection.Delete Shift:=xlUp
        i = i - 1
    End If
        
    i = i + 1
        
Loop


' Обратное объединение
i = 4
j = 3
Application.DisplayAlerts = False

Do While Cells(i - 1, "A") <> ""
    
    If (Cells(i, 3) <> Cells(j, 3) Or Cells(i, 4) <> Cells(j, 4)) And i - j > 1 Then
        For k = 1 To 8
            Range(Cells(i - 1, k), Cells(j, k)).Select
            Selection.Merge
        Next
        j = i
    End If
    
    If (Cells(i, 3) <> Cells(j, 3) Or Cells(i, 4) <> Cells(j, 4)) And i - j = 1 Then
        j = i
    End If
        
    i = i + 1
        
Loop
Application.DisplayAlerts = True


' Выравнивание высоты строк
i = 3
Do While Cells(i, 9) <> ""
    Rows(i).EntireRow.AutoFit
    i = i + 1
Loop

' Вставляем числа месяца
For j = 1 To n_day
    Cells(1, 10 + j) = j
Next

' Рисуем границы
Range(Cells(1, 11), Cells(i - 1, 10 + n_day)).Select
With Selection.Borders(xlEdgeLeft)
    .Weight = xlMedium
End With
With Selection.Borders(xlEdgeTop)
    .Weight = xlMedium
End With
With Selection.Borders(xlEdgeBottom)
    .Weight = xlMedium
End With
With Selection.Borders(xlEdgeRight)
    .Weight = xlMedium
End With
With Selection.Borders(xlInsideVertical)
    .Weight = xlThin
End With
With Selection.Borders(xlInsideHorizontal)
    .Weight = xlThin
End With

' Наводим порядок в шапке
Cells(1, 9) = Cells(2, 9)
Cells(1, 10) = Cells(2, 10)
Rows(2).Select
Selection.Delete Shift:=xlUp

Application.ScreenUpdating = True

End Sub
