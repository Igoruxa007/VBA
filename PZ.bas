Attribute VB_Name = "Module11"
Sub PZ()
Attribute PZ.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Создание производственного задания из ППР
' Надо удалить хвост

' Application.ScreenUpdating = False
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

' Application.ScreenUpdating = True

End Sub
