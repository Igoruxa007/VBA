Attribute VB_Name = "Module11"
Sub PZ()
Attribute PZ.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �������� ����������������� ������� �� ���
' ���� ������� �����

' Application.ScreenUpdating = False

Dim arrOld(15) As String
    arrOld(0) = "�"
    arrOld(1) = "���"
    arrOld(2) = "��"
    arrOld(3) = "��"
    arrOld(4) = "��������"
    arrOld(5) = "��� + �����������"
    arrOld(6) = "�������� ���������"
    arrOld(7) = "�����"
    arrOld(8) = "���"
    arrOld(9) = "��� (� ����. ����������)"
    arrOld(10) = "��-1"
    arrOld(11) = "��-2"
    arrOld(12) = ""
    arrOld(13) = ""
    arrOld(14) = ""
    arrOld(15) = ""
    
Dim arrNew(15) As String
    arrNew(0) = "������. "
    arrNew(1) = "������������ ������������. "
    arrNew(2) = "������� ������. "
    arrNew(3) = "����������� ������. "
    arrNew(4) = "��������. "
    arrNew(5) = "������������ ������������. "
    arrNew(6) = "�������� ���������. "
    arrNew(7) = "��������� ������������� ��������. "
    arrNew(8) = "��������� ���������� �����������. "
    arrNew(9) = "�������������� ��������. "
    arrNew(10) = "������� ������. "
    arrNew(11) = "������� ������. "
    arrNew(12) = ""
    arrNew(13) = ""
    arrNew(14) = ""
    arrNew(15) = ""

'' ������� ���������
'    Rows("1:10").Select
'    Selection.Delete Shift:=xlUp
'
'' ������� ���������
'    Columns("K:AO").Select
'    Selection.Delete Shift:=xlToLeft
'
'' ������� ������ �������
'
'    Columns("E:H").Select
'    Selection.Delete Shift:=xlToLeft
'    Columns("A:B").Select
'    Selection.Delete Shift:=xlToLeft

count_of_records = 1
Do While Cells(count_of_records + 1, "C") <> ""
    count_of_records = count_of_records + 1
Loop
    
'' ������ �����������
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

' �������� �������
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
