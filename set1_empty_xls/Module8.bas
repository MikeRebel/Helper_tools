Attribute VB_Name = "Module8"
Sub ��������_��������()
Attribute ��������_��������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��������_�������� ������
'

'

MyRowNumber = InputBox("������� ���������� �����", "���������� �����", 1)
If IsNumeric(MyRowNumber) Then
    With ActiveCell
        MyAddress = .Cells.Row
        
    End With
    MyRow = MyAddress + MyRowNumber
    With Cells(MyRow, "D")
        .Value = "1"
    End With
    
    Range("A" & MyAddress, "C" & MyRow).Select
    
    
    
    Selection.Copy
    
End If

End Sub
Sub ����������_�_��������()
Attribute ����������_�_��������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ����������_�_�������� ������
'

'
    ActiveCell.FormulaR1C1 = "=""81""&RC[-4]"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A65000")
    ActiveCell.Range("A1:A65000").Select
    ActiveCell.Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    ActiveCell.Offset(0, -4).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, 5).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RAND()"
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A65000")
    With Range("F1")
        .Value = "rand"
    End With
    With Range("E1")
        .Value = "pref"
    End With
    
    Range("B2").Select
    Selection.End(xlDown).Select
    With ActiveCell
        MyRows = .Row + 1
    End With
    
    Rows(MyRows & ":" & MyRows).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    
    Columns("A:F").Select
  
    ActiveWorkbook.Worksheets("����1").Sort.SortFields.Add Key:=Range("F2:F65001" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("����1").Sort
        .SetRange Range("A1:F65001")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("E:F").Select
    Selection.Delete
    

End Sub
