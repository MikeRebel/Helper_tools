Attribute VB_Name = "Module2"
Sub add_prefix()
Attribute add_prefix.VB_Description = "Добавить префикс 7 в телефон"
Attribute add_prefix.VB_ProcData.VB_Invoke_Func = " \n14"
'
' add_7 Макрос
' Добавить префикс 7 в телефон
'
' Сочетание клавиш: Ctrl+s
'
SelectionNumRows = Selection.Rows.Count

MyPrefix = InputBox("Введите префикс", "Префикс", 81)
If IsNumeric(MyRowNumber) Then

'''

    Range("B1").Select
    ActiveCell.FormulaR1C1 = MyPrefix
    Range("B2").Select
    ActiveCell.FormulaR1C1 = MyPrefix
    Range("B1:B2").Select
    Selection.AutoFill Destination:=Range("B1:B" & SelectionNumRows)
    Range("B1:B" & SelectionNumRows).Select
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]&RC[-2]"
    Range("C1").Select
    Selection.AutoFill Destination:=Range("C1:C" & SelectionNumRows)
    Range("C1:C" & SelectionNumRows).Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:C").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    
    End If
    
End Sub
Sub выделить_диапазон_2()
Attribute выделить_диапазон_2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' выделить_диапазон_2 Макрос
'
' Сочетание клавиш: Ctrl+e

MySourceWorkbookName = ActiveWorkbook.Name
Workbooks.Add
MyDestinationWorkbookName = ActiveWorkbook.Name
Windows(MySourceWorkbookName).Activate

MyRowNumber = InputBox("Введите количество строк", "Количество строк", 200)
If IsNumeric(MyRowNumber) Then
 For Each c In ActiveWorkbook.Sheets
 MyCurrentSheetName = c.Name
 
 Sheets(MyCurrentSheetName).Activate
' C.Columns.Select
 
    With c.Application.ActiveCell
        MyAddress = c.Application.ActiveCell.Cells.Row
        
    End With
    MyRow = MyAddress + MyRowNumber - 1
    With Cells(MyRow, "B")
        .Value = "1"
    End With
    
    Range("A" & MyAddress, "A" & MyRow).Select
    Selection.Copy
    Windows(MyDestinationWorkbookName).Activate
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    MyCurrentAddress = Selection.Row
    MyCurrentAddress = MyCurrentAddress + 1
    Range("A" & MyCurrentAddress).Select
    Windows(MySourceWorkbookName).Activate
    MyRow = MyRow + 1
    Range("A" & MyRow).Activate
    
    Next c
    
Windows(MyDestinationWorkbookName).Activate

End If
    

End Sub
