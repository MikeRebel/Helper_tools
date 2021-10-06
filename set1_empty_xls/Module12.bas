Attribute VB_Name = "Module12"

Sub Макрос1()
Attribute Макрос1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос1 Макрос
'

'
    d = ActiveCell
    w = ActiveCell.Value
    ActiveCell.Formula = w
    
    
    Windows("Книга1").Activate
    ActiveCell.FormulaR1C1 = "=SUMIF(var_1,RC[-6],a_1)"
    Range("F1").Select
    Selection.Copy
    Range("G1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
