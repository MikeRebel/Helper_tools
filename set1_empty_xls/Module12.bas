Attribute VB_Name = "Module12"

Sub ������1()
Attribute ������1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������1 ������
'

'
    d = ActiveCell
    w = ActiveCell.Value
    ActiveCell.Formula = w
    
    
    Windows("�����1").Activate
    ActiveCell.FormulaR1C1 = "=SUMIF(var_1,RC[-6],a_1)"
    Range("F1").Select
    Selection.Copy
    Range("G1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
