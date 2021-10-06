Attribute VB_Name = "Module1"
Sub base_generator()
Randomize

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

Set Sourcewb = ActiveWorkbook
Set Destwb = ActiveWorkbook

S = Selection
TotalNumberOfDefs = UBound(S, 1)
Destwb.Sheets(1).Select
Range("A1").Select
TotalBaseLength = Selection.Value


Destwb.Sheets(3).Select
Columns("A").Select
Selection.Delete

For i = 1 To TotalBaseLength


    StartAdress = (Int(1 + (Rnd() * TotalNumberOfDefs)))
    If IsEmpty(S(StartAdress, 2)) Then
        MyNumber = (S(StartAdress, 1) * 10000000) + (Int(1000000 + (Rnd() * 1000000)))
    ElseIf IsNumeric(S(StartAdress, 2)) Then
        MyDef = S(StartAdress, 1) * 1000 + S(StartAdress, 2)
        MyNumber = (MyDef * 10000) + (Int(1 + (Rnd() * 1000)))
    End If

    Destwb.Sheets(3).Select
    Range("A" & i).Select
    Selection.Value = MyNumber
    
    If i > 2 Then
        For j = 1 To i - 1
            Range("A" & j).Select
            CheckedValue = Selection.Value
            If CheckedValue = MyNumber Then
                i = i - 1
                MyCounter = MyCounter + 1
            End If
            If MyCounter > 100000 Then
                Exit Sub
            End If
        Next j
        j = 1
    End If




Next i



    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With


End Sub


Private Function WorksheetIsExist(iName$) As Boolean
    On Error Resume Next
    WorksheetIsExist = (TypeOf Worksheets(iName$) Is Worksheet)
End Function


Function CreateSheet(sSName As String, bVisible As Boolean)
Dim wsNewSheet As Worksheet

On Error GoTo errÍandle

Set wsNewSheet = ActiveWorkbook.Worksheets.Add(after:=Worksheets(Worksheets.Count))
  With wsNewSheet
   .Name = sSName
   .Visible = bVisible
  End With
Exit Function
errÍandle:
  MsgBox Err.Descriðtion, vbExclamation, "Error #" & Err.Number

End Function
