Attribute VB_Name = "MarketSimulator"
Private MyOptions() As Variant



Function CreateSheet(sSName As String, bVisible As Boolean)
Dim wsNewSheet As Worksheet

On Error GoTo err�andle

Set wsNewSheet = ActiveWorkbook.Worksheets.Add(after:=Worksheets(Worksheets.Count))
  With wsNewSheet
   .Name = sSName
   .Visible = bVisible
  End With
Exit Function
err�andle:
  MsgBox Err.Descri�tion, vbExclamation, "Error #" & Err.Number

End Function

Private Function WorksheetIsExist(iName$) As Boolean
    On Error Resume Next
    WorksheetIsExist = (TypeOf Worksheets(iName$) Is Worksheet)
End Function




Sub Market_simulator_rows()
'

' ���������� ������ ���������� ��������� � ��������� ���������
'
Dim MyStep As Long '��������� ������� ��� ���������� �����������. ������ ������� �����������
Dim MySdvig As Long '�������� ������� ��� ���������� �����������. ��������� ������� �����������
MyCounter = 0 '������� ��� ���������� �����

DefaultValue = "2"
DefaultValue1 = "5"
MyStep = InputBox("������� ����� ������ �� ����� comb, ������� � ������� ����������� ������ � ������������ ����������", "����� ������", DefaultValue)
MySdvig = InputBox("������� ���������� ����� �� ����� interface, �� ������� ���������� ����������� ��������", "����� ������", DefaultValue1)

    With Application
        'we do this for speed
        .ScreenUpdating = False
    End With
    'If you are in Page Break Preview Or Page Layout view go
    'back to normal view, we do this for speed
    With ActiveSheet

        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView

    'Turn off Page Breaks, we do this for speed
        .DisplayPageBreaks = False
    End With
S = Selection

SelectionCols = 0
For Each MyRow In Selection.Rows
    SelectionCols = SelectionCols + 1
Next MyRow
SelectionEndRow = UBound(S) + MyStep - 1
SelectionEndCol = UBound(S, 2)
SelectionStartCol = SelectionEndCol - SelectionCols

If UBound(S) Mod MySdvig = 0 Then
    For MyRange = MyStep To SelectionEndRow Step MySdvig
        Range(Cells(MyRange, SelectionStartCol), Cells(MyRange + MySdvig - 1, SelectionEndCol)).Select '�������� ��������
        
        Selection.Copy
        Sheets("Interface").Select
        Range("Market").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("Simulation").Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("comb").Select
        ActiveSheet.Cells(MyStep, 15).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        MyStep = MyStep + MySdvig
    Next MyRange
End If


ExitTheSub:
With ActiveSheet
    ActiveWindow.View = ViewMode
End With
With Application
    .ScreenUpdating = True
End With

ActiveWorkbook.Save

End Sub




Sub Label_replace()
e = ActiveSheet.Name
w = WorksheetIsExist("data")

For Each i In Selection.Rows
            i.Select
            CurCellValue = Selection.Value
            If w = True Then
                Sheets("data").Select
                Worksheets("data").Columns("C").Replace What:=CurCellValue(1, 1), Replacement:=CurCellValue(1, 2), LookAt:=xlWhole, SearchOrder:=xlByColumns, MatchCase:=True
                Sheets(e).Select
            End If
Next i
End Sub

Sub Market_simulator_for_mult_product()
'

' ���������� ������ ���������� ��������� � ��������� ��������� ��� ��������� ������������ ���������.
'
Dim MyStartCombination As Long '��������� ������� ��� ���������� �����������. ������ ������� ����������
Dim MyStopCombination As Long '�������� ������� ��� ���������� �����������. ��������� ������� ����������

For Each iName In ActiveWorkbook.Names
iName.Delete '�������� �����
Next

S = Selection
Selection.Name = "Market"
MyMarketSize = UBound(S, 2)

DefaultValue = "3"
DefaultValue1 = "3"
MyStartCombination = InputBox("������� �����  �� ����� comb, ������� � �������� ����������� ������ � ������������", "����� ������", DefaultValue)
MyStopCombination = InputBox("������� �����  ��������� ���������� �� ����� comb", "����� ������", DefaultValue1)



    With Application
        'we do this for speed
        .ScreenUpdating = False
    End With
    'If you are in Page Break Preview Or Page Layout view go
    'back to normal view, we do this for speed
    With ActiveSheet

        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView

    'Turn off Page Breaks, we do this for speed
        .DisplayPageBreaks = False
    End With

Sheets("comb").Activate


' ����������� ������ ���������� ��� ��������
For i = 1 To 100000
    MyValue = Range("A" & i).Value
    If MyValue = MyStartCombination Then
        MyStartCombinationRow = i ' ����� ������, � ������� ��������� ������ ����������
        Exit For
    End If
Next

MyEndCombinationRow = MyStartCombinationRow + (MyStopCombination - MyStartCombination) ' ����� ������, � ������� ��������� ��������� ����������


Sheets("interface").Activate

With ActiveSheet
    ActiveWindow.View = ViewMode
End With
With Application
    .ScreenUpdating = True
End With



Set S = Application.InputBox("�������� �������� � ������������ ������� ���� �����", "����� �������", _
 ActiveCell.Address, Type:=8)
 


    With Application
        'we do this for speed
        .ScreenUpdating = False
    End With
    'If you are in Page Break Preview Or Page Layout view go
    'back to normal view, we do this for speed
    With ActiveSheet

        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView

    'Turn off Page Breaks, we do this for speed
        .DisplayPageBreaks = False
    End With
 
 


S.Select
Selection.Name = "Simulation"
S = Selection

SimulationEndRow = UBound(S)
SimulationEndCol = UBound(S, 2)

SimulationFullSize = SimulationEndRow * SimulationEndCol '��������� ������ ������� �����������


On Error GoTo ExitTheSub



For MyCombCounter = MyStartCombinationRow To MyEndCombinationRow


    
    
        Sheets("comb").Activate
    
        Range(Cells(MyCombCounter, 2), Cells(MyCombCounter, MyMarketSize + 1)).Select '�������� ��������
        
        Selection.Copy
        Sheets("Interface").Select
        Range("Market").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("Simulation").Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("comb").Select
        ActiveSheet.Cells(MyCombCounter, MyMarketSize + 2).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        f = 0
        If SimulationEndRow > 1 Then
        
            For j = MyCombCounter + 1 To (MyCombCounter + SimulationEndRow - 1) ' ��������� ���������� �� ���������� ����� � ���� ���� ����� ���������.
                Range(Cells(j, MyMarketSize + 2), Cells(j, MyMarketSize + 1 + SimulationEndCol)).Select
            
                f = f + 1
                Selection.Copy
                ActiveSheet.Cells(MyCombCounter, MyMarketSize + 2 + (SimulationEndCol * f)).Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Next j
            Range(Cells(MyCombCounter + 1, MyMarketSize + 2), Cells(j, MyMarketSize + 1 + SimulationEndCol)).Select
            Selection.Delete
        End If
        
Next MyCombCounter


ExitTheSub:
With ActiveSheet
    ActiveWindow.View = ViewMode
End With
With Application
    .ScreenUpdating = True
End With


End Sub

