Attribute VB_Name = "GeneralSpecific"

Sub ��������_��������()
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
    With Cells(MyRow, "I")
        .Value = "1"
    End With
    
    Range("A" & MyAddress, "H" & MyRow).Select
    
    
    
    Selection.Copy
    
End If

End Sub
Sub ����������_�_��������()
'
' ����������_�_�������� ������
' ��������� ������� 81 � �������� � ������������� ����.
' ���� �������� ������ ������� A. ���� �� ����� ��� ������� � 2 ����� ��������� � �������� �.�. ������� E � F ������������ ��� ����������� �������� � ����������.

'
Dim MyPrefix As String

    MyPrefix = InputBox("������� �������", "�������", 81)
    ActiveCell.FormulaR1C1 = "=" + MyPrefix + "&RC[1]"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A65000")
    ActiveCell.Range("A1:A65000").Select
    Selection.Copy
    ActiveCell.Offset(0, 4).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, 1).Range("A1").Select
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
    Range("E2:E65000").Select
   
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Columns("E:F").Select
    Selection.Delete
    Columns("A").Select
    Selection.Delete
    
End Sub



Sub ��������_��������()
'
' �������� �������� �� ������ ���� � �������� ������ 1 �� ������  �������� �������� �������.
' ������� ���� ����� �� ���������� ��� ����������� ���� ������� ��� ������� � ����� �� ������, ������ �� ���� �������������
' ��������� ������: Ctrl+e

MySourceWorkbookName = ActiveWorkbook.Name
Workbooks.Add
MyDestinationWorkbookName = ActiveWorkbook.Name
Windows(MySourceWorkbookName).Activate

MyRowNumber = InputBox("������� ���������� �����", "���������� �����", 200)
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


Sub add_prefix()
'
' add_7 ������
' �������� ������ ������� � ����� ��������
'
' ��������� ������: Ctrl+s
'
SelectionNumRows = Selection.Rows.Count

MyPrefix = InputBox("������� �������", "�������", 81)
If IsNumeric(MyRowNumber) Then

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

Sub Maxdiff_solution()




For i = 0 To 1000000

    Worksheets("����1").Range("D73").Select
    Selection.Value = i
    
    Columns("A:B").Sort key1:=Range("A1")
    
   S = Worksheets("����1").Range("AE61").Value
   If S = 0 Then Exit Sub
 
   
   
   Worksheets("����1").Range("D73").Value = i
   
    
    
    
Next i


End Sub

Sub PhoneBase_clearing()

' ���� ���� ���� � ����������� ����������, ��������� ������ ���������� ��� ������� �����������.

e = Selection
If IsArray(e) Then

For Each c In Selection.Rows
    c.Select
    CurrentRow = Selection
    If IsArray(CurrentRow) Then
    
    
        TrimmedRow = trimRow(CurrentRow)
        
        If IsArray(TrimmedRow) Then
            CleanedRow = cleanRow(TrimmedRow)
        Else
            MsgBox ("������")
        End If
        
        i = 1
        For Each d In Selection.Cells
            d.Select
            Selection.Value = CleanedRow(1, i)
            i = i + 1
        Next d
    End If
    

Next c

Else
MsgBox ("�������� ��������")
End If


End Sub

Function trimRow(data)
i = 1
For Each c In data
   
    If c < 1000000000 Then
        c = Empty
        data(1, i) = c
    End If
    
i = i + 1
Next c

trimRow = data

    

End Function

Function cleanRow(data)
i = 2
For Each c In data
    For j = i To UBound(data, 2)
        If c = data(1, j) Then
            data(1, j) = Empty
            
        End If
        
    
    Next j
    i = i + 1
Next c


For Each e In data
    For j = 1 To UBound(data, 2) - 1
        If data(1, j) = Empty Then
            data(1, j) = data(1, j + 1)
            data(1, j + 1) = Empty
        
        End If
    
    Next j
Next e
cleanRow = data

End Function



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



Sub hyperlinks_correction()
'
' hyperlinks_correction ������
'
'������ �������� ��������� � ������ B1 �� ���������, � ������ C1 � ������ ������ ����������� ��  �������� �����. �������� ���� ��������� �������� ����� � ���� ������ ������ ��������� �� ����� ����

'
    mybook = Application.ActiveWorkbook.Name
    MyWorkSheet = ActiveSheet.Name
    
    ReplaceString = ActiveSheet.Cells(1, 2).Value
    StringToReplace = ActiveSheet.Cells(1, 3).Value
    
    SplitReplaceString = Split(ReplaceString, " ")
    SplitStringToReplace = Split(StringToReplace, " ")
    
    If UBound(SplitReplaceString) > 0 Then
        ReplaceString = "'" & ReplaceString & "'"
    End If

    
    If UBound(SplitStringToReplace) > 0 Then
        StringToReplace = "'" & StringToReplace & "'"
    End If
    
'    W = WorksheetIsExist("Data")
'    If W = True Then Worksheets("Data").Delete
'    MyTempSheet = CreateSheet("Data", True)
'    Sheets(MyWorkSheet).Activate
'    ActiveSheet.Columns(1).Select
    
'    Selection.Copy
'    Sheets("Data").Select
'    ActiveSheet.Cells(1, 1).Select
'    ActiveSheet.Paste
'    With Selection
'        .MergeCells = False
'    End With
    
'   Sheets("Data").Select
    ActiveSheet.Columns(1).Select
    
    For Each c In ActiveSheet.Hyperlinks
        MyLink = c.SubAddress
        MyFound = InStr(MyLink, StringToReplace)
        If MyFound = 0 Then
            
            NewLink = Replace(MyLink, ReplaceString, StringToReplace)
            c.SubAddress = NewLink
        End If
    
    
    Next c
    
    
    
End Sub

Sub hyperlinks_range_update()

MsgBox "������ �������� ��������, �� ������� ��������� ����������� �� ��������� ���������� ����� ����� �� �������� �����. ����������� �������� ���������� ��� ����� ��������. �� �������� ��� �����������, ����������� �� ���� ������."

'
' hyperlinks_correction ������
'
    Dim DefaultValue As Long
    Dim DefaultValue1 As Long
    Dim StartSheet As Variant
    Dim StopSheet As Variant
    Dim StartAddress As Integer
    Dim StopAddress As Integer
    Dim flag1 As Integer
    Dim flag2 As Integer
    flag1 = 0
    flag2 = 0
    
    mybook = Application.ActiveWorkbook.Name
    ReplaceString = ActiveSheet.Cells(1, 2).Value
    ReplaceNumber = ActiveSheet.Cells(1, 3).Value
    
    With Application
        .ScreenUpdating = False
    End With
    DefaultValue = "1"
    DefaultValue1 = "1"
    DefaultValue2 = "A"
    DefaultValue3 = "C"
    StartSheet = InputBox("������� ����� ������� �����. �� ������ ����� ������ ���� ������� 1) ���-�� �����, � ������� ��������� �� ���� � B1. 2) ���������� ��������� ����� � ������ ����� � C1. ��� ������ ���� ���������� ��� ���� ������ �� ������ ������� �� ������ ����������.", "����� �����", DefaultValue)
    StopSheet = InputBox("������� ����� ���������� �����", "����� �����", DefaultValue1)
    FirstRange = InputBox("������� ����� ������� ������� � ���������", "����� ������ ���������", DefaultValue2)
    LastRange = InputBox("������� ����� ���������� ������� � ���������", "����� ����� ���������", DefaultValue3)
    
    If IsNumeric(StartSheet) And IsNumeric(StopSheet) Then
    
    For i = StartSheet To StopSheet
    Sheets(i).Activate
     With ActiveSheet
        'If you are in Page Break Preview Or Page Layout view go
        'back to normal view, we do this for speed
        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView
        'Turn off Page Breaks, we do this for speed
        .DisplayPageBreaks = False
     End With
'
    MyWorkSheet = ActiveSheet.Name


    ActiveSheet.Columns(1).Select
    
    For Each c In ActiveSheet.Hyperlinks
        MyLink = c.SubAddress
        HyperlinkFound = InStr(MyLink, MyWorkSheet)
        If HyperlinkFound = 0 Then ' 0 ���� �������� ������ ���������� �� ������ ����� >0 ���� �� ���� �� ���-��

            RangeStartPointPosition = InStr(MyLink, FirstRange)
            RangeStopPointPosition = InStr(MyLink, LastRange)
            Delimiter = ":"
            MySplitArray = Split(MyLink, ":")
            MySplitArrayStartLength = Len(MySplitArray(0))
            MySplitArrayStopLength = Len(MySplitArray(1))
            MyStartString = MySplitArray(0)
            MyStopString = MySplitArray(1)
            StartAddresLength = MySplitArrayStartLength - RangeStartPointPosition
            StopAddressLength = MySplitArrayStopLength - 1
            StartAddress = Right(MyStartString, StartAddresLength)
            StopAddress = Right(MyStopString, StopAddressLength)
            CorrectedStartAddress = 0
            CorrectedStopAddress = 0
            If StartAddress > ReplaceString Then
                CorrectedStartAddress = StartAddress - ReplaceNumber
            End If
            If StopAddress > ReplaceString Then
                CorrectedStopAddress = StopAddress - ReplaceNumber
            End If
            If StartAddress > ReplaceString Then
                NewLink = Replace(MyLink, StartAddress, CorrectedStartAddress)
                flag1 = 1
            End If
            If StopAddress > ReplaceString Then
                NewLink = Replace(NewLink, StopAddress, CorrectedStopAddress)
                flag2 = 1
            End If
            If flag1 = 1 Or flag2 = 1 Then
                c.SubAddress = NewLink
                flag1 = 0
                flag2 = 0
            End If
        End If
    
    Next c
Next i


    ActiveWindow.View = ViewMode

    Else
        MsgBox ("�� ����� �����-�� �����")
    End If
    
    With Application
        .ScreenUpdating = True
    End With
    
    
    
End Sub

Sub Pokraska_yacheiki_po_usloviyu()


'������ �������� �� �������� ����� � ������������ �� ���������� � ����� test � ���-�� ������

MyAddress = 0
MyValue = 0
DefaultValue = 0.95

Dim UsloviePokraski As Double
UsloviePokraski = InputBox("������� ����������� ��������", "����� ��������", DefaultValue)

MySheet = ActiveSheet.Name
 For Each i In Selection.Cells ' ��� ��������� �������� ��� ��������� � [] [diapazon]
    MyAddress = i.Address
    Sheets("test").Select
    Range(MyAddress).Select
    MyValue = Selection.Value
    If IsError(MyValue) = False Then
        If IsNumeric(MyValue) Then
            If MyValue > UsloviePokraski Then ' ������� ��������. ������ ���� ����
'               With Selection.Font
'                    .ThemeColor = xlThemeColorDark1
'                   .TintAndShade = 0
'               End With
                Sheets(MySheet).Select
                Range(MyAddress).Select
                With Selection.Interior
                    .ColorIndex = 3 ' ������� ������ http://www.automateexcel.com/2004/08/18/excel_color_reference_for_colorindex/
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                End With
            End If
        End If
    End If
Next i
Sheets(MySheet).Select
End Sub


Sub �����������_���������������_�_��������()
' ���� ����� ���������� ������ ������������� �� ���� ���� �� ���������.
'MySheet = ActiveSheet.Name
'    For I = 1 To 6
'Range("A1").Activate
'Selection.EntireRow.Insert
'    Next I
Range("att").Select
RowNumber = Selection.Columns.Count
InitialRow = RowNumber + 6
InsertRow = InitialRow + 1

For Each c In Selection.Columns
    MyAddress = c.Address
    MyCollumn = c.Column
    MyRow = c.Row
    c.Select
    Selection.Copy
    ActiveSheet.Columns(InsertRow).Activate
    ActiveSheet.Paste
    InsertRow = InsertRow + 2
     
Next c
FinalRow = InsertRow - 2




End Sub
Sub Combinatorial_incomplete_gererator()

' ������ ��������������� ���������� ��������� ��������
'
' ��� ������ ����� ������� ���� res ��� �����������. �������� �������� �� ����� �����.

Dim RangeArray() As Integer
Dim PositionArray() As Integer
Dim ResultArray() As Integer
MaxCombination = 3 ' ������������ ���������� ���������
PresentValue = 1 ' ��������, ������� ������ ���� � ���������
MaxYSize = 0 ' ����������� ������ ������� ��������� ��������
IterationCounter = 1 ' ����� ������ ��� ��������������� ���������� �� ����� res
CombinationCounter = 0 ' ������� ���������� ��������� � ��������������� ����������
MaxIterationCoutner = 1 ' ����������� ��������� ���������� �������� ���  ������� �������
PreviousIterationCounter = 0 ' ���������� �������� ���  ���������� �������, ����� ������ �����

WorkAreaXSize = Selection.Columns.Count '������ ������� ��������� �������
ReDim RangeArray(WorkAreaXSize)
ReDim PositionArray(WorkAreaXSize + 1)
ReDim ResultArray(WorkAreaXSize)
RangeArray(0) = 1 ' �� ������ ������, ������ ������� �� ������������ ��� ������������ �������� ������� � ������� ����� ������

For Each Column In Selection.Columns ' ���������� � ������ RangeArray, ������� ��������� �������� ������� ��������� ������� � ������ �������.

    Column.SpecialCells(xlCellTypeConstants).Select
    WorkAreaYSize = Selection.Cells.Count
    If WorkAreaYSize > MaxYSize Then
        MaxYSize = WorkAreaYSize
    End If
    RangeArray(Column.Column) = WorkAreaYSize
    
Next ' ������� ������ ������ ������� � ������ RangeArray

Range(Cells(1, 1), Cells(MaxYSize, WorkAreaXSize)).Select
ValueArray = Selection ' ��� ������ ��������, ������� ���� ���������. ( todo ����� ����������, ����� ����� �����������)


' ������ ������ ��������� �������, ��� �����, ������ ����� �������� � ������� ��������� �������. �������� ������� -  ������, ����� �������� ������� - ������.

    For i = 1 To WorkAreaXSize
        PositionArray(i) = 1
    Next i






    For i = 1 To WorkAreaXSize  '������� ��������, ������� ��������
    
    For e = 1 To i ' ��������� ������������ ���������� ���������� ��� ������ ��������, ��� ������������ ���-�� ������� �� ������ ������� �������.
    
        MaxIterationCoutner = MaxIterationCoutner * RangeArray(e) ' ���������� ��������� ���� ��������� � ��������� ������� ������� ������ ����� ������������ ��������� ��������� ������� � �������.
        
    Next e
    
    For y = 1 To MaxIterationCoutner '����������� � ��������� ������ ResultArray() �������� �� ������ ��  PositionArray()
        If PreviousIterationCounter < y Then  ' ��� ��������� ������� �� ������ ���������� �������� �� �����������, ����� �� ���� ������.
            For j = 1 To WorkAreaXSize
                ResultArray(j) = ValueArray(PositionArray(j), j) '����������� � ��������� ������ ResultArray() �������� �� ������ ��  PositionArray()
                If ResultArray(j) = PresentValue Then
                    CombinationCounter = CombinationCounter + 1 '  ������� ���������� ��������� �������� PresentValue, ������� �� ������ ���������� ������ ��� CombinationCounter ���
                End If
        

            Next j ' ����� ResultArray(), ����������� ����������
            
        
        ' ��������� ResultArray() �� ���� res � ������ � ������� IterationCounter ���� �� �������� �� �������� ������
            If CombinationCounter <= MaxCombination Then
                
                For d = 1 To WorkAreaXSize
                    Sheets("res").Cells(IterationCounter, d).Value = ResultArray(d)
                    
                Next d
                IterationCounter = IterationCounter + 1 ' ��������� ���������� ��������� ��������� �� ��������� ������
            End If
        
            
        
        End If
        
        CombinationCounter = 0 ' �������� ���������� ��������� ����� ��� �� ������ ������.
        
        
        PositionArray(1) = PositionArray(1) + 1 ' �������� �������� ����������� ���������� ������, ������ ����� ��������. ���������� � ������� �������� ������� � �������� � �� ��� ���, ���� ��� ���������� ������� � RangeArray() ����� ������ ��� ����� ��������� PositionArray()
        For ErrorCorrections = 1 To WorkAreaXSize
            If PositionArray(ErrorCorrections) > RangeArray(ErrorCorrections) Then
                PositionArray(ErrorCorrections) = 1
                'If IterationCounter <= MaxIterationCoutner Then ' ��������, ����� �� ���� ������ �� ���������� ������ ������� PositionArray ��� ���������� ����� ������� ���� ����� ��������� ��������
                PositionArray(ErrorCorrections + 1) = PositionArray(ErrorCorrections + 1) + 1
               ' End If
            End If
        Next ErrorCorrections
        
        
        
    Next y
    
    PreviousIterationCounter = MaxIterationCoutner '����������, ������� �������� ���������� � ��������� ����.
    
    MaxIterationCoutner = 1 '�������� ��������� ������� ��� ��������� ��������
    For S = 1 To WorkAreaXSize
        PositionArray(S) = 1
    Next S

Next i ' ��������� ��������


'        PositionArray(y) = 1
'    Next y
'        For j = 1 To RangeArray(i)
        
'            If i = cell.Column Then
            
                            ' ������ �� ������, ����� ��������� � ������� �������
                            
'            Else
'            Sheets("res").Select
'            Range(Cells(IterationCounter, 1), Cells(IterationCounter, WorkAreaXSize)).Select
'            PositionArray(i) = j
            
'            For Each MyCell In Selection.Cells
            
'                MyCell.Value = Sheets("data").Cells(PositionArray(MyCell.Column), MyCell.Column).Value
'                If MyCell.Value = PresentValue Then
'                    CombinationCounter = CombinationCounter + 1
'                End If
'            Next
'            If CombinationCounter > MaxCombination Then
'                IterationCounter = IterationCounter - 1
'                Range(Cells(IterationCounter, 1), Cells(IterationCounter, WorkAreaXSize)).Select
'                Selection.Delete
'           End If
            
            
            
'            IterationCounter = IterationCounter + 1
'            End If
        


End Sub


Sub Combinatorial_Full_gererator()
' ��������� ��� ���������� ���� ��������. ������ �������� �������.
' ��� ������ ����� ������� ���� res ��� �����������. �������� �������� ������ ���� �� ����� Data

Dim RangeArray() As Integer
Dim CombinationCount As Long
CombinationCount = 1
Dim PositionArray() As Integer
Dim i As Long

WorkAreaXSize = Selection.Columns.Count

ReDim RangeArray(WorkAreaXSize)
i = 1
For Each Column In Selection.Columns
    
    
    Column.SpecialCells(xlCellTypeConstants).Select
    WorkAreaYSize = Selection.Cells.Count
    RangeArray(i) = WorkAreaYSize
    CombinationCount = CombinationCount * RangeArray(i)
    If CombinationCount > 1000000 Then
        MsgBox ("���-�� ���������� � ������ ����� = " & CombinationCount & " �� ���������� � excel.")
       Exit Sub
    End If
    i = i + 1
Next



Range(Cells(1, 1), Cells(WorkAreaYSize, WorkAreaXSize)).Select
WorkAreaYSize = Selection.Rows.Count

MaxEntryCounter = CombinationCount / WorkAreaYSize

Sheets("res").Select
Range(Cells(1, 1), Cells(CombinationCount, WorkAreaXSize)).Select

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



ReDim PositionArray(WorkAreaXSize)
For i = 1 To WorkAreaXSize
    PositionArray(i) = 1
Next i

For Each MyRow In Selection.Rows
    MyRow.Select
    ColumnCount = 1
    For Each MyCell In Selection.Cells
        
        MyCell.Value = Sheets("Data").Cells(PositionArray(MyCell.Column), ColumnCount).Value
        If MyCell.Column = WorkAreaXSize Then
            PositionArray(MyCell.Column) = PositionArray(MyCell.Column) + 1
            For i = WorkAreaXSize To 1 Step -1
                If PositionArray(i) > RangeArray(i) Then
                    If i > 1 Then
                        PositionArray(i) = 1
                        PositionArray(i - 1) = PositionArray(i - 1) + 1
                    End If
                End If
            Next i
       End If
       ColumnCount = ColumnCount + 1
    Next
Next
    
    
'Range(Cells(1, 1), Cells(WorkAreaYSize, WorkAreaXSize)).Select

ExitTheSub:
With ActiveSheet
    ActiveWindow.View = ViewMode
End With
With Application
    .ScreenUpdating = True
End With

ActiveWorkbook.Save
    
    
    
End Sub


Sub ����������_�_��������()
'
' ���� ���������� ����������, ��� ���������� ������������� ������ ������ ��������� ������ ������� �� �������� ��������� � ����� ����
' ��� ����� ��������� ���������.

    DownBorder = 7 ' 7 ��� ������ ������� ��������� � ������� ��� ��������. �������� ���� ����� ������
    Dim DeleteAdressArray() As Variant ' ������ � �������� �������� ��� ��������
    DelI = 1 ' ������ �������
    OldValue = 0
For Each MyCol In Selection.Columns
    MyCol.Select
        
    MyCell = Cells(DownBorder, MyCol.Column).Select
    S = Selection
    
    If IsArray(S) = True Then
        
        MyValue = S(1, 1)
        
    Else
    
        MyValue = Selection.Value
    
    End If
    
    
    If OldValue = MyValue Then
    
        For i = 1 To 1000000
        
            TestedValue = Cells(i, MyCol.Column).Value
            Cells(i, MyCol.Column).UnMerge
            If IsArray(TestedValue) = False Then
                
                s1 = Split(TestedValue, "%")
                If UBound(s1) = 1 Then
                    
                    Cells(i, MyCol.Column + 1).Select
                    p = Paint_Diff()
                    
                End If

            
            End If
            If TestedValue = "" Then
                StopCounter = StopCounter + 1
                If StopCounter > 100 Then
                    StopCounter = 0
 
                    Columns(MyCol.Column).Select
                    ReDim Preserve DeleteAdressArray(DelI)
                    DeleteAdressArray(DelI) = MyCol.Column
                    DelI = DelI + 1
                    Exit For
                End If
            End If
            
            
        Next i
    
    End If
    
    OldValue = MyValue
    



Next MyCol




For i = UBound(DeleteAdressArray) To 1 Step -1

    Columns(DeleteAdressArray(i)).Select
    d = Delete_col()

Next i


'
End Sub


Function Paint_Diff()
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
End Function


Function Delete_col()

    Selection.Delete Shift:=xlToLeft
    
End Function


Sub �������_��������_��������_����������()

For Each MyCell In Selection.Cells
    MyCell.Select
 
    MyValue = Selection.Value
    If IsArray(MyValue) = False Then
        If Selection.Interior.Color = 13421772 Then
            Selection.Value = ""
        End If
        
        If Selection.Interior.Color = 255 Then
            s1 = Split(MyValue, "%")
            If UBound(s1) = 1 Then
            
            ElseIf UBound(s1) = 0 Then
            
                s1 = Split(MyValue, " ")
            
            End If
            Selection.Value = s1(1)
            
            Erase s1
   
        End If
    
    End If
    
    
    


Next MyCell




End Sub

Sub ��������_���������()
Dim TotalPosition()
Dim LabelStartPosition()
a = 1
B = 1
MyTotalDifference = 0

e = Selection
If IsArray(e) Then
    EmptySpaceStartCol = UBound(e, 2) + 1
    EmptySpaceStartRow = UBound(e, 1) + 1
    For Each MyCol In Selection.Columns
        
        For Each MyRow In MyCol.Rows
            CurrentIndexCol = MyRow.Column
            CurrentIndexRow = MyRow.Row
            If CurrentIndexCol = 1 Then
                If MyRow.Formula = "Total" Then
                    i = i + 1
                    ReDim Preserve TotalPosition(i)
                    TotalPosition(i) = CurrentIndexRow
                End If
            End If
            
            If CurrentIndexCol = 2 Then
                
                If Not IsEmpty(MyRow.Value2) And Not MyRow.Value2 = "" And PreviousValue = "" Then
                    j = j + 1
                    ReDim Preserve LabelStartPosition(j)
                    LabelStartPosition(j) = CurrentIndexRow
                    
                End If
                PreviousValue = MyRow.Value2
                         
            End If
            
             If CurrentIndexCol > 2 Then
                    
                If MyRow.NumberFormat = "0%" And CurrentIndexRow >= LabelStartPosition(a) And CurrentIndexRow < TotalPosition(a) Then
                    CurrentProcent = e(CurrentIndexRow, CurrentIndexCol)
                    CurrentTotal = e(TotalPosition(a), CurrentIndexCol)
                    CurrentCount = CurrentProcent * CurrentTotal
                    MySumm = MySumm + CurrentCount
                    MyProcentSumm = MyProcentSumm + CurrentProcent
                    Cells(CurrentIndexRow, EmptySpaceStartCol + B).Value = CurrentCount
                  
                End If
                If CurrentIndexRow = TotalPosition(a) Then
                    
                    Cells(CurrentIndexRow, EmptySpaceStartCol + B).Value = MySumm
                    MyTableCounts = e(CurrentIndexRow, CurrentIndexCol)
                    MyCountDifference = MySumm - MyTableCounts
                    If MyCountDifference <= 1 And MyCountDifference >= -1 Then
                        MyCountDifference = 0
                    End If
                    If MyCountDifference <> 0 And MyProcentSumm > 0 Then
                        MyMultipleProcent = MySumm / MyProcentSumm
                        MyMultipleProcentDifference = MyMultipleProcent - MyTableCounts
                        If MyMultipleProcentDifference <= 1 And MyMultipleProcentDifference >= -1 Then
                            MyCountDifference = 0
                        End If
                    End If
                    If MyCountDifference <> 0 And MyProcentSumm = 0 Then
                        MyCountDifference = 0
                        Cells(CurrentIndexRow + 1, EmptySpaceStartCol + B).Interior.ColorIndex = 12
                    End If
                    
                    MyTotalDifference = MyTotalDifference + MyCountDifference
                    If MyTableCounts > Round(MySumm) Then
                        Cells(CurrentIndexRow, EmptySpaceStartCol + B).Interior.ColorIndex = 10
                    End If
                    If MyTableCounts < Round(MySumm) Then
                        Cells(CurrentIndexRow, EmptySpaceStartCol + B).Interior.ColorIndex = 9
                    End If
                    Cells(CurrentIndexRow + 1, EmptySpaceStartCol + B).Value = MyCountDifference
                    CurrentProcent = 0
                    CurrentTotal = 0
                    CurrentCount = 0
                    MyProcentSumm = 0
                    MySumm = 0
                    MyCountDifference = 0
                    MyMultipleProcent = 0
                    MyMultipleProcentDifference = 0
                    MyTableCounts = 0
                    
                    a = a + 1
                    
                    
                End If
             End If
             
        Next MyRow
        If CurrentIndexCol > 2 Then
        a = 1
        B = B + 1
        End If
    
    Next MyCol
    
Cells(EmptySpaceStartRow + 1, EmptySpaceStartCol).Value = MyTotalDifference
If EmptySpaceStartCol > 0 Then
    Cells(EmptySpaceStartRow + 1, EmptySpaceStartCol).Interior.ColorIndex = 9
End If
End If


End Sub


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



Sub �������_��������_�_������_������150()


'������ �������� �� �������� ����� � ������������ �� ���������� � ����� test � ���-�� ������

Set S = Application.InputBox("�������� ��������, �� �������� ��������� �������� � ��������, �� ����� �����. ��������������, ��� �� ���������� �� ������ �����. ���� �������� ������, ������� ������������. ��������������, ��� � ��������� ��� ������, ����� ���� � ���������.", "����� ���������", _
 ActiveCell.Address, Type:=8)
MyAddress = 0
MyValue = 0
DefaultValue = 150
Dim UslovieUdaleniya As Integer
UslovieUdaleniya = InputBox("������� ����������� �������� ����, ���� �������� ������ � �������� ���������. ��������������, ��� �������� ������.", "����� ��������", DefaultValue)

SelectionNumRows = S.Rows.Count
SelectionNumCols = S.Columns.Count
SelectionStartRow = S.Row
SelectionEndRow = SelectionStartRow + SelectionNumRows - 1
SelectionStartCol = S.Column
SelectionEndCol = SelectionStartCol + SelectionNumCols - 1

For Each c In ActiveWorkbook.Sheets
    MyCurrentSheetName = c.Name
 
    Sheets(MyCurrentSheetName).Activate

    Range(Cells(SelectionStartRow, SelectionStartCol), Cells(SelectionEndRow, SelectionEndCol)).Select


    


    
        For Each i In Selection.Cells ' ��� ��������� �������� ��� ��������� � [] [diapazon]
            MyAddress = i.Address
            MySplitAddress = Split(MyAddress, "$")
            MyPreviousAddress = MySplitAddress(2) - 1
            MyValue = i.Value
            If IsError(MyValue) = False Then
                If IsNumeric(MyValue) Then
                    If i.NumberFormatLocal = "0%" Then
                    Else
                        If MyValue < UslovieUdaleniya Then ' ������� ��������. ������ ���� ����
'               With Selection.Font
'                    .ThemeColor = xlThemeColorDark1
'                   .TintAndShade = 0
'               End With
                        
                        Range(MyAddress).Select
                        Selection.Value = "<" & UslovieUdaleniya
                        With Selection.Interior
                            .ColorIndex = 2 ' ������� ������ http://www.automateexcel.com/2004/08/18/excel_color_reference_for_colorindex/
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                        End With
                        Range(MySplitAddress(1) & MyPreviousAddress).Select
                        Selection.Value = ""
                        With Selection.Interior
                            .ColorIndex = 2 ' ������� ������ http://www.automateexcel.com/2004/08/18/excel_color_reference_for_colorindex/
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            
                        End With
                        Selection.ClearNotes
                    End If
                End If
        End If
    End If
Next i
Next c
End Sub

Sub �������������_����������_��������()
'
' ������1 ������
'

' ����������� ���������� �������� � ������ ������� (�����������, ������, ���������, ���) �� ����������������� ������� spss

Dim sdvig1 As Long
Dim start1 As Long
Dim CorrelationLabel As String
sdvig1 = 5 '���-�� �����, ������� ������������ ����� �������
start1 = 1 '��������� ������
start1sdvig = 2 '������ �� ���������� �����

LabelCol = 13 '���-�� ������� �������, ���� ����������� ���������
ValueCol = 14 '���-�� ������� �������, ���� ����������� ��������� �����

For Each c In Selection.Cells
 If c.Value = "Correlation Coefficient" Then
    c.Select
    Cells(c.Row, c.Column - 1).Select
    CorrelationLabelArray = Selection.Value
    CorrelationLabel = CorrelationLabelArray(1, 1) '����� �����
    Cells(c.Row, c.Column + 1).Select
    CorrelationValue = Selection.Value ' �������� �����
    
    Cells(start1, LabelCol).Select
    Selection.Value = CorrelationLabel
    Cells(start1, ValueCol).Select
    Selection.Value = CorrelationValue
    start1 = start1 + sdvig1
    
    If start1 > 25 Then ' �������� � ����������� �� ���������� ���������� � ���������� ����� ���-�� ���������� � ���������� �������� �� 5
        start1 = start1sdvig
        start1sdvig = start1sdvig + 1
    End If

    If start1sdvig = 6 Then ' ���� ��� ��������� �������� ���� � ��������� �������
        LabelCol = LabelCol + 2
        ValueCol = ValueCol + 2
        start1 = 1
        start1sdvig = 2
    End If
End If
 

'    Rows(RowStart1 & ":" & RowEnd1).Select
'    Selection.Delete Shift:=xlUp

'    Rows(RowStart2 & ":" & RowEnd2).Select
'    Selection.Delete Shift:=xlUp
    
'    RowStart1 = RowStart1 + 26
'    RowEnd1 = RowStart1 + sdvig1 - 1

'    RowStart2 = RowStart2 + 26
'    RowEnd2 = RowStart2 + sdvig2 - 1
    
Next c
    
End Sub


Sub ���������_������()
Dim MonthNamesArray(6, 31) As Variant
Dim VacationArray(30) As Variant
MyWorkbookName = ActiveWorkbook.Name
MySheetName = ActiveSheet.Name
i = 0
j = 1


For Each MyRows In Selection.Rows

    If MyRows.Row = 1 Then ' ���������� �������� �������, ������� � ��� ���� � � ����� ������� ��������� ����.
        Worksheets(MySheetName).Rows(1).Select
        For Each HeaderMonth In Selection.Cells
            If HeaderMonth.Column > 2 Then
                
                CurrentValue = HeaderMonth.Value
                
                If CurrentValue <> "" And i < 7 Then
                    If j > 1 Then
                        j = 1
                    End If
                    i = i + 1
                    MonthNamesArray(i, 0) = CurrentValue
                    
                    
                    
                End If
                If j > 31 Then
                    
                    FinalRowIndex = MonthNamesArray(i, j - 1)
                    Exit For
                    
                End If
                
                MonthNamesArray(i, j) = HeaderMonth.Column
                j = j + 1
                
            End If
            
        Next HeaderMonth
        
    ElseIf MyRows.Row > 3 Then
            
            Worksheets(MySheetName).Rows(MyRows.Row).Select
            v = 1
            For Each MyDay In Selection.Cells
                
                DayValue = MyDay.Value
                If DayValue = 1 Then
                    VacationColIndex = MyDay.Column
                    VacationDayCounter = VacationDayCounter + DayValue
                    
                    
                End If
                If VacationDayCounter = 1 Then
                    
                    For i = 1 To 6
                        For j = 1 To 31
                            
                            If VacationColIndex = MonthNamesArray(i, j) Then
                                VacationStartMonth = MonthNamesArray(i, 0)
                                VacationArray(v + 1) = VacationStartMonth
                                VacationArray(v) = Cells(2, MyDay.Column).Value
                                v = v + 2
                                
                            End If
                        Next j
                    Next i
                ElseIf VacationDayCounter > 1 And DayValue = "" Then
                    
                    
                    For i = 1 To 6
                        For j = 1 To 31
                    
                            If VacationColIndex = MonthNamesArray(i, j) Then
                                VacationStopMonth = MonthNamesArray(i, 0)
                                VacationArray(v + 2) = VacationDayCounter
                                VacationArray(v + 1) = VacationStopMonth
                                VacationArray(v) = " - " & Cells(2, MyDay.Column - 1).Value
                                v = v + 3
                            End If
                            
                        Next j
                    Next i
                            
                    VacationStartMonth = ""
                    VacationStopMonth = ""
                    VacationDayCounter = 0
                    VacationColIndex = 0
                    
                    
                End If
                If MyDay.Column = FinalRowIndex And VacationArray(1) <> "" Then
                    For i = 1 To 30
                        Select Case VacationArray(i)
                            Case "������"
                                VacationArray(i) = ".01.2020"
                            Case "�������"
                                VacationArray(i) = ".02.2020"
                            Case "����"
                                VacationArray(i) = ".03.2020"
                            Case "������"
                                VacationArray(i) = ".04.2020"
                            Case "���"
                                VacationArray(i) = ".05.2020"
                            Case "����"
                                VacationArray(i) = ".06.2020"
                            Case "����"
                                VacationArray(i) = ".07.2020"
                            Case "������"
                                VacationArray(i) = ".08.2020"
                            Case "��������"
                                VacationArray(i) = ".09.2020"
                            Case "�������"
                                VacationArray(i) = ".10.2020"
                            Case "������"
                                VacationArray(i) = ".11.2020"
                            Case "�������"
                                VacationArray(i) = ".12.2020"
             
         
                End Select
                    
                        Cells(MyDay.Row, FinalRowIndex + i).Select
                        MyValue = VacationArray(i)
                        Selection.Value = MyValue
                        VacationArray(i) = Empty
                        
                    Next i
                    Exit For
                End If
            Next MyDay
    End If
    
    
Next MyRows



End Sub


Function �����(sWord As String)
    Dim sSymbol As String, sInsertWord As String
    Dim i As Integer
 
       sInsertWord = ""
    sSymbol = ""
    For i = 1 To Len(sWord)
        sSymbol = Mid(sWord, i, 1)
        sSymbol2 = Mid(sWord, i + 1, 1)
        sSymbol3 = Mid(sWord, i + 2, 1)
        sSymbol4 = Mid(sWord, i + 3, 1)
        sSymbol5 = Mid(sWord, i + 4, 1)
        sSymbol6 = Mid(sWord, i + 5, 1)
        sSymbol7 = Mid(sWord, i + 6, 1)
        sSymbol8 = Mid(sWord, i + 7, 1)
        sSymbol9 = Mid(sWord, i + 8, 1)
        sSymbol10 = Mid(sWord, i + 9, 1)
        If LCase(sSymbol) = 9 And LCase(sSymbol2) Like "*[0-9]*" And LCase(sSymbol3) Like "*[0-9]*" And LCase(sSymbol4) Like "*[0-9]*" And LCase(sSymbol5) Like "*[0-9]*" And LCase(sSymbol6) Like "*[0-9]*" And LCase(sSymbol7) Like "*[0-9]*" And LCase(sSymbol8) Like "*[0-9]*" And LCase(sSymbol9) Like "*[0-9]*" And LCase(sSymbol10) Like "*[0-9]*" Then
        
       ����� = Mid(sWord, i, 10)

        End If
    Next i

End Function

