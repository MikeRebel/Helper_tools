Attribute VB_Name = "OldSimpleSimulator"
Sub perepbor_simulatora_column()
'

' Yota ������
'
Dim step As Integer
Dim sdvig As Integer

MsgBox "!!!��������!!! ����� �������� ������ ���������� �����: �������� ��������, ������� ���� ���������������, ������ ���� ������� ����� �������� �������. ������ �������� � �������������� ������ ������ ���������� �� ����� comb. ����, � ������� ���������� ���������, ������ ���������� interface. �������� �����, ���� ����������� �������� ��������� ������ ���������� Market. ��������, ������ ���������� ����������� ���� ������ ���������� Simulation. ����� ������ ������ � �������"
DefaultValue = "1"
DefaultValue1 = "1"
step = InputBox("������� ����� ������� �� ����� comb, ������� � ������� ����������� ������� � ������������ ����������", "����� �������", DefaultValue)
sdvig = InputBox("������� ���������� ������� �� ����� interface, �� ������� ���������� ����������� ��������", "����� �������", DefaultValue1)



For Each c In Selection.Columns
        c.Columns.Select
        Selection.Copy
    
    
'   Sheets("comb1").Select
'    Range("D" & 5 & ":D" & 127).Select
'    Selection.Copy
    Sheets("interface").Select
    Range("Market").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("Simulation").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("comb").Select
    
    
    ActiveSheet.Cells(2, step).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("comb").Select
    step = step + sdvig
    
Next c

ExitTheSub:

End Sub

Sub perepbor_simulatora_row()
'

' ���������� ������ ���������� ��������� � ��������� ���������
'
Dim MyStep As Long '��������� ������� ��� ���������� �����������. ������ ������� �����������
Dim MySdvig As Long '�������� ������� ��� ���������� �����������. ��������� ������� �����������
Dim shag As Long '���������� ����� � ���������� (��� ������� ��������� � �����������)
MyCounter = 0 '������� ��� ���������� �����

DefaultValue = "2"
DefaultValue1 = "5"
MyStep = InputBox("������� ����� ������ �� ����� comb, ������� � ������� ����������� ������ � ������������ ����������", "����� ������", DefaultValue)
MySdvig = InputBox("������� ���������� ����� �� ����� interface, �� ������� ���������� ����������� ��������", "����� ������", DefaultValue1)

    With Application
        'we do this for speed
        '.ScreenUpdating = False
    End With
    'If you are in Page Break Preview Or Page Layout view go
    'back to normal view, we do this for speed
    With ActiveSheet

        'ViewMode = ActiveWindow.View
        'ActiveWindow.View = xlNormalView

    'Turn off Page Breaks, we do this for speed
        '.DisplayPageBreaks = False
    End With
S = Selection
SelectionEndRow = UBound(S) + MyStep - 1
If UBound(S) Mod 5 = 0 Then
    For MyRange = MyStep To SelectionEndRow Step 5
        Range("B" & MyRange, "N" & MyRange + 4).Select '�������� ��������
        
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

For Each c In Selection.Rows
        c.Rows.Select
        Selection.Copy
        Sheets("simulation").Select
        shag = step + sdvig - 1
        For i = step To shag
            ActiveSheet.Cells(i, 1).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Next i
        
    Sheets("interface").Select
    Range("Market").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Range("Simulation").Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Sheets("simulation").Select
    ActiveSheet.Cells(step, 64).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Sheets("comb").Select
    step = step + sdvig
    MyCounter = MyCounter + 1
    MySaveCounter = MyCounter Mod 10000 ' ��������� ������ � ����������
    If MySaveCounter = 0 Then
        ActiveWorkbook.Save
    End If

Next c

ExitTheSub:
With ActiveSheet
    ActiveWindow.View = ViewMode
End With
With Application
    .ScreenUpdating = True
End With

ActiveWorkbook.Save

End Sub

