Attribute VB_Name = "MarketSimulator"
Private MyOptions() As Variant



Function CreateSheet(sSName As String, bVisible As Boolean)
Dim wsNewSheet As Worksheet

On Error GoTo errНandle

Set wsNewSheet = ActiveWorkbook.Worksheets.Add(after:=Worksheets(Worksheets.Count))
  With wsNewSheet
   .Name = sSName
   .Visible = bVisible
  End With
Exit Function
errНandle:
  MsgBox Err.Descriрtion, vbExclamation, "Error #" & Err.Number

End Function

Private Function WorksheetIsExist(iName$) As Boolean
    On Error Resume Next
    WorksheetIsExist = (TypeOf Worksheets(iName$) Is Worksheet)
End Function




Sub Market_simulator_rows()
'

' Перебераем список комбинаций конжоинта и сохраняем результат
'
Dim MyStep As Long 'Начальная позиция для сохранения результатов. Первая строчка результатов
Dim MySdvig As Long 'Конечная позиция для сохранения результатов. Последняя строчка результатов
MyCounter = 0 'Счетчик для сохранения файла

DefaultValue = "2"
DefaultValue1 = "5"
MyStep = InputBox("Введите номер строки на листе comb, начиная с которой вставляются строки с расчитанными значениями", "Выбор строки", DefaultValue)
MySdvig = InputBox("Введите количество строк на листе interface, из которых копируются расчитанные значения", "Выбор строки", DefaultValue1)

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
        Range(Cells(MyRange, SelectionStartCol), Cells(MyRange + MySdvig - 1, SelectionEndCol)).Select 'Диапазон значений
        
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

' Перебераем список комбинаций конжоинта и сохраняем результат для множества симулируемых продуктов.
'
Dim MyStartCombination As Long 'Начальная позиция для сохранения результатов. Первая строчка комбинаций
Dim MyStopCombination As Long 'Конечная позиция для сохранения результатов. Последняя строчка комбинаций

For Each iName In ActiveWorkbook.Names
iName.Delete 'удаление имени
Next

S = Selection
Selection.Name = "Market"
MyMarketSize = UBound(S, 2)

DefaultValue = "3"
DefaultValue1 = "3"
MyStartCombination = InputBox("Введите номер  на листе comb, начиная с которого вставляются строки с комбинациями", "Выбор строки", DefaultValue)
MyStopCombination = InputBox("Введите номер  финальной комбинации на листе comb", "Выбор строки", DefaultValue1)



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


' Определаяем первую комбинацию для расчётов
For i = 1 To 100000
    MyValue = Range("A" & i).Value
    If MyValue = MyStartCombination Then
        MyStartCombinationRow = i ' Номер строки, в которой находится первая комбинация
        Exit For
    End If
Next

MyEndCombinationRow = MyStartCombinationRow + (MyStopCombination - MyStartCombination) ' Номер строки, в которой находится последняя комбинация


Sheets("interface").Activate

With ActiveSheet
    ActiveWindow.View = ViewMode
End With
With Application
    .ScreenUpdating = True
End With



Set S = Application.InputBox("Выделите диапазон с результатами расчёта доли рынка", "Выбор столбца", _
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

SimulationFullSize = SimulationEndRow * SimulationEndCol 'Финальный размер вставки результатов


On Error GoTo ExitTheSub



For MyCombCounter = MyStartCombinationRow To MyEndCombinationRow


    
    
        Sheets("comb").Activate
    
        Range(Cells(MyCombCounter, 2), Cells(MyCombCounter, MyMarketSize + 1)).Select 'Диапазон значений
        
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
        
            For j = MyCombCounter + 1 To (MyCombCounter + SimulationEndRow - 1) ' Переносим результаты из нескольких строк в одну если строк несколько.
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

