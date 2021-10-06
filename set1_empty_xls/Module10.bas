Attribute VB_Name = "Module10"
Private MyOptions() As Variant

Sub export_generator()

w = WorksheetIsExist("template")
    If w = True Then

        Dim MyRangeSize As Long

        Sheets("template").Select
        MyRange = Range("E1: E1000").SpecialCells(xlCellTypeConstants).Select

        MyRangeSize = Selection.Cells.Count

'Заполняем массив с опциями экспорта.

        If MyRangeSize > 1 Then
            MyGetOptionsCorrect = My_Get_options(MyRangeSize, MyReturnedOptions:=MyOptions)
        Else
            MsgBox ("Пустые параметры выгрузки")
        End If

' Можно потом тут проверить через запрос SQL валидность названий баз данных.
'DatabaseName = "Northwind"
'QueryString = _
'    "SELECT * FROM product.dbf"
'Chan = SQLOpen("DSN=" & DatabaseName)
'SQLExecQuery Chan, QueryString
'Set Output = Worksheets("Sheet1").Range("A1")
'SQLRetrieve Chan, Output, , , True
'SQLClose Chan



        If MyGetOptionsCorrect = True Then
            MyResult = GenerateExtractTemplate(MyOptions, MyRangeSize)
        End If

        MyStop = Save_Data_File()

    End If


End Sub

Function My_Get_options(Size As Long, MyReturnedOptions As Variant) As Boolean
My_Get_options = True

ReDim Preserve MyOptions(1 To 5, 1 To Size)

Range("I1").Select
'Первым элементом массива добавляем путь в папку, куда сохранять шаблон.

For i = 1 To 5
    MyOptions(i, 1) = Selection.Value
Next i

Range("A2:E" & Size).Select
'Заносим в массив опции экспорта

For Each MyCol In Selection.Columns
    MyCol.Columns.Select
    For Each MyCel In Selection.Cells
         
         MyOptions(MyCel.Column, MyCel.Row) = MyCel.Value
         
    Next MyCel
    
Next MyCol

'Проверяем, что  нет пустых элементов.
For i = 1 To Size
    If IsEmpty(MyOptions(2, i)) Or IsEmpty(MyOptions(3, i)) Or IsEmpty(MyOptions(4, i)) Or IsEmpty(MyOptions(5, i)) Then
        My_Get_options = False
        MsgBox ("Ошибка в параметрах выгрузки")
    End If
Next i

End Function

Function GenerateExtractTemplate(ExtractOptions() As Variant, Size As Long)
Application.ScreenUpdating = False
Application.DisplayAlerts = False
w = WorksheetIsExist("data")
    If w = True Then Worksheets("data").Delete
MySheet = CreateSheet("data", True)
Sheets("data").Select

MyRow = 1
MyCol = 1

' На листе data генерится шаблон выгрузки для всех файлов если hold не равно 1

For i = 2 To Size
    If IsEmpty(ExtractOptions(1, i)) Then
        ActiveSheet.Cells(MyRow, MyCol).Select
        Selection.Value = MyCol
        MyRow = MyRow + 1
        ActiveSheet.Cells(MyRow, MyCol).Select
        Selection.Value = "*PN " & MyOptions(2, i) & "\" & MyOptions(3, i)
        MyRow = MyRow + 1
        ActiveSheet.Cells(MyRow, MyCol).Select
        Selection.Value = "*TY SAV"
        MyRow = MyRow + 1
        ActiveSheet.Cells(MyRow, MyCol).Select
        If MyOptions(4, i) = 1 Then
            Selection.Value = "*TX " & MyOptions(3, i) & "_used"
        ElseIf MyOptions(4, i) = 2 Then
            Selection.Value = "*TX " & MyOptions(3, i) & "_completes"
        End If
        MyRow = MyRow + 1
        ActiveSheet.Cells(MyRow, MyCol).Select
        Selection.Value = "*FI " & MyOptions(4, i)
        MyRow = MyRow + 1
        ActiveSheet.Cells(MyRow, MyCol).Select
        Selection.Value = "*DI " & MyOptions(1, 1)
        MyRow = MyRow + 1
        ActiveSheet.Cells(MyRow, MyCol).Select
        Selection.Value = "*OU"
        
        Sheets("template").Select
        MyRange = Range("J1: Z60000").SpecialCells(xlCellTypeConstants).Select
        For Each MySelColumn In Selection.Columns
            MySelColumn.Rows.Select
            MyValue = Selection.Cells(1, 1).Value
            If MyValue = MyOptions(5, i) Then
                Selection.Copy
                Sheets("data").Select
                MyRow = MyRow + 1
                ActiveSheet.Cells(MyRow, MyCol).Select
                ActiveSheet.Paste
                ActiveSheet.Cells(MyRow, MyCol).Select
                Selection.Delete
                Exit For
                
            End If
           
        Next MySelColumn
        
        MyCol = MyCol + 1
        MyRow = 1
    End If
    
Next i

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Function

Function Save_Data_File()
Dim BatArray() As Variant
Sheets("data").Select
Range(Range("A2").End(xlDown), Range("A2").End(xlToRight)).Select
Application.ScreenUpdating = False
Application.DisplayAlerts = False

' Каждый шаблон помещается в отдельный фаил *.lot и, при помощи BatArray(), создается исполняемый фаил *.bat
' Все файлы сохраняются в папку, указанную в настройках.

For Each MySelColumn In Selection.Columns
    ReDim Preserve BatArray(1 To MySelColumn.Column)
    BatArray(MySelColumn.Column) = MySelColumn.Column & ".lot"
    MySelColumn.Columns.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs Filename:=MyOptions(1, 1) & BatArray(MySelColumn.Column), FileFormat:= _
    xlText, CreateBackup:=False
    ActiveWorkbook.Close
Next MySelColumn

Workbooks.Add
j = 1

For i = 1 To UBound(BatArray())
    ActiveSheet.Cells(j, 1).Select
    Selection.Value = "extract.exe " & BatArray(i)
    j = j + 1
    ActiveSheet.Cells(j, 1).Select
    Selection.Value = "ping 127.0.0.1 -n 5"
    j = j + 1
Next i

ActiveWorkbook.SaveAs Filename:=MyOptions(1, 1) & "launch_extract.bat", FileFormat:= _
xlText, CreateBackup:=False
ActiveWorkbook.Close
Worksheets("data").Delete
MySheet = CreateSheet("data", True)
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Function

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
s = Selection

SelectionCols = 0
For Each MyRow In Selection.Rows
    SelectionCols = SelectionCols + 1
Next MyRow
SelectionEndRow = UBound(s) + MyStep - 1
SelectionEndCol = UBound(s, 2)
SelectionStartCol = SelectionEndCol - SelectionCols

If UBound(s) Mod MySdvig = 0 Then
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

Sub Sig_test_rows()
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
s = Selection

SelectionCols = 0
For Each MyRow In Selection.Columns
    SelectionCols = SelectionCols + 1
Next MyRow
SelectionEndRow = UBound(s) + MyStep - 1
SelectionEndCol = UBound(s, 2)
SelectionStartCol = SelectionEndCol - SelectionCols + 1

If UBound(s) Mod MySdvig = 0 Then
    For MyRange = MyStep To SelectionEndRow Step MySdvig
        Range(Cells(MyRange, SelectionStartCol), Cells(MyRange + MySdvig - 1, SelectionEndCol)).Select 'Диапазон значений
        
        Selection.Copy
        Sheets("Interface").Select
        Range(Cells(12, 4), Cells(19, 6)).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range(Cells(52, 3), Cells(59, 10)).Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("comb").Select
        ActiveSheet.Cells(MyStep, 4).Select
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


Sub Redo_Mult_Tables()
Dim MyStartRow As Long

MyTables = WorksheetIsExist("Tables")
MySigTables = WorksheetIsExist("SigTables")
    If MyTables = True Then
        Sheets("Tables").Select
        MyStartRow = InputBox("Введите номер первой строки", "Выбор строки", 1)
        TableUpdated = UpdateLabels(MyStartRow)
    End If
    If MySigTables = True And TableUpdated = True Then
        Sheets("SigTables").Select
        SigTableUpdated = UpdateLabels(MyStartRow)
    End If
    If TableUpdated = False Or SigTableUpdated = False Then
        MsgBox ("Упс! Что-то пошло не так.")
    End If
    
    

End Sub

Private Function UpdateLabels(StratFrom As Long) As Boolean
UpdateLabels = False
Columns(2).Select
s = Selection
For i = StratFrom To UBound(s, 1)
    MyLabel = s(i, 2)
    If MyLabel = "-" Then
        Selection.Cells(i, 2) = Selection.Cells(i, 1)
        Selection.Cells(i, 1) = Empty
    End If
    If MyLabel = Empty Then
        StopCounter = StopCounter + 1
        If StopCounter > 100 Then
            UpdateLabels = True
            Exit Function
        End If
    Else
        StopCounter = 0
    End If
Next i


End Function

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

s = Selection
Selection.Name = "Market"
MyMarketSize = UBound(s, 2)

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



Set s = Application.InputBox("Выделите диапазон с результатами расчёта доли рынка", "Выбор столбца", _
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
 
 


s.Select
Selection.Name = "Simulation"
s = Selection

SimulationEndRow = UBound(s)
SimulationEndCol = UBound(s, 2)

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

