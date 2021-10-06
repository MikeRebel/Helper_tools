Attribute VB_Name = "OldSimpleSimulator"
Sub perepbor_simulatora_column()
'

' Yota Макрос
'
Dim step As Integer
Dim sdvig As Integer

MsgBox "!!!ВНИМАНИЕ!!! чтобы сработал макрос необходимо чтобы: Диапазон значений, которые надо просимулировать, должен быть выделен перед запуском макроса. Списки значений и результирующие данные должны находиться на листе comb. Лист, с выводом результата симуляции, должен называться interface. Диапазон ячеек, куда вставляются значения атрибутов должен называться Market. Диапазон, откуда копируются расчитанные доли должен называться Simulation. Иначе макрос упадет с ошибкой"
DefaultValue = "1"
DefaultValue1 = "1"
step = InputBox("Введите номер колонки на листе comb, начиная с которой вставляются столбцы с расчитанными значениями", "Выбор колонки", DefaultValue)
sdvig = InputBox("Введите количество колонок на листе interface, из которых копируются расчитанные значения", "Выбор колонки", DefaultValue1)



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

' Перебераем список комбинаций конжоинта и сохраняем результат
'
Dim MyStep As Long 'Начальная позиция для сохранения результатов. Первая строчка результатов
Dim MySdvig As Long 'Конечная позиция для сохранения результатов. Последняя строчка результатов
Dim shag As Long 'Количество строк в результате (для вставки заголовка с комбинацией)
MyCounter = 0 'Счетчик для сохранения файла

DefaultValue = "2"
DefaultValue1 = "5"
MyStep = InputBox("Введите номер строки на листе comb, начиная с которой вставляются строки с расчитанными значениями", "Выбор строки", DefaultValue)
MySdvig = InputBox("Введите количество строк на листе interface, из которых копируются расчитанные значения", "Выбор строки", DefaultValue1)

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
        Range("B" & MyRange, "N" & MyRange + 4).Select 'Диапазон значений
        
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
    MySaveCounter = MyCounter Mod 10000 ' сохраняем каждые Х комбинаций
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

