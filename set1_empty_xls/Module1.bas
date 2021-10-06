Attribute VB_Name = "Module1"
Sub Maxdiff_solution()




For i = 0 To 1000000

    Worksheets("Лист1").Range("D73").Select
    Selection.Value = i
    
    Columns("A:B").Sort key1:=Range("A1")
    
   s = Worksheets("Лист1").Range("AE61").Value
   If s < 3 Then Exit Sub
 
   
   
   Worksheets("Лист1").Range("D73").Value = i
   
    
    
    
Next i


End Sub

Sub PhoneBase_clearing()
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
            MsgBox ("Ошибка")
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
MsgBox ("Выберите диапазон")
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

Sub update_tables()
'With Application
'        .ScreenUpdating = False
'        .EnableEvents = False
'        .Calculation = xlCalculationManual
'End With


Dim ListNamesArray As Variant
ListNamesArray = Array("Tables", "SigTables")

For Each c In ListNamesArray
    transformTable (c)
    
Next c



'With Application
'        .ScreenUpdating = True
'        .EnableEvents = True
'        .Calculation = xlCalculationAutomatic
'End With

End Sub

Sub update_table_links()
Dim NewLink As String
Dim MyAdress As Variant
Dim BlockNamesArray As Variant
Dim d As Variant

BlockNamesArray = Array("Соцдем", "Связь", "оценка", "Планшеты", "USB-модем", "Операторы", "Соц-дем")

BlockNamesCount = 0
Sheets("Contents").Activate
 

 
For Each c In ActiveSheet.Hyperlinks
    MyLink = c.SubAddress
    
 
        
    s1 = Split(MyLink, "!")
    s2 = Split(s1(1), ":")
    Dim StartAdressArray() As Variant
    Dim StopAdressArray() As Variant
    MyCol = c.Range.Column
    MyRow = c.Range.Row
      
    For i = 1 To Len(s2(0))
        ReDim Preserve StartAdressArray(i)
        StartAdressArray(i) = Mid(s2(0), i, 1)
    Next i
    For i = 1 To Len(s2(1))
        ReDim Preserve StopAdressArray(i)
        StopAdressArray(i) = Mid(s2(1), i, 1)
    Next i
    For i = 1 To UBound(StartAdressArray)
        If IsNumeric(StartAdressArray(i)) Then
            StartRow = StartRow & StartAdressArray(i)
        Else
            StartCol = StartCol & StartAdressArray(i)
        End If
    Next i
    For i = 1 To UBound(StopAdressArray)
        If IsNumeric(StopAdressArray(i)) Then
           StopRow = StopRow & StopAdressArray(i)
        Else
            StopCol = StopCol & StopAdressArray(i)
        End If
    Next i
    
    
 
    
    
    Range("C" & MyRow).Select
    MyValue = Selection.Value
    
    Set MyFind = Selection.Find("^", , xlValues, xlPart)
        
        If Not MyFind Is Nothing Then
            CutValue = Split(MyValue, "^")
            Selection.Value = CutValue(0) & CutValue(2)
        End If
    
    Set MyFind = Selection.Find("  ))", , xlValues, xlPart)
    If Not MyFind Is Nothing Then
        MyFindCount = MyFindCount + 1
        If MyFindCount < 3 Then
            MyError = 0
            StartRow = StartRow - MyError
            StopRow = StopRow - MyError
            
        Else
            If MyFindCount Mod 2 <> 0 Then
                MyError1 = StartRow
                StartRow = StartRow - MyError1 + 1
                StopRow = StopRow - MyError1 + 1

                BlockNamesCount = BlockNamesCount + 1
            Else
                MyError2 = StartRow
                StartRow = StartRow - MyError2 + 1
                StopRow = StopRow - MyError2 + 1
            End If
        End If
    Else
        If s1(0) = "SigTables" Then
            StartRow = StartRow - MyError1 + 1
            StopRow = StopRow - MyError1 + 1

                
        Else
            StartRow = StartRow - MyError2 + 1
            StopRow = StopRow - MyError2 + 1
            

        End If
    End If
      
    s1(0) = s1(0) & " " & BlockNamesArray(BlockNamesCount)
 
    
    NewLink = "!" & StartCol & Str(StartRow) & ":" & StopCol & Str(StopRow)
    NewLink = Replace(NewLink, " ", "", , 2)
    NewLink = "'" & s1(0) & "'" & NewLink
    c.SubAddress = NewLink
    
    StartRow = Null
    StopRow = Null
    StartCol = Null
    StopCol = Null
    
    For i = 2 To Worksheets.Count
        
        For Each d In Worksheets(i).Hyperlinks
        
            ReplacedLink = d.SubAddress
            If ReplacedLink = MyLink Then
                d.SubAddress = NewLink
            End If
            
        Next d
        
    Next i
    
Next c




For Each MyCell In Selection.Cells
    MyCell.Select
    Selection.Activate
    MyValue = Selection.Value
    Set MyFind = Selection.Find("  ))", , xlValues, xlPart)
    If Not MyFind Is Nothing Then
                MyReplace = Replace(MyValue, "  ))", "))")
                Selection.Value = MyReplace
    End If
Next MyCell


End Sub

Function transformTable(table As String)

Dim MyFind
Dim MyFilterFind

NameBlockAddres = 0
StartRow = 1
SortCounter = 0

Dim MyNewSheetName As String
Dim BlockNamesArray As Variant

BlockNamesArray = Array("Соцдем", "Связь", "оценка", "Планшеты", "USB-модем", "Операторы", "Соц-дем")



Sheets(table).Activate
MySourceSheetName = ActiveSheet.Name
Columns("A").Select

For Each MyCol In Selection.Cells
    MyCol.Select
    Selection.Activate
    MyValue = Selection.Value
    If Not IsEmpty(MyValue) And Not IsArray(MyValue) Then
        
        'определяем и убиваем фильтр, который не нужен клиенту(Перенесено в сортировку)
        
        Set MyFilterFind = Selection.Find("^", , xlValues, xlPart)

        If Not MyFilterFind Is Nothing Then
            CutValue = Split(MyValue, "^")
            MyValue = CutValue(0) & CutValue(2)
            Selection.Value = MyValue
        End If
        
        
        'Разбиваем на блоки
        
        Set MyFind = Selection.Find("  ))", , xlValues, xlPart)
        
            If Not MyFind Is Nothing Then
                MyReplace = Replace(MyValue, "  ))", "))")
                Selection.Value = MyReplace
                
                If StopTimer > 0 Then
                    
                    EndRow = MyCol.Row - 1
                    
                   
                    
                    MyNewSheetName = BlockNamesArray(NameBlockAddres)
                    MySheet = CreateSheet(table & " " & MyNewSheetName, True)
                    NameBlockAddres = NameBlockAddres + 1
                    MyDestSheetName = ActiveSheet.Name
                    Sheets(table).Activate
                    Range("A" & StartRow, "IV" & EndRow).Select
                    Selection.Cut
                    Sheets(MyDestSheetName).Activate
                    ActiveSheet.Columns("A:B").ColumnWidth = 30
                    ActiveSheet.Columns("C:C").ColumnWidth = 5
                    Range("A1").Select
                    ActiveSheet.Paste

                    Sheets(MySourceSheetName).Select
                    
                    StartRow = EndRow + 1
                    
                    StopTimer = 0
                End If
        
                
        
            Else
                
            End If
    End If

    If StopTimer > 10000 Then 'Сколько пустых строк просмотреть перед выходом (Максимальный возможный размер блока в строках)
        EndRow = MyCol.Row - 1
        MyReplace = Replace(MyValue, "  ))", "))")
        Selection.Value = MyReplace
                   
                
        MyNewSheetName = BlockNamesArray(NameBlockAddres)
        MySheet = CreateSheet(table & " " & MyNewSheetName, True)
                    
        MyDestSheetName = ActiveSheet.Name
        Sheets(table).Activate
        Range("A" & StartRow, "IV" & EndRow).Select
        Selection.Cut
        Sheets(MyDestSheetName).Activate
        ActiveSheet.Columns("A:B").ColumnWidth = 30
        ActiveSheet.Columns("C:C").ColumnWidth = 5
        Range("A1").Select
        ActiveSheet.Paste
        
        StartRow = EndRow + 1
                    
        StopTimer = 0
        Sheets(MySourceSheetName).Delete
    Exit For
    
    End If
    StopTimer = StopTimer + 1
Next MyCol


End Function

Sub Sort_by_Total() ' сортируем таблицу по значению столбца Total

Dim MyStartRow As Long
Dim ListNamesArray() As Variant
Dim ListNamesHeaderBlockRowsArray() As Variant


ListNamesArray = Array("Tables", "SigTables")
ListNamesHeaderBlockRowsArray = Array(1, 2)


For c = 0 To UBound(ListNamesArray)
 
        w = MySortTables(ListNamesArray(c), ListNamesHeaderBlockRowsArray(c))
        
        
Next c


    
End Sub


Function MySortTables(w, l)

Dim SortArray() As Variant
Dim ExcludeArray() As Variant
SortCount = 0
MyStartRow = 1
MyStopCounter = 0
SortArray = Array("8.  s7: Оператор сим-карты, на которую звоним^ (n=((Mytotal=1) & ((sample=1|sample=2)) ^ ((Все))", "10.  q0: Кто оплачивает расходы по номеру^ (n=((Mytotal=1) & ((sample=1|sample=2)) ^ ((Все))", "11.  q42: Кто оплачивает расходы по телефону 2^ (n=((Mytotal=1) & ((sample=1|sample=2)&q4>1 ) ^ ((Есть несколько телефонов))", _
"12.  q0q42_1: Кто оплачивает связь (unduplicated)^ (n=((Mytotal=1) & ((sample=1|sample=2)) ^ ((Все))", "17.  Oper_Sim2_1: Операторы других сим-карт телефона^ (n=((Mytotal=1) & ((sample=1|sample=2)&q1=1) ^ ((Использует более одной сим-карты))", "18.  q50: Другие устройства, которыми пользуется респондент^ (n=((Mytotal=1) & ((sample=1|sample=2)) ^ ((Все))", _
"25.  Oper_MI_phones1: Выход в МИ по операторам с любого телефона^ (n=((Mytotal=1) & ((sample=1|sample=2)) ^ ((Все))", "26.  Oper_MI_phones1: Выход в МИ по операторам с любого телефона^ (n=((Mytotal=1) & ((sample=1|sample=2)&Oper_MI_phones1_6>0) ^ ((Выходят в МИ хотя бы с одного телефона))", "29.  q6: Марка телефона 1^ (n=((Mytotal=1) & ((sample=1|sample=2)) ^ ((Все))", _
"30.  q40: Марка телефона 2 ^ (n=((Mytotal=1) & ((sample=1|sample=2)&q4>1 ) ^ ((Есть несколько телефонов))", "31.  Brand_ph: Марки телефонов^ (n=((Mytotal=1) & ((sample=1|sample=2)&not sysmis(Brand_ph)) ^ ((От базы ответивших по каждому телефону ))", "32.  q7: Тип мобильного телефона 1^ (n=((Mytotal=1) & ((sample=1|sample=2)) ^ ((Все))", _
"33.  q41: Тип мобильного телефона 2^ (n=((Mytotal=1) & ((sample=1|sample=2)&q4>1 ) ^ ((Есть несколько телефонов))", "34.  Tel_type: Тип мобильных телефонов^ (n=((Mytotal=1) & ((sample=1|sample=2)&not sysmis(Tel_type)) ^ ((От базы ответивших по каждому телефону ))", "36.  q11: Место покупки смартфона^ (n=((Mytotal=1) & ((sample=1|sample=2)&q7=3&(q10=1|q10=2|q10=3)) ^ ((Используют смартфон и купили его не более года назад))", _
"47.  q14: Тип доступа в Интернет с телефона 1^ (n=((Mytotal=1) & ((sample=1|sample=2)&q13=1) ^ ((Выходят в Интернет с телефона 1))", "48.  q45: Тип доступа в Интернет с телефона 2^ (n=((Mytotal=1) & ((sample=1|sample=2)&q4>1&q44=1) ^ ((Выходят в Интернет с телефона 2))", "50.  q15: Используют WiFi на телефоне 1^ (n=((Mytotal=1) & ((sample=1|sample=2)&DTq14m1_1=1) ^ ((Используют WiFi на данном телефоне))", _
"53.  Oper_MI_ustr1: МИ по операторам хотя бы на одном устройстве ^ (n=((Mytotal=1) & ((sample=1|sample=2)) ^ ((Все))", "54.  Oper_MI_ustr1: МИ по операторам хотя бы на одном устройстве ^ (n=((Mytotal=1) & ((sample=1|sample=2)&Oper_MI_ustr1_6>0) ^ ((Пользуются МИ хотя бы на одном устройстве))", "55.  Oper_MI_ph1_1: Выход в МИ по операторам с телефона 1^ (n=((Mytotal=1) & ((sample=1|sample=2)) ^ ((Все))", _
"56.  q46: Оператор выхода в МИ с телефона 2^ (n=((Mytotal=1) & ((sample=1|sample=2)&q4>1&DTq45m1_2=1) ^ ((Выходят в МИ с данного телефона))", "57.  Oper_MI_phones1: Выход в МИ по операторам с любого телефона^ (n=((Mytotal=1) & ((sample=1|sample=2)) ^ ((Все))", "58.  Oper_MI_phones1: Выход в МИ по операторам с любого телефона^ (n=((Mytotal=1) & ((sample=1|sample=2)&Oper_MI_phones1_6>0) ^ ((Выходят в МИ хотя бы с одного телефона))", _
"59.  q20: Места использования МИ на телефоне 1^ (n=((Mytotal=1) & ((sample=1|sample=2)&(DTq14m1_2=1&q1=2)|(q17=1) ) ^ ((Используют МИ на сим-карте, на которую звоним))", "66.  q26: Оператор, на которого планируют перейти ^ (n=((Mytotal=1) & ((sample=1|sample=2)&q24=1|q24=2) ^ ((Намерены сменить текущего оператора))", "72.  q30: Тип доступа в Интернет с телефона 1 раньше^ (n=((Mytotal=1) & ((sample=1|sample=2)&q29=1) ^ ((Пользовались раньше Интернетом на телефоне 1))", _
"78.  q33: Тип телефона, на который планируют сменить ^ (n=((Mytotal=1) & ((sample=1|sample=2)&q32=1|q32=2) ^ ((Планируют сменить телефон в ближайшие полгода ))", "80.  q47: Триггеры для подключения к МИ^ (n=((Mytotal=1) & ((sample=1|sample=2)&((q4=1)&(q8<>1|q13<>1))|(((q4>1)&(q8<>01|q13<>1))&(q43<>1|q44<>01))) ^ ((Не выходят в Интернет ни с одного телефона))", "83.  OPER_PL: Оператор МИ на планшете^ (n=((Mytotal=1) & ((sample=1)&DTq50m1_1=1) ^ ((Пользуются планшетом))", _
"84.  OPER_PL: Оператор МИ на планшете^ (n=((Mytotal=1) & ((sample=1)&OPER_PL<>999) ^ ((Пользуются МИ хотя бы на одном планшете))", "86.  p2: Марка основного планшета ^ (n=((Mytotal=1) & ((sample=1)&DTq50m1_1=1) ^ ((Пользуются планшетом))", "93.  p6: Тип доступа в Интернет с основного планшета^ (n=((Mytotal=1) & ((sample=1)&p5=1) ^ ((Выходят в Интернет с основного планшета))", _
"94.  p32: Тип доступа в Интернет с дополнительного планшета^ (n=((Mytotal=1) & ((sample=1)&(P1>1&p31=1)) ^ ((Выходят в Интернет с дополнительного планшета))", "95.  p8: Места использования WiFi на основном планшете^ (n=((Mytotal=1) & ((sample=1)&DTp6m1_1=1) ^ ((Используют WiFi на основном планшете))", "97.  p10: Оператор МИ на основном планшете^ (n=((Mytotal=1) & ((sample=1)&DTp6m1_2=1) ^ ((Используют МИ на основном планшете))", _
"98.  p34: Оператор МИ на дополнительном планшете^ (n=((Mytotal=1) & ((sample=1)&(P1>1&DTp32m1_2=1)) ^ ((Используют МИ на дополнительном планшете))", "99.  p11: Места использования МИ на основном планшете^ (n=((Mytotal=1) & ((sample=1)&DTp6m1_2=1) ^ ((Используют МИ на основном планшете))", "108.  p21: Тип доступа в Интернет с планшета раньше^ (n=((Mytotal=1) & ((sample=1)&p20=1) ^ ((Пользовались раньше Интернетом на планшете))", _
"117.  p42: Тип доступа в Интернет с планшета раньше^ (n=((Mytotal=1) & ((sample=1)&p41=1) ^ ((Пользовались раньше Интернетом на планшете))", "118.  p43: Оператор на планшете раньше^ (n=((Mytotal=1) & ((sample=1)&DTp42m1_2=1) ^ ((Пользовались раньше МИ оператора на планшете))", "122.  OPER_USB: Оператор USB-модема^ (n=((Mytotal=1) & ((sample=1)&DTq50m1_2=1) ^ ((Пользуются USB))", _
"124.  u3: Оператор основного USB-модема^ (n=((Mytotal=1) & ((sample=1)&DTq50m1_2=1) ^ ((Пользуются основным USB ))", "125.  u8: Оператор дополнительного USB-модема^ (n=((Mytotal=1) & ((sample=1)&u1>1) ^ ((Пользуются дополнительным USB))", "126.  u4: Места использования основного USB-модема^ (n=((Mytotal=1) & ((sample=1)&DTq50m1_2=1) ^ ((Пользуются основным USB ))", _
"127.  u9: Места использования дополнительного USB-модема^ (n=((Mytotal=1) & ((sample=1)&u1>1) ^ ((Пользуются дополнительным USB))", "134.  u21: Оператор предыдущегоUSB-модема^ (n=((Mytotal=1) & ((sample=1)&u20=1) ^ ((Пользовались раньше USB-модемом ))", "137.  w1: Оператор WiFi роутера^ (n=((Mytotal=1) & ((sample=1)&DTq50m1_4=1) ^ ((Пользуются WiFi-роутером))", _
"138.  o0: Знание операторов^ (n=((Mytotal=1) & ((sample=1|sample=2)) ^ ((Все))", "159.  o3: Лучший оператор для МИ^ (n=((Mytotal=1) & ((sample=1|sample=2)&not sysmis(o3)) ^ ((Знают операторов ))")




    
        Sheets(w).Select
        MySourceSheetName = ActiveSheet.Name
        Columns("A").Select

        For Each MyCol In Selection.Cells
            MyStopCounter = MyStopCounter + 1
            
            If MyStopCounter = 10000 Then 'Увеличить периуд неактивности.
                Exit Function
            End If
            
            
            MyCol.Select
            Selection.Activate
            MyValue = Selection.Value
            
'            Set MyFilterFind = Selection.Find("^", , xlValues, xlPart)
'
'            If Not MyFilterFind Is Nothing Then
'                CutValue = Split(MyValue, "^")
'                MyValue = CutValue(0) & CutValue(2)
'                Selection.Value = MyValue
'            End If

            Set MyBlockEndFind = Selection.Find("  ))", , xlValues, xlPart)
        
            If Not MyBlockEndFind Is Nothing Then
                MyValue = Replace(MyValue, "  ))", "))")
            End If
            
            If Not IsEmpty(MyValue) And Not IsArray(MyValue) Then
                For MySort = SortCount To UBound(SortArray)
                    
        
                    If MyValue = SortArray(MySort) Then
                    MyStopCounter = 0
                        MyTotalRow = MyCol.Row + 1 ' Ряд, в котором находится Total
                        Range("C" & MyTotalRow + 1).Select
                        LabelSelection = Selection ' Пропуск рядов шапки
                        
                        If IsArray(LabelSelection) Then
                            MyLabelRow = UBound(LabelSelection) ' Определение количества рядов в шапке.
                        Else
                            MyLabelRow = 1
                        End If
                        
                        MyAdressRow = MyTotalRow + l + MyLabelRow 'адрес начала меток. l - сдвиг в зависимости от типа таблиц.
                        Range("B" & MyAdressRow).Select
                        If Selection.Value = "-" Then  ' Если тут знак - значит это малтипл. Если нет - сингл. Разные алгоритмы опеределения диапазона сортировки.
                            Range(Range("A" & MyAdressRow), Range("B" & MyAdressRow).End(xlDown)).Select
                            s1 = Selection
                            MyEndRow = 0
                            Dim j As Integer
                            For j = 1 To UBound(s1)
                                If s1(j, 2) = "-" Then
                                    ExcludeThisSortRow = Array_Unsorted_Match(s1(j, 1))
                                    If ExcludeThisSortRow = False Then
                                        MyEndRow = MyEndRow + 1
                                    End If
                                End If
                            Next j
                            MyEndRow = MyEndRow + MyAdressRow - 1
                            Range("A" & MyAdressRow, "ZZ" & MyEndRow).Select 'Менять тут адрес, если изменится длинна таблицы.
                            Selection.Columns("A:ZZ").Sort key1:=Range("C" & MyAdressRow), order1:=xlDescending, Header:=xlNo 'Менять тут адрес, если изменится длинна таблицы.
                        Else
                
                            Range("A" & MyAdressRow).Select
                            s = Selection
                            MyEndRow = MyAdressRow + UBound(s) - 1 'адрес конца меток.
                            For NoOtherSort = MyAdressRow To MyEndRow
                                Range("B" & NoOtherSort).Select
                                ExcludeThisSortRow = Array_Unsorted_Match(Selection.Value)
                                If ExcludeThisSortRow = True Then
                                
                                        MyEndRow = MyEndRow - 1
                                End If
                            Next NoOtherSort
                    
                            Range("B" & MyAdressRow, "ZZ" & MyEndRow).Select 'Менять тут адрес, если изменится длинна таблицы.
                            Selection.Columns("A:ZZ").Sort key1:=Range("C" & MyAdressRow), order1:=xlDescending, Header:=xlNo 'Менять тут адрес, если изменится длинна таблицы.
                        End If
                        
                    End If
                    
                Next MySort
        
            End If
        
        Next MyCol
'
'        MyRange = Range("C" & MyStartRow & ": C100000").Select
'        For Each i In Selection.Cells
'            i.Select
'            CurCellValue = Selection.Value
'            If CurCellValue = "Total" Then
'                MyTotalRow = i.Row
'
'
'            End If
'
'        Next i
        
  

End Function

Function Array_Unsorted_Match(CurValue) As Boolean

Array_Unsorted_Match = False

ExcludeArray = Array("Затрудняюсь ответить", "Другой", "Ничем не пользуется", "Не отвечал на вопрос", "Другое", "Ничего, сам(а) планирую начать", "Ничего, я не планирую начать")

For i = 0 To UBound(ExcludeArray)
    
    If CurValue = ExcludeArray(i) Then
        Array_Unsorted_Match = True
    End If
    
Next i

End Function
Sub stacked()
Attribute stacked.VB_ProcData.VB_Invoke_Func = " \n14"
'
' stacked Макрос
'

'
'    Selection.Copy
'    MySelection = Selection
    Application.DisplayAlerts = False ' убираем предупреждения
    mybook = Application.ActiveWorkbook.Name
    
    MyBookToCopy = ActiveSheet.Cells(1, 3).Value
    If IsEmpty(MyBookToCopy) Then ErrorMessage = MsgBox("В клетке С1 напишите название книги, в которую хотите сохранять результат", vbCritical)
    
    
    
    rRange = ActiveSheet.Cells(1, 4).Value
    If IsEmpty(rRange) Then rRange = 1
    
    w = WorksheetIsExist("Data")
    If w = True Then Worksheets("Data").Delete
    
    MyCol = 1
    MyRow = 1
    PasteCounter = 1
    
    MySelectionRowCount = Selection.Rows.Count
    areaCount = Selection.Areas.Count
    MySelectionColumnCount = 0
    For areaNumber = 1 To areaCount
        MySelectionColumnCount = MySelectionColumnCount + Selection.Areas(areaNumber).Columns.Count
    Next areaNumber
'    Dim MyInput
'    MyInput = InputBox("Отделять каждые N столбиков пустым столбиком. Введите N: ")
    
' Создаем лист с данными
    MySheet = CreateSheet("Data", True)
    
    
'копируем каждый ряд выделенной области в новый лист, разъединяем слитые ячейки
        Sheets("Tables").Activate
    For Each c In Selection.Columns
'       MsgBox (c.Cells(1, 1))
'       Workbooks("set1.xls").Activate
        Sheets("Tables").Activate
        c.Columns.Select
        Selection.Copy
'        Windows("Книга1").Activate
        Sheets("Data").Select
        ActiveSheet.Cells(MyRow, MyCol).Select
        ActiveSheet.Paste
        With Selection
            .MergeCells = False
        End With
        
            If PasteCounter = 1 Then ' для первого столбика удаляем лишние пустые ячейки
                For j = MySelectionRowCount To 1 Step -1
                    MyVal = ActiveSheet.Cells(j, MyCol).Value
                    If j > 1 Then k = j - 1 Else k = 2 ' чтобы когда J=1 не вылетала изза невозможности  выбрать значение клетки №0
                    MyVal1 = ActiveSheet.Cells(k, MyCol).Value ' Значения ячейки на один ряд выше
                    If MyVal1 = MyVal Then ActiveSheet.Cells(k, MyCol).Delete Shift:=xlUp
                Next j
            End If
            
            If PasteCounter > 1 Then ' для следующих столбиков удаляем лишние пустые ячейки и ячеки с текстом кроме первой над процентами
                For j = MySelectionRowCount To 1 Step -1
                    MyVal = ActiveSheet.Cells(j, MyCol).Value
                    If j > 1 Then k = j - 1 Else k = 2 ' чтобы когда J=1 не вылетала изза невозможности  выбрать значение клетки №0
                    MyVal1 = ActiveSheet.Cells(k, MyCol).Value ' Значения ячейки на один ряд выше
                    If IsEmpty(MyVal) Then ActiveSheet.Cells(j, MyCol).Delete Shift:=xlUp
                    If Application.WorksheetFunction.IsText(MyVal1) = True And Application.WorksheetFunction.IsText(MyVal) = True Then ActiveSheet.Cells(k, MyCol).Delete Shift:=xlUp ' лишняя проверка, но пускай остается
                                        
                Next j
            End If
            
            MyCol = MyCol + 1
            PasteCounter = PasteCounter + 1
    Next c

' Надо бы поизящнее

'  Windows("Книга2").Activate
'    Range("A1:IV100").Select
'    With Selection
'         .MergeCells = False
'    End With

' Удаляем первый пустой ряд

'    Range("A1").Select
'    Selection.EntireRow.Delete

' удаляем остальные пустые клетки
    'For i = MySelectionRowCount To 1 Step -1
     '   For j = MySelectionColumnCount To 1 Step -1
      '      MyVal = ActiveSheet.Cells(i, j).Value
       '     If MyVal = Empty Then
        '        ActiveSheet.Cells(i, j).Delete Shift:=xlUp
         '   End If
      '  Next j
'    Next i

'вставляем ряд между подписями и данными если ни разу ряд не удалялся.
'    If MyDeleteCount = 0 Then
'            Range(Range("A1").End(xlDown), Range("A1").End(xlToRight)).Select
            
'            For Each MyCell In Selection.Cells
'            If MyCell.Value = "ToDelete" Then MyCell.Cells.Delete Shift:=xlUp
'            Next
            
            
            
            Rows(2).Insert
 '   End If
    
    Range("A1:A2").Select
    Selection.Value = " "
'    Range(Range("A1").End(xlDown), Range("A1").End(xlToRight)).Select
'    Selection.Copy
'    Sheets("Tables").Activate

     Application.Workbooks(MyBookToCopy).Activate
'    Range("A1:A2").Select
    
'    MyRange = Application.InputBox(Prompt:= _
                "Кликните по верхней правой ячейке для вставки", _
                    Title:="КУДА ВСТАВЛЯТЬ?", Type:=2)
'    MsgBox (MyRange)
'    rRange = Selection.Address
'     Dim rRange As Range
'     rRange = InputBox("Введите номер верхней строки для вставки")
     MyBookToCopy = Application.ActiveWorkbook.Name
     
     Application.Workbooks(mybook).Activate
     Range(Range("A1").End(xlDown), Range("A1").End(xlToRight)).Select
     Selection.Copy
     
     Application.Workbooks(MyBookToCopy).Activate
     ActiveSheet.Cells(rRange, 1).Select
     ActiveSheet.Paste
     
     Application.Workbooks(mybook).Activate
     Worksheets("Data").Delete
     Sheets("Tables").Activate
     rRange = rRange + MySelectionRowCount + 2
     ActiveSheet.Cells(1, 4).Value = rRange
    
     Application.Workbooks(MyBookToCopy).Activate
    
    Application.DisplayAlerts = True


End Sub

Sub pie()
'
' pie chart Макрос
'

'

    Application.DisplayAlerts = False ' убираем предупреждения
    mybook = Application.ActiveWorkbook.Name
    
    MyBookToCopy = ActiveSheet.Cells(1, 3).Value
    If IsEmpty(MyBookToCopy) Then ErrorMessage = MsgBox("В клетке С1 напишите название книги, в которую хотите сохранять результат", vbCritical)
    
    
    
    rRange = ActiveSheet.Cells(1, 4).Value
    If IsEmpty(rRange) Then rRange = 1
    
    w = WorksheetIsExist("Data")
    If w = True Then Worksheets("Data").Delete
    
    MyCol = 1
    MyRow = 1
    PasteCounter = 1
    
    MySelectionRowCount = Selection.Rows.Count
    areaCount = Selection.Areas.Count
    MySelectionColumnCount = 0
    For areaNumber = 1 To areaCount
        MySelectionColumnCount = MySelectionColumnCount + Selection.Areas(areaNumber).Columns.Count
    Next areaNumber
    If MySelectionColumnCount <> 2 Then ErrorMessage = MsgBox("Выделите сначала метки а потом один столбец с данными. Несколько столбцов не помещаются на графике типа Pie.", vbCritical)

' Создаем лист с данными
    MySheet = CreateSheet("Data", True)
    
    
'копируем каждый ряд выделенной области в новый лист, разъединяем слитые ячейки
        Sheets("Tables").Activate
    For Each c In Selection.Columns
        Sheets("Tables").Activate
        c.Columns.Select
        Selection.Copy
        Sheets("Data").Select
        ActiveSheet.Cells(MyRow, MyCol).Select
        ActiveSheet.Paste
        With Selection
            .MergeCells = False
        End With
        
            If PasteCounter = 1 Then ' для первого столбика удаляем лишние пустые ячейки
                For j = MySelectionRowCount To 1 Step -1
                    MyVal = ActiveSheet.Cells(j, MyCol).Value
                    If j > 1 Then k = j - 1 Else k = 2 ' чтобы когда J=1 не вылетала изза невозможности  выбрать значение клетки №0
                    MyVal1 = ActiveSheet.Cells(k, MyCol).Value ' Значения ячейки на один ряд выше
                    If MyVal1 = MyVal Then ActiveSheet.Cells(k, MyCol).Delete Shift:=xlUp
                Next j
            End If
            
            If PasteCounter > 1 Then ' для следующих столбиков удаляем лишние пустые ячейки и ячеки с текстом кроме первой над процентами
                For j = MySelectionRowCount To 1 Step -1
                    MyVal = ActiveSheet.Cells(j, MyCol).Value
                    If j > 1 Then k = j - 1 Else k = 2 ' чтобы когда J=1 не вылетала изза невозможности  выбрать значение клетки №0
                    MyVal1 = ActiveSheet.Cells(k, MyCol).Value ' Значения ячейки на один ряд выше
                    If IsEmpty(MyVal) Then ActiveSheet.Cells(j, MyCol).Delete Shift:=xlUp
                    If Application.WorksheetFunction.IsText(MyVal1) = True And Application.WorksheetFunction.IsText(MyVal) = True Then ActiveSheet.Cells(k, MyCol).Delete Shift:=xlUp ' лишняя проверка, но пускай остается
                                        
                Next j
            End If
            
            MyCol = MyCol + 1
            PasteCounter = PasteCounter + 1
    Next c
    
    Range("A1").Select
    Selection.Value = " "
    
    Application.Workbooks(MyBookToCopy).Activate
    MyBookToCopy = Application.ActiveWorkbook.Name
    
    Application.Workbooks(mybook).Activate
    Range(Range("A1").End(xlDown), Range("A1").End(xlToRight)).Select
    Selection.Copy
    
    Application.Workbooks(MyBookToCopy).Activate
    ActiveSheet.Cells(rRange, 1).Select
    ActiveSheet.Paste
     
    Application.Workbooks(mybook).Activate
    Worksheets("Data").Delete
    Sheets("Tables").Activate
    rRange = rRange + MySelectionRowCount + 2
    ActiveSheet.Cells(1, 4).Value = rRange
    
    Application.Workbooks(MyBookToCopy).Activate

    Application.DisplayAlerts = True


End Sub


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

Sub total_remove()
'
' total_remove Макрос
'

'
    Dim DelArray(16)
    MyRow = 1
    del_counter = 1
    
    w = WorksheetIsExist("Data")
    If w = True Then Worksheets("Data").Delete
    MySheet = CreateSheet("Data", True)
    Sheets("Tables").Activate
    ActiveSheet.Columns(1).Select
    
    Selection.Copy
    Sheets("Data").Select
    ActiveSheet.Cells(1, 1).Select
    ActiveSheet.Paste
    With Selection
        .MergeCells = False
    End With
    
    Sheets("Data").Select
    ActiveSheet.Columns(1).Select
    
    For Each c In Selection.Rows
    c.Rows.Select
    MyValue = Selection.Cells(1, 1).Value
    MyAdress = Selection.Cells(1, 1).Address
    If MyValue = "Всего" Then
        DelArray(del_counter) = MyAdress
'        Range(MyAdress).Select
'        Selection.EntireRow.Delete
        del_counter = del_counter + 1
    End If
    If del_counter = 17 Then
        Exit For
    End If
    
    MyRow = MyRow + 1
    Next c
    
    Sheets("Tables").Activate
    For i = 16 To 1 Step -1
    MyRow = DelArray(i)
    Range(MyRow).Select
    Selection.EntireRow.Delete
    Next i

    Worksheets("Data").Delete
    
End Sub

Sub hyperlinks_correction()
'
' hyperlinks_correction Макрос
'
MsgBox "Макрос заменяет выражение в ячейке B1 на выражение, в ячейке C1 в адресе каждой гиперссылки на  активном листе."
'
    mybook = Application.ActiveWorkbook.Name
    MyWorkSheet = ActiveSheet.Name
    ReplaceString = ActiveSheet.Cells(1, 2).Value
    StringToReplace = ActiveSheet.Cells(1, 3).Value

    
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
            NewLink = Replace(MyLink, ReplaceString, MyWorkSheet)
            c.SubAddress = NewLink
        End If
    
    Next c
    
    
    
End Sub

Sub summary()
    Application.DisplayAlerts = False
    w = WorksheetIsExist("Data")
    If w = True Then Worksheets("Data").Delete
    MySheet = CreateSheet("Data", True)
    Sheets("Summary tables").Select
    SummObjectPlace = ActiveSheet.Cells(1, 2).Value ' Сколько раз встречается выделенная подпись в табличках сверху (номер таблички)
    StartSheet = ActiveSheet.Cells(1, 3).Value ' Номер листа, с которого начинать суммирование
    EndSheet = ActiveSheet.Cells(1, 4).Value ' Номер листа, которым закончить суммирование

On Error GoTo ExitTheSub

For Each summari In Selection.Cells
    summari.Cells.Select
    SummObject = Selection.Cells(1, 1).Value ' Выделенная подпись значения, по нахождению которой будет создаваться саммари.
    InitialSummObjectAddress = Selection.Cells(1, 1).Address ' Начальное полжение выделенной подписи значения
    

    Dim SummArray(255) As Variant ' массив суммирования
    
    
' определяем точный адрес ячейки, в которой лежит нужное значение.
    For i = StartSheet To EndSheet
    Sheets(i).Activate
    ActiveSheet.Columns(2).Select
    Selection.Copy
    Sheets("Data").Select
    ActiveSheet.Cells(1, 1).Select
    ActiveSheet.Paste
    With Selection
        .MergeCells = False
    End With
    Sheets("Data").Select
    ActiveSheet.Columns(2).Select

    MyPlace = 1
    FullRow = 0
    EmptyRow = 0
    SummObjectAddress = ""
    For Each c In Selection.Rows
        c.Rows.Select
        MyValue = Selection.Cells(1, 1).Value
        MyAdress = Selection.Cells(1, 1).Address
        If IsEmpty(MyValue) Then ' 3 ифа внизу нужны, чтобы пропускать таблички регионов, в которых нет нужного оператора
            EmptyRow = 1
        End If
        If IsEmpty(MyValue) = False Then
            FullRow = 1
        End If
        If EmptyRow = 1 And FullRow = 1 And IsEmpty(MyValue) Then
            EmptyRow = 0
            FullRow = 0
            MyPlace = MyPlace + 1
        End If
        If MyValue = SummObject And MyPlace = SummObjectPlace Then
            SummObjectAddress = MyAdress
            Exit For
        If MyValue = SummObject Then
            MyPlace = MyPlace + 1
        End If
        End If
        If MyPlace > SummObjectPlace Then
            Exit For
        End If
                
    Next c

' Точный адрес ячейки для суммирования на листе в переменной SummObjectAddress
    Sheets("Data").Select
    ActiveSheet.Columns(1).Select
    Selection.Clear
    ActiveSheet.Columns(2).Select
    Selection.Clear
' почистили лист Data
    
    If SummObjectAddress <> "" Then
    Sheets(i).Activate
    Range(SummObjectAddress, Range(SummObjectAddress).End(xlToRight)).Select
    pos = 1
    For Each Sum In Selection.Cells
        SummArray(pos) = SummArray(pos) + Sum
        pos = pos + 1
    Next Sum
    End If
Next i

    Sheets(1).Activate
    Range(InitialSummObjectAddress, Range(InitialSummObjectAddress).End(xlToRight)).Select
        pos = 2
        MyCounter = 1
    For Each t In Selection.Cells
        If MyCounter > 1 Then
            t.Value = SummArray(pos)
            pos = pos + 1
        End If
        MyCounter = MyCounter + 1
    Next t
    Erase SummArray
    
Next summari
Worksheets("Data").Delete
ExitTheSub:
Application.DisplayAlerts = True
End Sub

Sub Operator4_Add_EmptyRow()
Range("A715").Select
Lrow = 715
For i = 1 To 9 Step 1
    Rows(Lrow).Insert
Next i
Range("A575").Select
Lrow = 575
For i = 1 To 9 Step 1
    Rows(Lrow).Insert
Next i
Range("A435").Select
Lrow = 435
For i = 1 To 9 Step 1
    Rows(Lrow).Insert
Next i
Range("A295").Select
Lrow = 295
For i = 1 To 9 Step 1
    Rows(Lrow).Insert
Next i
Range("A123").Select
Lrow = 123
For i = 1 To 12 Step 1
    Rows(Lrow).Insert
Next i
Range("A1").Select
Lrow = 1
For i = 1 To 83 Step 1
    Rows(Lrow).Insert
Next i

End Sub

Sub Operators4_Total_Delete()
    Dim Firstrow As Long
    Dim Lastrow As Long
    Dim Lrow As Long
    Dim CalcMode As Long
    Dim ViewMode As Long

    With Application
        .ScreenUpdating = False
    End With

    'We use the ActiveSheet but you can replace this with
    'Sheets("MySheet")if you want
    With ActiveSheet

        'We select the sheet so we can change the window view
        .Select

        'If you are in Page Break Preview Or Page Layout view go
        'back to normal view, we do this for speed
        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView

        'Turn off Page Breaks, we do this for speed
        .DisplayPageBreaks = False

        'Set the first and last row to loop through
        Firstrow = 24
        Lastrow = 133

        'We loop from Lastrow to Firstrow (bottom to top)
        For Lrow = Lastrow To Firstrow Step -1

            'We check the values in the A column in this example
            With .Cells(Lrow, "A")

                If Not IsError(.Value) Then

                    If .Value = "Всего" Then .EntireRow.Delete
                    'This will delete each row with the Value "ron"
                    'in Column A, case sensitive.

                End If

            End With

        Next Lrow

    End With

    ActiveWindow.View = ViewMode
    With Application
        .ScreenUpdating = True
    End With

End Sub
Sub Month_merge()
    
    Dim Firstrow As Long
    Dim Lastrow As Long
    Dim Lrow As Long
    Dim CalcMode As Long
    Dim ViewMode As Long
    Dim DefaultValue As Long
    Dim DefaultValue1 As Long
    Dim StartSheet As Variant
    Dim StopSheet As Variant
    
    With Application
        .ScreenUpdating = False
    End With
    DefaultValue = "3"
    DefaultValue1 = "45"
    StartSheet = InputBox("Введите номер первого листа", "Выбор листа", DefaultValue)
    StopSheet = InputBox("Введите номер последнего листа", "Выбор листа", DefaultValue1)

    If IsNumeric(StartSheet) And IsNumeric(StopSheet) Then
    
    For i = StartSheet To StopSheet
    Sheets(i).Activate
    
    'We use the ActiveSheet but you can replace this with
    'Sheets("MySheet")if you want
    With ActiveSheet

        'We select the sheet so we can change the window view
        .Select

        'If you are in Page Break Preview Or Page Layout view go
        'back to normal view, we do this for speed
        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView

        'Turn off Page Breaks, we do this for speed
        .DisplayPageBreaks = False

        'Set the first and last row to loop through
        Firstrow = 1
        Lastrow = 2000

        'We loop from Lastrow to Firstrow (bottom to top)
        For Lrow = Lastrow To Firstrow Step -1

            'We check the values in the A column in this example
            With .Cells(Lrow, "D")

                If Not IsError(.Value) Then

                    If .Value = "Месяц" Or .Value = "Месяц опроса" Then
                        Range("D" & Lrow, "I" & Lrow).Select
                        With Selection
                             .MergeCells = True
                             
                            
                        End With
                    End If

                End If

            End With

        Next Lrow

    End With
    Next i


    ActiveWindow.View = ViewMode

    Else
        MsgBox ("Вы ввели какую-то фигню")
    End If
    
    With Application
        .ScreenUpdating = True
    End With
End Sub
Sub Month_UNmerge()
    Dim Firstrow As Long
    Dim Lastrow As Long
    Dim Lrow As Long
    Dim CalcMode As Long
    Dim ViewMode As Long
    Dim DefaultValue As Long
    Dim DefaultValue1 As Long
    Dim StartSheet As Variant
    Dim StopSheet As Variant

    With Application
        .ScreenUpdating = False
    End With

    DefaultValue = "3"
    DefaultValue1 = "45"
    StartSheet = InputBox("Введите номер первого листа", "Выбор листа", DefaultValue)
    StopSheet = InputBox("Введите номер последнего листа", "Выбор листа", DefaultValue1)

    If IsNumeric(StartSheet) And IsNumeric(StopSheet) Then
    
    For i = StartSheet To StopSheet
    Sheets(i).Activate

    'We use the ActiveSheet but you can replace this with
    'Sheets("MySheet")if you want
    With ActiveSheet

        'We select the sheet so we can change the window view
        .Select

        'If you are in Page Break Preview Or Page Layout view go
        'back to normal view, we do this for speed
        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView

        'Turn off Page Breaks, we do this for speed
        .DisplayPageBreaks = False

        'Set the first and last row to loop through
        Firstrow = 1
        Lastrow = 2000

        'We loop from Lastrow to Firstrow (bottom to top)
        For Lrow = Lastrow To Firstrow Step -1

            'We check the values in the A column in this example
            With .Cells(Lrow, "D")

                If Not IsError(.Value) Then

                    If .Value = "Месяц" Or .Value = "Месяц опроса" Then
                        Range("D" & Lrow, "I" & Lrow).Select
                        With Selection
                             .MergeCells = False
                             
                            
                        End With
                    End If

                End If

            End With

        Next Lrow

    End With
    Next i
    
    ActiveWindow.View = ViewMode

    Else
        MsgBox ("Вы ввели какую-то фигню")
    End If

    With Application
        .ScreenUpdating = True
    End With

End Sub
Sub Operators3_Total_Delete()
    Dim Firstrow As Long
    Dim Lastrow As Long
    Dim Lrow As Long
    Dim CalcMode As Long
    Dim ViewMode As Long

    With Application
        .ScreenUpdating = False
    End With

    'We use the ActiveSheet but you can replace this with
    'Sheets("MySheet")if you want
    With ActiveSheet

        'We select the sheet so we can change the window view
        .Select

        'If you are in Page Break Preview Or Page Layout view go
        'back to normal view, we do this for speed
        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView

        'Turn off Page Breaks, we do this for speed
        .DisplayPageBreaks = False

        'Set the first and last row to loop through
        Firstrow = 21
        Lastrow = 121

        'We loop from Lastrow to Firstrow (bottom to top)
        For Lrow = Lastrow To Firstrow Step -1

            'We check the values in the A column in this example
            With .Cells(Lrow, "A")

                If Not IsError(.Value) Then

                    If .Value = "Всего" Then .EntireRow.Delete
                    'This will delete each row with the Value "ron"
                    'in Column A, case sensitive.

                End If

            End With

        Next Lrow

    End With

    ActiveWindow.View = ViewMode
    With Application
        .ScreenUpdating = True
    End With

End Sub


Sub Operators3_Zero_Value_Row_Delete()
    Dim Firstrow As Long
    Dim Lastrow As Long
    Dim Lrow As Long
    Dim CalcMode As Long
    Dim ViewMode As Long
    Dim MySumm As Single
    Dim CheckCount As Long
    Dim FirstRowArray As Variant
    Dim LastRowArray As Variant
    
    CheckCount = 0
    
    FirstRowArray = Array(518, 372, 195, 1)
    LastRowArray = Array(536, 390, 213, 17)
    
    With Application
        .ScreenUpdating = False
    End With

    'We use the ActiveSheet but you can replace this with
    'Sheets("MySheet")if you want
    With ActiveSheet

        'We select the sheet so we can change the window view
        .Select

        'If you are in Page Break Preview Or Page Layout view go
        'back to normal view, we do this for speed
        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView

        'Turn off Page Breaks, we do this for speed
        .DisplayPageBreaks = False
For i = 0 To 3

        'Set the first and last row to loop through
        Firstrow = FirstRowArray(i)
        Lastrow = LastRowArray(i)

        'We loop from Lastrow to Firstrow (bottom to top)
        For Lrow = Lastrow To Firstrow Step -1
            MySumm = 0
            's1 = .Cells(Lrow, "C").Value
            's2 = .Cells(Lrow, "D").Value
            's3 = .Cells(Lrow, "E").Value
            ' Проверяем, что в нужном диапазоне все значения - числа и складываем их
            If IsNumeric(.Cells(Lrow, "C").Value) And IsNumeric(.Cells(Lrow, "D").Value) And IsNumeric(.Cells(Lrow, "E").Value) Then
                 MySumm = .Cells(Lrow, "C").Value + .Cells(Lrow, "D").Value + .Cells(Lrow, "E").Value 'Добавить  в проверку и сумму колонку для нового месяца
            Else
                MySumm = 1
            End If
            With .Cells(Lrow, "B")

                If Not IsError(MySumm) Then
                    'Удаляем ряд если сумма равна 0 и имеется метка строки во втором столбце
                    If MySumm = 0 And IsEmpty(.Value) = False And .Value <> "Нет ответа, затрудняется ответить" And .Value <> "Билайн" And .Value <> "МТС" And .Value <> "Мегафон" Then
                        .EntireRow.Delete
                        CheckCount = CheckCount + 1
                    'This will delete each row with MySumm = 0
                     End If
                End If

            End With

        Next Lrow
    Next i
    End With
'If CheckCount <> 28 Then
'    MsgBox ("Макрос удалил неверное количество строк")
'End If


    ActiveWindow.View = ViewMode
    With Application
        .ScreenUpdating = True
    End With

End Sub

Sub Operator3_Add_EmptyRow()
Range("A556").Select
Lrow = 556
For i = 1 To 8 Step 1
    Rows(Lrow).Insert
Next i
Range("A418").Select
Lrow = 418
For i = 1 To 8 Step 1
    Rows(Lrow).Insert
Next i
Range("A280").Select
Lrow = 280
For i = 1 To 8 Step 1
    Rows(Lrow).Insert
Next i
Range("A111").Select
Lrow = 111
For i = 1 To 11 Step 1
    Rows(Lrow).Insert
Next i
Range("A1").Select
Lrow = 1
For i = 1 To 77 Step 1
    Rows(Lrow).Insert
Next i

End Sub

Sub Operator5_Add_EmptyRow()
Range("A878").Select
Lrow = 878
For i = 1 To 10 Step 1
    Rows(Lrow).Insert
Next i
Range("A736").Select
Lrow = 736
For i = 1 To 10 Step 1
    Rows(Lrow).Insert
Next i
Range("A594").Select
Lrow = 594
For i = 1 To 10 Step 1
    Rows(Lrow).Insert
Next i
Range("A452").Select
Lrow = 452
For i = 1 To 10 Step 1
    Rows(Lrow).Insert
Next i
Range("A310").Select
Lrow = 310
For i = 1 To 10 Step 1
    Rows(Lrow).Insert
Next i
Range("A135").Select
Lrow = 135
For i = 1 To 13 Step 1
    Rows(Lrow).Insert
Next i
Range("A1").Select
Lrow = 1
For i = 1 To 89 Step 1
    Rows(Lrow).Insert
Next i

End Sub


Sub Operators5_Zero_Value_Row_Delete()
    Dim Firstrow As Long
    Dim Lastrow As Long
    Dim Lrow As Long
    Dim CalcMode As Long
    Dim ViewMode As Long
    Dim MySumm As Single
    Dim CheckCount As Long
    Dim FirstRowArray As Variant
    Dim LastRowArray As Variant
    
    CheckCount = 0
    
    FirstRowArray = Array(842, 694, 546, 398, 217, 1)
    LastRowArray = Array(860, 712, 564, 416, 235, 17)
    
    With Application
        .ScreenUpdating = False
    End With

    'We use the ActiveSheet but you can replace this with
    'Sheets("MySheet")if you want
    With ActiveSheet

        'We select the sheet so we can change the window view
        .Select

        'If you are in Page Break Preview Or Page Layout view go
        'back to normal view, we do this for speed
        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView

        'Turn off Page Breaks, we do this for speed
        .DisplayPageBreaks = False
For i = 0 To 5

        'Set the first and last row to loop through
        Firstrow = FirstRowArray(i)
        Lastrow = LastRowArray(i)

        'We loop from Lastrow to Firstrow (bottom to top)
        For Lrow = Lastrow To Firstrow Step -1
            MySumm = 0
            's1 = .Cells(Lrow, "C").Value
            's2 = .Cells(Lrow, "D").Value
            's3 = .Cells(Lrow, "E").Value
            ' Проверяем, что в нужном диапазоне все значения - числа и складываем их
            If IsNumeric(.Cells(Lrow, "C").Value) And IsNumeric(.Cells(Lrow, "D").Value) And IsNumeric(.Cells(Lrow, "E").Value) Then
                 MySumm = .Cells(Lrow, "C").Value + .Cells(Lrow, "D").Value + .Cells(Lrow, "E").Value 'Добавить  в проверку и сумму колонку для нового месяца
            Else
                MySumm = 1
            End If
            With .Cells(Lrow, "B")

                If Not IsError(MySumm) Then
                    'Удаляем ряд если сумма равна 0 и имеется метка строки во втором столбце
                    If MySumm = 0 And IsEmpty(.Value) = False And .Value <> "Нет ответа, затрудняется ответить" And .Value <> "Билайн" And .Value <> "МТС" And .Value <> "Мегафон" Then
                        .EntireRow.Delete
                        CheckCount = CheckCount + 1
                    'This will delete each row with MySumm = 0
                     End If
                End If

            End With

        Next Lrow
    Next i
    End With
'If CheckCount <> 28 Then
'    MsgBox ("Макрос удалил неверное количество строк")
'End If


    ActiveWindow.View = ViewMode
    With Application
        .ScreenUpdating = True
    End With

End Sub
Sub Operators5_Total_Delete()
    Dim Firstrow As Long
    Dim Lastrow As Long
    Dim Lrow As Long
    Dim CalcMode As Long
    Dim ViewMode As Long

    With Application
        .ScreenUpdating = False
    End With

    'We use the ActiveSheet but you can replace this with
    'Sheets("MySheet")if you want
    With ActiveSheet

        'We select the sheet so we can change the window view
        .Select

        'If you are in Page Break Preview Or Page Layout view go
        'back to normal view, we do this for speed
        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView

        'Turn off Page Breaks, we do this for speed
        .DisplayPageBreaks = False

        'Set the first and last row to loop through
        Firstrow = 26
        Lastrow = 145

        'We loop from Lastrow to Firstrow (bottom to top)
        For Lrow = Lastrow To Firstrow Step -1

            'We check the values in the A column in this example
            With .Cells(Lrow, "A")

                If Not IsError(.Value) Then

                    If .Value = "Всего" Then .EntireRow.Delete
                    'This will delete each row with the Value "ron"
                    'in Column A, case sensitive.

                End If

            End With

        Next Lrow

    End With

    ActiveWindow.View = ViewMode
    With Application
        .ScreenUpdating = True
    End With

End Sub

Sub Operators4_Zero_Value_Row_Delete()
    Dim Firstrow As Long
    Dim Lastrow As Long
    Dim Lrow As Long
    Dim CalcMode As Long
    Dim ViewMode As Long
    Dim MySumm As Single
    Dim CheckCount As Long
    Dim FirstRowArray As Variant
    Dim LastRowArray As Variant
    
    CheckCount = 0
    
    FirstRowArray = Array(671, 524, 206, 1)
    LastRowArray = Array(689, 542, 224, 17)
    
    With Application
        .ScreenUpdating = False
    End With

    'We use the ActiveSheet but you can replace this with
    'Sheets("MySheet")if you want
    With ActiveSheet

        'We select the sheet so we can change the window view
        .Select

        'If you are in Page Break Preview Or Page Layout view go
        'back to normal view, we do this for speed
        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView

        'Turn off Page Breaks, we do this for speed
        .DisplayPageBreaks = False
For i = 0 To 3

        'Set the first and last row to loop through
        Firstrow = FirstRowArray(i)
        Lastrow = LastRowArray(i)

        'We loop from Lastrow to Firstrow (bottom to top)
        For Lrow = Lastrow To Firstrow Step -1
            MySumm = 0
            's1 = .Cells(Lrow, "C").Value
            's2 = .Cells(Lrow, "D").Value
            's3 = .Cells(Lrow, "E").Value
            ' Проверяем, что в нужном диапазоне все значения - числа и складываем их
            If IsNumeric(.Cells(Lrow, "C").Value) And IsNumeric(.Cells(Lrow, "D").Value) And IsNumeric(.Cells(Lrow, "E").Value) Then
                 MySumm = .Cells(Lrow, "C").Value + .Cells(Lrow, "D").Value + .Cells(Lrow, "E").Value 'Добавить  в проверку и сумму колонку для нового месяца
            Else
                MySumm = 1
            End If
            With .Cells(Lrow, "B")

                If Not IsError(MySumm) Then
                    'Удаляем ряд если сумма равна 0 и имеется метка строки во втором столбце
                    If MySumm = 0 And IsEmpty(.Value) = False And .Value <> "Нет ответа, затрудняется ответить" And .Value <> "Билайн" And .Value <> "МТС" And .Value <> "Мегафон" Then
                        .EntireRow.Delete
                        CheckCount = CheckCount + 1
                    'This will delete each row with MySumm = 0
                     End If
                End If

            End With

        Next Lrow
    Next i
    End With
If CheckCount <> 28 Then
    MsgBox ("Макрос удалил неверное количество строк")
End If


    ActiveWindow.View = ViewMode
    With Application
        .ScreenUpdating = True
    End With

End Sub



Sub hyperlinks_range_update()
MsgBox "Макрос сдвигает диапазон, на который ссылается гиперссылка на указанное количество строк вверх на активном листе. Внимательно прочтите инструкции при вводе значений. Не работает для гиперссылок, ссылающихся на одну ячейку."


'
' hyperlinks_correction Макрос
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
    DefaultValue = "3"
    DefaultValue1 = "45"
    StartSheet = InputBox("Введите номер первого листа. На первом листе должно быть указано 1) кол-во строк, в которых изменения не было в B1. 2) Количество удаленных строк в каждом листе в C1. Оно должно быть одинаковым для всех листов от номера первого до номера последнего.", "Выбор листа", DefaultValue)
    StopSheet = InputBox("Введите номер последнего листа", "Выбор листа", DefaultValue1)

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
        HyperlinkFound = InStr(MyLink, MyWorkSheet)
        If HyperlinkFound > 0 Then
            FirstRange = "A"
            LastRange = "K"
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
        MsgBox ("Вы ввели какую-то фигню")
    End If
    
    With Application
        .ScreenUpdating = True
    End With
    
    
    
End Sub

Sub Pokraska_yacheiki_po_usloviyu()


'Красит значения на активном листе в соответствии со значениями в листе test в той-же ячейке

MyAddress = 0
MyValue = 0
DefaultValue = 0.95
UsloviePokraski = InputBox("Введите тестируемое значение", "Выбор значения", DefaultValue)

MySheet = ActiveSheet.Name
 For Each i In Selection.Cells ' для диапазона вставить имя диапазона в [] [diapazon]
    MyAddress = i.Address
    Sheets("test").Select
    Range(MyAddress).Select
    MyValue = Selection.Value
    If IsError(MyValue) = False Then
        If IsNumeric(MyValue) And MyValue > UsloviePokraski Then ' условие покраски. Менять если надо
'            With Selection.Font
'                .ThemeColor = xlThemeColorDark1
'                .TintAndShade = 0
'            End With
            Sheets(MySheet).Select
            Range(MyAddress).Select
            With Selection.Interior
                .ColorIndex = 3 ' таблица цветов http://www.automateexcel.com/2004/08/18/excel_color_reference_for_colorindex/
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
            End With
        End If
    End If
Next i

End Sub


Sub SimExcell()

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


Dim RangeArray() As Integer
Dim PositionArray() As Integer
Dim ResultArray() As Integer
MaxCombination = 3 ' максимальное количество сочетаний
PresentValue = 1 ' Значение, которое должно быть в сочетании
MaxYSize = 0 ' Определение нижней границы диапазона значений
IterationCounter = 1 ' Адрес строки для сгенерированной комбинации на листе res
CombinationCounter = 0 ' Подсчёт количества сочетаний в сгенерированной комбинации
MaxIterationCoutner = 1 ' Максимально возможное количество итераций для  текущей позиции
PreviousIterationCounter = 0 ' количество итераций для  предыдущей позиции, чтобы убрать дубли

WorkAreaXSize = Selection.Columns.Count 'Длинна матрицы начальных условий
ReDim RangeArray(WorkAreaXSize)
ReDim PositionArray(WorkAreaXSize + 1)
ReDim ResultArray(WorkAreaXSize)
RangeArray(0) = 1 ' На всякий случай, данный элемент не используется для соответствия индексов массива и номеров ячеек Экселя

For Each Column In Selection.Columns ' Определяем в массив RangeArray, сколько элементов содержит матрица начальных условий в каждом столбце.

    Column.SpecialCells(xlCellTypeConstants).Select
    WorkAreaYSize = Selection.Cells.Count
    If WorkAreaYSize > MaxYSize Then
        MaxYSize = WorkAreaYSize
    End If
    RangeArray(Column.Column) = WorkAreaYSize
    
Next ' занесли длинну каждой колонки в массив RangeArray

Range(Cells(1, 1), Cells(MaxYSize, WorkAreaXSize)).Select
ValueArray = Selection ' Это массив значений, которые надо перебрать. ( todo Нужно переделать, чтобы точно определялся)


' Создаём массив начальной позиции, Это адрес, откуда брать значения в матрице начальных условий. Значения массива -  клетка, номер элемента массива - строка.

    For i = 1 To WorkAreaXSize
        PositionArray(i) = 1
    Next i






    For i = 1 To WorkAreaXSize  'Сколько столбцов, столько итераций
    
    For e = 1 To i ' Вычисляем максимальное количество комбинаций для данной итерации, как произведение кол-ва колонок на длинну каждого столбца.
    
        MaxIterationCoutner = MaxIterationCoutner * RangeArray(e) ' Количество сочетаний всех вариантов в начальной матрице условий задачи равно произведению количесва элементов каждого её столбца.
        
    Next e
    
    For y = 1 To MaxIterationCoutner 'Подставляем в финальный массив ResultArray() значения по адресу из  PositionArray()
        If PreviousIterationCounter < y Then  ' При следующем проходе не должны повторятся итерации из предыдущего, чтобы не было дублей.
            For j = 1 To WorkAreaXSize
                ResultArray(j) = ValueArray(PositionArray(j), j) 'Подставляем в финальный массив ResultArray() значения по адресу из  PositionArray()
                If ResultArray(j) = PresentValue Then
                    CombinationCounter = CombinationCounter + 1 '  считаем количество сочетаний значения PresentValue, которые не должны повторятся больше чем CombinationCounter раз
                End If
        

            Next j ' имеем ResultArray(), заполненный значениями
            
        
        ' сохранить ResultArray() на лист res в строку с адресом IterationCounter если он подходит по условиям задачи
            If CombinationCounter <= MaxCombination Then
                
                For d = 1 To WorkAreaXSize
                    Sheets("res").Cells(IterationCounter, d).Value = ResultArray(d)
                    
                Next d
                IterationCounter = IterationCounter + 1 ' Следующий подходящий результат сохранять на следующую строку
            End If
        
            
        
        End If
        
        CombinationCounter = 0 ' обнуляем количество сочетаний после ифа на всякий случай.
        
        
        PositionArray(1) = PositionArray(1) + 1 ' Тупейший алгоритм определения следующего адреса, откуда брать значения. прибавляем к первому значению единицу и сдвигаем её до тех пор, пока все предельные условия в RangeArray() будут меньше или равны значениям PositionArray()
        For ErrorCorrections = 1 To WorkAreaXSize
            If PositionArray(ErrorCorrections) > RangeArray(ErrorCorrections) Then
                PositionArray(ErrorCorrections) = 1
                'If IterationCounter <= MaxIterationCoutner Then ' Проверка, чтобы не было ошибки на превышение длинны массива PositionArray при выполнении этого участка кода после финальной итерации
                PositionArray(ErrorCorrections + 1) = PositionArray(ErrorCorrections + 1) + 1
               ' End If
            End If
        Next ErrorCorrections
        
        
        
    Next y
    
    PreviousIterationCounter = MaxIterationCoutner 'запоминаем, сколько итераций пропустить в следующем тике.
    
    MaxIterationCoutner = 1 'Обнуляем начельные условия для следующей итерации
    For s = 1 To WorkAreaXSize
        PositionArray(s) = 1
    Next s

Next i ' Следующая итерация


'        PositionArray(y) = 1
'    Next y
'        For j = 1 To RangeArray(i)
        
'            If i = cell.Column Then
            
                            ' ничего не делать, адрес совпадает с текущей ячейкой
                            
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
Dim RangeArray() As Integer
Dim CombinationCount As Long
CombinationCount = 1
Dim PositionArray() As Integer


WorkAreaXSize = Selection.Columns.Count

ReDim RangeArray(WorkAreaXSize)

For Each Column In Selection.Columns
    
    
    Column.SpecialCells(xlCellTypeConstants).Select
    WorkAreaYSize = Selection.Cells.Count
    RangeArray(Column.Column) = WorkAreaYSize
    CombinationCount = CombinationCount * RangeArray(Column.Column)
    If CombinationCount > 1000000 Then
        MsgBox ("Кол-во комбинаций в данной точке = " & CombinationCount & " не помещается в excel.")
       Exit Sub
    End If
    
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


Dim PositionArray() As Integer
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
    
    
Range(Cells(1, 1), Cells(CombinationCount, WorkAreaXSize)).Select

ExitTheSub:
With ActiveSheet
    ActiveWindow.View = ViewMode
End With
With Application
    .ScreenUpdating = True
End With

ActiveWorkbook.Save
    
    
    
End Sub

Sub Sort_by_Olesya() ' сортируем таблицу по Олесе

Dim MyStartRow As Long
        
w = OlesyaSortSelection()
    
End Sub


Function OlesyaSortSelection()

MyMainArray = Selection
For j = 2 To UBound(MyMainArray, 2)
    For i = 1 To UBound(MyMainArray, 1)
    
        
        MySelectionValue = MyMainArray(i, j)
        CutValue = Split(MySelectionValue, "[")
        
            If UBound(CutValue) > 0 Then
                MyValue = Replace(MyValue, "  ))", "))")
            End If
    Next i
Next j
    




End Function

Sub Перекраска_и_удаление()
'
' Ищет совпадение заголовков, при совпадении перекрашивает каждую клетку соседнего справа столбца со значимым различием в синий цвет
' как метку значимого понижения.
    DownBorder = 7 ' 7 это нижняя граница заголовка с метками для проверки. поменять если шапка другая
    Dim DeleteAdressArray() As Variant ' Массив с номерами столбцов для удаления
    DelI = 1 ' Индекс Массива
    OldValue = 0
For Each MyCol In Selection.Columns
    MyCol.Select
        
    MyCell = Cells(DownBorder, MyCol.Column).Select
    s = Selection
    
    If IsArray(s) = True Then
        
        MyValue = s(1, 1)
        
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


Sub удалить_проценты_оставить_значимости()

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



Sub apply_autofilter_across_worksheets()
'Updateby Extendoffice
    Dim xWs As Worksheet
    On Error Resume Next

    
    For Each xWs In Worksheets
        Current_sheet_name = xWs.Name
        If Current_sheet_name = "образец" Or Current_sheet_name = "NPS DATA" Then
            
        Else
            'xWs.Range("A2").AutoFilter 1, Criteria1:=Array("(tns - 900, maxima - 1465, vp - 200)", "vp",   "Maxima",   "socis",    "tomsk",    "cair", "gepi", "for",  "ИТОГО",    "Подрядчик",    "Москва vp",    "Москва maxima") Operator:=xlAnd, Criteria2:="<>vp"
            'основная выборка
            'xWs.Range("A2").AutoFilter 1, Criteria1:=Array("(tns - 900, maxima - 1465, vp - 200)", "vp", "Maxima", "socis", "tomsk", "cair", "gepi", "for", "ИТОГО", "Подрядчик", "Москва vp", "Москва maxima", "mi-50", "Москва mi-50"), Operator:=xlFilterValues ' всё
            'xWs.Range("A2").AutoFilter 1, Criteria1:=Array("(tns - 900, maxima - 1465, vp - 200)", "vp", "Maxima", "socis", "tomsk", "gepi", "for", "Москва vp", "Москва maxima", "mi-50", "Москва mi-50"), Operator:=xlFilterValues ' cair
            'xWs.Range("A2").AutoFilter 1, Criteria1:=Array("(tns - 900, maxima - 1465, vp - 200)", "vp", "Maxima", "socis", "tomsk", "cair", "gepi", "Москва vp", "Москва maxima", "mi-50", "Москва mi-50"), Operator:=xlFilterValues ' for
            'xWs.Range("A2").AutoFilter 1, Criteria1:=Array("(tns - 900, maxima - 1465, vp - 200)", "vp", "Maxima", "socis", "tomsk", "cair", "for", "Москва vp", "Москва maxima", "mi-50", "Москва mi-50"), Operator:=xlFilterValues ' gepi
            'xWs.Range("A2").AutoFilter 1, Criteria1:=Array("(tns - 900, maxima - 1465, vp - 200)", "vp", "Maxima", "socis", "cair", "gepi", "for", "Москва vp", "Москва maxima", "mi-50", "Москва mi-50"), Operator:=xlFilterValues ' tomsk
            'xWs.Range("A2").AutoFilter 1, Criteria1:=Array("(tns - 900, maxima - 1465, vp - 200)", "vp", "Maxima", "tomsk", "cair", "gepi", "for", "Москва vp", "Москва maxima", "mi-50", "Москва mi-50"), Operator:=xlFilterValues ' socis
            'xWs.Range("A2").AutoFilter 1, Criteria1:=Array("Maxima", "socis", "tomsk", "cair", "gepi", "for", "Москва maxima", "mi-50", "Москва mi-50"), Operator:=xlFilterValues ' vp
            'xWs.Range("A2").AutoFilter 1, Criteria1:=Array("vp", "socis", "tomsk", "cair", "gepi", "for", "Москва vp", "mi-50", "Москва mi-50"), Operator:=xlFilterValues ' maxima
            'xWs.Range("A2").AutoFilter 1, Criteria1:=Array("Maxima", "vp", "socis", "tomsk", "cair", "gepi", "for", "Москва maxima", "Москва vp"), Operator:=xlFilterValues ' mi-50
            'буст
            'xWs.Range("A2").AutoFilter 1, Criteria1:=Array("tomsk", "cair", "socis", "Общая", "nari", "ИТОГО", "Подрядчик", "Саратовская область vp", "Саратовская область socis"), Operator:=xlFilterValues ' всё
            'xWs.Range("A2").AutoFilter 1, Criteria1:=Array("tomsk", "socis", "Общая", "nari", "Саратовская область vp", "Саратовская область socis"), Operator:=xlFilterValues ' cair
            'xWs.Range("A2").AutoFilter 1, Criteria1:=Array("tomsk", "cair", "socis", "Общая", "Саратовская область vp", "Саратовская область socis"), Operator:=xlFilterValues ' nari
            'xWs.Range("A2").AutoFilter 1, Criteria1:=Array("cair", "socis", "Общая", "nari", "Саратовская область vp", "Саратовская область socis"), Operator:=xlFilterValues ' tomsk
            'xWs.Range("A2").AutoFilter 1, Criteria1:=Array("tomsk", "cair", "nari", "Саратовская область vp"), Operator:=xlFilterValues ' socis
            xWs.Range("A2").AutoFilter 1, Criteria1:=Array("tomsk", "cair", "socis", "nari", "Саратовская область socis"), Operator:=xlFilterValues ' vp
        End If
    Next
    Call удаление
        
End Sub


Sub unset_autofilter_across_worksheets()
'Updateby Extendoffice
    Dim xWs As Worksheet
    On Error Resume Next
    For Each xWs In Worksheets
        If Current_sheet_name = "образец" Or Current_sheet_name = "NPS DATA" Then
            
        Else

            'xWs.Range("A2").AutoFilter 1, "<> "
            xWs.Range("$A$2:$B$92").AutoFilter Field:=1
        End If
    Next
    
    
End Sub


Sub удаление()

Dim xWs As Worksheet
On Error Resume Next
For Each xWs In Worksheets
        xWs.Activate
        Current_sheet_name = xWs.Name
        If Current_sheet_name = "образец" Or Current_sheet_name = "NPS DATA" Then
            
        Else
            xWs.Rows("3:3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Delete Shift:=xlUp
        
        End If
        
         
        
        'xWs.Rows("3:3").Select
        'Range(Selection, Selection.End(xlDown)).Select
        'Selection.Delete Shift:=xlUp
        
Next
    
    Call unset_autofilter_across_worksheets
    End Sub
    
    
Sub правка_вординга()

Dim xWs As Worksheet
On Error Resume Next
For Each xWs In Worksheets
        xWs.Activate
        Current_sheet_name = xWs.Name
        If Current_sheet_name = "образец" Or Current_sheet_name = "NPS DATA" Then
            
        Else
            xWs.Rows("3:3").Select
        
            For Each xCell In Selection.Cells
                MyValue = xCell.Value
                If MyValue = "(tns - 900, maxima - 1465, vp - 200)" Then
                    xCell.Select
                    Selection.Value = "Общая"
                End If
                If MyValue = "(tns - 900, maxima - 1465, vp - 200)" Then
                    xCell.Select
                    Selection.Value = "Общая"
                End If
            Next
            
        End If
        
        
Next
    
    End Sub
    
    
    
Sub comments_edit()
'
' comments_edit Макрос
'
SearchValue = "Кутузовская"
ReplaceValue = "Кутузовский"

For Each xCell In Selection.Cells
    If xCell.Comment Is Nothing Then
    Else
        xCell.Comment.Text
        MyValue = xCell.Comment.Text
    
        hasSearchValue = InStr(MyValue, SearchValue)
        If hasSearchValue > 0 Then
            MyValue = Replace(MyValue, SearchValue, ReplaceValue)
            xCell.Comment.Text MyValue
        
        End If
    End If
        
    
    
Next

End Sub
