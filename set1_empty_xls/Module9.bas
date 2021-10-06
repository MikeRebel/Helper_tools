Attribute VB_Name = "Module9"
Sub test_100_percents()
Dim MyAdressCol, MyWorkSheet As Variant ' MyAdressCol - колонка В с метками, для определения где начинается или кончается диапазон для суммирования
Dim MyStartArray, MyStopArray As Variant ' MyStartArray - массив с началом диапазонов, MyStopArray - массив с концами диапазонов
Dim BaseWks As Workbook
Dim MyAdressIndex, i As Long
ReDim MyStartArray(1)
ReDim MyStopArray(1)
MyAdressIndex = 0
i = 1
d = 0

Set BaseWks = ActiveWorkbook
Set MyWorkSheet = BaseWks.Sheets(BaseWks.Sheets.Count)

MyWorkSheet.Copy after:=BaseWks.Sheets(BaseWks.Sheets.Count)

Set MyTempSheet = BaseWks.Sheets(BaseWks.Sheets.Count)
Set MyAdressCol = MyTempSheet.Columns(2)

Set BaseWks = Nothing
Set MyWorkSheet = Nothing
Set MyTempSheet = Nothing

'Удаляем слитые ячейки

'For Each c In MyAdressCol.Cells
'c.MergeCells = False
'Next c



' Определяем диапазон расчёта
For Each cell In MyAdressCol.Value2

    MyAdressIndex = MyAdressIndex + 1
    CurrentValue = cell
    
    If MyAdressIndex > 1 Then
        upperborder = isupperborder(MyAdressCol, MyAdressIndex)
    End If
    lowerborder = islowerborder(MyAdressCol, MyAdressIndex)
    
    If upperborder = False And lowerborder = False Then
        d = d + 1
    Else
        d = 0
    End If
    
'Если вподряд идёт 200 строк без нижней или верхней границы - считаем, что фаил закончен.

    If d > 200 Then
        Exit For
    End If
    
    If upperborder = True Then
        ReDim Preserve MyStartArray(i)
        MyStartArray(i) = MyAdressIndex
    ElseIf lowerborder = True Then
        ReDim Preserve MyStopArray(i)
        MyStopArray(i) = MyAdressIndex
        i = i + 1
    End If
    
Next cell
'Считаем сумму по диапазонам из массивов от MyStartArray до MyStopArray по каждой колонке на листе, начиная с 3-й


MyValue = 0
t = 0

    For Each MyRange In ActiveSheet.Columns
        If MyRange.Column > 2 Then
            For testValue = 1 To UBound(MyStartArray)

                MyRange.Range(Cells(MyStartArray(testValue), 1), Cells(MyStopArray(testValue), 1)).Select
                MyCurrentWorkRange = Selection
                For Each c In Selection.Cells
                    MyValue = MyValue + c.Value
                Next c
                MyValue = MyValue * 100
                MyValue = Round(MyValue, 0)
                
                'If MyValue = 100 Then
                '    Selection.Interior.Color = RGB(0, 255, 0)
                '    t = 0
                'Else
                If MyValue = 0 Then
                
                    t = t + 1
                    
                ElseIf MyValue < 100 Then
                    Selection.Interior.Color = RGB(255, 0, 0)
                    t = 0
                End If
                
               
                
            MyValue = 0
        
            Next testValue
        End If
        If t > 200 Then
            Exit Sub
        End If
       
    Next MyRange
    
    ActiveSheet.Name = "Tables Errors"

End Sub


'Верхняя граница True, если в текущей клетке CurrentValue есть значение, а в прошлой клетке (PreviousValue) было пусто.
Function isupperborder(TargetCollumn As Variant, CellIndex As Variant)


    CurrentValue = TargetCollumn.Cells(CellIndex, 1).Value
    PreviousValue = TargetCollumn.Cells(CellIndex - 1, 1).Value


    If IsEmpty(PreviousValue) = True And IsEmpty(CurrentValue) = False Then
        isupperborder = True
        Exit Function
    End If

isupperborder = False

End Function

'Нижняя граница True, если в текущей клетке CurrentValue есть значение, а в следующей клетке (NextValue) пусто или прочерк.

Function islowerborder(TargetCollumn As Variant, CellIndex As Variant)
    CurrentValue = TargetCollumn.Cells(CellIndex, 1).Value
    NextValue = TargetCollumn.Cells(CellIndex + 1, 1).Value


    If IsEmpty(NextValue) = True And IsEmpty(CurrentValue) = False And CurrentValue <> "-" Then
        islowerborder = True
        Exit Function
    ElseIf NextValue = "-" And IsEmpty(CurrentValue) = False And CurrentValue <> "-" Then
    
        islowerborder = True
        Exit Function
    End If

islowerborder = False

End Function


