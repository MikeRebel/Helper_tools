Attribute VB_Name = "Module1"
Private MyOptions() As Variant

Sub Cho_template()

w = WorksheetIsExist("CHO")
    If w = True Then

        Dim MyRangeSize As Long

        Sheets("CHO").Select
        Myrange = Range("A2: A60000").SpecialCells(xlCellTypeConstants).Select
        MyRangeSize = Selection.Cells.Count

'Заполняем массив с опциями шаблона.

        If MyRangeSize > 1 Then
            MyGetOptionsCorrect = My_Get_options(Myrange, MyRangeSize, MyReturnedOptions:=MyOptions)
        Else
            MsgBox ("Пустые параметры выгрузки")
        End If
        
        If MyGetOptionsCorrect = True Then
            MyResult = GenerateCHOTemplate(MyOptions, MyRangeSize)
        End If

       MyStop = Save_Data_File()

    End If

End Sub


Function My_Get_options(WorkRange As Variant, Size As Long, MyReturnedOptions As Variant) As Boolean
My_Get_options = True


Range("b2").Select
X = Selection.Value
Range("c2").Select
Y = Selection.Value
Range("d2").Select
Z = Selection.Value
n = X * Y * Z + (4 * (X - 1) + 8 + 2)

If n = Size Then
    MyRowsInCho = Z * X + (2 * (X - 1)) + 3 + 1

    ReDim Preserve MyOptions(1 To MyRowsInCho)

    MyOptions(1) = 5
    MyOptions(2) = 1
    MyOptions(3) = 2
    j = 0
    K = 0
    For i = 4 To MyRowsInCho
        j = j + 1
        If j <= Z Then
            MyOptions(i) = Y
        Else
            K = K + 1
            If K <= 1 Then
                MyOptions(i) = 2
            Else
                MyOptions(i) = 2
                K = 0
                j = 0
            End If
                
        End If
    
    Next i

    MySumm = 0
    For i = 1 To MyRowsInCho
        MySumm = MySumm + MyOptions(i)
    Next i

    If Not (MySumm = Size) Then
        My_Get_options = False
        MsgBox ("Массив параметров создан с ошибками. Обратитесь к программисту.")
    End If

Else
    My_Get_options = False
    MsgBox ("количество переменных не совпадает с расчётным")
End If

End Function

Function GenerateCHOTemplate(ExtractOptions() As Variant, Size As Long)
Range("A2: A60000").SpecialCells(xlCellTypeConstants).Select


Application.ScreenUpdating = False
Application.DisplayAlerts = False
w = WorksheetIsExist("data")
    If w = True Then Worksheets("data").Delete
MySheet = CreateSheet("data", True)


MyRow = 0
MyCol = 0


Sheets("CHO").Select
Range("A2: A60000").SpecialCells(xlCellTypeConstants).Select


' На листе data генерится шаблон выгрузки для cho файла
MyRow = MyRow + 1

        For Each MySelColumn In Selection.Rows
        
            MySelColumn.Rows.Select
            MyValue = Selection.Cells(1, 1).Value
            ChoRowLength = ExtractOptions(MyRow)
            MyCol = MyCol + 1
            
            If MyCol = 1 And ChoRowLength > 1 Then
                MyValue = "/" & MyValue & "'"
            ElseIf MyCol = 1 And ChoRowLength = 1 Then
                MyValue = "/" & MyValue
            ElseIf MyCol = ChoRowLength Then
                MyValue = "''" & MyValue
            Else
                MyValue = "''" & MyValue & "'"
            End If
            
                
            
            
            If MyCol < ChoRowLength Then
                
                Sheets("data").Select
                ActiveSheet.Cells(MyRow, MyCol).Select
                Selection.Value = MyValue
                
            Else
                Sheets("data").Select
                ActiveSheet.Cells(MyRow, MyCol).Select
                Selection.Value = MyValue
                MyRow = MyRow + 1
                MyCol = 0
            End If
            Sheets("CHO").Select
        Next MySelColumn
        
   


Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Function

Function Save_Data_File()

Sheets("data").Select
'Range(Range("A2").End(xlDown), Range("A2").End(xlToRight)).Select
'Application.ScreenUpdating = False
'Application.DisplayAlerts = False


' Каждый шаблон помещается в отдельный фаил *.lot и, при помощи BatArray(), создается исполняемый фаил *.bat
' Все файлы сохраняются в папку, указанную в настройках.


Do
    fName = Application.GetSaveAsFilename
Loop Until fName <> False


ActiveWorkbook.SaveAs Filename:=fName, FileFormat:= _
xlText, CreateBackup:=False
'ActiveWorkbook.Close
'Worksheets("data").Delete
'MySheet = CreateSheet("data", True)

'Application.DisplayAlerts = True
'Application.ScreenUpdating = True
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

