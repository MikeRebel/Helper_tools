Attribute VB_Name = "Module3"
Private myFiles() As String
Private Fnum As Long

Function Get_File_Names(MyPath As String, Subfolders As Boolean, _
                        ExtStr As String, myReturnedFiles As Variant) As Long
                        

    Dim Fso_Obj As Object, RootFolder As Object
    Dim SubFolderInRoot As Object, file As Object

    'Add a slash at the end if the user forget it
    If Right(MyPath, 1) <> "\" Then
        MyPath = MyPath & "\"
    End If

    'Create FileSystemObject object
    Set Fso_Obj = CreateObject("Scripting.FileSystemObject")

    Erase myFiles()
    Fnum = 0

    'Test if the folder exist and set RootFolder
    If Fso_Obj.FolderExists(MyPath) = False Then
        Exit Function
    End If
    Set RootFolder = Fso_Obj.GetFolder(MyPath)

    'Fill the array(myFiles)with the list of Excel files in the folder(s)
    'Loop through the files in the RootFolder
    For Each file In RootFolder.Files
        If LCase(file.Name) Like LCase(ExtStr) Then
            Fnum = Fnum + 1
            ReDim Preserve myFiles(1 To Fnum)
            myFiles(Fnum) = MyPath & file.Name
        End If
    Next file

    'Loop through the files in the Sub Folders if SubFolders = True
    If Subfolders Then
        Call ListFilesInSubfolders(OfFolder:=RootFolder, FileExt:=ExtStr)
    End If

    myReturnedFiles = myFiles
    Get_File_Names = Fnum
End Function

Sub ListFilesInSubfolders(OfFolder As Object, FileExt As String)
'Origenal SubFolder code from Chip Pearson
'http://www.cpearson.com/Excel/RecursionAndFSO.htm
'Changed by Ron de Bruin, 27-March-2008
    Dim SubFolder As Object
    Dim fileInSubfolder As Object

    For Each SubFolder In OfFolder.Subfolders
        ListFilesInSubfolders OfFolder:=SubFolder, FileExt:=FileExt

        For Each fileInSubfolder In SubFolder.Files
            If LCase(fileInSubfolder.Name) Like LCase(FileExt) Then
                Fnum = Fnum + 1
                ReDim Preserve myFiles(1 To Fnum)
                myFiles(Fnum) = SubFolder & "\" & fileInSubfolder.Name
            End If
        Next fileInSubfolder

    Next SubFolder
End Sub

Function GetFolderPath(Optional ByVal Title As String = "Выберите папку", _
                       Optional ByVal InitialPath As String = "c:\") As String
    ' функция выводит диалоговое окно выбора папки с заголовком Title,
   ' начиная обзор диска с папки InitialPath
   ' возвращает полный путь к выбранной папке, или пустую строку в случае отказа от выбора
   Dim PS As String: PS = Application.PathSeparator
    With Application.FileDialog(msoFileDialogFolderPicker)
        If Not Right$(InitialPath, 1) = PS Then InitialPath = InitialPath & PS
        .ButtonName = "Выбрать": .Title = Title: .InitialFileName = InitialPath
        If .Show <> -1 Then Exit Function
        GetFolderPath = .SelectedItems(1)
        If Not Right$(GetFolderPath, 1) = PS Then GetFolderPath = GetFolderPath & PS
    End With


'Sub ПримерИспользования_GetFolderPath()
'   ПутьКПапке = GetFolderPath("Заголовок окна", ThisWorkbook.Path)   ' запрашиваем имя папки
'   If ПутьКПапке = "" Then Exit Sub    ' выход, если пользователь отказался от выбора папки
'   MsgBox "Выбрана папка: " & ПутьКПапке, vbInformation
'End Sub

End Function


Sub Get_Sheet(PasteAsValues As Boolean, StartColumnIndex As Integer, ColumnsCount As Integer, SourceShName As String, _
              SourceShIndex As Integer, myReturnedFiles As Variant)
    Dim mybook As Workbook, BaseWks As Worksheet
    Dim CalcMode As Long
    Dim SourceSh As Variant
    Dim sh As Worksheet
    Dim CurrentWorksheetForAdd As Worksheet
    Dim i As Long

    'Change ScreenUpdating, Calculation and EnableEvents
    With Application
'        CalcMode = .Calculation
'        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    On Error GoTo ExitTheSub

    'Add a new workbook with one sheet
    Set BaseWks = Workbooks.Add(xlWBATWorksheet).Worksheets(1)


    'Check if we use a named sheet or the index
    If SourceShName = "" Then
        SourceSh = SourceShIndex
    Else
        SourceSh = SourceShName
    End If
    'Если надо скопировать каждый фаил на отдельный лист
    If ColumnsCount = 0 And StartColumnIndex = 0 Then
    'Loop through all files in the array(myFiles)
    For i = LBound(myReturnedFiles) To UBound(myReturnedFiles)
        Set mybook = Nothing
        On Error Resume Next
        Set mybook = Workbooks.Open(myReturnedFiles(i))
        On Error GoTo 0

        If Not mybook Is Nothing Then

            'Set sh and check if it is a valid
            On Error Resume Next
            Set sh = mybook.Sheets(SourceSh)

            If Err.Number > 0 Then
                Err.Clear
                Set sh = Nothing
            End If
            On Error GoTo 0

            If Not sh Is Nothing Then
                sh.Copy after:=BaseWks.Parent.Sheets(BaseWks.Parent.Sheets.Count)

                On Error Resume Next
                ActiveSheet.Name = mybook.Name
                On Error GoTo 0

                If PasteAsValues = True Then
                    With ActiveSheet.UsedRange
                        .Value = .Value
                    End With
                End If

            End If
            'Close the workbook without saving
            mybook.Close savechanges:=False
        End If

        'Open the next workbook
    Next i
    BaseWks.Delete
    'Если надо скопировать несколько столбцов из всех файлов на один лист. Новый фаил подставляется снизу
    ElseIf ColumnsCount > 0 And StartColumnIndex = 0 Then
    MyRowCount = 1
        'Loop through all files in the array(myFiles)
    For i = LBound(myReturnedFiles) To UBound(myReturnedFiles)
        
        Set mybook = Nothing
        On Error Resume Next
        Set mybook = Workbooks.Open(myReturnedFiles(i))
        On Error GoTo 0

        If Not mybook Is Nothing Then

            'Set sh and check if it is a valid
            On Error Resume Next
            Set sh = mybook.Sheets(SourceSh)

            If Err.Number > 0 Then
                Err.Clear
                Set sh = Nothing
            End If
            On Error GoTo 0

            If Not sh Is Nothing Then
                sh.Copy after:=BaseWks.Parent.Sheets(BaseWks.Parent.Sheets.Count)

                On Error Resume Next
                
                'WorkFileName = mybook.Name
                'WorkFileName = Mid(mybook.Path, 22, Len(mybook.Path) - 20)
                WorkFileName = mybook.Path & mybook.Name
                MyWorkSheet = ActiveSheet.Name
                
                On Error GoTo 0
                
                Range(Range("A1").End(xlDown), Range("A1").End(xlToRight)).Select
                Selection.Copy
                MySelection = Selection
                
                Sheets("Лист1").Activate
                ActiveSheet.Cells(MyRowCount, 1).Select
                ActiveSheet.Paste
                MyRowCount = MyRowCount + UBound(MySelection)
                If MyRowCount > 1000000 Then
                    MsgBox ("Эксель больше не может вместить в себя. Остановился на файле " & myReturnedFiles(i))
                    GoTo ExitTheSub
                End If
                
                On Error Resume Next
              
                If PasteAsValues = True Then
                    With ActiveSheet.UsedRange
                        .Value = .Value
                    End With
                End If
                For Each MyRow In Selection.Rows
                    MyRow.Rows.Select
                    Selection.Cells(1, ColumnsCount + 1).Value = WorkFileName
                Next MyRow
                
                Application.DisplayAlerts = False
                Sheets(MyWorkSheet).Delete
                Application.DisplayAlerts = True
            End If
            'Close the workbook without saving
            mybook.Close savechanges:=False
        End If
                
        'Open the next workbook
    Next i
            
    'Если надо скоприровать диапазон столбцов от столбца StartColumnIndex до ColumnsCount. Реализован вариант только для одного столбца.
    'Копируется транспонировано из столбца в строку.
    ElseIf ColumnsCount > 0 And StartColumnIndex > 0 Then
        NextRowIndex = 1 'следующий ряд, куда вставляются результаты
        MyStartRowCount = 1 'Номер начальной колонки, куда вставляется результат
        For i = LBound(myReturnedFiles) To UBound(myReturnedFiles)
            Set mybook = Nothing
            On Error Resume Next
            'итый фаил
            Set mybook = Workbooks.Open(myReturnedFiles(i))
            On Error GoTo 0

            If Not mybook Is Nothing Then

                'Set sh and check if it is a valid
                On Error Resume Next
                'первый лист файла, куда всё накапливается. Рабочий фаил.
                Set sh = mybook.Sheets(SourceSh)

                If Err.Number > 0 Then
                    Err.Clear
                    Set sh = Nothing
                End If
                On Error GoTo 0

                If Not sh Is Nothing Then
                    sh.Copy after:=BaseWks.Parent.Sheets(BaseWks.Parent.Sheets.Count)

                    On Error Resume Next
'                   CurrentWorksheetForAdd это лист, скопированный в рабочий фаил BaseWks новый лист с данными из файла i из него добываем
'                   в массив MyReturnedCol() нужный диапазон от StartColumnIndex до ColumnsCount, и копируем на новую строку в BaseWks его значения.

                    Set CurrentWorksheetForAdd = BaseWks.Parent.Sheets(BaseWks.Parent.Sheets.Count)
'                    ActiveSheet.Name = mybook.Name
                    On Error GoTo 0
'                   Если надо скопировать один столбец. Копирует транспонировано. Нужно дописать для диапазона.

                    If StartColumnIndex = ColumnsCount Then
                        
                        Dim MyCurrentCol As Variant
                        Dim MyReturnedCol As Variant
                        ' Массив со значениями копируемой колонки
                        MyCurrentCol = CurrentWorksheetForAdd.Columns(StartColumnIndex)
                        ReDim MyReturnedCol(1)
                        CurrentRowIndex = 0
                        For Each cell In MyCurrentCol
                            'Добавляем в MyReturnedCol непустые значения пока Empty не разлучит нас 50 раз.
                            If cell <> "" Then
                                CurrentRowIndex = CurrentRowIndex + 1
                                ReDim Preserve MyReturnedCol(CurrentRowIndex)
                                MyReturnedCol(CurrentRowIndex) = cell
                                StopCounter = 1
                            End If
                            If cell = "" Then
                                CurrentRowIndex = CurrentRowIndex + 1
                                StopCounter = StopCounter + 1
                            End If
                            If StopCounter > 50 Then
                                CurrentRowIndex = CurrentRowIndex - 50
                                Exit For
                            End If
                        Next cell
                        MyWorkSheet = ActiveSheet.Name
                        MyReturnedCol(0) = MyWorkSheet
                        BaseWks.Activate
                        
                        MyFinalRowCount = CurrentRowIndex + 1
                        If NextRowIndex > 0 And MyStartRowCount > 0 And MyFinalRowCount > 0 Then
                            Range(Cells(NextRowIndex, MyStartRowCount), Cells(NextRowIndex, MyFinalRowCount)).Select
                            j = 0
                            For Each cell In Selection.Cells
                                cell.Value = MyReturnedCol(j)
                                j = j + 1
                            Next cell
                        
                            NextRowIndex = NextRowIndex + 1
                        Else
                            MsgBox ("Данные на листе" & MyWorkSheet & " повреждены.")
                        End If
                        Erase MyReturnedCol
                    End If
                    
                    If PasteAsValues = True Then
                        With ActiveSheet.UsedRange
                            .Value = .Value
                        End With
                    
                    End If

                End If
                
            End If
                'Close the workbook without saving
                mybook.Close savechanges:=False
                'Удаляем промежуточный рабочий лист
                Application.DisplayAlerts = False
                CurrentWorksheetForAdd.Delete
                Application.DisplayAlerts = True
        'Open the next workbook
        Next i

        
    Else
        MsgBox ("Неверно заданы значения ColumnsCount =" & ColumnsCount & " или StartColumnIndex =" & StartColumnIndex)
        
    End If
        


   
    

    
    ' delete the first sheet in the workbook
    Application.DisplayAlerts = False
    On Error Resume Next
    
    On Error GoTo 0
    Application.DisplayAlerts = True

ExitTheSub:
    'Restore ScreenUpdating, Calculation and EnableEvents
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
'        .Calculation = CalcMode
    End With
End Sub

Sub Copy_Sheet_From_Folder()

'Перед работой установить ColumnsCount и StartColumnIndex в соответствии с инструкцией.

'First we call the function "Get_File_Names" to fill a array with all file names
'There are three arguments in this function that we can change.

'1) MyPath
'The folder where the files are
'Note: There is also a macro example "RDB_Merge_Data_Browse" that let you browse to the folder

'2) Subfolders
'True if you want to include subfolders

'3) ExtStr
'File extension of the files you want to merge.
'Examples are: "*.xls" , "*.csv" , "*.xlsx"
'"*.xlsm" ,"*.xlsb" , for all Excel file formats use "*.xl*"

'Then if there are files in the folder we call the macro "Get_Sheet"
'There are three arguments in this macro that we can change

'1) PasteAsValues
'True to paste as values (recommend)

'2) SourceShName
'Enter the name of the sheet that you have in every workbook
'If "" it will use the SourceShIndex

'3) SourceShIndex
'To avoid problems with different sheet names you can use the index
'If you use 1 it use the first worksheet in each workbook.

'4) StartColumnIndex номер начального столбца

' 5)ColumnsCount если 0 работает по алгоритму слить все файлы из папки в один на разные листы.
' если больше 0 - сливает указанное количество столбцов.


    Dim ПутьКПапке As String
    Dim myFiles As Variant
    Dim myCountOfFiles As Long
    ПутьКПапке = GetFolderPath("Выбери папку", ThisWorkbook.Path)   ' запрашиваем имя папки
    If ПутьКПапке = "" Then Exit Sub    ' выход, если пользователь отказался от выбора папки

    myCountOfFiles = Get_File_Names( _
                     MyPath:=ПутьКПапке, _
                     Subfolders:=True, _
                     ExtStr:="*.xlsx", _
                     myReturnedFiles:=myFiles)

    If myCountOfFiles = 0 Then
        MsgBox "No files that match the ExtStr in this folder"
        Exit Sub
    End If

    Get_Sheet _
            PasteAsValues:=True, _
            StartColumnIndex:=0, _
            ColumnsCount:=20, _
            SourceShName:="", _
            SourceShIndex:=1, _
            myReturnedFiles:=myFiles





End Sub
