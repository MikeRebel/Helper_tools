Attribute VB_Name = "Module4"
Sub Copy_Every_Sheet_To_New_Workbook()
Attribute Copy_Every_Sheet_To_New_Workbook.VB_ProcData.VB_Invoke_Func = " \n14"
'Working in 97-2010
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    Dim sh As Worksheet
    Dim DateString As String
    Dim FolderName As String

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    'Copy every sheet from the workbook with this macro
    
    Set Sourcewb = ActiveWorkbook

    'Create new folder to save the new files in
    DateString = Format(Now, "yyyy-mm-dd hh-mm-ss")
    'FolderName = Sourcewb.Path & "\" & DateString & " " & Sourcewb.Name
    FolderName = Sourcewb.Path
    'MkDir FolderName

    'Copy every visible sheet to a new workbook
    For Each sh In Sourcewb.Worksheets

        'If the sheet is visible then copy it to a new workbook
        If sh.Visible = -1 Then
            sh.Copy

            'Set Destwb to the new workbook
            Set Destwb = ActiveWorkbook

            'Determine the Excel version and file extension/format
            With Destwb
                If Val(Application.Version) < 12 Then
                    'You use Excel 97-2003
                    FileExtStr = ".xls": FileFormatNum = -4143
                Else
                    'You use Excel 2007-2010
                    If Sourcewb.Name = .Name Then
                        MsgBox "Your answer is NO in the security dialog"
                        GoTo GoToNextSheet
                    Else
                        Select Case Sourcewb.FileFormat
                        Case 51: FileExtStr = ".xlsx": FileFormatNum = 51
                        Case 52:
                            If .HasVBProject Then
                                FileExtStr = ".xlsm": FileFormatNum = 52
                            Else
                                FileExtStr = ".xlsx": FileFormatNum = 51
                            End If
                        Case 56: FileExtStr = ".xls": FileFormatNum = 56
                        Case Else: FileExtStr = ".xlsb": FileFormatNum = 50
                        End Select
                    End If
                End If
            End With

            'Change all cells in the worksheet to values if you want
            If Destwb.Sheets(1).ProtectContents = False Then
                With Destwb.Sheets(1).UsedRange
                    .Cells.Copy
                    .Cells.PasteSpecial xlPasteValues
                    .Cells(1).Select
                End With
                Application.CutCopyMode = False
            End If


            'Save the new workbook and close it
            With Destwb
                .SaveAs FolderName _
                      & "\" & Destwb.Sheets(1).Name & FileExtStr, _
                        FileFormat:=FileFormatNum
                .Close False
            End With

        End If
GoToNextSheet:
    Next sh

    'MsgBox "You can find the files in " & FolderName




    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
Application.DisplayAlerts = False
Sourcewb.Close
Application.DisplayAlerts = True

'Information

'If you run the code in Excel 2007-2010 it will look at the FileFormat of the parent workbook and
'save the new file in that format.
'Only if the parent workbook is an xlsm file and if there is no code in the new workbook it will
'save the new file as xlsx, If the parent workbook is not an xlsx, xlsm, or xls then it will be saved as xlsb.

'This are the main formats in Excel 2007-2010 :

'51 = xlOpenXMLWorkbook (without macro's in 2007-2010, xlsx)
'52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2010, xlsm)
'50 = xlExcel12 (Excel Binary Workbook in 2007-2010 with or without macro’s, xlsb)
'56 = xlExcel8 (97-2003 format in Excel 2007-2010, xls)

'If you always want to save in a certain format you can replace this part of the macro

'                Select Case Sourcewb.FileFormat
'                Case 51: FileExtStr = ".xlsx": FileFormatNum = 51
'                Case 52:
'                    If .HasVBProject Then
'                        FileExtStr = ".xlsm": FileFormatNum = 52
'                    Else
'                        FileExtStr = ".xlsx": FileFormatNum = 51
'                    End If
'                Case 56: FileExtStr = ".xls": FileFormatNum = 56
'                Case Else: FileExtStr = ".xlsb": FileFormatNum = 50
'                End Select


'With one of the one liners from this list

'FileExtStr = ".xlsb": FileFormatNum = 50
'FileExtStr = ".xlsx": FileFormatNum = 51
'FileExtStr = ".xlsm": FileFormatNum = 52
'FileExtStr = ".xls": FileFormatNum = 56


'Or maye you want to save the one sheet workbook to csv, txt or prn.
'(you can use this also if you run it in 97-2003)

'FileExtStr = ".csv": FileFormatNum = 6
'FileExtStr = ".txt": FileFormatNum = -4158
'FileExtStr = ".prn": FileFormatNum = 36

End Sub
