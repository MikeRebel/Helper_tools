Attribute VB_Name = "Module5"
Option Explicit

Public Const Mname As String = "Несколько полезных функций"

Sub DeletePopUpMenu()
    ' Delete the popup menu if it already exists.
    On Error Resume Next
    Application.CommandBars(Mname).Delete
    On Error GoTo 0
End Sub

Sub CreateDisplayPopUpMenu()
    ' Delete any existing popup menu.
    Call DeletePopUpMenu

    ' Create the popup menu.
    Call Custom_PopUpMenu_1

    ' Display the popup menu.
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopup
    On Error GoTo 0
End Sub

Sub Custom_PopUpMenu_1()
    Dim MenuItem As CommandBarPopup
    Dim MenuItem1 As CommandBarPopup
    Dim HyperItem As CommandBarPopup
    Dim FileItem As CommandBarPopup
    Dim CBCItem As CommandBarPopup
    
    ' Add the popup menu.
    With Application.CommandBars.Add(Name:=Mname, Position:=msoBarPopup, _
         MenuBar:=False, Temporary:=True)
        
        Set HyperItem = .Controls.Add(Type:=msoControlPopup)
            With HyperItem
                .Caption = "Работа с гиперссылками"
            
        ' Два макроса для изменения гиперссылок.
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "Изменить адрес гиперссылок."
                    .FaceId = 71
                    .OnAction = "'" & ThisWorkbook.Name & "'!" & "hyperlinks_correction"
                End With

                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "Сдвинуть диапазон адресов в гиперссылках вверх."
                    .FaceId = 72
                    .OnAction = "'" & ThisWorkbook.Name & "'!" & "hyperlinks_range_update"
                End With
        
            End With

        ' Меню для работы с табличками на 3 оператора.
            Set MenuItem = .Controls.Add(Type:=msoControlPopup)
            With MenuItem
                .Caption = "Ретеншн"
                Set MenuItem1 = .Controls.Add(Type:=msoControlPopup)
                With MenuItem1
                    .Caption = "Почистить таблички"

                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "3 оператора Удалить Total из табличек"
                        .FaceId = 71
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operators3_Total_Delete"
                    End With

                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "3 оператора Удалить опреаторов с нулевым каунтом из табличек"
                        .FaceId = 72
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operators3_Zero_Value_Row_Delete"
                    End With
            
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "3 оператора Добавить пустые строки в таблички"
                        .FaceId = 73
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operator3_Add_EmptyRow"
                    End With
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "4 оператора Удалить Total из табличек"
                        .FaceId = 74
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operators4_Total_Delete"
                    End With

                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "4 оператора Удалить опреаторов с нулевым каунтом из табличек"
                        .FaceId = 75
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operators4_Zero_Value_Row_Delete"
                    End With
            
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "4 оператора Добавить пустые строки в таблички"
                        .FaceId = 76
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operator4_Add_EmptyRow"
                    End With
            
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "5 операторов Удалить Total из табличек"
                        .FaceId = 77
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operators5_Total_Delete"
                    End With

                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "5 операторов Удалить опреаторов с нулевым каунтом из табличек"
                        .FaceId = 78
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operators5_Zero_Value_Row_Delete"
                    End With
            
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "5 операторов Добавить пустые строки в таблички"
                        .FaceId = 79
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operator5_Add_EmptyRow"
                    End With
                End With
            End With
        
        
        Set FileItem = .Controls.Add(Type:=msoControlPopup)
        With FileItem
            .Caption = "Работа с файлами"
            
        ' Два макроса для изменения гиперссылок.
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Слить несколько файлов эксель из папки в один фаил в несколькими листами."
            .FaceId = 71
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "Copy_Sheet_From_Folder"
        End With

        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Сохранить каждый лист отдельным файлом."
            .FaceId = 72
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "Copy_Every_Sheet_To_New_Workbook"
        End With
        
        End With

        Set CBCItem = .Controls.Add(Type:=msoControlPopup)
        With CBCItem
            .Caption = "Конжоинт"
            
        ' Для конжоинта.
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Перебрать варианты в столбик"
                .FaceId = 71
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "perepbor_simulatora_column"
            End With
        
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Перебрать варианты в строку"
                .FaceId = 72
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "perepbor_simulatora_row"
            End With
        End With

    End With
End Sub

Sub TestMacro()
    MsgBox "Hi there! Greetings from the Netherlands."
End Sub

