Attribute VB_Name = "Module5"
Option Explicit

Public Const Mname As String = "��������� �������� �������"

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
                .Caption = "������ � �������������"
            
        ' ��� ������� ��� ��������� �����������.
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "�������� ����� �����������."
                    .FaceId = 71
                    .OnAction = "'" & ThisWorkbook.Name & "'!" & "hyperlinks_correction"
                End With

                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "�������� �������� ������� � ������������ �����."
                    .FaceId = 72
                    .OnAction = "'" & ThisWorkbook.Name & "'!" & "hyperlinks_range_update"
                End With
        
            End With

        ' ���� ��� ������ � ���������� �� 3 ���������.
            Set MenuItem = .Controls.Add(Type:=msoControlPopup)
            With MenuItem
                .Caption = "�������"
                Set MenuItem1 = .Controls.Add(Type:=msoControlPopup)
                With MenuItem1
                    .Caption = "��������� ��������"

                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "3 ��������� ������� Total �� ��������"
                        .FaceId = 71
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operators3_Total_Delete"
                    End With

                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "3 ��������� ������� ���������� � ������� ������� �� ��������"
                        .FaceId = 72
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operators3_Zero_Value_Row_Delete"
                    End With
            
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "3 ��������� �������� ������ ������ � ��������"
                        .FaceId = 73
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operator3_Add_EmptyRow"
                    End With
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "4 ��������� ������� Total �� ��������"
                        .FaceId = 74
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operators4_Total_Delete"
                    End With

                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "4 ��������� ������� ���������� � ������� ������� �� ��������"
                        .FaceId = 75
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operators4_Zero_Value_Row_Delete"
                    End With
            
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "4 ��������� �������� ������ ������ � ��������"
                        .FaceId = 76
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operator4_Add_EmptyRow"
                    End With
            
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "5 ���������� ������� Total �� ��������"
                        .FaceId = 77
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operators5_Total_Delete"
                    End With

                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "5 ���������� ������� ���������� � ������� ������� �� ��������"
                        .FaceId = 78
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operators5_Zero_Value_Row_Delete"
                    End With
            
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "5 ���������� �������� ������ ������ � ��������"
                        .FaceId = 79
                        .OnAction = "'" & ThisWorkbook.Name & "'!" & "Operator5_Add_EmptyRow"
                    End With
                End With
            End With
        
        
        Set FileItem = .Controls.Add(Type:=msoControlPopup)
        With FileItem
            .Caption = "������ � �������"
            
        ' ��� ������� ��� ��������� �����������.
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "����� ��������� ������ ������ �� ����� � ���� ���� � ����������� �������."
            .FaceId = 71
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "Copy_Sheet_From_Folder"
        End With

        With .Controls.Add(Type:=msoControlButton)
            .Caption = "��������� ������ ���� ��������� ������."
            .FaceId = 72
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "Copy_Every_Sheet_To_New_Workbook"
        End With
        
        End With

        Set CBCItem = .Controls.Add(Type:=msoControlPopup)
        With CBCItem
            .Caption = "��������"
            
        ' ��� ���������.
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "��������� �������� � �������"
                .FaceId = 71
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "perepbor_simulatora_column"
            End With
        
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "��������� �������� � ������"
                .FaceId = 72
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "perepbor_simulatora_row"
            End With
        End With

    End With
End Sub

Sub TestMacro()
    MsgBox "Hi there! Greetings from the Netherlands."
End Sub

