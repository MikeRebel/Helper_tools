Attribute VB_Name = "Module7"
Private Sub Auto_Open()
With Application.CommandBars("Cell")
     .Protection = msoBarNoProtection
     With .Controls.Add(Before:=1, Temporary:=True)
          .Caption = "Полезные макросы"
          .OnAction = "CreateDisplayPopUpMenu"
     End With
     .Enabled = True
End With
End Sub


