Attribute VB_Name = "Module13"
Sub upd(ByVal Target As Range)

'    Application.OnTime Now() + TimeSerial(0, 0, 59), "upd"

    Range("J2:L2").Select
    Selection.AutoFill Destination:=Range("J2:L257")
   
    
End Sub

Public Function diconnect(MyAdress)
'

    Range(MyAdress).Select
    Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
End Function

Sub testZ()

Application.OnTime Now() + TimeSerial(0, 5, 0), "testZ"

For Each MyWorkSheet In Application.Worksheets

MyWorkSheet.Activate
Range("B1").Select
TestVal = Selection.Value

If Selection.Value = " UserName" Then

Columns("B").Select

For Each MyRow In Selection.Rows

    w = MyRow.DisplayFormat.Interior.ColorIndex
        If w < 0 Then
            Exit For
        End If
        
        MyCurrentTime = Now
'        MyFormatedCurrentTime = FormatDateTime(MyCurrentTime, vbShortTime)
        If MyRow.Row > 1 Then
            Range("I" & MyRow.Row).Select
            MyTimeValue = Selection.Value
            If Not (IsEmpty(MyTimeValue)) Then
'                MyFormatedTime = FormatDateTime(MyTimeValue, vbShortTime)
                CutFormatedTime = MyCurrentTime - CDate(MyTimeValue)
                CutFormatedTime = CDate(CutFormatedTime)
'                CutCurrentTimeValue = Split(MyFormatedCurrentTime, ":")
'                CutTimeValue = Split(MyFormatedTime, ":")
'                MinuteInt = CutCurrentTimeValue(1) - CutTimeValue(1)
'                MinuteInt = CInt(MinuteInt)
                 MinuteInt = CutFormatedTime
                 
                Range("I" & MyRow.Row).Select
                
                If Selection.Value = "Deleted" Then
                    Selection.Value = ""
                End If
                If w = 38 Then
                    If MinuteInt > "00:07:00" Then
                        d = diconnect("H" & MyRow.Row)
                    End If
                End If
                If MinuteInt > "00:14:00" Then
                    d = diconnect("H" & MyRow.Row)
                    Range("I" & MyRow.Row).Select
                    Selection.Value = "Deleted"
                End If
                
            End If
            
         End If
         
        Next MyRow
        
ElseIf Selection.Value = " Record" Then
    d = diconnect("G2")
    
Else
    Worksheets(MyWorkSheet).Delete
    
End If

Next MyWorkSheet

End Sub

