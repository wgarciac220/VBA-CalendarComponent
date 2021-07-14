Public Sub AdvancedCalendar()
    
    With New CalendarForm
    .FontSize = 30
    .UseTheme Red
    .Show
    
    MsgBox .DateSelection, vbInformation, ThisWorkbook.Name
    End With
    
    MsgBox "Task Run Successfully!", vbInformation, ThisWorkbook.Name

End Sub

Sub ToRGB(ByVal Value As Long)


    B = Value \ 65536
    G = (Value - B * 65536) \ 256
    R = Value - B * 65536 - G * 256
    
    Debug.Print "RGB(" & R & "," & G & "," & B & ")"
End Sub


