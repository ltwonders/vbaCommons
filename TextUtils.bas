Attribute VB_Name = "TextUtils"
Public Function trimChar(original As String, Optional asc_num As Long = 32)
    trimChar = original
    
    Do While trimChar <> ""
        If Asc(Left(trimChar, 1)) = asc_num Then
            trimChar = Mid(trimChar, 2)
        ElseIf Asc(Right(trimChar, 1)) = asc_num Then
            trimChar = Mid(trimChar, 1, Len(trimChar) - 1)
        Else
            Exit Do
        End If
    Loop
    
End Function
Private Sub test()
    a = trimChar(";", Asc(";"))
    
End Sub

