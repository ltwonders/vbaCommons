Attribute VB_Name = "TextUtils"
Public Function trimChar(original As String, Optional asc_num As Long = 32)
    trimChar = original
    If trimChar = "" Then Exit Function
    
    Do While Asc(Left(trimChar, 1)) = asc_num
        trimChar = Right(trimChar, Len(trimChar) - 1)
    Loop
    Do While Asc(Right(trimChar, 1)) = asc_num
        trimChar = Left(trimChar, Len(trimChar) - 1)
    Loop
End Function

