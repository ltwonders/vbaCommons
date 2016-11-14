Attribute VB_Name = "DateUtils"
Function getNextWorkday(Optional start_date As Date) As Date
    Dim result As Date
    
    If start_date = DateValue("00:00:00") Then start_date = Date
        If DateTime.Weekday(start_date + 1, vbMonday) < 6 Then            'Ìí¼ÓÖÜÄ©ÅÐ¶Ï
             result = start_date + 1
        Else
            If DateTime.Weekday(start_date + 2, vbMonday) < 6 Then
                result = start_date + 2
            ElseIf DateTime.Weekday(start_date + 3, vbMonday) < 6 Then
                result = start_date + 3
            End If
        End If
        
        getNextWorkday = result
End Function
Function getLastWorkday(Optional start_date As Date) As Date
    Dim result As Date
    If start_date = DateValue("00:00:00") Then start_date = Date
    
    If DateTime.Weekday(start_date - 1, vbMonday) < 6 Then            'Ìí¼ÓÖÜÄ©ÅÐ¶Ï
         result = start_date - 1
    Else
        If DateTime.Weekday(start_date - 2, vbMonday) < 6 Then
            result = start_date - 2
        ElseIf DateTime.Weekday(start_date - 3, vbMonday) < 6 Then
            result = start_date - 3
        End If
    End If

    getLastWorkday = result
End Function
Private Sub test()
    a = getNextWorkday(DateValue("2016/10/10"))
    b = getLastWorkday(DateValue("2016/10/10"))
End Sub
