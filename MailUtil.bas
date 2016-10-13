Attribute VB_Name = "MailUtil"
Public Function mailByRange(rng As Range, intro As String, subject As String, mailTo As String, Optional mailCC As String, Optional mailBCC As String) As Boolean
    Dim shtActive As Worksheet
    Dim rngSend As Range
    
On Error Resume Next
    
With Application
    .ScreenUpdating = False
    .EnableEvents = False
End With

If rng Is Nothing Then
    Set rngSend = ActiveSheet.UsedRange
Else
    Set rngSend = rng
End If

Set shtActive = ActiveSheet

With rngSend
    .Parent.Select
    .Select
    ActiveWorkbook.EnvelopeVisible = True
    With .Parent.MailEnvelope
        .introduction = intro
        With .Item
            .To = mailTo
            .cc = mailCC
            .bcc = mailBCC
            .subject = subject
            .Send
        End With
    End With
    
End With

shtActive.Select

If Err.Number > 0 Then Err.Clear Else mailByRange = True

stopMacro:
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    ActiveWorkbook.EnvelopeVisible = False
    
End Function
Public Function mailBySheet(sht_name As String, attach_name As String, mail_to As String, mail_sub As String) As Boolean
    Application.DisplayAlerts = False
    Sheets(sht_name).Copy
    ActiveWorkbook.SaveAs attach_name & ".xlsx"
    
    On Error Resume Next
    ActiveWorkbook.SendMail mail_to, mail_sub
    If Err.Number > 0 Then mailBySheet = False Else mailBySheet = True
    
    Err.Clear
    ActiveWorkbook.ChangeFileAccess xlReadOnly
    Kill ActiveWorkbook.FullName
    ActiveWorkbook.Close False
    Application.DisplayAlerts = True
End Function
