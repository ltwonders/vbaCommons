Attribute VB_Name = "RangeUtils"
Public Sub copyRows(sht_from As String, row_start As Long, row_end As Long, sht_to As String, Optional row_to As Long = 0)
    If row_to = 0 Then row_to = getLastRow(sht_to) + 1
    Sheets(sht_from).Rows(row_start & ":" & row_end).Copy
    Sheets(sht_to).Rows(row_to).PasteSpecial Paste:=xlPasteValues
End Sub
Public Sub copyRange(rng As Range, sht_to As String, col_to As Long, Optional row_to As Long = 0)
    If row_to = 0 Then row_to = getLastRow(sht_to, col_to) + 1
    rng.Copy
    Sheets(sht_to).Cells(row_to, col_to).PasteSpecial
End Sub
Public Sub copyCols(title_of_col As String, sht_from As String, row_title_at As Long, sht_to As String, col_to As Long, _
                    Optional row_to As Long = 0, Optional copy_with_title As Boolean = False)
    On Error Resume Next
    Dim colMatch As Long
        
    colMatch = Application.Match(title_of_col, Sheets(sht_from).Rows(row_title_at), 0)
    If Err.Number > 0 Then
        Err.Clear
        Exit Sub
    End If
    If Not copy_with_title And getLastRow(sht_from, colMatch) <= row_title_at Then Exit Sub         'If only title row and no title copied then exit sub
    If copy_with_title Then rowStart = row_title_at Else rowStart = row_title_at + 1
    strrngcopy = getColId(colMatch) & rowStart & ":" & getColId(colMatch) & getLastRow(sht_from, colMatch)   'Copy the entire column from sht_from
    If row_to = 0 Then row_to = getLastRow(sht_to, col_to) + 1
    copyRange rng:=Sheets(sht_from).Range(strrngcopy), sht_to:=sht_to, col_to:=col_to, row_to:=row_to
    
End Sub
Public Sub clearRows(sht As String, row_start As Long, Optional row_end As Long = 0)
    If row_end = 0 Then row_end = getLastRow(sht)
    If row_end >= row_start Then Sheets(sht).Rows(row_start & ":" & row_end).Clear Else Exit Sub
End Sub
Public Sub formatRange(rng As Range, Optional line_style As XlLineStyle = xlContinuous, Optional font_size As Long = 10, _
                Optional font_color As Long = 1, Optional font_bold As Boolean = False, Optional interior_color As Long = 2, Optional column_width As Long = 0)
    With rng
        .Borders.LineStyle = line_style
        .Interior.ColorIndex = interior_color
        If column_width > 0 Then .ColumnWidth = column_width
        With .Font
            .ColorIndex = font_color
            .Size = font_size
            .Bold = font_bold
        End With
    End With
End Sub
Public Function getLastRow(Optional sht_name As String, Optional col_index As Long = 1) As Long
    If sht_name = "" Then
        getLastRow = Cells(Cells.Rows.Count, col_index).End(xlUp).Row
    Else
        getLastRow = Sheets(sht_name).Cells(Cells.Rows.Count, col_index).End(xlUp).Row
    End If
End Function
Public Function getLastCol(Optional sht_name As String, Optional row_index As Long = 1) As Long
    If sht_name = "" Then
        getLastCol = Cells(row_index, Cells.Columns.Count).End(xlToLeft).Column
    Else
        getLastCol = Sheets(sht_name).Cells(row_index, Cells.Columns.Count).End(xlToLeft).Column
    End If
End Function
Public Function getColId(pure_num As Long) As String
    If pure_num Mod 26 = 0 Then
        getColId = VBA.IIf(pure_num \ 26 = 1, "", VBA.Chr(pure_num \ 26 + 63)) & "Z"
    Else
        getColId = VBA.IIf(pure_num \ 26 = 0, "", Chr(pure_num \ 26 + 64)) & Chr(pure_num Mod 26 + 64)
    End If
End Function
Public Function hasRepeatValue(sht As String, col_index As Long, Optional row_start As Long = 2, Optional row_end As Long = 0) As Boolean
    
    If row_end = 0 Then row_end = getLastRow(sht, col_index)
    If row_end < row_start Then Exit Function
    
    Dim dictTemp As New Dictionary, strKeyValue
    
    For i = row_start To row_end
        strKeyValue = Sheets(sht).Cells(i, col_index)
        If dictTemp.Exists(strKeyValue) Then
            hasRepeatValue = True
            Exit Function
        Else
            dictTemp.Add strKeyValue, ""
        End If
    Next i
    
End Function
