Attribute VB_Name = "ArrayUtil"
Public Function writeTwoDimenArray(arr, sht_to As String, col_to As Long, Optional row_to As Long = 0, Optional asTranspose As Boolean = True) As Boolean
    'Write a one-dimen to the sht_to's col_to, first dimen in row, seconde dimen in col
    'Two-dimen array is default vertical first,then horizontal
    'Default row_to will be the lastrow of col_to

    If Not IsArray(arr) Then Exit Function
    On Error Resume Next
    
    If row_to = 0 Then row_to = getLastRow(sht_to, col_to) + 1
    
    Dim iFirstDimen As Long, iSecondDimen As Long
    iSecondDimen = UBound(arr, 2) + 1
    iFirstDimen = UBound(arr, 1) + 1
    If asTranspose Then
        Sheets(sht_to).Cells(row_to, col_to).Resize(iSecondDimen, iFirstDimen) = Application.Transpose(arr)
    Else
        Sheets(sht_to).Cells(row_to, col_to).Resize(iFirstDimen, iSecondDimen) = arr
    End If
    If Err.Number > 0 Then Err.Clear Else writeArray = True
End Function
Public Function writeOneDimenArray(arr, sht_to As String, col_to As Long, Optional row_to As Long = 0, Optional asHorizontal As Boolean = True) As Boolean
    'Write a one-dimen a sheet
    'One-dimen array is default horizontal
    'Default row_to is the lastrow of sht_to's col_to
    
    If Not IsArray(arr) Then Exit Function
    On Error Resume Next
    
    If row_to = 0 Then row_to = getLastRow(sht_to) + 1
    If asHorizontal Then
        Sheets(sht_to).Cells(row_to, col_to).Resize(1, UBound(arr, 1) + 1) = arr
    Else
        Sheets(sht_to).Cells(row_to, col_to).Resize(UBound(arr, 1) + 1, 1) = Application.Transpose(arr)
    End If
    If Err.Number > 0 Then Err.Clear Else writeArray = True
End Function
Public Function getSum(arr, Optional only_positive As Boolean = False) As Long
    If Not IsArray(arr) Then Exit Function
    Dim result As Long
    On Error Resume Next
    
    For i = LBound(arr) To UBound(arr)
        If only_positive And arr(i) < 0 Then GoTo continue
        
        result = result + arr(i)

continue:
    Next i
    
    If Err.Number > 0 Then Err.Clear Else getSum = result
End Function
Public Function getMax(arr) As Long
    If Not IsArray(arr) Then Exit Function
    Dim result As Long
    On Error Resume Next
    
    result = arr(0)
    For i = LBound(arr) To UBound(arr)
        If arr(i) > result Then result = arr(i)
    Next i
    
    If Err.Number > 0 Then Err.Clear Else getMax = result
End Function
Public Function getMin(arr) As Long
    If Not IsArray(arr) Then Exit Function
    Dim result As Long
    On Error Resume Next
    
    result = arr(0)
    For i = LBound(arr) To UBound(arr)
        If arr(i) < result Then result = arr(i)
    Next i
    
    If Err.Number > 0 Then Err.Clear Else getMin = result
End Function
Public Sub sortArray(arr, Optional Ascending As Boolean = True)
    If Not IsArray(arr) Then Err.Raise 23456, "Can't sort a non-array"
    
    Dim result()
    bound = UBound(arr)
    ReDim result(bound)
    
    Dim valueCurrent, valueTemp
    For i = LBound(arr) To UBound(arr)
        
        For j = i To UBound(arr)
            valueCurrent = arr(i)
            If arr(j) <= valueCurrent Then
                valueCurrent = arr(j)
                arr(j) = arr(i)
                arr(i) = valueCurrent
            End If
        Next j
        
        If Ascending Then result(i) = valueCurrent Else result(bound - i) = valueCurrent
    Next i
    
End Sub
