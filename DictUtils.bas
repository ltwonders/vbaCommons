Attribute VB_Name = "DictUtils"
Public Function genSingleDict(sht As String, col_key As Long, col_item As Long, Optional row_start As Long = 2, _
                            Optional row_end As Long = 0, Optional join_item As Boolean = False) As Dictionary
                            
    'Generate a one-to-one dict,default to keep the item newest
    'Structure: {key:item} if not join_item else {key:item|item}
    
    If row_end = 0 Then row_end = getLastRow(sht, col_key)
    
    Dim dictTarget As New Dictionary
    Dim strKey As String, strItem As String, oldValueItem
    
    For i = row_start To row_end
    With Sheets(sht)
        strKey = Trim(CStr(.Cells(i, col_key)))
        strItem = Trim(CStr(.Cells(i, col_item)))
    End With
        If strKey <> "" Then
            oldValueItem = dictTarget.Item(strKey)
            If join_item And oldValueItem <> "" Then
                dictTarget.Item(strKey) = oldValueItem & "|" & strItem
            Else
                dictTarget.Item(strKey) = strItem
            End If
        End If
        
    Next i
    
    Set genSingleDict = dictTarget
    Set dictTarget = Nothing
End Function
Public Function genDualDict(sht As String, col_key As Long, col_item1 As Long, col_item2 As Long, _
                            Optional row_start As Long = 2, Optional row_end As Long = 0) As Dictionary
                            
    'Generate a dict with two items,Keep the item newest
    'Structure: {key:[item1,item2]}
    
    If row_end = 0 Then row_end = getLastRow(sht, col_key)
    
    Dim dictTarget As New Dictionary
    
    Dim strKey As String, arrItem
    Dim itemValue1 As String, itemValue2 As String
    
    For i = row_start To row_end
    With Sheets(sht)
        strKey = Trim(CStr(.Cells(i, col_key)))
        itemValue1 = Trim(CStr(.Cells(i, col_item1).Value))
        itemValue2 = Trim(CStr(.Cells(i, col_item2)))
    End With
        arrItem = Array(itemValue1, itemValue2)

        If strKey <> "" Then
            If dictTarget.Exists(strKey) Then
                dictTarget.Item(strKey) = arrItem
            Else
                dictTarget.Add strKey, arrItem
            End If
        End If
        
    Next i
    
    Set genDualDict = dictTarget
    Set dictTarget = Nothing
End Function
Public Function genNestedDict(sht As String, col_key1 As Long, col_key2 As Long, col_item As Long, _
                            Optional row_start As Long = 2, Optional row_end As Long = 0, Optional join_item As Boolean = False) As Dictionary
    'Default repeatable allowed but keep newest the item value, if join_item then merge item with "|"
    'All value kepted as string
    'Structure: {key1:{key2:item}} if not join_item else {key1:{key2:item|item}}
    
    If row_end = 0 Then row_end = getLastRow(sht, col_key1)
    Dim valueKey1 As String, valueKey2 As String, valueItem, oldValueItem
    Dim dictTarget As New Dictionary
    
    For i = row_start To row_end
    With Sheets(sht)
        valueKey1 = Trim(CStr(.Cells(i, col_key1)))
        valueKey2 = Trim(CStr(.Cells(i, col_key2)))
        valueItem = Trim(CStr(.Cells(i, col_item)))
    End With
        If valueKey1 <> "" And Not dictTarget.Exists(valueKey1) Then Set dictTarget.Item(valueKey1) = New Dictionary
        If valueKey1 <> "" And valueKey2 <> "" Then
            oldValueItem = dictTarget.Item(valueKey1).Item(valueKey2)
            If join_item And oldValueItem <> "" Then
                dictTarget.Item(valueKey1).Item(valueKey2) = oldValueItem & "|" & valueItem
            Else
                dictTarget.Item(valueKey1).Item(valueKey2) = valueItem
            End If
        End If
        
    Next i
    Set genNestedDict = dictTarget
    Set dictTarget = Nothing
    
End Function
Public Function mergeSingleDict(merge_to As Dictionary, merge_from As Dictionary) As Dictionary
    'Only support a {str:long} dict to add then item
    
    Dim result As New Dictionary, arrTemp
    If merge_to Is Nothing And merge_from Is Nothing Then Exit Function
    If Not merge_to Is Nothing Then Set result = mergeSingleDict(Nothing, merge_to)
    If merge_from Is Nothing Then
        Set mergeSingleDict = result
        Exit Function
    End If
    arrTemp = merge_from.Keys
    For i = LBound(arrTemp) To UBound(arrTemp)
        If result.Exists(arrTemp(i)) Then
            result.Item(arrTemp(i)) = CLng(result.Item(arrTemp(i))) + merge_from.Item(arrTemp(i))
        Else
            result.Item(arrTemp(i)) = merge_from.Item(arrTemp(i))
        End If
    Next i
    Set mergeSingleDict = result
    Set result = Nothing
    
End Function
Public Function genSingleSumDict(sht As String, col_key As Long, col_item As Long, Optional row_start As Long = 2, Optional row_end As Long = 0, Optional only_positive As Boolean = True) As Dictionary
    'Generate a single summary dict,only supports item's type should be long
    
    If row_end = 0 Then row_end = getLastRow(sht, col_key)
    
    Dim dictTarget As New Dictionary
    Dim strKey As String, lngItem As Long, oldValueItem
    
    For i = row_start To row_end
        strKey = Trim(CStr(Sheets(sht).Cells(i, col_key)))
        lngItem = CLng(Sheets(sht).Cells(i, col_item))
        
        If strKey = "" Or (only_positive And lngItem < 0) Then GoTo continue
        If Not dictTarget.Exists(strKey) Then
            dictTarget.Add strKey, lngItem
        Else
            dictTarget.Item(strKey) = dictTarget.Item(strKey) + lngItem
        End If
continue:
    Next i
    
    Set genSingleSumDict = dictTarget
    Set dictTarget = Nothing
    
End Function
Public Function genNestedSumDict(sht As String, col_key1 As Long, col_key2 As Long, col_item As Long, Optional row_start As Long = 2, _
                                    Optional row_end As Long = 0, Optional only_positive As Boolean = True) As Dictionary
    
    If row_end = 0 Then row_end = getLastRow(sht, col_key1)
    Dim valueKey1 As String, valueKey2 As String, valueItem
    Dim dictTarget As New Dictionary
    
    For i = row_start To row_end
    With Sheets(sht)
        valueKey1 = Trim(CStr(.Cells(i, col_key1)))
        valueKey2 = Trim(CStr(.Cells(i, col_key2)))
        valueItem = CLng(.Cells(i, col_item))
    End With
        If valueKey1 <> "" And Not dictTarget.Exists(valueKey1) Then Set dictTarget.Item(valueKey1) = New Dictionary
        If valueKey2 = "" Or (only_positive And valueItem < 0) Then GoTo continue
        
        If Not dictTarget.Item(valueKey1).Exists(valueKey2) Then
            dictTarget.Item(valueKey1).Add valueKey2, valueItem
        Else
            dictTarget.Item(valueKey1).Item(valueKey2) = valueItem + dictTarget.Item(valueKey1).Item(valueKey2)
        End If
continue:
    Next i
    Set genNestedSumDict = dictTarget
    Set dictTarget = Nothing
    
End Function
Public Function expandDict(ByRef dict As Dictionary, multiple As Long) As Dictionary
    If dict Is Nothing Then Err.Raise 34567, "Can't expand a empty dictionary"

    Dim result As New Dictionary
    
    arrKeys = dict.Keys()
    For i = LBound(arrKeys) To UBound(arrKeys)
        strKey = arrKeys(i)
'        Debug.Print "Before exapanding: " & dict.Item(strKey)
        result.Item(strKey) = dict.Item(strKey) * multiple
'        Debug.Print "After exapanding: " & result.Item(strKey)
    Next i
    
    Set expandDict = result
    
End Function

