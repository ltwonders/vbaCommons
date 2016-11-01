Attribute VB_Name = "test"
Private Sub test_mergesingledict()
    Dim a As New Dictionary, b As New Dictionary, c As Dictionary
    a.Add "a", 1
    b.Add "b", 1
    Set c = mergeSingleDict(a, b)
End Sub
Private Sub test_sortedArray()
    b = Array("A0908099321", "A0908099330", "A0908099325", "A0908099328", "A0908099322", "A0908099323", "A0908099326")
    sortArray b
End Sub
Private Sub test_mail()

End Sub
Private Sub test_range()

End Sub
Private Sub test_text()

End Sub
Private Sub test_singletextdict()
    Dim a As New SingleTextDict
    a.addContent "使用量汇总", 1, 2
    Set b = a.getContent
End Sub

