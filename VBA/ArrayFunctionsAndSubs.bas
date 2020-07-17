Attribute VB_Name = "ArrayFunctionsAndSubs"
'Identifies the next available/free slot in an Array
    Function nextAvaliableArrayLocation(ary As Variant) As Integer
        Dim i, j As Integer
        For i = 1 To UBound(ary)
        free = False
            If ary(i) = "" Then Exit For
        Next i
        nextAvaliableArrayLocation = i
    End Function
