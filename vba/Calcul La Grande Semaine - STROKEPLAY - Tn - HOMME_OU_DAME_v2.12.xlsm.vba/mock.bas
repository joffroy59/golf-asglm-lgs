Function getSerieMock(index As Double) As Variant
    result = ""
    If index < 15.5 Then
        result = "Serie1"
    Else
        result = "Serie2"
    End If
    getSerieMock = result
End Function