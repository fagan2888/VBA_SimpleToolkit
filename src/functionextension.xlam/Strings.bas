Attribute VB_Name = "Strings"
Public Function fexConcatStringsWithComma(Strings As Range, Optional ByVal withInitial = "")
    tmpVal = withInitial
    For Each s In Strings
        If s <> "" Then tmpVal = tmpVal & s.value & ", "
    Next
    fexConcatStringsWithComma = Left(tmpVal, Len(tmpVal) - 2)
End Function

Public Function fexConcatStringsWithSpace(Strings As Range, Optional ByVal withInitial = "")
    tmpVal = withInitial
    For Each s In Strings
        If s <> "" Then tmpVal = tmpVal & s.value & " "
    Next
    fexConcatStringsWithSpace = Left(tmpVal, Len(tmpVal) - 1)
End Function

Public Function fexStringArray(ParamArray s() As Variant) As String()

    Dim retVal() As String
    
    addOne = 0
    If LBound(s) = 0 Then addOne = 1
    
    ReDim retVal(1 To (UBound(s) + addOne))
    
    For i = addOne To (UBound(s) + addOne)
        retVal(i) = s(i - addOne)
    Next
    fexStringArray = retVal
End Function
