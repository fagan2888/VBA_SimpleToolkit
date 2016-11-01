Attribute VB_Name = "Ranges"
'값 찾기 함수들
Public Function fexFindOrElse(targets As Range, value As Variant, _
                Optional elseThen As Variant = "")

    For Each r In targets
        If r.value = value Then
            fexFindOrElse = value
            Exit Function
        End If
    Next
        
    findOrElse = elseThen
End Function

Public Function fexContains(rng As Range, ByVal obj As Variant) As Boolean

    Dim r As Range
    For Each r In rng
        If r.value = obj Then
            fexContains = True
            Exit Function
        End If
    Next
    
    contains = False
End Function

'셀 선택 관련 함수들.
Public Function fexDownAndNRightRange(rngInit As Range, n As Integer) As Range
         
    If rngInit(2, 1).value = "" Then
        Set fexDownAndNRightRange = rngInit
    Else
        Set fexDownAndNRightRange = Range(rngInit, rngInit.End(xlDown)(1, n))
    End If

End Function


Public Function fexDownRightRange(rngInit As Range) As Range
         
    If rngInit(2, 1).value = "" And rngInit(1, 2).value = "" Then
        Set fexDownRightRange = rngInit
    ElseIf rngInit(2, 1).value = "" Then
        Set fexDownRightRange = rngInit.Worksheet.Range(rngInit, rngInit.End(xlToRight))
    ElseIf rngInit(1, 2).value = "" Then
        Set fexDownRightRange = rngInit.Worksheet.Range(rngInit, rngInit.End(xlDown))
    Else
        Set fexDownRightRange = rngInit.Worksheet.Range(rngInit, rngInit.End(xlDown).End(xlToRight))
    End If

End Function

Public Function fexDownRange(rngInit As Range) As Range
         
    If rngInit(2, 1).value = "" Then
        Set fexDownRange = rngInit
    Else
        Set fexDownRange = rngInit.Worksheet.Range(rngInit, rngInit.End(xlDown))
    End If

End Function

Public Function fexRightRange(rngInit As Range) As Range
         
    If rngInit(1, 2).value = "" Then
        Set fexRightRange = rngInit
    Else
        Set fexRightRange = rngInit.Worksheet.Range(rngInit, rngInit.End(xlToRight))
    End If

End Function

' Cell address 관련 functions
Public Function fexLastValueIndex(rng As Range) As Integer
    Application.Volatile True
    Dim c As Range
        
    lastValueIndex = 0
    For Each c In rng
         If c.Text <> "" Then
            fexLastValueIndex = fexLastValueIndex + 1
        End If
    Next

End Function

Public Function fexLastRecordCellReference(rngInit As Range) As Range
    
    k = 0
    Do
        k = k + 1
    Loop Until rngInit(k + 1, 1).value = ""
    
    Set fexLastRecordCellReference = rngInit(k, 1)
End Function



Public Function fexFindNotDuplicatedValues(rng1 As Range, rng2 As Range, _
    Optional ByVal label1 As String = "Table2 누락", Optional ByVal label2 As String = "Table1 누락", Optional ByVal rngSize = 100) As Variant()
        
    Dim retVal() As Variant
    ReDim retVal(1 To rngSize, 1 To 2) As Variant
    For j = 1 To rngSize
        retVal(j, 1) = ""
        retVal(j, 2) = ""
    Next
    
    i = 1
    For Each r1 In rng1
        If Application.WorksheetFunction.CountIf(rng2, r1.value) = 0 Then
            retVal(i, 1) = r1.value
            retVal(i, 2) = label1
            i = i + 1
        End If
    Next

    For Each r2 In rng2
        If Application.WorksheetFunction.CountIf(rng1, r2.value) = 0 Then
            retVal(i, 1) = r2.value
            retVal(i, 2) = label2
            i = i + 1
        End If
    Next

    fexFindNotDuplicatedValues = retVal
End Function

Public Function fexRangeToStringArray(rngTarget As Range) As String()

    Dim retVal() As String
    ReDim retVal(rngTarget.Cells.Count - 1) As String
    
    i = 0
    For Each r In rngTarget
        retVal(i) = r.value
        i = i + 1
    Next
    
    fexRangeToStringArray = retVal
End Function

