Attribute VB_Name = "CreateObject"
Public Function dcaCreateDataField(ByVal name As String, value As Variant) As DataField
    Set dcaCreateDataField = New DataField
    dcaCreateDataField.initialize name, value
End Function

Public Function dcaCreateDataRecord(fieldNames() As String, values As Variant, Optional keyIndex As Integer = 0) As DataRecord
    Dim retVal As New DataRecord
    retVal.initialize keyIndex
    For i = LBound(fieldNames) To UBound(fieldNames)
        retVal.add fieldNames(i), values(i)
    Next
    Set dcaCreateDataRecord = retVal
End Function

Public Function dcaCreateEmptyDataRecord(Optional keyIndex As Integer = 0) As DataRecord
    Set dcaCreateEmptyDataRecord = New DataRecord
    dcaCreateEmptyDataRecord.initialize keyIndex
End Function

Public Function dcaCreateEmptyDataRecordSet() As DataRecordSet
    Set dcaCreateEmptyDataRecordSet = New DataRecordSet
End Function


Public Function dcaCreateDataRecordSetFromRangeWithHeader(rng As Range, Optional keyIndex As Integer = 0) As DataRecordSet
    Dim drs As DataRecordSet
    Set drs = dcaCreateEmptyDataRecordSet()
    For i = 2 To rng.Rows.Count
        Dim dr As DataRecord
        Set dr = dcaCreateEmptyDataRecord()
        dr.initialize keyIndex
        For j = 1 To rng.Columns.Count
            If Not rng(1, j).value = "" Then dr.add rng(1, j).value, rng(i, j).value
        Next
        drs.add dr
    Next
    Set dcaCreateDataRecordSetFromRangeWithHeader = drs
End Function

Public Function dcaCreateDataRecordSetFromRangeWithOutHeader(rngHeader As Range, rngData As Range, Optional keyIndex As Integer = 0) As DataRecordSet
    If rngHeader.Columns.Count <> rngData.Columns.Count Then Err.Raise 9
    
    Dim drs As DataRecordSet
    Set drs = dcaCreateEmptyDataRecordSet()
    For i = 1 To rngData.Rows.Count
        Dim dr As DataRecord
        Set dr = dcaCreateEmptyDataRecord()
        dr.initialize keyIndex
        For j = 1 To rngHeader.Columns.Count
            If Not rngHeader(1, j).value = "" Then dr.add rngHeader(1, j).value, rngData(i, j).value
        Next
        drs.add dr
    Next
    
    Set dcaCreateDataRecordSetFromRangeWithOutHeader = drs
End Function

Public Function dcaCreateDataRecordSetFromADODBRecordSet(rs As Recordset, Optional keyIndex As Integer = 0) As DataRecordSet

    Dim drs As DataRecordSet
    Set drs = dcaCreateEmptyDataRecordSet()
    Do While Not rs.EOF
        Dim r As DataRecord
        Set r = dcaCreateEmptyDataRecord
        r.initialize keyIndex
        For i = 0 To rs.fields.Count - 1
            r.add rs.fields(i).name, rs.fields(i).value
        Next
        drs.add r
        rs.MoveNext
    Loop
    Set dcaCreateDataRecordSetFromADODBRecordSet = drs
End Function
