Attribute VB_Name = "ForeachRecord"
'@ijlee
'rs는 recordset type 이나 rs와 동일한 인터페이스를 가지고 있는 dataset을 사용합니다.
'다음의 멤버는 반드시 포함해야 합니다.
' - Property EOF As Boolean
' - Sub MoveNext

Public Sub dcaForeachRecordByOp(rs As Variant, op As IRecordOperation, _
         ParamArray args() As Variant)
    
    Dim i As Integer
    i = 1
    
    Do While Not rs.EOF
        op.operation rs.fields, i, args
        i = i + 1
        rs.MoveNext
    Loop

End Sub

Public Sub dcaForeachRecordByRun(rs As Variant, op As String, _
            ParamArray args() As Variant)
            
    Dim i As Integer
    i = 1
    Do While Not rs.EOF
        Application.run op, rs.fields, i, args
        i = i + 1
        rs.MoveNext
    Loop

End Sub



