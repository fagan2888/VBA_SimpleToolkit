Attribute VB_Name = "ForeachRecord"
'@ijlee
'rs�� recordset type �̳� rs�� ������ �������̽��� ������ �ִ� dataset�� ����մϴ�.
'������ ����� �ݵ�� �����ؾ� �մϴ�.
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



