VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataRecordSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'cursor의 key는 DataRecord의 첫번째 field로함.
Dim cursor_ As Cursor
Dim rs_ As Collection

Public Sub Class_Initialize()
    Set rs_ = New Collection
    cursor_.idx = 1
    cursor_.key = ""
End Sub

Public Sub Class_Terminate()
    Set rs_ = Nothing
End Sub

'cursor level properties
Public Property Get Cursor() As Cursor
    Cursor = cursor_
End Property

Public Property Get BOF() As Boolean
    BOF = False
    If cursor_.idx = 1 Then BOF = True
End Property

Public Property Get EOF() As Boolean
    EOF = False
    If rs_.Count + 1 = cursor_.idx Then EOF = True
End Property

Public Sub Move(idxOrKey As Variant)
    If TypeName(idxOrKey) = "Integer" Then
        cursor_.idx = idxOrKey
        cursor_.key = rs_(idxOrKey).keyValue
    ElseIf TypeName(idxOrKey) = "String" Then
        cursor_.key = idxOrKey
        cursor_.idx = indexOf(rs_, rs_(idxOrKey))
    End If
End Sub
Public Sub MoveNext()
    If Not EOF() Then
        cursor_.idx = cursor_.idx + 1
        cursor_.key = getCursorKey(cursor_, rs_)
    End If
End Sub
Public Sub MovePrevious()
    If Not BOF() Then
        cursor_.idx = cursor_.idx - 1
        cursor_.key = getCursorKey(cursor_, rs_)
    End If
End Sub
Public Sub MoveFirst()
    cursor_.idx = 1
    cursor_.key = getCursorKey(cursor_, rs_)
End Sub
Public Sub MoveLast()
    cursor_.idx = rs_.Count
    cursor_.key = getCursorKey(cursor_, rs_)
End Sub

Public Property Get names() As Variant
    If rs_.Count = 0 Then
        Err.Raise 0
    Else
        Dim r As DataRecord
        Set r = rs_(1)
        names = r.names
    End If
End Property

Public Property Get fields(idxOrKey As Variant) As DataField
    Dim key As String
    If TypeName(idxOrKey) = "Integer" Then
        key = rs_(cursor_.idx).name(CInt(idxOrKey))
    Else
        key = idxOrKey
    End If
    Set fields = dcaCreateDataField(key, rs_(cursor_.idx)(key))
End Property

Public Property Get Count() As Long
    Count = rs_.Count
End Property

'dataset level
Public Property Get Item(idxOrKey As Variant) As DataRecord
Attribute Item.VB_UserMemId = 0
    Set Item = rs_(idxOrKey)
End Property

Public Property Set Item(idxOrKey As Variant, value As Variant)
Attribute Item.VB_UserMemId = 0
    Set rs_(idxOrKey) = value
End Property
Public Property Let Item(idxOrKey As Variant, value As Variant)
Attribute Item.VB_UserMemId = 0
    rs_(idxOrKey) = value
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = rs_.[_NewEnum]
End Function

Public Sub add(record As DataRecord)
    rs_.add record, record.keyValue
End Sub
Public Sub remove(idxOrKey As Variant)
    rs_.remove idxOrKey
End Sub

Public Function Contains(key As String) As Boolean
    
    Contains = Contains(rs_, key)
    
End Function

Public Function filterEquals(header As String, value As Variant) As DataRecordSet
    
    Set filterEquals = New DataRecordSet
    For i = 1 To rs_.Count
        If rs_(i)(header) = value Then
            filterEquals.add Item(i)
        End If
    Next

End Function

Public Function getCursorKey(Cursor As Cursor, rs As Collection) As String
    If Cursor.idx = rs.Count + 1 Then
        getCursorKey = "EOF"
    Else
        getCursorKey = rs(cursor_.idx).keyValue
    End If
End Function

Public Sub spread(rngStart As Range, Optional withHeader As Boolean = False)
    
    k = 0
    Dim r As DataRecord
    If rs_.Count = 0 Then Exit Sub
    ns = names()
    If withHeader Then
        For j = LBound(ns) To UBound(ns)
            rngStart(1, j + 1).value = names(j)
        Next
        k = 1
    End If
    
    For i = 1 To rs_.Count
        For j = LBound(ns) To UBound(ns)
            rngStart(i + k, j + 1).value = rs_(i)(j)
        Next
    Next
    
End Sub


'merge
Public Function merge(rs As DataRecordSet) As DataRecordSet
    Dim rsNew As DataRecordSet
    Set rsNew = New DataRecordSet
    
    Dim r As DataRecord
    For Each r In rs_
        rsNew.add r
    Next
        
    For i = 1 To rs.Count
        rsNew.add rs(i)
    Next
    
    Set merge = rsNew
End Function

Public Sub mergeInPlace(rs As DataRecordSet)
    
    For i = 1 To rs.Count
        add rs(i)
    Next
    
End Sub


' basic data manipulation methods
Public Function map(f As IUnaryFunction) As DataRecordSet
    
    Set map = New DataRecordSet
    
    For i = 1 To rs_.Count
        map.add f.evaluate(Item(i))
    Next
    
End Function

Public Function mapFields(fields() As String) As DataRecordSet
    
    Set mapFields = New DataRecordSet
    
    For i = 1 To rs_.Count
        Dim tmpRecord As DataRecord
        Set tmpRecord = New DataRecord
        
        For j = LBound(fields) To UBound(fields)
            tmpRecord.add fields(j), Item(i)(fields(j))
        Next
        mapFields.add tmpRecord
    Next
    
End Function

Public Function reduce(f As IBinaryFunction) As DataRecord

    Dim tmpRecord As DataRecord
    Set tmpRecord = record(1)
    For i = 2 To rs_.Count
        Set tmpRecord = f.evaluate(tmpRecord, Item(i))
    Next
    
    reduce = tmpRecord
End Function

Public Function filter(f As IPredicator) As DataRecordSet
    Set filter = New DataRecordSet
    
    For i = 1 To rs_.Count
        If f.evaluate(Item(i)) Then filter.add Item(i)
    Next
End Function

Public Sub forEach(f As IUnaryFunction)
    For i = 1 To rs_.Count
        f.evaluate Item(i)
    Next
End Sub

Public Function findFirst(header As Variant, value As Variant, Optional r As DataRecord = Nothing) As Boolean
    
    For i = 1 To rs_.Count
        If rs_(i)(header) = value Then
            If Not r Is Nothing Then Set r = Item(i)
            findFirst = True
            Exit Function
        End If
    Next
    
    findFirst = False
End Function

Public Function distinct(header As String) As Collection
    Set distinct = CollectionExtension.distinct(asCollection(header))
End Function

Public Function asCollection(header As String) As Collection
    
    Set asCollection = New Collection
    For i = 1 To rs_.Count
        asCollection.add rs_(i)(header)
    Next
    
End Function
