VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IEquatable

Private d_ As Dictionary
Private keyIndex_ As Integer

Public Sub initialize(keyIndex As Integer)
    keyIndex_ = keyIndex
End Sub

Public Property Get keyIndex() As Integer
    keyIndex = keyIndex_
End Property

Public Sub Class_Initialize()
    Set d_ = New Dictionary
    keyIndex_ = 0
End Sub

Public Sub Class_Terminate()
    Set d_ = Nothing
End Sub

Public Property Get Count() As Long
    Count = d_.Count
End Property

Public Property Get names() As Variant
    names = d_.Keys
End Property

Private Function idxToKey(ByVal idx As Integer) As String
    idxToKey = d_.Keys(idx)
End Function

Private Function getKeyValue(idxOrFieldName As Variant) As String
    If TypeName(idxOrFieldName) = "Integer" Or TypeName(idxOrFieldName) = "Long" Then
        getKeyValue = idxToKey(idxOrFieldName)
        Exit Function
    End If
    getKeyValue = idxOrFieldName
End Function

Public Property Get keyName() As String
    keyName = getKeyValue(keyIndex_)
End Property
Public Property Get keyValue() As String
    keyValue = d_(getKeyValue(keyIndex_))
End Property

Public Property Get name(idx As Integer) As String
    name = getKeyValue(idx)
End Property

Public Property Get Item(idxOrFieldName As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Item = d_(getKeyValue(idxOrFieldName))
End Property

Public Property Set Item(idxOrFieldName As Variant, value As Variant)
Attribute Item.VB_UserMemId = 0
    d_(getKeyValue(idxOrFieldName)) = value
End Property

Public Property Let Item(idxOrFieldName As Variant, value As Variant)
Attribute Item.VB_UserMemId = 0
    d_(getKeyValue(idxOrFieldName)) = value
End Property

Public Sub add(fieldName As String, value As Variant)
    d_.add fieldName, value
End Sub

Public Sub remove(idxOrFieldName As Variant)
    d_.remove getKeyValue(idxOrFieldName)
End Sub

Public Function IEquatable_equals(other As Variant) As Boolean
    If Not TypeOf other Is DataRecord Then
        Err.Raise 13
    Else
        Dim o As DataRecord
        Set o = other
        IEquatable_equals = keyValue = o.keyValue
    End If
End Function



