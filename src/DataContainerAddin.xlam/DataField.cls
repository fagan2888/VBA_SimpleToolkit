VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim v As Variant
Dim n As String

Public Sub initialize(name As String, value As Variant)
    n = name
    v = value
End Sub

Public Property Get getType() As Variant
    getType = TypeName(v)
End Property

Public Property Let value(val As Variant)
Attribute value.VB_UserMemId = 0
    v = val
End Property

Public Property Get value() As Variant
Attribute value.VB_UserMemId = 0
    value = v
End Property

Public Property Get name() As String
    name = n
End Property
