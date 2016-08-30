VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private provider_ As String
Private connectionString_ As String

Public Sub initialize(provider As String, connectionString As String)
    provider_ = provider
    connectionString_ = connectionString
End Sub

Public Property Get provider() As String
    provider = provider_
End Property

Public Property Get connectionString() As String
    connectionString = connectionString_
End Property

