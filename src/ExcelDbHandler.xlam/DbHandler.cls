VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim con_ As ADODB.Connection
Dim rs_ As ADODB.Recordset
Dim dbConfig_ As dbConfig

Public Sub initialize(config As dbConfig)
    Set dbConfig_ = config
    Set con_ = New ADODB.Connection
    Set rs_ = New ADODB.Recordset
    
    With con_
        .provider = dbConfig_.provider
        .connectionString = dbConfig_.connectionString
        .CursorLocation = adUseClient
        .Open
    End With
End Sub

Private Sub Class_Terminate()
    If rs_.State = adStateOpen Then rs_.Close
    If con_.State = adStateOpen Then con_.Close
End Sub

' recordset handler
Public Function getRecordSet(ByVal sql As String) As ADODB.Recordset
    
    If rs_.State = adStateOpen Then rs_.Close

    rs_.Open sql, con_
    Set getRecordSet = rs_

End Function


Public Sub insertData(ByVal sql As String)

    con_.Execute sql

End Sub
