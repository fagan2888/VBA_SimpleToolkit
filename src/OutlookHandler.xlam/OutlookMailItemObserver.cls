VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutlookMailItemObserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private olApp_ As Outlook.application
Private ns_ As Outlook.Namespace
Public WithEvents items As Outlook.items
Attribute items.VB_VarHelpID = -1
Private func_ As IMailItemFunction

Private Sub Class_Initialize()
    
End Sub

'ol folder must be subfolder of inbox
Public Sub start(olFolder As String, func As IMailItemFunction)
    Set olApp_ = Outlook.application
    Set ns_ = olApp_.GetNamespace("MAPI")
    Set func_ = func
    Set items = ns_.GetDefaultFolder(olFolderInbox).folders(olFolder).items
End Sub

Public Sub cease()
    Set items = Nothing
End Sub

Private Sub Items_ItemAdd(ByVal item As Object)

    On Error GoTo ErrorHandler
    Dim msg As Outlook.MailItem
    If TypeName(item) = "MailItem" Then
        Set msg = item
        func_.update msg
    End If
ProgramExit:
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " - " & Err.Description
    Resume ProgramExit
End Sub


