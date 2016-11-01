VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutlookObserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private olApp_ As Outlook.application
Private ns_ As Outlook.Namespace

'TODO: collection 형태로 바꿀것
Private e_ As IEventHandler

Public Sub initialize()
    Set olApp_ = Outlook.application
    Set ns_ = olApp_.GetNamespace("MAPI")
End Sub

Public Property Get olNameSpace() As Outlook.Namespace
    Set olNameSpace = ns_
End Property

Public Property Get olApp() As Outlook.application
    Set olApp = olApp_
End Property

'ol folder must be subfolder of inbox
Public Sub start(e As IEventHandler)
    Set e_ = e
    e_.initialize Me
End Sub

Public Sub cease()
    Set e_ = Nothing
End Sub


Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    Set olApp_ = Nothing
    Set ns = Nothing
    Set e_ = Nothing
End Sub
