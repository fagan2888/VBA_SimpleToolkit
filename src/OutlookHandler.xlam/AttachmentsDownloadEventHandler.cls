VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AttachmentsDownloadEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IMailItemFunction
Private pattern_ As String
Private targetDir_ As String
Private regEx As New RegExp
Public Sub initialize(pattern As String, targetDir As String)
    
    pattern_ = pattern
    targetDir_ = targetDir
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = pattern_
    End With
End Sub

Public Sub IMailItemFunction_update(msg As MailItem)
    
    If regEx.test(msg.subject) Then
        Dim oh As MAPIHandler
        Set oh = New MAPIHandler
        fs = oh.downloadAttachmentsFromMailItem(msg, targetDir_)
        s = "다음의 파일을 " & targetDir_ & "에 저장하였습니다.\r\n"
        For i = LBound(fs) To UBound(fs)
            s = s & i & " : " & extractFileName(s(i)) & "\r\n"
        Next
        
End Sub

Private Function extractFileName(path As String)

    extractFileName = Right(path, Len(path) - Len(targetDir_))

End Function
