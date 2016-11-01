VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MAPIHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim olApp_ As Outlook.application

Private Sub Class_Initialize()
    Set olApp_ = Nothing
End Sub

Private Sub Class_Terminate()
End Sub

Public Sub initialize()
    Set olApp_ = GetObject("", "Outlook.Application")
End Sub

Public Function application() As Outlook.application
    Set application = olApp_
End Function

'Folder
Public Function getMailFolder(olFolderName As String) As Outlook.folder
    Dim ns As Outlook.Namespace
    Set ns = olApp_.GetNamespace("mapi")
    If Not isFolderExist(olFolderName) Then Err.Raise 5, Description:="invalid folder name"
    Set getMailFolder = ns.GetDefaultFolder(olFolderInbox).folders(olFolderName)
End Function

Public Function isFolderExist(olFolderName As String) As Boolean
    Dim ns As Outlook.Namespace
    Set ns = olApp_.GetNamespace("mapi")
    isFolderExist = checkFolder(ns.GetDefaultFolder(olFolderInbox).folders, olFolderName)
End Function

Private Function checkFolder(folders As Outlook.folders, value As String) As Boolean
    checkFolder = False
    For i = 1 To folders.Count
        If folders(i).name = value Then
            checkFolder = True
            Exit Function
        End If
    Next
End Function

'mail item
Public Function hasMailReceived(olFolderName As String, pattern As String, _
    Optional ByRef mail As Outlook.MailItem = Nothing) As Boolean
       
    Dim regEx As New RegExp
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = pattern
    End With
    
    hasMailReceived = False
    Set mail = Nothing
    
    For Each m In getMailFolder(olFolderName).items
        If regEx.test(m.subject) Then
            hasMailReceived = True
            Set mail = m
            Exit Function
        End If
    Next
        
End Function

Public Function getMailBody(olFolderName As String, pattern As String) As String

    Dim mail As Outlook.MailItem
    Set mail = checkMailRecieved(olFolderName, pattern)
    
    getMailBody = mail.Body
    
End Function
'attachment

Public Function downloadAttachments(targetDir As String, olFolderName As String, _
            mailNamePattern As String) As String()

    Dim mail As Outlook.MailItem
    Set mail = checkMailRecieved(olFolderName, mailNamePattern)

    downloadAttachments = downloadAttachmentsFromMailItem(mail, targetDir)
End Function

Public Function downloadAttachmentsFromMailItem(mail As Outlook.MailItem, targetDir As String) As String()
    ReDim retVal(1 To mail.Attachments.Count) As String
    
    i = 1
    For Each a In mail.Attachments
        a.SaveAsFile targetDir & "\" & a.Filename
        retVal(i) = targetDir & "\" & a.Filename
        i = i + 1
    Next
    
    downloadAttachmentsFromMailItem = retVal

End Function

Public Function downloadAttachmentsWithPattern(targetDir As String, olFolderName As String, mailNamePattern As String, _
        AttachmentPattern() As String) As String()

    Dim mail As Outlook.MailItem
    Set mail = checkMailRecieved(olFolderName, mailNamePattern)

    Dim regEx As New RegExp
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
    End With
    
    Dim c As Collection
    Set c = New Collection
    For j = 1 To UBound(AttachmentPattern)
        regEx.pattern = AttachmentPattern(j)
        For Each a In mail.Attachments
            If regEx.test(a.Filename) Then
                a.SaveAsFile (targetDir & "\" & a.Filename)
                c.Add targetDir & "\" & a.Filename
            End If
        Next
    Next
    
    ReDim retVal(1 To c.Count) As String
    For i = 1 To c.Count
        retVal(i) = c.item(i)
    Next
    
    downloadAttachmentsWithPattern = retVal
End Function

Private Function checkMailRecieved(olFolderName As String, mailNamePattern As String) As Outlook.MailItem
    If Not hasMailReceived(olFolderName, mailNamePattern, checkMailRecieved) Then
        Err.Raise 5, Description:="mail not founded"
    End If
End Function
        
'send mail
Public Function createMailItem() As Outlook.MailItem
    Set createMailItem = olApp_.CreateItem(olMailItem)
End Function

Public Function sendMail(rngMessage() As Range, recipients As String, cc As String, subject As String, Optional delimiter As String = ",")
    
    Dim m As Outlook.MailItem
    Set m = createMailItem()
    
    recipients = Replace(recipients, " ", "")
    splittedRecipients = Split(recipients, delimiter)
    
    For i = 0 To UBound(splittedRecipients)
        m.recipients.Add splittedRecipients(i)
    Next
    m.cc = cc
    m.subject = subject
    m.Display
    
    Set wrdedit = olApp_.ActiveInspector.WordEditor
    
    For i = 1 To UBound(rngMessage)
        rngMessage(i).Copy
        wrdedit.application.Selection.Paste
    Next
    
    rngMessage(1).application.CutCopyMode = False
End Function
