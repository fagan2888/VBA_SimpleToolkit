Attribute VB_Name = "CreateObject"
Public Function ohCreateMAPIHandler() As MAPIHandler
    Set ohCreateMAPIHandler = New MAPIHandler
    ohCreateMAPIHandler.initialize
End Function

Public Function ohCreateOutlookObserver() As OutlookObserver
    Set ohCreateOutlookObserver = New OutlookObserver
End Function

Public Function ohCreateOutlookMailItemObserver() As OutlookMailItemObserver
    Set ohCreateOutlookMailItemObserver = New OutlookMailItemObserver
End Function

Public Function ohCreateAttachmentDownloadEventHandler() As AttachmentsDownloadEventHandler
    Set ohCreateAttachmentDownloadEventHandler = New AttachmentsDownloadEventHandler
End Function
