Attribute VB_Name = "CreateObject"
Public Function ehCreateEmptyWorkBookHandler() As WorkBookHandler
    Set ehCreateEmptyWorkBookHandler = New WorkBookHandler
End Function

'열려있는 워크북은 이름(+확장자)만 새로 열거는 full경로
Public Function ehCreateWorkBookHandlerWithOpenWorkbook(pathOrName As String, _
        Optional inBackground As Boolean = True) As WorkBookHandler

    Set ehCreateWorkBookHandlerWithOpenWorkbook = New WorkBookHandler
    ehCreateWorkBookHandlerWithOpenWorkbook.openWorkbook pathOrName, inBackground
    
End Function

Public Function ehCreateWorkBookHandlerWithNewWorkbook(pathOrName As String, _
        Optional inBackground As Boolean = True) As WorkBookHandler

    Set ehCreateWorkBookHandlerWithNewWorkbook = New WorkBookHandler
    ehCreateWorkBookHandlerWithNewWorkbook.newWorkBook pathOrName, inBackground
    
End Function


