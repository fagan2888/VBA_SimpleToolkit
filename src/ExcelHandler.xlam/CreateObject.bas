Attribute VB_Name = "CreateObject"
Public Function ehCreateEmptyWorkBookHandler() As WorkBookHandler
    Set ehCreateEmptyWorkBookHandler = New WorkBookHandler
End Function

'�����ִ� ��ũ���� �̸�(+Ȯ����)�� ���� ���Ŵ� full���
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


