Attribute VB_Name = "Utils"
Public Function fexAddInDirectory() As String

    fexAddInDirectory = ThisWorkbook.Path

End Function

Public Function fexCurrentDirectory() As String

    fexCurrentDirectory = Application.ActiveWorkbook.Path

End Function

Public Sub CopySelectionNotepad(ByVal message As String)
    
    ClipBoard_SetData (message)
    With Application
        Call Shell("Notepad.exe", vbNormalFocus)
        SendKeys "^v"
    End With
    
End Sub


