Attribute VB_Name = "HyperLink"
Public Sub fexSetEmptyHyperlink(ByVal targetRng As Range, TextToDisplay As String)

    With targetRng.Worksheet
        .Hyperlinks.Add Anchor:=targetRng, _
        Address:="", _
        ScreenTip:="클릭하세요", _
        TextToDisplay:=TextToDisplay
    End With
End Sub

