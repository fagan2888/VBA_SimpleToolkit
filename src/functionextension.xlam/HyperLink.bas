Attribute VB_Name = "HyperLink"
Public Sub fexSetEmptyHyperlink(ByVal targetRng As Range, TextToDisplay As String)

    With targetRng.Worksheet
        .Hyperlinks.Add Anchor:=targetRng, _
        Address:="", _
        ScreenTip:="Ŭ���ϼ���", _
        TextToDisplay:=TextToDisplay
    End With
End Sub

