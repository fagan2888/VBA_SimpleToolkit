Attribute VB_Name = "CollectionExtension"
Public Function distinct(c As Collection) As Collection
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To c.Count
        d(c(i)) = 1
    Next i
    
    Dim v As Variant
    Dim retVal As Collection
    Set retVal = New Collection
    For Each v In d.Keys()
        retVal.add (v)
    Next v
    
    Set distinct = retVal

End Function

Public Function indexOf(c As Collection, value As Variant) As Long
    For i = 1 To c.Count
        If IsObject(value) Then
            If Not isEquatable(value) Then
                Err.Raise 13, Description = "indexOf는 equatable인 object만 허용됩니다"
            Else
                Dim v As IEquatable
                Set v = value
                If v.equals(c(i)) Then
                    indexOf = i
                    Exit Function
                End If
            End If
        Else
            If value = c.Item(i) Then
                indexOf = i
                Exit Function
            End If
        End If
    Next
    indexOf = -1
End Function

Public Function Contains(c As Collection, s As String) As Boolean
    For i = 1 To c.Count
        If c(i) = s Then
            Contains = True
            Exit Function
        End If
    Next
        
    Contains = False
End Function

Public Function collectionToArray(c As Collection) As Variant
    Dim retVal() As Variant
    ReDim retVal(c.Count) As Variant
    
    For i = 0 To c.Count - 1
        retVal(i) = c(i + 1)
    Next
    collectionToArray = retVal
End Function

