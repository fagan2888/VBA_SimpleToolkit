Attribute VB_Name = "TypeCheck"
Dim defaultType_ As New Collection
Dim isInitialized As Boolean

Public Function isComparable(value As Variant)
    If Not isDefaultType(value) Then
         isComparable = TypeOf value Is IComparable
    End If
End Function

Public Function isEquatable(value As Variant)
    If Not isDefaultType(value) Then
         isEquatable = TypeOf value Is IEquatable
    End If
End Function

Private Function isDefaultType(value As Variant)
    If Not isInitialized Then setDefaultType
    isDefaultType = Contains(defaultType_, TypeName(value))
End Function

Private Sub setDefaultType()
    defaultType_.add "Integer"
    defaultType_.add "Long"
    defaultType_.add "Single"
    defaultType_.add "Double"
    defaultType_.add "Currency"
    isInitialized = True
End Sub
