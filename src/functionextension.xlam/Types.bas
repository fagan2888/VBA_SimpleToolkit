Attribute VB_Name = "Types"
Function getVariableType(myVar) As String

' ---------------------------------------------------------------
' Written By Shanmuga Sundara Raman for http://vbadud.blogspot.com
' Modified by ijlee
' ---------------------------------------------------------------

    If VarType(myVar) = vbNull Then
        getVariableType = "Null"
    ElseIf VarType(myVar) = vbInteger Then
        getVariableType = "Integer"
    ElseIf VarType(myVar) = vbLong Then
        getVariableType = "Long"
    ElseIf VarType(myVar) = vbSingle Then
        getVariableType = "Single"
    ElseIf VarType(myVar) = vbDouble Then
        getVariableType = "Double"
    ElseIf VarType(myVar) = vbCurrency Then
        getVariableType = "Currency"
    ElseIf VarType(myVar) = vbDate Then
        getVariableType = "Date"
    ElseIf VarType(myVar) = vbString Then
        getVariableType = "String"
    ElseIf VarType(myVar) = vbObject Then
        getVariableType = "Object"
    ElseIf VarType(myVar) = vbError Then
        getVariableType = "Error"
    ElseIf VarType(myVar) = vbBoolean Then
        getVariableType = "Boolean"
    ElseIf VarType(myVar) = vbVariant Then
        getVariableType = "Variant (used only with arrays of variants) "
    ElseIf VarType(myVar) = vbDataObject Then
        getVariableType = "DataObject"
    ElseIf VarType(myVar) = vbDecimal Then
        getVariableType = "Decimal"
    ElseIf VarType(myVar) = vbByte Then
        getVariableType = "Byte"
    ElseIf VarType(myVar) = vbUserDefinedType Then
        getVariableType = "UserDefinedType"
    ElseIf VarType(myVar) = vbArray Then
        getVariableType = "Array"
    Else
        getVariableType = VarType(myVar)
    End If

' Excel VBA, Visual Basic, Get Variable Type, VarType
End Function
