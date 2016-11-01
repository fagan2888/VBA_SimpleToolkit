Attribute VB_Name = "CellValue"
Public Function ehGetRelativePositionValue(sht As Worksheet, value As Variant, _
    Optional r As Integer = 1, Optional c As Integer = 2) As Variant
    
    ehGetRelativePositionValue = sht.Cells.Find(value)(r, c).value
    
End Function


