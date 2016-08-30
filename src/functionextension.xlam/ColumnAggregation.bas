Attribute VB_Name = "ColumnAggregation"
Public Function fexCSum(rngInit As Range) As Double
    
    fexCSum = WorksheetFunction.Sum(Range(rngInit, fexLastRecordCellReference(rngInit)))

End Function
