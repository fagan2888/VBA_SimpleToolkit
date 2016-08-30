Attribute VB_Name = "AdvancedFilter"
Public Sub fexAdvancedFilterInPlace(rngTarget As Range, rngCondition As Range)
    rngTarget.AdvancedFilter Action:=xlFilterInPlace, Criteriarange:=rngCondition, Unique:=False
End Sub


