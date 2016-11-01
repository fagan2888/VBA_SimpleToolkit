Attribute VB_Name = "evaluation"
Public Function fexEval(func As String)
    Application.Volatile
    fexEval = evaluate(func)
End Function
