VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkBookHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim wkb As Workbook
Dim app As Excel.Application
Dim inBackground_ As Boolean
Dim path_ As String
Dim mode As HandlerMode

Enum HandlerMode
    NewWkb
    OpenWkb
End Enum

Private Sub Class_Initialize()
    initialize
End Sub

Private Sub Class_Terminate()
    initialize
End Sub

Public Sub initialize()

    If (Not wkb Is Nothing) And inBackground_ Then
        wkb.Close savechanges:=False
        Set wkb = Nothing
        If Not app Is Nothing Then app.Quit
        Set app = Nothing
    End If
        
End Sub

Private Function isInitialized() As Boolean
    isInitialized = (wkb Is Nothing) And (app Is Nothing)
End Function

Private Sub initializedCheck()
    If Not isInitialized() Then Err.Raise 5, Description = "object not initialized"
End Sub

Public Sub newWorkBook(path As String, Optional inBackground As Boolean = True)

    initializedCheck
    mode = NewWkb
    inBackground_ = inBackground
    Set app = New Excel.Application
    app.Visible = Not inBackground
    Set wkb = app.Workbooks.Add
    wkb.saveAs path
    
End Sub

Public Sub openWorkbook(pathOrName As String, Optional inBackground As Boolean = True)
    
    initializedCheck
    mode = OpenWkb
    inBackground_ = inBackground
    '열려있는 워크북은 이름(+확장자)만 새로 열거는 full경로
    If InStr(1, pathOrName, "\") = 0 Then
        inBackground_ = False
        Set wkb = Application.Workbooks(pathOrName)
    Else
        Set app = New Excel.Application
        app.Visible = Not inBackground
        Set wkb = app.Workbooks.Open(pathOrName)
    End If
End Sub

Public Function sheetNames() As String()
    Dim n() As String
    ReDim n(1 To wkb.Sheets.Count) As String
    
    i = 1
    For Each s In wkb.Sheets
        n(i) = s.name
        i = i + 1
    Next
        
    sheetNames = n
End Function

Public Function sheetCount() As Integer
    sheetCount = wkb.Sheets.Count
End Function

Public Function name() As String
    name = wkb.name
End Function

Public Property Get getWorkBook() As Workbook
    Set getWorkBook = wkb
End Property

Public Property Get inBackground() As Boolean
    inBackground = inBackground_
End Property

Public Sub save()
    wkb.save
End Sub

Public Sub saveAndInitialize()
    wkb.save
    initialize
End Sub

Public Sub saveAs(path As String)
    wkb.saveAs path
End Sub

Public Sub saveAsAndInitialize(path As String)
    wkb.saveAs path
    initialize
End Sub

Public Sub saveAsPdf(path As String, targetSheets() As String, Optional openAfterPublish As Boolean = True)
    
    Application.ScreenUpdating = False
    Set currentSelection = wkb.ActiveSheet
    
    wkb.Sheets(targetSheets).Select
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=path, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        openAfterPublish:=openAfterPublish
    
    currentSelection.Select
    
    Application.ScreenUpdating = True
End Sub
Private Sub copy(rngTarget As Range, rngPastePoint As Range, asValue As Boolean)
    rngTarget.copy
    If asValue Then
        rngPastePoint.PasteSpecial xlPasteValues
    Else
        rngPastePoint.PasteSpecial xlPasteFormats
        rngPastePoint.PasteSpecial xlPasteValues
    End If
End Sub
 
Public Sub copyFromByName(sheetNameOrIdx As Variant, rangeName As String, rngPastePoint As Range, _
                        Optional asValue As Boolean = False)
    
    Dim rngTarget As Range
    Set rngTarget = wkb.Sheets(sheetNameOrIdx).Range(rangeName)
    
    copy rngTarget, rngPastePoint, asValue
End Sub

Public Sub copyFromByCellPoint(sheetNameOrIdx As Variant, rowStart As Integer, colStart As Integer, _
        rowEnd As Integer, colEnd As Integer, rngPastePoint As Range, Optional asValue As Boolean = False)
    Dim rngTarget As Range
    Set rngTarget = wkb.Sheets(sheetNameOrIdx).Range(Cells(rowStart, colStart), Cells(rowEnd, colEnd))
    
    copy rngTarget, rngPastePoint, asValue
End Sub

Public Sub copyFromByNameRightDownRange(sheetNameOrIdx As Variant, rangeStart As String, rngPastePoint As Range, _
                        Optional asValue As Boolean = False)
                        
    Dim rngTarget As Range
    Set rngTarget = fexDownRightRange(wkb.Sheets(sheetNameOrIdx).Range(rangeStart))
    
    copy rngTarget, rngPastePoint, asValue
End Sub

Public Sub copyFromByCellPointRightDownRange(sheetNameOrIdx As Variant, rowStart As Integer, colStart As Integer, _
        rngPastePoint As Range, Optional asValue As Boolean = False)
    Dim rngTarget As Range
    Set rngTarget = fexDownRightRange(wkb.Sheets(sheetNameOrIdx).Cells(rowStart, colStart))
    
    copy rngTarget, rngPastePoint, asValue
End Sub

Public Sub copyFromByNameRightRange(sheetNameOrIdx As Variant, rangeStart As String, rngPastePoint As Range, _
                        Optional asValue As Boolean = False)
    Dim rngTarget As Range
    Set rngTarget = fexRightRange(wkb.Sheets(sheetNameOrIdx).Range(rangeStart))
    
    copy rngTarget, rngPastePoint, asValue
End Sub

Public Sub copyFromByCellPointRightRange(sheetNameOrIdx As Variant, rowStart As Integer, colStart As Integer, _
        rngPastePoint As Range, Optional asValue As Boolean = False)
    Dim rngTarget As Range
    Set rngTarget = fexRightRange(wkb.Sheets(sheetNameOrIdx).Cells(rowStart, colStart))
    
    copy rngTarget, rngPastePoint, asValue
End Sub

Public Sub copyFromByNameDownRange(sheetNameOrIdx As Variant, rangeStart As String, rngPastePoint As Range, _
                        Optional asValue As Boolean = False)
    Dim rngTarget As Range
    Set rngTarget = fexDownRange(wkb.Sheets(sheetNameOrIdx).Range(rangeStart))
    
    copy rngTarget, rngPastePoint, asValue
End Sub

Public Sub copyFromByCellPointDownRange(sheetNameOrIdx As Variant, rowStart As Integer, colStart As Integer, _
        rngPastePoint As Range, Optional asValue As Boolean = False)
    Dim rngTarget As Range
    Set rngTarget = fexDownRange(wkb.Sheets(sheetNameOrIdx).Cells(rowStart, colStart))
    
    copy rngTarget, rngPastePoint, asValue
End Sub

Public Sub copyToByName(rngTarget As Range, sheetNameOrIdx As Variant, rngStart As String, _
        Optional asValue As Boolean = False)
    Dim rngPastePoint As Range
    Set rngPastePoint = wkb.Sheets(sheetNameOrIdx).Range(rngStart)
    
    copy rngTarget, rngPastePoint, asValue
    wkb.save
    Application.CutCopyMode = False
End Sub

Public Sub copyToByCellPoint(rngTarget As Range, sheetNameOrIdx As Variant, rowStart As Integer, _
        colStart As Integer, Optional asValue As Boolean = False)
    
    Dim rngPastePoint As Range
    Set rngPastePoint = wkb.Sheets(sheetNameOrIdx).Cells(rowStart, colStart)
    
    copy rngTarget, rngPastePoint, asValue
    wkb.save
    Application.CutCopyMode = False
End Sub

Public Function getReletivePositionValue(sheetNameOrIdx As Variant, value As Variant, _
    Optional r As Integer = 1, Optional c As Integer = 2) As Variant
    getReletivePositionValue = ehGetRelativePositionValue(wkb.Sheets(sheetNameOrIdx), value, r, c)
End Function

Public Function getRangeByName(sheetNameOrIdx As Variant, rngName As String) As Range
    Set getRangeByName = wkb.Sheets(sheetNameOrIdx).Range(rngName)
End Function

Public Function getRangeByCellPoint(sheetNameOrIdx As Variant, rowStart As Integer, colStart As Integer, _
        rowEnd As Integer, colEnd As Integer) As Range
    Set getRangeByCellPoint = wkb.Sheets(sheetNameOrIdx).Range(wkb.Sheets(sheetNameOrIdx).Cells(rowStart, colStart), wkb.Sheets(sheetNameOrIdx).Cells(rowEnd, colEnd))
End Function

Public Function getDownRightRangeByName(sheetNameOrIdx As Variant, rangeStart As String) As Range
    Set getDownRightRangeByName = fexDownRightRange(wkb.Sheets(sheetNameOrIdx).Range(rangeStart))
End Function

Public Function getDownRangeByName(sheetNameOrIdx As Variant, rangeStart As String) As Range
    Set getDownRangeByName = fexDownRange(wkb.Sheets(sheetNameOrIdx).Range(rangeStart))
End Function
Public Function getRightRangeByName(sheetNameOrIdx As Variant, rangeStart As String) As Range
    Set getRightRangeByName = fexRightRange(wkb.Sheets(sheetNameOrIdx).Range(rangeStart))
End Function

Public Function getDownRightRangeByCellPoint(sheetNameOrIdx As Variant, rowStart As Integer, colStart As Integer)
    Set getDownRightRangeByCellPoint = fexDownRightRange(wkb.Sheets(sheetNameOrIdx).Cells(rowStart, colStart))
End Function

Public Function getDownRangeByCellPoint(sheetNameOrIdx As Variant, rowStart As Integer, colStart As Integer)
    Set getDownRangeByCellPoint = fexDownRange(wkb.Sheets(sheetNameOrIdx).Cells(rowStart, colStart))
End Function

Public Function getRightRangeByCellPoint(sheetNameOrIdx As Variant, rowStart As Integer, colStart As Integer)
    Set getRightRangeByCellPoint = fexRightRange(wkb.Sheets(sheetNameOrIdx).Cells(rowStart, colStart))
End Function

