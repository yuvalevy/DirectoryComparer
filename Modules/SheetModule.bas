Attribute VB_Name = "SheetModule"

' ##Marking the whole range with spesific color
Public Sub MarkCells(rng As Range, color As Integer)
    
    For Each currentCell In rng
        If currentCell.row > 3 Then
            currentCell.Interior.ColorIndex = color
        End If
    Next currentCell
    
End Sub

' ##Fixes all sheet bottons
Public Sub FixAllButtons()
    Call FixButton(shtActive.CalcButton, 50, 70)
    Call FixButton(shtActive.CopyButton, 85, 70)
    Call FixButton(shtActive.ArrangeButton, 120, 70)
    Call FixButton(shtActive.RedButton, 155, 70)
End Sub

' ##Fixes a spesific botton to the right dimention
Private Sub FixButton(button As MSForms.CommandButton, top As Integer, left As Integer)
    button.Height = 25
    button.Width = 110
    button.top = top
    button.left = left
End Sub

' ##Delete the whole row
Public Sub DeleteRow(rwIndex As Long)
    Call LoggingModule.MessageBox("Deleting row  " & rwIndex)
   shtActive.Range("A" & rwIndex).EntireRow.Delete
End Sub

' ##Gets the last row system wrote at
Public Function GetLastRow() As Long
    Dim last As Long
    last = shtActive.UsedRange.Rows.Count
    
    If last < 4 Then
        last = 4
    End If
    
    GetLastRow = last
End Function

' ##Enable/Disable the ability to edit the sheet
Public Sub EditingSheet(setState As Boolean)
    
    If (InDebugMode = False) Then
        With Application
            .Calculation = xlCalculationManual
            .ScreenUpdating = setState
        End With
    End If
    
End Sub

'## Sets value for spesific cell
Public Sub SetCellValue(cellRange As String, setsValue As String)
    shtActive.Range(cellRange).value = setsValue
End Sub

' ##Returns whether the spesific 'B' & rwIndex is colored in red
Public Function IsRowMarkedRed(rwIndex As Long) As Boolean

    Dim sttsCell As String
    sttsCell = Utils.CellSrcLetter & rwIndex
   
    IsRowMarkedRed = shtActive.Range(sttsCell).Interior.ColorIndex = 3
    
End Function

Public Sub WriteStatusResult(statusIndex As Integer, rwIndex As Long)

    Dim sttsCell As String
    Dim setedValue As String
    
    sttsCell = Utils.CellStatusLetter & rwIndex
   
    Select Case statusIndex
        Case 0
            setedValue = "Not yet copied"
        Case 1
            setedValue = "Source file missing"
        Case 2
            setedValue = "Destination file missing"
        Case 3
            setedValue = "Source file does not exists"
        Case 100
            setedValue = "Copied"
        'Case Else
            'setedValue= "Copied"
    End Select
    
    Call SheetModule.SetCellValue(sttsCell, setedValue)
    
End Sub
