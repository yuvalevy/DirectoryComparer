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
    Call FixButton(CalcButton, 50, 70)
    Call FixButton(CopyButton, 85, 70)
    Call FixButton(ArrangeButton, 120, 70)
    Call FixButton(RedButton, 155, 70)
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
    Call MessageBox("Deleting row  " & rwIndex)
    Range("A" & rwIndex).EntireRow.Delete
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
    
    If (inDebugMode = False) Then
        With Application
            .Calculation = xlCalculationManual
            .ScreenUpdating = setState
        End With
    End If
    
End Sub

'## Sets value for spesific cell
Public Sub SetCellValue(cellRange As String, setsValue As String)
    Range(cellRange).value = setsValue
End Sub
