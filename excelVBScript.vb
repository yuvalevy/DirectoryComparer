Public inDebugMode As Boolean

Private Sub RedButton_Click()
    inDebugMode = False
    Call EditingSheet(False)
     
    Call DeleteRedLines
    
    Call DeleteEmptyRows
  
    Call FixAllButtons
    Call EditingSheet(True)
End Sub

Private Sub CalcButton_Click()
    inDebugMode = False
    Call EditingSheet(False)
     
    Call AddOrRemoveFiles
    
    Call FixAllButtons
    Call DeleteEmptyRows
    Call EditingSheet(True)
End Sub

Private Sub ArrangeButton_Click()
    inDebugMode = False
    Call EditingSheet(False)
    
    Call DeleteEmptyRows
        
    Call FixAllButtons
    Call EditingSheet(True)
End Sub


' ##Buisness Subs


Sub DeleteRedLines()
    
    Set myCells = shtActive.UsedRange
    
    Dim rowCount As Long
    Dim colCount As Long
    Dim rwIndex As Long
    Dim cellSrcLetter As String
    Dim currentCell As String
    
    cellSrcLetter = shtConfig.Cells(2, 2).Value  'source cell letter
    rowCount = myCells.Rows.Count
    colCount = myCells.Columns.Count
     
    For rwIndex = rowCount To 4 Step -1
        currentCell = cellSrcLetter & rwIndex
        If Range(currentCell).Interior.ColorIndex = 3 Then
            Call DeleteRow(rwIndex)
        End If
    Next
    
End Sub


Sub DeleteEmptyRows()
    
    Set myCells = shtActive.UsedRange
    
    Dim rowCount As Long
    Dim colCount As Long
    Dim rwIndex As Long
    Dim isWholeRowEmpty As Boolean
    
    rowCount = myCells.Rows.Count
    colCount = myCells.Columns.Count
    
    Call MessageBox("Last Row: " & rowCount & " Last Col: " & colCount)
    
    ReDim myArray(rowCount, colCount) As Object
    
    
    ' ## check the bowndries
    ' If cell2.Row > 3 And (cell2.Column = 3 Or cell2.Column = 4) Then
      
    For rwIndex = rowCount To 4 Step -1
        wholeRowEmpty = True
       
        For clIndex = 2 To 3
            Call MessageBox(myCells(rwIndex, clIndex).Value)
            
            If myCells(rwIndex, clIndex) <> "" Then
                Call MessageBox(clIndex & "<>")
                wholeRowEmpty = False
            End If
        
        Next
        
        If wholeRowEmpty = True Then
            Call DeleteRow(rwIndex)
        End If
    
    Next

End Sub


Sub AddOrRemoveFiles()
    
    Dim lastWritenInRow As String
    
    Dim srcStartCell As String
    Dim desStartCell As String
    Dim srcEndCell As String
    Dim srcRangeValue As String
    
    Dim cellSrcLetter As String
    'Dim cellDestLetter As String
    
    Dim directoryCell As String
    Dim directoryPath As String
    Dim directoryPattern As String
    Dim fileName As String
     
    Dim markingRow As String
    
    
    lastWritenInRow = shtActive.UsedRange.Rows.Count
    
    ' ##Get Configuration
    directoryCell = shtConfig.Cells(1, 2).Value 'source directory cell
    cellSrcLetter = shtConfig.Cells(2, 2).Value  'source cell letter
    directoryPattern = shtConfig.Cells(3, 2).Value 'directory pattern
    srcStartCell = shtConfig.Cells(4, 2).Value 'srcStartCell
    desStartCell = shtConfig.Cells(5, 2).Value 'desStartCell
    'cellDestLetter = shtConfig.Cells(6, 2).Value 'destination cell letter
    srcEndCell = cellSrcLetter & lastWritenInRow
    srcRangeValue = srcStartCell & ":" & srcEndCell
    
    ' ##Get directory path from sheet
    directoryPath = Range(directoryCell).Value
    
    ' ##Checks if last char in directory is '\'
    If Right(directoryPath, 1) <> "\" Then
        directoryPath = directoryPath & "\"
    End If
     
    fileName = Dir(directoryPath & "*")
    
    ' ##Marking all cells in color red (3)
    Call MarkCells(Range(srcRangeValue), 3)

    Dim i As Integer
    Dim foundVal As String

    i = 1
    Do While fileName <> ""
        Call MessageBox(fileName)
        Set vals = Range(srcRangeValue).Find(fileName)
        
        If (vals Is Nothing) Then
            ' ##Adds new file to the last row
            lastWritenInRow = lastWritenInRow + 1
            markingRow = cellSrcLetter & lastWritenInRow
            Range(markingRow).Value = fileName
        Else
            markingRow = cellSrcLetter & vals.Row
        End If
        
        ' ##Mark existing file in default color
        Call MarkCells(Range(markingRow), 0)

        i = i + 1
        fileName = Dir()
    Loop

    'fileName = Dir(directoryPath & directoryPattern)
      
    'Range(cellSrcLetter & Index).Value (fileName)
    
    MessageBox "Finish CalcFilesByDirectory"
    

End Sub

' ### HELP SUBs

Private Sub MessageBox(messageToDesplay As String)
    If inDebugMode = True Then
        MsgBox messageToDesplay
    End If
End Sub

Private Sub EditingSheet(setState As Boolean)
    
    If (inDebugMode = False) Then
   
        With Application
            .Calculation = xlCalculationManual
            .ScreenUpdating = setState
        End With
        
    End If
    
End Sub

Private Sub FixAllButtons()
    Call FixButton(CalcButton, 50, 70)
    Call FixButton(CopyButton, 85, 70)
    Call FixButton(ArrangeButton, 120, 70)
    Call FixButton(RedButton, 155, 70)
End Sub

Private Sub FixButton(button As MSForms.CommandButton, top As Integer, left As Integer)
    button.Height = 25
    button.Width = 110
    button.top = top
    button.left = left
End Sub

Private Sub MarkCells(rng As Range, color As String)
    For Each cell In rng
        
       cell.Interior.ColorIndex = color
        
    Next cell
End Sub

Private Sub DeleteRow(rwIndex As Long)
    Call MessageBox("Deleting row  " & rwIndex)
    Range("A" & rwIndex).EntireRow.Delete
End Sub
