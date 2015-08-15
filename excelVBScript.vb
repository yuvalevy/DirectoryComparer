Public inDebugMode As Boolean

Private Sub CalcButton_Click()
    inDebugMode = False
    Call EditingSheet(False)
    
    Dim directoryPath As String
    directoryPath = Range("B2")
     
    Call CalcFilesByDirectory
    
    Call FixButton(CalcButton, 87, 70)
    Call DeleteEmptyRows
    Call EditingSheet(True)
End Sub

Private Sub ArrangeButton_Click()
    inDebugMode = False
    Call EditingSheet(False)
    
    Call DeleteEmptyRows
    
    Call FixButton(ArrangeButton, 50, 70)
    Call EditingSheet(True)
End Sub

Sub DeleteEmptyRows()
    
    Set myCells = shtActive.UsedRange
    
    Dim rowCount As Long
    Dim colCount As Long
    Dim isWholeRowEmpty As Boolean
    
    rowCount = myCells.Rows.Count
    colCount = myCells.Columns.Count
    
    Call MessageBox("Last Row: " & rowCount & "Last Col: " & colCount)
    
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
            Call MessageBox("Deleting row  " & rwIndex)
            myCells(rwIndex, 0).EntireRow.Delete
        End If
    
    Next

End Sub

Sub CalcFilesByDirectory()
    
    Dim lastWritenInRow As String
    Dim srcStartCell As String
    Dim desStartCell As String
    Dim srcEndCell As String
    Dim srcLst As String
    Dim srcRangeValue As String
    Dim fileName As String
    Dim cellLetter As String
    Dim directoryCell As String
    Dim directoryPath As String
    Dim directoryPattern As String
    Dim srcRangeLst As Range
    
    lastWritenInRow = shtActive.UsedRange.Rows.Count
    
    ' ##Get Configuration
    directoryCell = shtConfig.Cells(1, 2).Value 'source directory cell
    cellLetter = shtConfig.Cells(2, 2).Value 'source cell letter
    directoryPattern = shtConfig.Cells(3, 2).Value 'directory pattern
    srcStartCell = shtConfig.Cells(4, 2).Value 'srcStartCell
    desStartCell = shtConfig.Cells(5, 2).Value 'desStartCell
    srcEndCell = cellLetter & lastWritenInRow
    srcRangeValue = srcStartCell & ":" & srcEndCell
    
    ' ##Get directory path from sheet
     
    directoryPath = Range(directoryCell).Value
    
    ' ##Checks if last char in directory is '\'
    If Right(directoryPath, 1) <> "\" Then
        directoryPath = directoryPath & "\"
    End If
     
    fileName = Dir(directoryPath & "*")
    

    Dim i As Integer
    Dim foundVal As String

    i = 1
    Do While fileName <> ""
        Call MessageBox(fileName)
        Set vals = Range(srcRangeValue).Find(fileName)
        
        If (vals Is Nothing) Then
            ' ##Adds new file to the last row
            lastWritenInRow = lastWritenInRow + 1
            Range(cellLetter & lastWritenInRow).Value = fileName
        End If
        i = i + 1
        fileName = Dir()
    Loop

    'fileName = Dir(directoryPath & directoryPattern)
      
'   Range(cellLetter & Index).Value (fileName)
    
    MessageBox "Finish CalcFilesByDirectory"
    

End Sub
' ### HELP SUBs

Private Sub MessageBox(messageToDesplay As String)
    If inDebugMode = True Then
        MsgBox messageToDesplay
    End If
End Sub

Private Sub EditingSheet(setState As Boolean)
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = setState
    End With
End Sub


Private Sub FixButton(button As MSForms.CommandButton, top As Integer, left As Integer)
    button.Height = 25
    button.Width = 100
    button.top = top
    button.left = left
End Sub

