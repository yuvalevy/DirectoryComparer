Public inDebugMode As Boolean

Private Sub CalcButton_Click()
    inDebugMode = True
    Call EditingSheet(False)
    
    Dim directoryPath As String
    directoryPath = Range("B2")
     
    Call CalcFilesByDirectory(directoryPath)
    
    Call FixButton(CalcButton, 87, 70)
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
    
    ' Dim cellsForDelete
    Set myCells = shtActive.UsedRange
    
    Dim rowCount As Long
    Dim colCount As Long
    
    rowCount = myCells.Rows.Count
    colCount = myCells.Columns.Count
    
    Call MessageBox("Row: " & rowCount & " Col: " & colCount)
    
    ReDim myArray(rowCount, colCount) As Object
    
    
    ' ## check the bowndries
    ' If cell2.Row > 3 And (cell2.Column = 3 Or cell2.Column = 4) Then
    
    
    Dim isWholeRowEmpty As Boolean
    
    For RowIndex = 4 To rowCount
         wholeRowEmpty = True
       
        For colIndex = 2 To 3
            Call MessageBox(myCells(RowIndex, colIndex).Value)
            
            If myCells(RowIndex, colIndex) <> "" Then
            
                Call MessageBox(colIndex & "<>")
                 wholeRowEmpty = False
            
            End If
        Next
        
        If wholeRowEmpty = True Then
            Call MessageBox("Deleting row  " & RowIndex)
            myCells(RowIndex, 0).EntireRow.Delete
        End If
    Next
    
    Call EditingSheet(True)

End Sub

Sub CalcFilesByDirectory(directory As String)
    
    Dim cellLetter As String
    
    ' ##Get Configuration
    Set directoryPattern = shtConfig.Cells(3, 2).Value
    Set srcStartCell = shtConfig.Cells(4, 2).Value
    Set desStartCell = shtConfig.Cells(5, 2).Value
    Set lastWritenInRow = shtConfig.Cells(1, 2).Value
    
    ' ##Checks if last char in directory is '\'
    If Right(directoryPath, 1) <> "\" Then
        directoryPath = directoryPath & "\"
    End If
    
    
    Set srcLst = Range()
    
   ' fileName = Dir(directoryPath & directoryPattern)
    
   ' cellLetter = "B"
   ' Index = 1
   ' Do While fileName <> ""
   '     Call MessageBox(fileName)
   '     Range(cellLetter & Index).Value = fileName
   '     Index = Index + 1
   '     fileName = Dir()
   ' Loop
    
   ' cellLetter = "C"
   ' fileName = Dir(directory & directoryPattern)
      
   'Range(cellLetter & Index).Value (fileName)
      
        
    
    
    MsgBox "Finish"
    

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
