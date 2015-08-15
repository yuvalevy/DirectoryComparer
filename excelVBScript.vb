Public inDebugMode As Boolean
Public lastWritenInRow As String

Private Sub CopyButton_Click()
    
    
    Call FixAllButtons
End Sub

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
     
    Call AddAndMarkFiles
    
    Call FixAllButtons
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


Sub AddAndMarkFiles()
       
    Dim srcStartCell As String
    Dim desStartCell As String
    Dim srcEndCell As String
    Dim srcRangeValue As String
    
    Dim cellSrcLetter As String
    'Dim cellDestLetter As String
    
    Dim directoryCell As String
    Dim directoryPath As String
    Dim filesPattern As String
    
    
    lastWritenInRow = shtActive.UsedRange.Rows.Count
    
    ' ##Get Configuration
    directoryCell = shtConfig.Cells(1, 2).Value 'source directory cell
    cellSrcLetter = shtConfig.Cells(2, 2).Value  'source cell letter
    filesPattern = shtConfig.Cells(3, 2).Value 'directory pattern
    srcStartCell = shtConfig.Cells(4, 2).Value 'srcStartCell
    desStartCell = shtConfig.Cells(5, 2).Value 'desStartCell
    'cellDestLetter = shtConfig.Cells(6, 2).Value 'destination cell letter
    srcEndCell = cellSrcLetter & lastWritenInRow
    srcRangeValue = srcStartCell & ":" & srcEndCell
    
    ' ##Get directory path from sheet
    directoryPath = Range(directoryCell).Value
    
    If GetAttr(directoryPath) = vbDirectory Then
    
        If lastWritenInRow < 4 Then
            lastWritenInRow = 4
        End If
        
    
        ' ##Checks if last char in directory is '\'
        If Right(directoryPath, 1) <> "\" Then
            directoryPath = directoryPath & "\"
        End If
        
        ' ##Marking all cells in color red (3)
        Call MarkCells(Range(srcRangeValue), 3)
    
        
        Call MarkDirectory(directoryPath, filesPattern, srcRangeValue, cellSrcLetter)

    Else
        MsgBox "Directory at " & directoryCell & " does not exists"
    End If
    MessageBox "Finish CalcFilesByDirectory"
    

End Sub

' ### HELP SUBs

Private Sub MarkFiles(subDirectory As String, filesPattern As String, srcRangeValue As String, cellSrcLetter As String)
    
    Dim markingRow As String
    Dim foundVal As String
     
    fileName = GetFiles(subDirectory, filesPattern)
    
    For Each fn In fileName
        If fn <> "" Then
            Set vals = Range(srcRangeValue).Find(fn)
            
            If (vals Is Nothing) Then
                ' ##Adds new file to the last row
                lastWritenInRow = lastWritenInRow + 1
                markingRow = cellSrcLetter & lastWritenInRow
            
                Range(markingRow).Value = fn
            Else
                markingRow = cellSrcLetter & vals.Row
            End If
            
            ' ##Mark existing file in default color
            Call MarkCells(Range(markingRow), 0)
        End If
    Next
 
End Sub

Private Sub MarkDirectory(directory As String, filesPattern As String, srcRangeValue As String, cellSrcLetter As String)
    
    Dim drctVal As String
    Call MarkFiles(directory, filesPattern, srcRangeValue, cellSrcLetter)
      
     myDirectories = GetDirectories(directory)
      
    For i = 0 To UBound(myDirectories) Step 1
        drctVal = myDirectories(i)
        If drctVal <> "" Then
            Call MarkFiles(drctVal, filesPattern, srcRangeValue, cellSrcLetter)
            Call MarkDirectory(drctVal, filesPattern, srcRangeValue, cellSrcLetter)
        End If
    Next

End Sub

Private Function GetDirectories(path As String) As String()
    
    Dim fullpath As String
    Dim folderName As String
    Dim i As Long
    Dim myArray(1000) As String
  
    folderName = Dir(path, vbDirectory)
    Do While folderName <> ""
        If folderName <> ".." And folderName <> "." Then
       fullpath = path & folderName
        
            If GetAttr(fullpath) = vbDirectory Then
                myArray(i) = fullpath & "\"
                i = i + 1
            End If
        End If
        folderName = Dir
    Loop
    
    GetDirectories = myArray
End Function

Private Function GetFiles(path As String, pattern As String) As String()
    
    Dim fullpath As String
    Dim fileName As String
    Dim i As Long
    Dim myArray(1000) As String
 
    fileName = Dir(path & pattern)
    
    Do While fileName <> ""
        If fileName <> ".." And fileName <> "." Then
            fullpath = path & fileName
            If GetAttr(fullpath) = 32 Then
                myArray(i) = fullpath
                i = i + 1
            End If
        End If
        fileName = Dir()
    Loop
    
    GetFiles = myArray
End Function


Private Sub MessageBox(messageToDesplay As String)
    If inDebugMode = True Then
        MsgBox messageToDesplay
    End If
End Sub

Private Sub EditingSheet(setState As Boolean)
    
    If (inDebugMode = False) Then
   
        'With Application
        '    .Calculation = xlCalculationManual
       '     .ScreenUpdating = setState
        'End With
        
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
        If cell.Row <> 3 Then
            cell.Interior.ColorIndex = color
        End If
    Next cell
End Sub

Private Sub DeleteRow(rwIndex As Long)
    Call MessageBox("Deleting row  " & rwIndex)
    Range("A" & rwIndex).EntireRow.Delete
End Sub
