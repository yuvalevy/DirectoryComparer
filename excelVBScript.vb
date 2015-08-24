Public inDebugMode As Boolean
Public lastWritenInRow As Long
Public WorkingDirectory As String
Public filesPattern As String
Public cellSrcLetter As String
Public cellDestLetter As String
Public cellStatusLetter As String

Private Sub CopyButton_Click()
    inDebugMode = False
    'Call SheetModule.EditingSheet(False)
     
    Call CopyFiles

    Call FixAllButtons
    'Call SheetModule.EditingSheet(True)
End Sub

Private Sub RedButton_Click()
    inDebugMode = False
    Call SheetModule.EditingSheet(False)
     
    Call DeleteRedLines
    
    Call DeleteEmptyRows
  
    Call FixAllButtons
    Call SheetModule.EditingSheet(True)
End Sub

Private Sub CalcButton_Click()
    inDebugMode = False
    Call SheetModule.EditingSheet(False)
     
    Call AddAndMarkFiles
    
    Call FixAllButtons
    Call SheetModule.EditingSheet(True)
End Sub

Private Sub ArrangeButton_Click()
    inDebugMode = False
    Call SheetModule.EditingSheet(False)
    
    Call DeleteEmptyRows
        
    Call FixAllButtons
    Call SheetModule.EditingSheet(True)
End Sub


' ##Buisness Subs


Sub DeleteRedLines()

    Dim rowCount As Long
    Dim rwIndex As Long
    Dim currentCell As String
    
    Call SetLetters
    
    rowCount = SheetModule.GetLastRow
      
    For rwIndex = rowCount To 4 Step -1
        currentCell = cellSrcLetter & rwIndex
        If IsRowMarkedRed(rwIndex) Then
            Call SheetModule.DeleteRow(rwIndex)
        End If
    Next
    
End Sub


Sub DeleteEmptyRows()
   
    Dim rowCount As Long
    Dim rwIndex As Long
    Dim isWholeRowEmpty As Boolean
    
    rowCount = SheetModule.GetLastRow
    
    Call MessageBox("Last Row: " & rowCount)
    
    For rwIndex = rowCount To 4 Step -1
        wholeRowEmpty = True
       
        'For colIndex = 2 To 3
        '    If Cells(rwIndex, colIndex) <> "" Then
        '        wholeRowEmpty = False
        '    End If
        'Next
        
        If Cells(rwIndex, 2) <> "" Or Cells(rwIndex, 3) <> "" Then
            wholeRowEmpty = False
        End If
    
        
        If wholeRowEmpty = True Then
            Call SheetModule.DeleteRow(rwIndex)
        End If
    
    Next

End Sub


Sub AddAndMarkFiles()

    Dim srcRangeValue As String
    
    srcRangeValue = GetCurrentActiveRange
    
    If Utils.IsDirectoryOK Then
        ' ##Marking all cells in color red (3)
        Call SheetModule.MarkCells(Range(srcRangeValue), 3)
        
        ' ##Starting to check all files in directory and subdirectory
        Call MarkFiles(WorkingDirectory, srcRangeValue)
        Call MarkDirectory(WorkingDirectory, srcRangeValue)
    Else
        MsgBox "Source folder does not exists"
    End If

End Sub

Private Sub CopyFiles()
    
    Dim rwIndex As Long
    Call SetLetters
    
    lastWritenInRow = SheetModule.GetLastRow
    
    For rwIndex = 4 To lastWritenInRow
        If GetStatusResult(rwIndex) <> 100 And Not IsRowMarkedRed(rwIndex) Then
            
            Call CopySingleFile(rwIndex)
           
        End If
    Next

End Sub

Private Sub CopySingleFile(rwIndex As Long)
    
    Dim srcCell As String
    Dim dstCell As String
    
    Dim sourceFile As String
    Dim destinationFile As String
    
    Dim statusIndex As Integer
    Dim rslt As Boolean
  
    srcCell = cellSrcLetter & rwIndex
    dstCell = cellDestLetter & rwIndex
    
    sourceFile = Range(srcCell).value
    destinationFile = Range(dstCell).value
    
    ' ##Making sure program can access files
    statusIndex = 0
    rslt = True
    Do While (statusIndex < 4 And rslt)
        statusIndex = statusIndex + 1
        rslt = IOModule.ChecksBeforeCopy(statusIndex, sourceFile, destinationFile, rwIndex)
    Loop
    
    If rslt = True Then
        Call FileSystem.FileCopy(sourceFile, destinationFile)
        statusIndex = 100
    End If
    Call WriteStatusResult(statusIndex, rwIndex)

End Sub

' ### HELP SUBs

Private Function GetCurrentActiveRange() As String
    
    Dim srcStartCell As String
    Dim desStartCell As String
    Dim srcEndCell As String
    Dim srcRangeValue As String
    
    ' ##Get Configuration
    Call SetLetters
    srcStartCell = shtConfig.Cells(4, 2).value 'srcStartCell
    desStartCell = shtConfig.Cells(5, 2).value 'desStartCell
    
    lastWritenInRow = SheetModule.GetLastRow
    srcEndCell = cellSrcLetter & lastWritenInRow
    srcRangeValue = srcStartCell & ":" & srcEndCell

    GetCurrentActiveRange = srcRangeValue
End Function

Private Sub SetLetters()
     cellSrcLetter = shtConfig.Cells(2, 2).value  'source cell letter
     cellDestLetter = shtConfig.Cells(6, 2).value  'destination cell letter
     cellStatusLetter = shtConfig.Cells(7, 2).value  'status cell letter
End Sub



Private Sub WriteStatusResult(statusIndex As Integer, rwIndex As Long)

    Dim sttsCell As String
    Dim setedValue As String
    
    sttsCell = cellStatusLetter & rwIndex
   
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


Private Function GetStatusResult(rwIndex As Long) As Integer

    Dim sttsCell As String
    sttsCell = cellStatusLetter & rwIndex
   
    Select Case Range(sttsCell).value
        Case "Not yet copied"
            GetStatusResult = 0
        Case "Source file missing"
            GetStatusResult = 1
        Case "Destination file missing"
            GetStatusResult = 2
        Case "Source file does not exists"
            GetStatusResult = 3
        Case "Copied"
            GetStatusResult = 100
        Case Else
            GetStatusResult = -1
    End Select
    
End Function

Private Function IsRowMarkedRed(rwIndex As Long) As Boolean

    Dim sttsCell As String
    sttsCell = cellSrcLetter & rwIndex
   
    IsRowMarkedRed = Range(sttsCell).Interior.ColorIndex = 3
    
End Function


Private Sub MarkDirectory(directory As String, srcRangeValue As String)
    
    Dim drctVal As String
  
    myDirectories = GetDirectories(directory)
      
    For i = 0 To UBound(myDirectories) Step 1
        drctVal = myDirectories(i)
        If drctVal <> "" Then
            Call MarkFiles(drctVal, srcRangeValue)
            Call MarkDirectory(drctVal, srcRangeValue)
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

Private Sub MessageBox(messageToDesplay As String)
    If inDebugMode = True Then
        MsgBox messageToDesplay
    End If
End Sub

