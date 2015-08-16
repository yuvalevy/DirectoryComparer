Public inDebugMode As Boolean
Public lastWritenInRow As Long
Public workingDirectory As String
Public filesPattern As String
Public cellSrcLetter As String
Public cellDestLetter As String
Public cellStatusLetter As String

Private Sub CopyButton_Click()
    inDebugMode = False
    'Call EditingSheet(False)
     
    Call CopyFiles

    Call FixAllButtons
    'Call EditingSheet(True)
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

    Dim rowCount As Long
    Dim rwIndex As Long
    Dim currentCell As String
    
    Call SetLetters
    
    rowCount = GetLastRow
      
    For rwIndex = rowCount To 4 Step -1
        currentCell = cellSrcLetter & rwIndex
        If IsRowMarkedRed(rwIndex) Then
            Call DeleteRow(rwIndex)
        End If
    Next
    
End Sub


Sub DeleteEmptyRows()
   
    Dim rowCount As Long
    Dim rwIndex As Long
    Dim isWholeRowEmpty As Boolean
    
    rowCount = GetLastRow
    
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
            Call DeleteRow(rwIndex)
        End If
    
    Next

End Sub


Sub AddAndMarkFiles()

    Dim srcRangeValue As String '''
    
    srcRangeValue = GetCurrentActiveRange
    
   If IsDirectoryOK Then
        ' ##Marking all cells in color red (3)
        Call MarkCells(Range(srcRangeValue), 3)
        
        ' ##Starting to check all files in directory and subdirectory
        Call MarkFiles(workingDirectory, srcRangeValue)
        Call MarkDirectory(workingDirectory, srcRangeValue)
    Else
        MsgBox "Source folder does not exists"
    End If

End Sub

Private Sub CopyFiles()
    
    Dim rwIndex As Long
    Call SetLetters
    
    lastWritenInRow = GetLastRow
    
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
    
    sourceFile = Range(srcCell).Value
    destinationFile = Range(dstCell).Value
    
    ' ##Making sure program can access files
    statusIndex = 1
    rslt = True
    Do While (statusIndex < 4 And rslt)
        rslt = ChecksBeforeCopy(statusIndex, sourceFile, destinationFile, rwIndex)
        statusIndex = statusIndex + 1
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
    srcStartCell = shtConfig.Cells(4, 2).Value 'srcStartCell
    desStartCell = shtConfig.Cells(5, 2).Value 'desStartCell
    
    lastWritenInRow = GetLastRow
    srcEndCell = cellSrcLetter & lastWritenInRow
    srcRangeValue = srcStartCell & ":" & srcEndCell

    GetCurrentActiveRange = srcRangeValue
End Function

Private Sub SetLetters()
     cellSrcLetter = shtConfig.Cells(2, 2).Value  'source cell letter
     cellDestLetter = shtConfig.Cells(6, 2).Value  'destination cell letter
     cellStatusLetter = shtConfig.Cells(7, 2).Value  'status cell letter
End Sub


Private Function IsDirectoryOK() As Boolean
  
    Dim directoryCell As String
    Dim directoryPath As String
   
    directoryCell = shtConfig.Cells(1, 2).Value 'source directory cell
    filesPattern = shtConfig.Cells(3, 2).Value  'directory pattern
  

    ' ##Get directory path from sheet
    directoryPath = Range(directoryCell).Value
    
    ' ##Checks if last char in directory is '\'
    If Right(directoryPath, 1) <> "\" Then
        directoryPath = directoryPath & "\"
    End If
    
     
    temp = Dir(directoryPath, vbDirectory)
    
    If temp <> "" Then
        workingDirectory = directoryPath
        IsDirectoryOK = True
    Else
       IsDirectoryOK = False
    End If
    
End Function

Private Function ChecksBeforeCopy(statusIndex As Integer, src As String, des As String, row As Long) As Boolean
            
    ChecksBeforeCopy = True

    Select Case statusIndex
        Case 1
        ' ##Checks if user spesified source file
            If src = "" Then
                MessageBox ("Missing source file in row " & row)
                ChecksBeforeCopy = False
            End If
        Case 2
        ' ##Checks if user spesified destination file
            If des = "" Then
                MessageBox ("Missing destination file in row " & row)
                ChecksBeforeCopy = False
            End If
        Case 3
            ' ##Checks if source file exists
            If Dir(src) = "" Then
                MessageBox ("Source file does not exists in row " & row)
                ChecksBeforeCopy = False
            End If
        Case Else
            ChecksBeforeCopy = True
        
    End Select
End Function

Private Sub WriteStatusResult(statusIndex As Integer, rwIndex As Long)

    Dim sttsCell As String
    sttsCell = cellStatusLetter & rwIndex
   
    Select Case statusIndex
        Case 0
            Range(sttsCell).Value = "Not yet copied"
        Case 1
            Range(sttsCell).Value = "Source file missing"
        Case 2
            Range(sttsCell).Value = "Destination file missing"
        Case 3
            Range(sttsCell).Value = "Source file does not exists"
        Case 100
            Range(sttsCell).Value = "Copied"
        'Case Else
            'Range(sttsCell).Value = "Copied"
    End Select
End Sub


Private Function GetStatusResult(rwIndex As Long) As Integer

    Dim sttsCell As String
    sttsCell = cellStatusLetter & rwIndex
   
    Select Case Range(sttsCell).Value
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

Private Sub MarkFiles(subDirectory As String, srcRangeValue As String)
    
    Dim markingRow As String
    Dim foundVal As String
     
    fileName = GetFiles(subDirectory)
    ' TODO: add if the first one is empy..
    For Each fn In fileName
        If fn <> "" Then
            Set vals = Range(srcRangeValue).Find(fn)
            
            If (vals Is Nothing) Then
                ' ##Adds new file to the last row
                lastWritenInRow = lastWritenInRow + 1
                markingRow = cellSrcLetter & lastWritenInRow
            
                Range(markingRow).Value = fn
            Else
                markingRow = cellSrcLetter & vals.row
            End If
            
            ' ##Mark existing file in default color
            Call MarkCells(Range(markingRow), 0)
        End If
    Next
 
End Sub

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

Private Function GetFiles(path As String) As String()
    
    Dim fullpath As String
    Dim fileName As String
    Dim i As Long
    Dim myArray(1000) As String
 
    fileName = Dir(path & filesPattern)
    
    Do While fileName <> ""
        If fileName <> ".." And fileName <> "." Then
            fullpath = path & fileName
            If GetAttr(fullpath) <> vbDirectory Then
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
        If cell.row <> 3 Then
            cell.Interior.ColorIndex = color
        End If
    Next cell
End Sub

Private Sub DeleteRow(rwIndex As Long)
    Call MessageBox("Deleting row  " & rwIndex)
    Range("A" & rwIndex).EntireRow.Delete
End Sub

Private Function GetLastRow() As Long
    Dim last As Long
    last = shtActive.UsedRange.Rows.Count
    
    If last < 4 Then
        last = 4
    End If
    
    GetLastRow = last
End Function
