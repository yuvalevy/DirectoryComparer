Attribute VB_Name = "IOModule"
Private FilesPattern As String
Public LastWritenInRow As Long

' ##Checks if 'directoryPath' is valid
Public Function ProjectDirectoryExists() As Boolean

    Dim directoryCell As String
    Dim directoryPath As String
    
    directoryCell = ConfigurationModule.GetSourceDirectoryCell
    FilesPattern = ConfigurationModule.GetDirectoryPattern
    
    ' ##Get directory path from sheet
    directoryPath = shtActive.Range(directoryCell).value
    directoryPath = FixDirectoryPath(directoryPath)
    
    If DirectoryExists(directoryPath) Then
        Utils.WorkingDirectory = directoryPath
        ProjectDirectoryExists = True
    Else
        ProjectDirectoryExists = False
    End If
 
End Function

' ##Whether a certain directory exists
Public Function DirectoryExists(directoryPath As String) As Boolean
' ################## REPLACE
    Dim att As VbFileAttribute
    att = GetAttr(directoryPath)
    
    If ((att = vbDirectory) Or (att = vbDirectory + vbReadOnly) Or (att = vbDirectory + vbSystem + vbHidden)) Then
        DirectoryExists = True
    Else
        DirectoryExists = False
    End If
    
   Dim FSO As FileSystemObject
    FSO.FolderExists (DirectoryExists)
    
End Function

' ##Checks if last char in directory is '\'
' ##If not, it adds it
Private Function FixDirectoryPath(directoryPath As String) As String
' ################## REPLACE not sure this will be needed
    If Right(directoryPath, 1) <> "\" Then
        directoryPath = directoryPath & "\"
    End If
    
    FixDirectoryPath = directoryPath
    
End Function

Public Function GetLastRow() As String
    LastWritenInRow = SheetModule.GetLastRow
    GetLastRow = LastWritenInRow
End Function

' ##List of validations made on the src & des files
Public Function ChecksBeforeCopy(statusIndex As Integer, src As String, des As String, row As Long) As Boolean
' ################## REPLACE
    ChecksBeforeCopy = True

    Select Case statusIndex
        Case 1
        ' ##Checks if user spesified source file
            If src = "" Then
                LoggingModule.MessageBox ("Missing source file in row " & row)
                ChecksBeforeCopy = False
            End If
        Case 2
        ' ##Checks if user spesified destination file
            If des = "" Then
                LoggingModule.MessageBox ("Missing destination file in row " & row)
                ChecksBeforeCopy = False
            End If
        Case 3
            ' ##Checks if source file not exists
            If Dir(src) = "" Then
                LoggingModule.MessageBox ("Source file does not exists in row " & row)
                ChecksBeforeCopy = False
            End If
        Case 4
                    
        Case Else
            ChecksBeforeCopy = True
        
    End Select
End Function

Public Sub CopySingleFile(rwIndex As Long)
' ################## REPLACE
    Dim srcCell As String
    Dim dstCell As String
    
    Dim sourceFile As String
    Dim destinationFile As String
    
    Dim statusIndex As Integer
    Dim rslt As Boolean
  
    srcCell = Utils.CellSrcLetter & rwIndex
    dstCell = Utils.CellDestLetter & rwIndex
    
    sourceFile = shtActive.Range(srcCell).value
    destinationFile = shtActive.Range(dstCell).value
    
    ' ##Making sure program can access files
    statusIndex = 0
    rslt = True
    Do While (statusIndex < 4 And rslt)
        statusIndex = statusIndex + 1
        rslt = ChecksBeforeCopy(statusIndex, sourceFile, destinationFile, rwIndex)
    Loop
    
    If rslt = True Then
        Call FileSystem.FileCopy(sourceFile, destinationFile)
        statusIndex = 100
    End If
    Call SheetModule.WriteStatusResult(statusIndex, rwIndex)

End Sub

' ######################### GET fuctions and subs

' ##Directory.GetFolders(path,SearchOptions.AllDirectories)
' note: recursion function
Public Function GetDirectories(directoryPath As String) As String()
' ################## REPLACE
    Dim fullPath As String
    Dim folderName As String
    Dim i As Long
    Dim myArray(1000) As String
  
    folderName = Dir(directoryPath, vbDirectory)
    Do While folderName <> ""
        '## When Dir() the first two results are 'this' and 'previous'
        If folderName <> ".." And folderName <> "." Then
        fullPath = directoryPath & folderName
        
            If DirectoryExists(fullPath) Then
                myArray(i) = fullPath & "\"
                i = i + 1
            End If
        End If
        folderName = Dir
    Loop
    
    GetDirectories = myArray
End Function


' ##Directory.GetFiles(path)
' note: pattern included in path parameter
Private Function GetFiles(path As String) As String()
' ################## REPLACE
    Dim fullPath As String
    Dim fileName As String
    Dim i As Long
    Dim myArray(1000) As String
 
    fileName = Dir(path & FilesPattern)
    
    Do While fileName <> ""
        If fileName <> ".." And fileName <> "." Then
            fullPath = path & fileName
            If GetAttr(fullPath) <> vbDirectory Then
                myArray(i) = fullPath
                i = i + 1
            End If
        End If
        fileName = Dir()
    Loop
    
    GetFiles = myArray
End Function

' ################## Marking

' ##Finds all directory sub FOLDERs and call AddAndMarkFiles()
Public Sub MarkSubDirectory(directory As String, srcRangeValue As String)
    
    Dim currentDirectory As String
  
    myDirectories = GetDirectories(directory)
      
    For i = LBound(myDirectories) To UBound(myDirectories) Step 1
        currentDirectory = myDirectories(i)
        If currentDirectory <> "" Then
            ' ##Mark all current directory files
            Call AddAndMarkFiles(currentDirectory, srcRangeValue)
            ' ##Find all subdirectory
            Call MarkSubDirectory(currentDirectory, srcRangeValue)
        End If
    Next

End Sub

' ##Marks new file with green. Marks old files with default color.
Public Sub AddAndMarkFiles(subDirectory As String, srcRangeValue As String)
    
    Dim markingRow As String
    Dim markingColor As Integer
    Dim files() As String
    Dim tempStr As String
    
    files = GetFiles(subDirectory)
    
    ' TODO: add if the first one is empy..
    For Each fn In files
        If fn <> "" Then
            Set vals = shtActive.Range(srcRangeValue).Find(fn)
            
            If (vals Is Nothing) Then
                tempStr = fn
                ' ##Adds new file to the last row
                LastWritenInRow = LastWritenInRow + 1
                markingRow = Utils.CellSrcLetter & LastWritenInRow
                markingColor = 4
                Call SheetModule.SetCellValue(markingRow, tempStr)
            Else
                markingRow = Utils.CellSrcLetter & vals.row
                markingColor = 0
            End If
            
            ' ##Mark existing file in default color
            Call MarkCells(shtActive.Range(markingRow), markingColor)
        End If
    Next
 
End Sub

Private Sub FixFileDirectory()
    'filePath As String
    Dim splitedStr() As String
    Dim combinedPath As String
    
    Dim temp As String
    temp = "C:\Users\Yuval\Documents\Projects\YuvalProject\Yuki\yuvaley.txt"
    
    splitedStr = Split(temp, "\")

    For i = LBound(splitedStr) To UBound(splitedStr) - 2
        combinedPath = combinedPath & splitedStr(i) & "\"
        
        'If Not DirectoryExists(combinedPath) Then
        '    MkDir (combinedPath)
        'End If
    Next i
    
    MkDir (combinedPath)
    
     combinedPath = combinedPath & splitedStr(UBound(splitedStr) - 1) & "\"
      MkDir (combinedPath)
   
End Sub
