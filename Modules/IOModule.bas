Attribute VB_Name = "IOModule"
Private FilesPattern As String
Public LastWritenInRow As Long
Public FSO As New FileSystemObject

' ##Checks if 'B2' is valid
Public Function ProjectDirectoryExists() As Boolean

    Dim directoryCell As String
    Dim directoryPath As String
    
    directoryCell = ConfigurationModule.GetSourceDirectoryCell
    FilesPattern = ConfigurationModule.GetDirectoryPattern
    
    ' ##Get directory path from sheet
    directoryPath = shtActive.Range(directoryCell).value

    If DirectoryExists(directoryPath) Then
        Utils.WorkingDirectory = directoryPath
        ProjectDirectoryExists = True
    Else
        ProjectDirectoryExists = False
    End If
 
End Function

' ##Whether a certain directory exists
Private Function DirectoryExists(directoryPath As String) As Boolean
   DirectoryExists = FSO.FolderExists(directoryPath)
End Function

Private Function FileExists(filePath As String, wishToOverride As Boolean) As Boolean
   FileExists = FSO.FileExists(filePath)
   If FileExists = True And wishToOverride = True Then
        
        Dim objFile As file
        Set objFile = FSO.GetFile(filePath)
        
        If objFile.Attributes And ReadOnly Then
            objFile.Attributes = objFile.Attributes Xor ReadOnly
        End If
   End If
   
   
End Function


Public Function GetLastRow() As String
    LastWritenInRow = SheetModule.GetLastRow
    GetLastRow = LastWritenInRow
End Function

' ##List of validations made on the src & des files
Public Function ChecksBeforeCopy(statusIndex As Integer, src As String, des As String, row As Long) As Boolean

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
            If Not FileExists(src, False) Then
                LoggingModule.MessageBox ("Source file does not exists in row " & row)
                ChecksBeforeCopy = False
            End If
        Case 4
                    
        Case Else
            ChecksBeforeCopy = True
        
    End Select
End Function

Public Sub CopySingleFile(rwIndex As Long)
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
        Call CreateFileDirectory(destinationFile)
        If FileExists(destinationFile, True) Then
            Call FSO.DeleteFile(destinationFile)
        End If
        
        Call FSO.CopyFile(sourceFile, destinationFile)
        statusIndex = 100
    End If
    Call SheetModule.WriteStatusResult(statusIndex, rwIndex)

End Sub

Public Sub AddAndMarkFiles()
    Dim objFolder As Folder
     
    Set objFolder = FSO.GetFolder(WorkingDirectory)
    
    Call AddMarkFiles(objFolder.files)
    Call IterateDirectory(objFolder)
End Sub

' ##Directory.GetFolders(path ,SearchOptions.AllDirectories)
' ##and call MarkFiles(..)
' note: recursion function
Private Sub IterateDirectory(objCurrentFolder As Folder)
    Dim i As Long
    
    Dim objFolders As Folders
    Dim objSubFolder As Folder
    
    Set objFolders = objCurrentFolder.SubFolders
    
    For Each objSubFolder In objFolders
       Call AddMarkFiles(objSubFolder.files)
       Call IterateDirectory(objSubFolder)
    Next objSubFolder
    
End Sub

' ##Marks new file with green. Marks old files with default color.
Private Sub AddMarkFiles(currentFiles As files)
    
    Dim markingCell As String
    Dim markingColor As Integer
    Dim fileName As String
    Dim currentFile As file
    Dim srcRangeValue As String
    
    srcRangeValue = Utils.GetCurrentActiveRange
     
    For Each currentFile In currentFiles
        fileName = currentFile.path
        Set vals = shtActive.Range(srcRangeValue).Find(fileName)
        
        If (vals Is Nothing) Then
            
            ' ##Adds new file to the last row
            LastWritenInRow = LastWritenInRow + 1
            markingCell = Utils.CellSrcLetter & LastWritenInRow
            markingColor = 4
            Call SheetModule.SetCellValue(markingCell, fileName)
        Else
            markingCell = Utils.CellSrcLetter & vals.row
            markingColor = 0
        End If
        
        ' ##Mark existing file in default color
        Call MarkCells(shtActive.Range(markingCell), markingColor)
        
    Next
End Sub


Private Sub CreateDirectory(directoryPath As String)

    If DirectoryExists(directoryPath) = False Then
        Dim startPoint As String
        Dim pos As Long
        
        pos = InStrRev(directoryPath, "\") - 1
        startPoint = left(directoryPath, pos)
        
        Call CreateDirectory(startPoint)
        Call MkDir(directoryPath)
    End If

End Sub
    
Private Sub CreateFileDirectory(directoryPath As String)
    
    Dim startPoint As String
    Dim pos As Long
      
    pos = InStrRev(directoryPath, "\") - 1
    startPoint = left(directoryPath, pos)
    
    Call CreateDirectory(startPoint)
    
End Sub




Public Sub test()

Dim tempPath As String
Dim temp1 As String

temp1 = "C:\tes\New.Girl.S01E01.HDTV.XviD-LOL.[VTV].avi"
tempPath = "C:\tes\New\Girl.S01E01.HDTV.XviDTV.avi"


'Call FSO.CopyFile(temp1, tempPath)

If FileExists(tempPath, True) Then
    FSO.DeleteFile (tempPath)
End If
    
    


End Sub

