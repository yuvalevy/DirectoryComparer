Attribute VB_Name = "IOModule"
Private filesPattern As String

' ##Checks if 'directoryPath' is valid
Public Function ProjectDirectoryExists() As Boolean

    Dim directoryCell As String
    Dim directoryPath As String
    
    directoryCell = ConfigurationModule.GetSourceDirectoryCell
    filesPattern = ConfigurationModule.GetDirectoryPattern
    
    ' ##Get directory path from sheet
    directoryPath = Range(directoryCell).value
    directoryPath = IOModule.FixDirectoryPath(directoryPath)
    
    temp = Dir(directoryPath, vbDirectory)
    
    If temp <> "" Then
        ProjectDirectoryExists = True
    Else
        ProjectDirectoryExists = False
    End If
    
End Function

' ##Checks if last char in directory is '\'
' ##If not, it adds it
Private Function FixDirectoryPath(directoryPath As String) As String
    
    If Right(directoryPath, 1) <> "\" Then
        directoryPath = directoryPath & "\"
    End If
    
    FixDirectoryPath = directoryPath
    
End Function

' ##List of validations made on the src & des files
Public Function ChecksBeforeCopy(statusIndex As Integer, src As String, des As String, row As Long) As Boolean
            
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
            ' ##Checks if source file not exists
            If Dir(src) = "" Then
                MessageBox ("Source file does not exists in row " & row)
                ChecksBeforeCopy = False
            End If
        Case Else
            ChecksBeforeCopy = True
        
    End Select
End Function

' ##Directory.GetFiles(path)
' note: pattern included in path parameter
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

' ################## Marking

' ##Marks new file with green. Marks old files with default color.
Private Sub MarkFiles(subDirectory As String, srcRangeValue As String)
    
    Dim markingRow As String
    Dim markingColor As Integer
     
    fileName = GetFiles(subDirectory)
    
    ' TODO: add if the first one is empy..
    For Each fn In fileName
        If fn <> "" Then
            Set vals = Range(srcRangeValue).Find(fn)
            
            If (vals Is Nothing) Then
                
                ' ##Adds new file to the last row
                lastWritenInRow = lastWritenInRow + 1
                markingRow = cellSrcLetter & lastWritenInRow
                markingColor = 4
                
                sCall SheetModule.SetCellValue(markingRow, fn)
            Else
                markingRow = cellSrcLetter & vals.row
                markingColor = 0
            End If
            
            ' ##Mark existing file in default color
            Call MarkCells(Range(markingRow), markingColor)
        End If
    Next
 
End Sub

