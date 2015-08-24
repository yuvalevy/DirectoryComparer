Attribute VB_Name = "BusinessModule"


' ############## BUSINESS FUNCTIONs & SUBs

' ##Iterate all directory & files and mark them
Public Sub AddAndMarkFiles()
    
    Dim srcRangeValue As String
    
    srcRangeValue = Utils.GetCurrentActiveRange
    
    If IOModule.ProjectDirectoryExists Then
        ' ##Marking all cells in color red (3)
        Call SheetModule.MarkCells(shtActive.Range(srcRangeValue), 3)
        
        ' ##Starting to check all files in directory and subdirectory
        Call IOModule.AddAndMarkFiles
    Else
        MsgBox "Source folder does not exists"
    End If
    
End Sub

' ##Copy all files from sheet
' ## Source copied to destination
Public Sub CopyFiles()
    
    Dim rwIndex As Long
    Call Utils.SetLetters
    
    For rwIndex = 4 To IOModule.GetLastRow
        If Not SheetModule.IsRowMarkedRed(rwIndex) Then
            
            Call IOModule.CopySingleFile(rwIndex)
           
        End If
    Next

End Sub

' ##Remove all rows which their source cell is colored in red
Sub DeleteRedLines()

    Dim rowCount As Long
    Dim rwIndex As Long
    Dim currentCell As String
    
    Call Utils.SetLetters
    
    rowCount = SheetModule.GetLastRow
      
    For rwIndex = rowCount To 4 Step -1
        currentCell = Utils.CellSrcLetter & rwIndex
        If SheetModule.IsRowMarkedRed(rwIndex) Then
            Call SheetModule.DeleteRow(rwIndex)
        End If
    Next
    
End Sub

' ##Delted all rows where source & destination cells are empty
Sub DeleteEmptyRows()
   
    Dim rowCount As Long
    Dim rwIndex As Long
    
    rowCount = SheetModule.GetLastRow
    
    Call LoggingModule.MessageBox("Last Row: " & rowCount)
    
    For rwIndex = rowCount To 4 Step -1
        wholeRowEmpty = True
       
        If shtActive.Cells(rwIndex, 2) <> "" Or shtActive.Cells(rwIndex, 3) <> "" Then
            wholeRowEmpty = False
        End If
    
        
        If wholeRowEmpty = True Then
            Call SheetModule.DeleteRow(rwIndex)
        End If
    
    Next

End Sub
