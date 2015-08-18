Attribute VB_Name = "Utils"
Public WorkingDirectory As String

Public Function IsDirectoryOK() As Boolean
  
    Dim funcResult As Boolean
    
    funcResult = IOModule.DirectoryExists(directoryPath)
    IsDirectoryOK = functionResult
    
    If funcResult Then
       WorkingDirectory = directoryPath
    End If
    
End Function

