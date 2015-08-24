Attribute VB_Name = "ConfigurationModule"
Private Str As String

'Exstracts the value of 'directory pattern' from configuration sheet
Public Function GetSourceDirectoryCell() As String
    Str = shtConfig.Cells(1, 2).value
    GetSourceDirectoryCell = Str
End Function

'Exstracts the value of 'source cell letter' from configuration sheet
Public Function GetSourceCellLetter() As String
    Str = shtConfig.Cells(2, 2).value
    GetSourceCellLetter = Str
End Function

'Exstracts the value of 'directory pattern' from configuration sheet
Public Function GetDirectoryPattern() As String
    Str = shtConfig.Cells(3, 2).value
    GetDirectoryPattern = Str
End Function

'Exstracts the value of 'srcStartCell' from configuration sheet
Public Function GetSrcStartCell() As String
    Str = shtConfig.Cells(4, 2).value
    GetSrcStartCell = Str
End Function

'Exstracts the value of 'desStartCell' from configuration sheet
Public Function GetDesStartCell() As String
    Str = shtConfig.Cells(5, 2).value
    GetDesStartCell = Str
End Function


'Exstracts the value of 'destination cell letter' from configuration sheet
Public Function GetDestinationCellLetter() As String
    Str = shtConfig.Cells(6, 2).value
    GetDestinationCellLetter = Str
End Function


'Exstracts the value of 'status cell letter' from configuration sheet
Public Function GetStatusCellLetter() As String
    Str = shtConfig.Cells(7, 2).value
    GetStatusCellLetter = Str
End Function
