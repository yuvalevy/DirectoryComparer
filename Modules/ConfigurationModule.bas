Attribute VB_Name = "ConfigurationModule"
Private str As String

'Exstracts the value of 'source directory cell' from configuration sheet
Public Function GetSourceDirectoryCell() As String
    str = shtConfig.Cells(1, 2).value
    GetSourceDirectoryCell = str
End Function

'Exstracts the value of 'directory pattern' from configuration sheet
Public Function GetDirectoryPattern() As String
    str = shtConfig.Cells(3, 2).value
    GetDirectoryPattern = str
End Function

'Exstracts the value of 'directory pattern' from configuration sheet
Public Function GetSourceDirectoryCell() As String
    str = shtConfig.Cells(1, 2).value
    GetSourceDirectoryCell = str
End Function

