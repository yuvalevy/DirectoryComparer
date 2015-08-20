Attribute VB_Name = "Utils"
Public WorkingDirectory As String
Public CellSrcLetter As String
Public CellDestLetter As String
Public CellStatusLetter As String

' ##Sets all cell letter value
Public Sub SetLetters()
     CellSrcLetter = ConfigurationModule.GetSourceCellLetter
     CellDestLetter = ConfigurationModule.GetDestinationCellLetter
     CellStatusLetter = ConfigurationModule.GetStatusCellLetter
End Sub

' ##Get the range of the active cells
Public Function GetCurrentActiveRange() As String
    Dim srcStartCell As String
    Dim desStartCell As String
    Dim srcEndCell As String
    Dim srcRangeValue As String
    
    ' ##Get Configuration
    Call SetLetters
    srcStartCell = ConfigurationModule.GetSrcStartCell
    desStartCell = ConfigurationModule.GetDesStartCell
    
    LastWritenInRow = SheetModule.GetLastRow
    srcEndCell = Utils.CellSrcLetter & LastWritenInRow
    srcRangeValue = srcStartCell & ":" & srcEndCell

    GetCurrentActiveRange = srcRangeValue
End Function

' ##Returns the status index
' note: lookto Function ChecksBeforeCopy
Public Function GetStatusResult(rwIndex As Long) As Integer

    Dim sttsCell As String
    sttsCell = CellStatusLetter & rwIndex
   
    Select Case shtActive.Range(sttsCell).value
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
