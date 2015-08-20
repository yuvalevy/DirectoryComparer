Attribute VB_Name = "LoggingModule"
Private InDebugMode As Boolean

Public Sub MessageBox(messageToDesplay As String)
    If InDebugMode = True Then
        MsgBox messageToDesplay
    End If
End Sub

Public Sub SetDebugMode(mode As Boolean)
    InDebugMode = mode
End Sub
