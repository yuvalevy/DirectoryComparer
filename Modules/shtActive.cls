VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtActive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CopyButton_Click()
    Call LoggingModule.SetDebugMode(False)
    'Call SheetModule.EditingSheet(False)
     
    Call BusinessModule.CopyFiles

    Call FixAllButtons
    'Call SheetModule.EditingSheet(True)
End Sub

Private Sub RedButton_Click()
    Call LoggingModule.SetDebugMode(False)
    Call SheetModule.EditingSheet(False)
     
    Call BusinessModule.DeleteRedLines
  
    Call SheetModule.FixAllButtons
    Call SheetModule.EditingSheet(True)
End Sub

Private Sub CalcButton_Click()
    Call LoggingModule.SetDebugMode(False)
    Call SheetModule.EditingSheet(False)
     
    Call BusinessModule.AddAndMarkFiles
    
    Call SheetModule.FixAllButtons
    Call SheetModule.EditingSheet(True)
End Sub

Private Sub ArrangeButton_Click()
    Call LoggingModule.SetDebugMode(False)
    Call SheetModule.EditingSheet(False)
    
    Call BusinessModule.DeleteEmptyRows
        
    Call SheetModule.FixAllButtons
    Call SheetModule.EditingSheet(True)
End Sub
