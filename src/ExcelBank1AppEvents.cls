VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelBank1AppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents mxlApp As Application        'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in
Attribute mxlApp.VB_VarHelpID = -1

Public Property Set App(xlApp As Application)   'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in
    Set mxlApp = xlApp
End Property

Private Sub mxlApp_SheetCalculate(ByVal Sh As Object)
    
    If Not SB1command = "" Then
    
        Select Case SB1command
            Case "SB1getAllPersonalTransactions"
                Range(activeCellAddress).Select
                Call SB1getAllPersonalTransactions(SB1commandArg, activeCellAddress)
                
            Case "SB1openSettingsForm"
                 Range(activeCellAddress) = ""
                 Range(activeCellAddress).Select
                 frmSB1settings.Show
                 
            Case "SB1getAllPersonalAccounts"
                Range(activeCellAddress).Select
                Call SB1getAllPersonalAccounts(activeCellAddress)
        End Select
    
    End If
    
    SB1command = ""
    
End Sub
