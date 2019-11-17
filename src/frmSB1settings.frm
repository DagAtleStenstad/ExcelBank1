VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSB1settings 
   Caption         =   "Innstillinger"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8220.001
   OleObjectBlob   =   "frmSB1settings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSB1settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
    End
End Sub

Private Sub cmdOk_Click()
    SaveSetting "ExcelBank1", "Settings", "Token", txtToken
    SaveSetting "ExcelBank1", "Settings", "TokenDeveloper", txtTokenDeveloper
    SaveSetting "ExcelBank1", "Settings", "DeveloperMode", chkDeveloperMode
    
    Unload Me
End Sub

Private Sub lblExcelBank1_Click()
 ActiveWorkbook.FollowHyperlink "https://bitbucket.org/Stenstad/excelbank1/src/master/"
End Sub

Private Sub UserForm_Initialize()
    txtToken = GetSetting("ExcelBank1", "Settings", "Token")
    txtTokenDeveloper = GetSetting("ExcelBank1", "Settings", "TokenDeveloper")
    chkDeveloperMode = GetSetting("ExcelBank1", "Settings", "DeveloperMode")
    
    If chkDeveloperMode Then frmSB1settings.MultiPage1.value = 1
End Sub
