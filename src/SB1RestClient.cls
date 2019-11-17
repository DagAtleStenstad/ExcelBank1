VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SB1RestClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private SB1ApiUrl As String
Private SB1Token As String

Private Sub Class_initialize()
    
    Do Until SB1Token <> vbNullString
        
        If GetSetting("ExcelBank1", "Settings", "DeveloperMode") Then
            SB1Token = GetSetting("ExcelBank1", "Settings", "TokenDeveloper")
            SB1ApiUrl = "https://developer-api.sparebank1.no/"
        Else
            SB1Token = GetSetting("ExcelBank1", "Settings", "Token")
            SB1ApiUrl = "https://api.sparebank1.no/"
        End If
        
        If SB1Token = vbNullString Then
            frmSB1settings.Show
        End If
    Loop
    
End Sub

Private Function SB1Service(methods As String, url As String, Optional Data As String) As String
    
    Dim SB1Client As Object
    Dim SB1Response As String

    Set SB1Client = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    With SB1Client
        .Open methods, SB1ApiUrl & url, False
        .setRequestHeader "Authorization", "Bearer " & SB1Token
        .Send Data
           
        SB1Response = .ResponseText
    End With
    
    Set SB1Client = Nothing
    
    SB1Service = SB1Response
End Function

Public Function getAllPersonalAccounts() As Collection

    Dim SB1Response As String
    Dim json As Object
    Dim item As Variant
    Dim account As account
    Dim accounts As New Collection

    SB1Response = SB1Service("GET", "open/personal/banking/accounts/all")
    
    Set json = JsonConverter.ParseJson(SB1Response)
    
    For Each item In json("accounts")
        Set account = New account
        
        account.id = item("id")
        account.accountNumber = item("accountNumber")("value")
        account.name = item("name")
        account.balance = item("balance")("amount")
        account.balanceCurrentyCode = item("balance")("currencyCode")
        account.availableBalance = item("availableBalance")("amount")
        account.availableBalanceCurrentyCode = item("availableBalance")("currencyCode")
        account.ownerFirstName = item("owner")("firstName")
        account.ownerLastName = item("owner")("lastName")
        
        accounts.Add account
    Next
    
    Set getAllPersonalAccounts = accounts

End Function

Public Function getAccountTransactions(accountNumber) As Collection

    Dim SB1Response As String
    Dim json As Object
    Dim item As Variant
    Dim transaction As transaction
    Dim transactions As New Collection

    SB1Response = SB1Service("GET", "open/personal/banking/accounts/" & getAccountID(accountNumber) & "/transactions")
    
    Set json = JsonConverter.ParseJson(SB1Response)
    
    For Each item In json("transactions")
        Set transaction = New transaction
        
        transaction.amount = item("amount")("amount")
        transaction.currencyCode = item("amount")("currencyCode")
        transaction.accountingDate = item("accountingDate")
        transaction.description = item("description")
        transaction.transactionCode = item("transactionCode")
        transaction.transactionType = item("transactionType")

        transactions.Add transaction
    Next
    
    Set getAccountTransactions = transactions

End Function

Private Function getAccountID(ByVal accountNumber As String) As String

    Dim account As account
    Dim accounts As New Collection

    Set accounts = getAllPersonalAccounts
     
    For Each account In accounts
        If account.accountNumber = accountNumber Then
            getAccountID = account.id
            Exit For
        End If
    Next
       
End Function
