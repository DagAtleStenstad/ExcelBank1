Attribute VB_Name = "Core"
Option Explicit

Public activeCellAddress As String
Public SB1command As String
Public SB1commandArg As String

Public gclsAppEvents As ExcelBank1AppEvents     'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in

Sub Auto_Open()                                 'https://stackoverflow.com/questions/24683155/including-thisworkbook-code-in-excel-add-in
    Set gclsAppEvents = New ExcelBank1AppEvents
    Set gclsAppEvents.App = Application
End Sub

Function SB1hentAllePeronligeKontoer() As String
    activeCellAddress = ActiveCell.Address
    SB1command = "SB1getAllPersonalAccounts"
    
    SB1hentAllePeronligeKontoer = "Henter kontoer. Venligst vent..."
End Function

Function SB1hentAllePeronligeTransaksjoner(konto As String) As String
    activeCellAddress = ActiveCell.Address
    SB1command = "SB1getAllPersonalTransactions"
    SB1commandArg = konto
    
    SB1hentAllePeronligeTransaksjoner = "Henter transaksjoner. Venligst vent..."
End Function

Function SB1Innstillinger() As String
    activeCellAddress = ActiveCell.Address
    SB1command = "SB1openSettingsForm"
End Function

Public Sub SB1getAllPersonalAccounts(StartingPosition As String)

    Dim SB1Client As New SB1RestClient
    Dim accounts As Collection
    Dim account As account
    
    Set accounts = SB1Client.getAllPersonalAccounts
    
    Dim row As Integer
    row = 0
    
    Dim bdTable As New BreakDownTable
    bdTable.StartingPosition = Range(StartingPosition)
    
    bdTable.Cell(row, 0).NumberFormat = "General"
    bdTable.Cell(row, 0) = "Kontonummer"
    
    
    bdTable.Cell(row, 1).NumberFormat = "General"
    bdTable.Cell(row, 1) = "Navn"
    
    bdTable.Cell(row, 2).NumberFormat = "General"
    bdTable.Cell(row, 2) = "Saldo"
    
    
    For Each account In accounts
        row = row + 1
        bdTable.Cell(row, 0).NumberFormat = "General"
        bdTable.Cell(row, 0) = account.accountNumber
                
        bdTable.Cell(row, 1).NumberFormat = "General"
        bdTable.Cell(row, 1) = account.name
        
        bdTable.Cell(row, 2).Style = "Currency"
        bdTable.Cell(row, 2) = account.availableBalance
    Next

End Sub

Public Sub SB1getAllPersonalTransactions(account As String, StartingPosition As String)

    Dim SB1Client As New SB1RestClient
    Dim transactions As Collection
    Dim transaction As transaction
    
    Set transactions = SB1Client.getAccountTransactions(account)
    
    Dim row As Integer
    row = 0
    
    Dim bdTable As New BreakDownTable
    bdTable.StartingPosition = Range(StartingPosition)
    
    bdTable.Cell(row, 0).NumberFormat = "General"
    bdTable.Cell(row, 0) = "Dato"
        
    bdTable.Cell(row, 1).NumberFormat = "General"
    bdTable.Cell(row, 1) = "Beskrivelse"
        
    bdTable.Cell(row, 2).NumberFormat = "General"
    bdTable.Cell(row, 2) = "Beløp"
        
    For Each transaction In transactions
        row = row + 1
        bdTable.Cell(row, 1).NumberFormat = "Short Date"
        bdTable.Cell(row, 0) = transaction.accountingDate
                
        bdTable.Cell(row, 1).NumberFormat = "General"
        bdTable.Cell(row, 1) = transaction.description
                
        bdTable.Cell(row, 2).Style = "Currency"
        bdTable.Cell(row, 2) = transaction.amount
    Next

End Sub
