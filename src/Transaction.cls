VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pAmount As Currency
Private pCurrencyCode As String
Private pAccountingDate As Date
Private pDescription As String
Private pTransactionCode As String
Private pTransactionType As String

Public Property Get amount() As Currency
    amount = pAmount
End Property

Public Property Let amount(value As Currency)
    pAmount = value
End Property

Public Property Get currencyCode() As String
    currencyCode = pCurrencyCode
End Property

Public Property Let currencyCode(value As String)
    pCurrencyCode = value
End Property

Public Property Get accountingDate() As Date
    accountingDate = pAccountingDate
End Property

Public Property Let accountingDate(value As Date)
    pAccountingDate = value
End Property

Public Property Get description() As String
    description = pDescription
End Property

Public Property Let description(value As String)
    pDescription = value
End Property

Public Property Get transactionCode() As String
    transactionCode = pTransactionCode
End Property

Public Property Let transactionCode(value As String)
    pTransactionCode = value
End Property

Public Property Get transactionType() As String
    transactionType = pTransactionType
End Property

Public Property Let transactionType(value As String)
    pTransactionType = value
End Property
