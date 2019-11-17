VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As String
Private pAccountNumber As String
Private pName As String
Private pBalance As Currency
Private pBalanceCurrentyCode As String
Private pavailableBalance As Currency
Private pAvailableBalanceCurrentyCode As String
Private pOwnerFirstName As String
Private pOwnerLastName As String

Public Property Get id() As String
    id = pId
End Property

Public Property Let id(value As String)
    pId = value
End Property

Public Property Get accountNumber() As String
    accountNumber = pAccountNumber
End Property

Public Property Let accountNumber(value As String)
    pAccountNumber = value
End Property

Public Property Get name() As String
    name = pName
End Property

Public Property Let name(value As String)
    pName = value
End Property

Public Property Get balance() As Currency
    balance = pBalance
End Property

Public Property Let balance(value As Currency)
    pBalance = value
End Property

Public Property Get balanceCurrentyCode() As String
    balanceCurrentyCode = pBalanceCurrentyCode
End Property

Public Property Let balanceCurrentyCode(value As String)
    pBalanceCurrentyCode = value
End Property

Public Property Get availableBalance() As Currency
    availableBalance = pavailableBalance
End Property

Public Property Let availableBalance(value As Currency)
    pavailableBalance = value
End Property

Public Property Get availableBalanceCurrentyCode() As String
    availableBalanceCurrentyCode = pAvailableBalanceCurrentyCode
End Property

Public Property Let availableBalanceCurrentyCode(value As String)
    pAvailableBalanceCurrentyCode = value
End Property

Public Property Get ownerFirstName() As String
    ownerFirstName = pOwnerFirstName
End Property

Public Property Let ownerFirstName(value As String)
    pOwnerFirstName = value
End Property

Public Property Get ownerLastName() As String
    ownerLastName = pOwnerLastName
End Property

Public Property Let ownerLastName(value As String)
    pOwnerLastName = value
End Property
