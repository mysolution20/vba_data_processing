Option Explicit

Private openingBalance_ As Currency
Private depositoryName_ As String
Private cashDeposit_ As Currency
Private cashWithdrawal_ As Currency

Public Property Let letOpeningBalance(ByVal openingBalance As Currency)

    openingBalance_ = openingBalance

End Property

Public Property Get getOpeningBalance() As Currency

    getOpeningBalance = openingBalance_

End Property


Public Property Let letDepositoryNameId(ByVal depositoryName As String)

    depositoryName_ = depositoryName
        
End Property

Public Property Get getDepositoryNameId() As String

    getDepositoryNameId = depositoryName_
        
End Property

Public Property Let letCashDeposit(ByVal cashDeposit As Currency)

    cashDeposit_ = cashDeposit

End Property

Public Property Get getCashDeposit() As Currency

    getCashDeposit = cashDeposit_

End Property

Public Property Let LetCashWithdrawal(ByVal cashWithdrawal As Currency)

    cashWithdrawal_ = cashWithdrawal

End Property

Public Property Get getCashWithdrawal() As Currency

    getCashWithdrawal = cashWithdrawal_

End Property
