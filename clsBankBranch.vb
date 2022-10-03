Option Explicit

Private dcBankDepositaries_ As New Dictionary
Private sheet_ As Worksheet
Private strAddress_ As String
Private key_ As Variant
Private arr_ As Variant
Private nameOfShape_ As String

Private Enum enumHeaders

    depositoryNameId = 1
    openingBalance = depositoryNameId + 1
    cashDeposit = openingBalance + 1
    cashWithdrawal = cashDeposit + 1
    
End Enum

Public Sub mDepositaryInfo()

    mLoadArray

    mLoadDataOperation
    
    mAvailableCashInfo

End Sub

Public Property Let letSheet(ByVal sheet As Worksheet)
    
    Set sheet_ = sheet
    
End Property

Private Property Get getSheet() As Worksheet
    
    Set getSheet = sheet_
    
End Property

Public Property Let letNameOfShape(ByVal nameOfShape As String)

    nameOfShape_ = nameOfShape

End Property

Private Property Get getNameOfShape() As String
    
    getNameOfShape = nameOfShape_
    
End Property

Public Property Let letAddress(ByVal strAddress As String)

    strAddress_ = strAddress

End Property

Private Property Get getAddress() As String
    
    getAddress = strAddress_
    
End Property

Private Property Get getDepositary() As clsDepositary

    Dim depositary As New clsDepositary

    Set getDepositary = depositary

End Property

Private Sub mLoadArray()

    Dim rg As Range

    Set rg = getSheet.Range(getAddress).CurrentRegion
    
    arr_ = rg.Offset(1).Resize(rg.Rows.Count - 1)

End Sub

Private Sub mLoadDataOperation()

    Dim nextRow As Integer, previousDepositary As clsDepositary, currentDepositary As clsDepositary

    For nextRow = LBound(arr_) To UBound(arr_)
     
       Set currentDepositary = getDepositary

            With currentDepositary
            
                .letDepositoryNameId = arr_(nextRow, enumHeaders.depositoryNameId)
                .letOpeningBalance = arr_(nextRow, enumHeaders.openingBalance)
                .letCashDeposit = arr_(nextRow, enumHeaders.cashDeposit)
                .LetCashWithdrawal = arr_(nextRow, enumHeaders.cashWithdrawal)
                
            End With
            
                key_ = arr_(nextRow, enumHeaders.depositoryNameId)
    
                    If dcBankDepositaries_.Exists(key_) Then
                        
                        Set previousDepositary = dcBankDepositaries_(key_)
                        
                            With previousDepositary
                                 
                                .letOpeningBalance = .getOpeningBalance + currentDepositary.getOpeningBalance
                                .letCashDeposit = .getCashDeposit + currentDepositary.getCashDeposit
                                .LetCashWithdrawal = .getCashWithdrawal + currentDepositary.getCashWithdrawal
                                    
                            End With
                    Else
                    
                        dcBankDepositaries_.Add key_, currentDepositary
                        
                    End If
                
    Next nextRow

End Sub

Private Function fnNegativeBalancePublication(Optional depositoryName As String = "", Optional openingBalance As Currency = 0, Optional cashDeposit As Currency = 0) As String

    fnNegativeBalancePublication = "Not enough funds to pay out for " + depositoryName + vbLf + _
                                   "The maximum amount that can be collected by " + depositoryName + " is:  " + _
                                   CStr(FormatCurrency(openingBalance + cashDeposit))

End Function

Private Function fnPositiveBalancePublication(Optional depositoryName As String = "", Optional balance As Currency = 0) As String

    fnPositiveBalancePublication = "Available cash has an account for " + depositoryName + ":  " + _
                                    CStr(FormatCurrency(balance))
                            
End Function

Private Sub mAvailableCashInfo()

    Dim depositary As clsDepositary, mainShape As shape, infoShape As TextRange2, publication As String, infoLoad As String, balance As Currency

    Set mainShape = getSheet.Shapes(getNameOfShape)
    Set infoShape = mainShape.TextFrame2.TextRange.Characters
    
    With mainShape
    
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
    
    End With

    With infoShape
    
        .text = infoLoad
        .Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorText1
    
    End With
    
    
        For Each key_ In dcBankDepositaries_
        
            Set depositary = dcBankDepositaries_(key_)
                          
                With depositary
                    
                    balance = (.getOpeningBalance + .getCashDeposit - .getCashWithdrawal)
    
                        If balance < 0 Then
                             
                             publication = fnNegativeBalancePublication(.getDepositoryNameId, .getOpeningBalance, .getCashDeposit)
    
                         Else
                         
                             publication = fnPositiveBalancePublication(.getDepositoryNameId, balance)
                             
                         End If
            
                            If infoLoad = "" Then
                               
                                infoLoad = publication
                                   
                            Else
                               
                                infoLoad = infoLoad + vbLf + vbLf + publication
                                   
                            End If
                         
                 End With
                
        Next key_
    
                infoShape.text = infoLoad
                
End Sub
