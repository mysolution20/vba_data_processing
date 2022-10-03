Option Explicit

Public Sub bankAccountInformation()

    Dim bankOperation As New clsBankBranch

    With bankOperation
       
        .letSheet = shBankOperation
        
        .letAddress = "A1"
         
        .letNameOfShape = "infoShape"
        
        .mDepositaryInfo
             
     End With
     
End Sub
