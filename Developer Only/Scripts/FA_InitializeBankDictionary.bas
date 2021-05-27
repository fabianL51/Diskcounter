Attribute VB_Name = "FA_InitializeBankDictionary"
Sub InitializeBankDictionary()
    
    
    
    'BANK COLUMN FOR BALANCE
    Set BankDict = CreateObject("Scripting.Dictionary")
    
    For iColumn = 7 To 9
        Bank = CStr(rng_his.Cells(2, iColumn).Value)
        If Bank <> "Inactive" Then
            BankDict.Add Bank, iColumn
        End If
    Next
    
    
End Sub
