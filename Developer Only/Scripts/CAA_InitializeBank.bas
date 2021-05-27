Attribute VB_Name = "CAA_InitializeBank"
Sub InitializeBank()
    
    Dim iRow As Integer
    
    Call InitializeBankDictionary
    
    If Trim(rng_his.Cells(3, "A").Value & vbNullString) = vbNullString Then
        rng_his.Cells(3, "D").Value = "Starting Point"
        rng_his.Cells(3, "A").Value = DateSerial(Year(Now), Month(Now), 1)
        rng_his.Cells(4, "M").Value = rng_his.Cells(4, "M").Value
        rng_his.Cells(2, "M").Value = rng_his.Cells(Rows.Count, "A").End(xlUp).Row + 1
        
        For iBank = 2 To 4 'ROW 2 = COLUMN 7
            BankName = CStr(rng_his.Cells(iBank, "O").Value)
            If rng_his.Cells(2, iBank + 5).Value <> "Inactive" Then
                rng_his.Cells(3, iBank + 5).Value = rng_his.Cells(iBank, "P").Value
            End If
        Next
    Else
        For iBank = 2 To 4 'ROW 2 = COLUMN 7
            BankName = CStr(rng_his.Cells(iBank, "O").Value)
            If BankDict.exists(BankName) Then
                iColumn = BankDict(BankName)
                If rng_his.Cells(3, iColumn).Value <> rng_his.Cells(iBank, "P").Value Then
                    rng_his.Cells(3, iColumn).Value = CCur(rng_his.Cells(iBank, "P").Value)
                    Reupdate = 1
                End If
            End If
        Next
    End If
    
    
    
    If Reupdate = 1 Then
        iRow = 3
    ElseIf Reupdate = 0 Then
        iRow = 0
    End If
    
    Call ReplaceOldBankName
End Sub
