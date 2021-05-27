Attribute VB_Name = "CAAA_ReplaceOldBankName"
Sub ReplaceOldBankName()
    
    If OldNewBank.Count > 0 Then
    
        For iRowBank = 4 To rng_his.Cells(Rows.Count, "A").End(xlUp).Row
            StrDet = CStr(rng_his.Cells(iRowBank, "D").Value)
            For Each OldBankName In OldNewBank.keys
                'MsgBox (OldNewBank(OldBankName))
                If InStr(1, StrDet, OldBankName) Then
                    StrDet = Replace(StrDet, OldBankName, OldNewBank(OldBankName))
                End If
            Next
            rng_his.Cells(iRowBank, "D").Value = StrDet
            BankName = CStr(rng_his.Cells(iRowBank, "F").Value)
            If OldNewBank.exists(BankName) Then rng_his.Cells(iRowBank, "F").Value = OldNewBank(BankName)
            
            For iRowMoon = 3 To rng_moon.Cells(Rows.Count, "A").End(xlUp).Row
                OldBank = CStr(rng_moon.Cells(iRowMoon, "F").Value)
                If OldNewBank.exists(OldBank) Then rng_moon.Cells(iRowMoon, "F").Value = OldNewBank(CStr(rng_moon.Cells(iRowMoon, "F").Value))
            Next
            
        Next
    End If
    

End Sub
