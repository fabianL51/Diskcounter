Attribute VB_Name = "CD_MonthlyExpenseAdder"
Sub MonthlyExpenseAdder()
    
    Dim iRow_his As Integer
    
    iRow_his = rng_his.Cells(Rows.Count, "A").End(xlUp).Row + 1
    
    'CHECK MONTHLY EXPENSES
    Do While rng_moon.Cells(3, "E").Value = "DUE" 'IF THERE ARE MONTHLY PAYMENTS STILL DUE TO BE PAID
            'IF DUE DATE ALREADY PASSED
            If rng_moon.Cells(3, "E").Value = "DUE" And Day(Date) >= rng_moon.Cells(3, "D").Value Then
                    rng_his.Cells(iRow_his, "A").Value = DateSerial(Year(Now), Month(Now), rng_moon.Cells(3, "D").Value)
                    rng_his.Cells(iRow_his, "C").Value = rng_moon.Cells(3, "B").Value
                    rng_his.Cells(iRow_his, "D").Value = rng_moon.Cells(3, "A").Value
                    rng_his.Cells(iRow_his, "E").Value = rng_moon.Cells(3, "C").Value
                    rng_his.Cells(iRow_his, "F").Value = rng_moon.Cells(3, "F").Value
                rng_moon.Cells(3, "E").Value = "PAID"
                iRow_his = iRow_his + 1
            Call MoonSort
            Else
                Exit Do
            End If
   Loop
   Call UpdateMoonDictionary
    

    
    
    
End Sub
