Attribute VB_Name = "DAB_CalculateSavingPlan"
Sub CalculateSavingPlan(iColumn As Integer)
    
    iRow_Income = rng_over.Cells(4, "Q") + 2
    iRow_T1 = rng_over.Cells(5, "Q")
    iRow_T2 = iRow_T1 + 1
    iRow_RC = iRow_T2 + 1
    
            
    'CHECK THE NET INCOME THIS MONTH
    Income = rng_over.Cells(iRow_Income, iColumn).Value
    If Income > 0 Then 'IF POSITIVE NET INCOME
        'IF FED IS NEGATIVE THEN MOBILIZE INCOME TO REVIVE T1
        If rng_over.Cells(19, "Q") < 0 Then
            rng_over.Cells(iRow_T1, iColumn).Value = Income
        'IF FED IS POSITIIVE
        Else
            'INCOME HIGHER AS IN FINANCIAL PLANNER
            If Income >= rng_nancy.Cells(3, "B").Value Then
                'COLLECT REST IS CHOSEN
                If rng_nancy.Cells(5, "D").Value = "Collect Rest" Then  'NET INCOME AT LEAST LIKE EXPECTED THEN TRANSFER THE AMOUNT
                    rng_over.Cells(iRow_T1, iColumn).Value = rng_nancy.Cells(7, "B").Value
                    rng_over.Cells(iRow_T2, iColumn).Value = rng_nancy.Cells(8, "B").Value
                'NO REST IS CHOSEN
                ElseIf rng_nancy.Cells(5, "D").Value = "No Rest" Then 'USE THE PERCENTAGE
                    rng_over.Cells(iRow_T1, iColumn).Value = rng_nancy.Cells(7, "C").Value * Income / 100
                    rng_over.Cells(iRow_T2, iColumn).Value = rng_nancy.Cells(8, "C").Value * Income / 100
                End If
            'INCOME LESS THAN IN FINANCIAL PLANNER
            ElseIf Income < rng_nancy.Cells(3, "B").Value Then 'FIFTY FIFTY
                rng_over.Cells(iRow_T1, iColumn).Value = 0.5 * Income
                rng_over.Cells(iRow_T2, iColumn).Value = 0.5 * Income
            End If
        End If
    Else 'NEGATIVE NET INCOME
        RestCollector = rng_over.Cells(17, "Q").Value
        Tier2 = rng_over.Cells(16, "Q").Value
        Tier1 = rng_over.Cells(15, "Q").Value
        
        If RestCollector >= Abs(Income) Then  'CHECK IF REST COLLECTOR CAN PAY THE LOSS
            rng_over.Cells(iRow_RC, iColumn).Value = Income 'ALREADY NEGATIVE
            rng_over.Cells(iRow_T2, iColumn).Value = 0
            rng_over.Cells(iRow_T1, iColumn).Value = 0
        ElseIf RestCollector + Tier2 >= Abs(Income) Then 'SECOND TIER EMERGENCY MONEY
            rng_over.Cells(iRow_RC, iColumn).Value = -RestCollector 'USE ALL BALANCE OF REST COLLECTOR
            rng_over.Cells(iRow_T2, iColumn).Value = Income - RestCollector 'REST PAID BY TIER2
            rng_over.Cells(iRow_T1, iColumn).Value = 0
        'USE OTHER BEFORE USE SAVINGS
        ElseIf RestCollector + Tier2 + Tier1 >= Abs(Income) Then
            rng_over.Cells(iRow_RC, iColumn).Value = -RestCollector 'USE ALL BALANCE OF REST COLLECTOR
            rng_over.Cells(iRow_T2, iColumn).Value = -Tier2  'USE ALL TIER2
            rng_over.Cells(iRow_T1, iColumn).Value = Income - RestCollector - Tier2 'REST PAID BY OTHER
        Else 'SAVING HAS TO GO DOWN THE HILL
            rng_over.Cells(iRow_RC, iColumn).Value = -RestCollector 'USE ALL BALANCE OF REST COLLECTOR
            rng_over.Cells(iRow_T2, iColumn).Value = -Tier2  'USE ALL TIER2
            'REST PAID BY SAVING
            rng_over.Cells(iRow_T1, iColumn).Value = Income - RestCollector - Tier2 - Tier1 'TIER1 INTO NEGATIVES
        End If
    End If
            
            
                    
            
    
End Sub
