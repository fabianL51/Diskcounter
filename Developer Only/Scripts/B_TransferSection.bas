Attribute VB_Name = "B_TransferSection"
Sub TransferSection()

    'Variable
    Dim overviewLast As String
    overviewLast = "Overview" + CStr(Year(Date) - 1)
    
    'SET RANGES FOR THIS YEAR FOR TRANSFER
    Set rng_newhis = MyWB.Worksheets(HistoryNow).Range("A:Q")
    Set rng_newover = MyWB.Worksheets(OverviewNow).Range("A:Q")
    
    
    If IsExistSheet(overviewLast) = True Then 'IF WORKSHEET FROM LAST YEAR EXISTS
        historyLast = "History" + CStr(Year(Date) - 1)
        'SET RANGES FOR LAST YEAR
        Set rng_his = MyWB.Worksheets(historyLast).Range("A:Q")
        Set rng_over = Worksheets(overviewLast).Range("A:Q") 'TRANSFER BALANCE FREE CASH FLOW IN OVERVIEW
        Set cutter_his = MyWB.Worksheets(historyLast)
        Set overrange = MyWB.Worksheets(overviewLast)
        
        MinRowCat = 7
        
        'SET DYNAMIC DICTIONARIES
        Set HisCatDict = CreateObject("Scripting.Dictionary")
        Set HisCatVal = CreateObject("Scripting.Dictionary")
        Set OldNewBank = CreateObject("Scripting.Dictionary")
        
        'CALCULATE FOR THE DECEMBER LAST YEAR
        Call InitializeBank
        Call UpdateCategoryValue
        Call ProcessOverview(1, 12)
        
        'SET RANGES FOR THIS YEAR FOR TRANSFER
        Set rng_newhis = MyWB.Worksheets(HistoryNow).Range("A:Q")
        Set rng_newover = MyWB.Worksheets(OverviewNow).Range("A:Q")
        
        For iRow = 10 To 11 'TRANSFER INCOME BALANCE FOR ALL TIERS FROM OVERVIEW LAST YEAR TO OVERVIEW THIS YEAR: Q10 This Year = Q15 Last Year
            rng_newover.Cells(iRow, "Q").Value = rng_over.Cells(iRow + 5, "Q").Value
        Next
        'SET REST COLLECTOR TO ZERO
        rng_newover.Cells(12, "Q").Value = 0
        
        For iRow = 2 To 4
        Next
        
        For iRow = 26 To 28 'TRANSFER BANK BALANCE
            'TRANSFER BANK BALANCE FROM OVERVIEW LAST YEAR TO HISTORY THIS YEAR: O2,P2 in History = P26,R26 in Overview
            rng_newhis.Cells(iRow - 24, "P").Value = rng_over.Cells(iRow, "R").Value
            rng_newhis.Cells(iRow - 24, "O").Value = rng_over.Cells(iRow, "P").Value
            
            Set oversheet = MyWB.Worksheets(OverviewNow)
            oversheet.Cells(iRow, "Q").Formula = "=" & HistoryNow & "!P" & iRow - 24
            oversheet.Cells(iRow, "P").Formula = "=" & HistoryNow & "!O" & iRow - 24
        Next
    Else
        'START BALANCE FOR ALL TIERS & REST SET TO ZERO
        rng_newover.Cells(10, "Q").Value = 0
        rng_newover.Cells(11, "Q").Value = 0
        rng_newover.Cells(12, "Q").Value = 0
    End If

    Transfer = 1
    
End Sub
