Attribute VB_Name = "A_Start"
Public OverviewNow, HistoryNow As String
Public rng_his, rng_over, rng_nancy, rng_moon As Range
Public MyWB As Workbook
Public IncomeDict, ExpenseDict, MoonPosDict, FFDict, BalDict, BankDict, BankBalDict, CatCountDict As Object
Public HisCatDict, HisCatVal, OldNewBank As Object
Public CurrentMoon, MinRowCat, MaxRowInc, MaxRowExp, MinRowHis As Integer
Public IncCatFull, ExpCatFull As Integer
Public MonthNow As Integer
Public Transfer, BankChanged As Integer




Sub Start()
    
    'DEFINE IMPORTANT VARIABLES
    CurrentYear = CStr(Year(Date))
    OverviewNow = "Overview" + CurrentYear
    HistoryNow = "History" + CurrentYear
    Set MyWB = ThisWorkbook
    
    'SET RANGES
    Set rng_moon = MyWB.Worksheets("Moonspense").Range("A:Q")
    Set rng_nancy = MyWB.Worksheets("Nancyplan").Range("A:Q")
    
    Call UpdateMoonDictionary
    
    Transfer = 0
    'Check Worksheet History this year
    If IsExistSheet(HistoryNow) = False Then 'if Worksheet doesn't exist
        'CREATE SHEETS
        Sheets("Overview_Template").Copy after:=Sheets("LemonTree")
        ActiveSheet.Name = OverviewNow
        Sheets("History_Template").Copy after:=Sheets(OverviewNow)
        ActiveSheet.Name = HistoryNow
        
        'CHECK TRANSFER SECTION
        Call TransferSection
    End If
    
    
    'SET RANGES FOR ACTUAL CASE
    Set rng_his = MyWB.Worksheets(HistoryNow).Range("A:Q")
    Set rng_over = MyWB.Worksheets(OverviewNow).Range("A:S")
    
    For iRow = 26 To 28 'TRANSFER BANK BALANCE

        rng_over.Cells(iRow, "Q").Formula = "=" & HistoryNow & "!P" & iRow - 24
        rng_over.Cells(iRow, "P").Formula = "=" & HistoryNow & "!O" & iRow - 24
    Next
    
    Set OldNewBank = CreateObject("Scripting.Dictionary")
    
    BankChanged = 0
    
    Call Initializer
    
    Call Updater
    
    'Worksheets(OverviewNow).Activate
    Worksheets(HistoryNow).Activate
    'Worksheets("Moonspense").Activate
    
    
    Shadowcounter.Show
    
    

    
    
End Sub
