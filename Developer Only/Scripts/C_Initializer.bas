Attribute VB_Name = "C_Initializer"
Public Reupdate As Integer, cutter_his As Object, overrange As Object
Sub Initializer()

    Set cutter_his = MyWB.Worksheets(HistoryNow)
    Set overrange = MyWB.Worksheets(OverviewNow)
    
    Reupdate = 0
    
    
    'SET INTEGER LIMITS
    MinRowCat = 7
    
    Call CheckBank

    
    MinRowHis = rng_his.Cells(4, "M").Value
    
    Call CheckValidMoonspense
    
    'SET DYNAMIC DICTIONARIES
    Set HisCatDict = CreateObject("Scripting.Dictionary")
    Set HisCatVal = CreateObject("Scripting.Dictionary")
    
    Call UpdateCategoryValue

End Sub
