Attribute VB_Name = "CDA_MoonSort"
Sub MoonSort()
    Set rng_temp = MyWB.Worksheets("Moonspense")
    
    rng_temp.Range("A3:F" & (rng_temp.Cells(Rows.Count, "A").End(xlUp).Row)).Sort key1:=rng_temp.Range("E3"), order1:=xlAscending, _
    key2:=rng_temp.Range("D3"), order2:=xlAscending
    
End Sub
