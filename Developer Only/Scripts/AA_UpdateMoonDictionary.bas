Attribute VB_Name = "AA_UpdateMoonDictionary"
Sub UpdateMoonDictionary()
    
    
    Set MoonPosDict = CreateObject("Scripting.Dictionary")
    'MOON
    CurrentMoon = rng_moon.Cells(Rows.Count, "A").End(xlUp).Row + 1
    Set rng_temp = Worksheets("Moonspense")
    If CurrentMoon > 2 Then
        rng_temp.Range("A3:F" & CurrentMoon).Sort key1:=rng_temp.Range("E3"), order1:=xlAscending, _
        key2:=rng_temp.Range("D3"), order2:=xlAscending
         
        For iRow_moon = CurrentMoon - 1 To 3 Step -1
            If Not IsEmpty(rng_moon.Cells(RowIndex:=iRow_moon, ColumnIndex:="A").Value) Then
                Component = CStr(rng_moon.Cells(RowIndex:=iRow_moon, ColumnIndex:="A").Value)
                If Not MoonPosDict.exists(Component) Then
                    MoonPosDict.Add Component, iRow_moon
                End If
            End If
        Next
    End If

End Sub
