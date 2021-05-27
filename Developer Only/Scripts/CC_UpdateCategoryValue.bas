Attribute VB_Name = "CC_UpdateCategoryValue"
Sub UpdateCategoryValue()
    For iRow = MinRowCat To rng_his.Cells(Rows.Count, "M").End(xlUp).Row
        Category = CStr(rng_his.Cells(iRow, "L").Value)
        If Not HisCatVal.exists(Category) Then
            HisCatVal.Add Category, rng_his.Cells(iRow, "M").Value
        Else
            HisCatVal(Category) = rng_his.Cells(iRow, "M").Value
        End If
    Next
End Sub
