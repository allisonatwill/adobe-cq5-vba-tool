Attribute VB_Name = "Module3"
Sub CheckMultiSubsets()
Attribute CheckMultiSubsets.VB_ProcData.VB_Invoke_Func = " \n14"
'Check for multiple subsets
'
Application.ScreenUpdating = False
    
    Dim SKUSubset As String
    Dim skuCells As Collection, shtExp As Worksheet, shtImp As Worksheet
    Dim rollReq As Worksheet
    Dim skuExp As Worksheet, skuImp As Worksheet, flagExp As Worksheet, flagImp As Worksheet, attExp As Worksheet
    Dim attImp As Worksheet, deactSku As Worksheet
    Dim checkMulti As Worksheet
    Dim skuCell

    Set rollReq = Sheets("Rollover Request")
    Set shtExp = Sheets("Subset Exporter")
    Set shtImp = Sheets("Subset Importer")
    Set skuExp = Sheets("SKU Exporter")
    Set skuImp = Sheets("New SKU Importer")
    Set flagExp = Sheets("SKU Flag Exporter")
    Set flagImp = Sheets("SKU Flag Importer")
    Set attExp = Sheets("Attribute Exporter")
    Set attImp = Sheets("Attribute Importer")
    Set deactSku = Sheets("Deactivate Old SKU Importer")
    Set checkMulti = Sheets("Check Multi Subset")

    Dim OldSKU As Long
    Dim NewSKU As Long

    Dim rLastSKU As Long
    Dim rLastSub As Long
    Dim rLastMulti As Long

    Dim i As Range
    Dim subRange As Range
    Dim multiRange As Range
    
    Dim cel As Variant
    Dim subsetString As String
    
rollReq.Activate
    
    rLastSKU = rollReq.Cells(Rows.Count, "A").End(xlUp).Row
    

For Each i In rollReq.Range("B2:B" & rLastSKU).Cells

    NewSKU = i.Value
    
    shtImp.Activate

    shtImp.Range("A1").AutoFilter Field:=2, Criteria1:=NewSKU
    
    rLastSub = shtImp.Cells(Rows.Count, "A").End(xlUp).Row
    
    Set subRange = shtImp.Range("A1:A" & rLastSub)
    
    subRange.Copy
    
    checkMulti.Range("A1").PasteSpecial Paste:=xlPasteValues
    
    rLastMulti = checkMulti.Cells(Rows.Count, "A").End(xlUp).Row
    
    Set multiRange = checkMulti.Range("A2:A" & rLastMulti)
    
        For Each cel In multiRange
            

            If Application.WorksheetFunction.CountIf(multiRange, cel) > 1 Then
            
            subsetString = cel
            
             MsgBox "Duplicate Subset: " + subsetString
    
            End If
    
        Next cel
    
    checkMulti.Range("A:A").Clear
    
    Next i
    
shtImp.ShowAllData

MsgBox "Done!"

rollReq.Activate

End Sub
