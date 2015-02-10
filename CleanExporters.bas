Attribute VB_Name = "Module2"
Option Explicit
Sub CleanImporters()

   Application.ScreenUpdating = False
    
    Dim SKUSubset As String
    Dim skuCells As Collection, shtExp As Worksheet, shtImp As Worksheet
    Dim rollReq As Worksheet
    Dim skuExp As Worksheet, skuImp As Worksheet, flagExp As Worksheet, flagImp As Worksheet, attExp As Worksheet
    Dim attImp As Worksheet, deactSku As Worksheet
    Dim OldSKU As Worksheet, oldSub As Worksheet, oldAtt As Worksheet, oldFlag As Worksheet
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
    Set OldSKU = Sheets("SKU (old)")
    Set oldSub = Sheets("Subset (old)")
    Set oldAtt = Sheets("Attribute (old)")
    Set oldFlag = Sheets("SKU Flag (old)")

    skuImp.Activate

    skuImp.Range(skuImp.Cells(2, 1), skuImp.UsedRange.SpecialCells(xlLastCell)).Cut (OldSKU.Cells(Rows.Count, 1).End(xlUp).Offset(1))


    shtImp.Activate

    shtImp.Range(shtImp.Cells(2, 1), shtImp.UsedRange.SpecialCells(xlLastCell)).Cut (oldSub.Cells(Rows.Count, 1).End(xlUp).Offset(1))


    flagImp.Activate

    flagImp.Range(flagImp.Cells(2, 1), flagImp.UsedRange.SpecialCells(xlLastCell)).Cut (oldFlag.Cells(Rows.Count, 1).End(xlUp).Offset(1))

    
    attImp.Activate

    attImp.Range(attImp.Cells(2, 1), attImp.UsedRange.SpecialCells(xlLastCell)).Cut (oldAtt.Cells(Rows.Count, 1).End(xlUp).Offset(1))

    deactSku.Range(deactSku.Cells(2, 1), deactSku.UsedRange.SpecialCells(xlLastCell)).Cut (oldAtt.Cells(Rows.Count, 1).End(xlUp).Offset(1))


    rollReq.Activate

MsgBox "Done!"

End Sub
