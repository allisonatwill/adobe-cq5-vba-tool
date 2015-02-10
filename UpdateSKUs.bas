Attribute VB_Name = "Module1"
Option Explicit
Sub UpdateSKU()
'FIND ALL OLDSKU VALUES AND UPDATE FOR EACH

   Application.ScreenUpdating = False
    
    Dim SKUSubset As String
    Dim skuCells As Collection, shtExp As Worksheet, shtImp As Worksheet
    Dim rollReq As Worksheet
    Dim skuExp As Worksheet, skuImp As Worksheet, flagExp As Worksheet, flagImp As Worksheet, attExp As Worksheet
    Dim attImp As Worksheet, deactSku As Worksheet
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

    Dim OldSKU As Long
    Dim NewSKU As Long

    Dim rLastSKU As Long

    Dim i As Range

    rLastSKU = rollReq.Cells(Rows.Count, "A").End(xlUp).Row

For Each i In Worksheets("Rollover Request").Range("A2:A" & rLastSKU).Cells
    OldSKU = i.Value
    NewSKU = i.Offset(0, 1).Value


'UPDATE SKU IMPORTER

skuExp.Range("A1").AutoFilter _
                                   Field:=1, Criteria1:=OldSKU
                                   
skuExp.Range(skuExp.Cells(2, 1), skuExp.UsedRange. _
                       SpecialCells(xlLastCell)).Copy
                       
skuImp.Cells(Rows.Count, 1).End(xlUp).Offset(1).PasteSpecial _
           Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False

skuImp.Columns("A:A").Replace What:=OldSKU, Replacement:=NewSKU, LookAt:=xlPart _

skuExp.ShowAllData
        
'UPDATE ATTRIBUTE EXPORTER

attExp.Range("A1").AutoFilter _
                                   Field:=1, Criteria1:=OldSKU
        
attExp.Range(attExp.Cells(2, 1), attExp.UsedRange. _
                       SpecialCells(xlLastCell)).Copy

attImp.Cells(Rows.Count, 1).End(xlUp).Offset(1).PasteSpecial _
           Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False

attImp.Columns("A:A").Replace What:=OldSKU, Replacement:=NewSKU, LookAt:=xlPart _

attExp.ShowAllData

'UPDATE SKU FLAG IMPORTER

flagExp.Range("A1").AutoFilter _
                                   Field:=1, Criteria1:=OldSKU
        
flagExp.Range(flagExp.Cells(2, 1), flagExp.UsedRange. _
                       SpecialCells(xlLastCell)).Copy

flagImp.Cells(Rows.Count, 1).End(xlUp).Offset(1).PasteSpecial _
           Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False

flagImp.Columns("A:A").Replace What:=OldSKU, Replacement:=NewSKU, LookAt:=xlPart _

flagExp.ShowAllData

'UPDATE DEACTIVE OLD SKU IMPORTER

deactSku.Cells(Rows.Count, 1).End(xlUp).Offset(1).Value = OldSKU
deactSku.Cells(Rows.Count, 23).End(xlUp).Offset(1).Value = "deactivate"




'UPDATE SKU SUBSET IMPORTER

Set skuCells = FindAll(shtExp.Columns(2), OldSKU) 'get all instances of SKU

shtExp.Activate

    For Each skuCell In skuCells

        SKUSubset = skuCell.Offset(0, -1).Value

        shtExp.Range("A1", Cells(1, 1).SpecialCells(xlLastCell)).AutoFilter _
                                   Field:=1, Criteria1:=SKUSubset

        shtExp.Range(shtExp.Cells(2, 1), shtExp.UsedRange. _
                       SpecialCells(xlLastCell)).Copy

        shtImp.Cells(Rows.Count, 1).End(xlUp).Offset(1).PasteSpecial _
           Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False

        shtExp.ShowAllData

    Next skuCell



Next i


For Each i In rollReq.Range("A2:A" & rLastSKU).Cells
OldSKU = i.Value
NewSKU = i.Offset(0, 1).Value

shtImp.Columns("B:B").Replace What:=OldSKU, Replacement:=NewSKU, LookAt:=xlPart _

Next i




rollReq.Activate

MsgBox "Done!"


End Sub




'return a Collection containing all cells with value [findWhat]

Function FindAll(rngToSearch As Range, findWhat As Long) As Collection
Dim rv As New Collection, f As Range, add1 As String
    Set f = rngToSearch.Find(What:=findWhat, LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then
        add1 = f.Address()
        Do While Not f Is Nothing
            rv.Add f
            Set f = rngToSearch.FindNext(After:=f)
            If f.Address = add1 Then Exit Do
        Loop
    End If
    Set FindAll = rv
End Function


