VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ButtonSysUp_Click()

   Dim strPasswd

    strPasswd = InputBox("Enter Password", "Update Sys Info")

    'Check to see if there is any entry made to input box, or if
    'cancel button is pressed. If no entry made then exit sub.

    If strPasswd = "" Or strPasswd = Empty Then
        MsgBox "No Input Provided", vbInformation, "Required Data"
        Exit Sub
    End If

    'If correct password is entered open form
    'If incorrect password entered give message and exit sub

    If strPasswd = "ax" Then
            
    DoCmd.OpenForm "SysUpdate"
    
    Else
        MsgBox "Sorry, you do not have access to this form", vbOKOnly, "Important Information"
        Exit Sub
    End If


End Sub

'Import from item template'
Private Sub ButtonImp_Click()

If DCount("*", "SourceDocumentTrim") = 0 Then
MsgBox ("Missing Company")

Else

    RunSharePointFinishedGoodsImport

End If

End Sub


'Validate imported data'
Private Sub ButtonVal_Click()

If Me.LstImp < Me.LstVal Or DCount("*", "SourceDocumentData") = 0 Then

MsgBox ("Import Data")

Else

If DCount("*", "CompanyPassFail") > 1 Then
MsgBox ("Failed Company Validation")

Else

'Apply the agreed Item Group 3100 / Customer Group exception before validation.
'Clear the invalid placeholder only for matching records, then continue validation.
If DCount("*", "SourceDocumentData", "AXItemGroup = '3100' AND CustGroup = 'N/A - N/A - Not Applicable'") > 0 Then
    CurrentDb.Execute "UPDATE SourceDocumentData SET CustGroup = Null WHERE AXItemGroup = '3100' AND CustGroup = 'N/A - N/A - Not Applicable'", dbFailOnError
    MsgBox "Warning: Customer Group was cleared for Item Group 3100 because it was set to 'N/A - N/A - Not Applicable'. Validation will continue.", vbExclamation, "Customer Group Warning"
End If

If DCount("*", "AllPassFail", "ItemTypeVal = 'Fail'") > 0 Then
MsgBox ("Failed Company Validation")

Else

If DCount("*", "AllPassFail", "CompanyVal = 'Fail'") > 0 Then
MsgBox ("Failed Company Validation")

Else

If DCount("*", "AllPassFail", "SiteVal = 'Fail'") > 0 Then
MsgBox ("Failed Site Validation")

Else

If DCount("*", "AllPassFail", "ItemDescVal = 'Fail'") > 0 Then
MsgBox ("Failed Item Description Validation")

Else

If DCount("*", "AllPassFail", "ItemDescLenVal = 'Fail'") > 0 Then
MsgBox ("Failed Item Description Length (>60 Char)")

Else


If DCount("*", "AllPassFail", "Sys21ItemCodeVal = 'Fail'") > 0 Then
MsgBox ("Failed S21 Item Code Validation")

Else

If DCount("*", "AllPassFail", "CategoryVal = 'Fail'") > 0 Then
MsgBox ("Failed Category Heirachy Validation")

Else

If DCount("*", "AllPassFail", "CustGrpVal = 'Fail'") > 0 Then
MsgBox ("Failed Customer Group Validation")

Else

If DCount("*", "AllPassFail", "S21ItemExistVal = 'Fail'") > 0 Then
MsgBox ("Failed S21 Item Exist Validation")

Else

If DCount("*", "AllPassFail", "SupplierVal = 'Fail'") > 0 Then
MsgBox ("Failed AX Supplier Validation")

Else

If DCount("*", "AllPassFail", "VatGrpVal = 'Fail'") > 0 Then
MsgBox ("Failed VAT Group Validation")

Else

If DCount("*", "AllPassFail", "WHSLocVal = 'Fail'") > 0 Then
MsgBox ("Failed Warehouse Location Validation")

Else

If DCount("*", "AllPassFail", "StockUoMVal = 'Fail'") > 0 Then
MsgBox ("Failed Stock UOM Validation")

Else

If DCount("*", "AllPassFail", "PurchUoMVal = 'Fail'") > 0 Then
MsgBox ("Failed Purch UOM Validation")

Else

If DCount("*", "AllPassFail", "AXItemCatVal = 'Fail'") > 0 Then
MsgBox ("Failed Item Category Validation")

Else

If DCount("*", "AllPassFail", "AXItemGroupVal = 'Fail'") > 0 Then
MsgBox ("Failed Item Group Validation")

Else

If DCount("*", "AllPassFail", "ExistingAXItemCodeVal = 'Fail'") > 0 Then
MsgBox ("Failed Existing AX Item Validation")

Else

If DCount("*", "AllPassFail", "PurchPriceNumVal = 'Fail'") > 0 Then
MsgBox ("Failed Purch Price Number Validation")

Else

If DCount("*", "AllPassFail", "ItemPackSizeNumVal = 'Fail' AND ItemPackSize Is Not Null AND ItemPackSize <> ''") > 0 Then
MsgBox ("Failed Pack Size Number Validation")

Else

If DCount("*", "AllPassFail", "NetWeightNumVal = 'Fail' AND NetWeight Is Not Null AND NetWeight <> ''") > 0 Then
MsgBox ("Failed Net Weight Number Validation")

Else


If DCount("*", "AllPassFail", "GTINVal = 'Fail'") > 0 Then
MsgBox ("Failed GTIN Validation")

Else

If DCount("*", "AllPassFail", "UomSeqGrpIdVal = 'Fail'") > 0 Then
MsgBox ("Failed UoMSeqGrp Validation")

Else


If DCount("*", "AllPassFail", "PurchPriceNumVal = 'Warning'") > 0 Then
MsgBox ("Missing Purch Price")

End If

If DCount("*", "SourceDocumentData", "ItemType = 'FIN' AND (ItemPackSize Is Null OR ItemPackSize = '')") > 0 Then
MsgBox ("Warning: Missing Pack Size")

End If

If DCount("*", "SourceDocumentData", "ItemType = 'FIN' AND (NetWeight Is Null OR NetWeight = '')") > 0 Then
MsgBox ("Warning: Missing Net Weight")

End If

If DCount("*", "AllPassFail", "GTINVal = 'Warning'") > 0 Then
MsgBox ("Missing GTIN or already Exists In That Company")
DoCmd.OpenForm "GTIN"

End If

If DCount("*", "UoMConvCheck", "Check = 'Check'") > 0 Then
MsgBox ("Check UoM Conversion")
DoCmd.OpenForm "UoMCheck"

End If


DoCmd.SetWarnings False

DoCmd.RunSQL "UPDATE SysInfo Set LastValidate = Now()"
Forms!main.Refresh
DoCmd.OpenForm "Validation"
MsgBox ("Validation Complete")


End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If


DoCmd.SetWarnings True

End Sub

'Open Validation Check'
Private Sub ValChck_Click()

DoCmd.OpenForm "Validation"

End Sub





Private Sub ButtonUoM_Click()

If Me.LstVal < Me.LstImp Then

MsgBox ("Not Validated")

Else

DoCmd.OpenForm "UoMCheck"

End If


End Sub


'Update AX ID'
Private Sub ButtonUpAxId_Click()

If Me.LstVal < Me.LstImp Then

MsgBox ("Not Validated")

Else

DoCmd.SetWarnings False

If DCount("*", "ExistingAXItemRelease") > 0 Then
DoCmd.RunSQL "UPDATE SysInfo Set LastItemUpdate = Now()"
DoCmd.OpenForm "AXItemCheck"
MsgBox ("Update Complete")
Forms!main.Refresh

DoCmd.SetWarnings True

Else

Dim strAXID As String

strAXID = InputBox("Please Enter Next AX Item Code", "AX Item Code")


    If strAXID = "" Or strAXID = Empty Or Len(strAXID) <> 6 Or DCount("*", "ItemsLive", "Item = '" & strAXID & "'") > 0 Then
        MsgBox "Invalid AX Code", vbInformation, "Required Data"
        Exit Sub
    End If

DoCmd.SetWarnings False

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim i As Long
    
    Set db = CurrentDb

    ' Clear the NewAXItemCode column (in case you want to run this more than once)
    sql = "UPDATE NewAXItems Set NewAXItemCode = Null"
    db.Execute sql

    ' Open your table with the Transaction_Ref and NewAXItemCode fields
    Set rs = db.OpenRecordset("NewAXItems")

    ' Zero the counter
    i = strAXID - 1

    ' Start the loop (set it to end when it gets to the last record)
    Do While Not rs.EOF

        ' Only update NewAXItemCodes that haven't got a value yet
        If IsNull(rs!NewAXItemCode) Then

            ' Push the counter on
            i = i + 1
            
            ' Update New Items with AX code
            sql = "UPDATE NewAXItems Set NewAXItemCode = " & i & " WHERE NewAXItems.RecID = '" & rs!RecID & "'"
                 
            db.Execute sql

        End If

        ' Move to the next record
        rs.MoveNext

    Loop

    'clean up
    rs.Close
    Set rs = Nothing
    Set db = Nothing

DoCmd.SetWarnings True

'Add in New AX Item check

If DCount("*", "NewAXItemCodeCheck") > 0 Then
MsgBox ("AX Item Code(s) Already Exist")

Else

DoCmd.SetWarnings False

DoCmd.RunSQL "UPDATE SysInfo Set LastItemUpdate = Now()"
DoCmd.OpenForm "AXItemCheck"
MsgBox ("Update Complete")
Forms!main.Refresh

End If
End If
End If

DoCmd.SetWarnings True

End Sub

Private Sub ButtonUpSearchName_Click()

If Me.LstAX < Me.LstVal Then

MsgBox ("AX Item Codes Not Updated")

Else

DoCmd.OpenForm "SearchNameUpdate"

End If

End Sub

'Export to BIS template'
Private Sub ButtonExp_Click()

On Error Resume Next
'MsgBox ("TEST0")

'MsgBox (Me.LstVal)
'MsgBox (Me.LstImp)
'MsgBox (Me.LstAX)

If Me.LstVal < Me.LstImp Then
'MsgBox ("TEST1")
MsgBox ("Not Validated")

Else

If Me.LstAX < Me.LstVal Then
'MsgBox ("TEST2")
MsgBox ("AX Item Codes Not Updated")

Else

If Me.LstAX < Me.LstExp Then
'MsgBox ("TEST3")
MsgBox ("Already Exported")

Else

DoCmd.SetWarnings False

If IsDate(DLookup("[LastExport]", "SysInfo", "format([LastExport], 'dd/mm/yyyy') <> Date()")) Then
DoCmd.RunSQL "Update SysInfo set LastExport = Now()"
DoCmd.RunSQL "Update SysInfo set SequenceNumber = 1"

DoCmd.TransferSpreadsheet acExport, 10, "Item-Main", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "Item-MainExist", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "Item-OrderSettings", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "Item-InventModule", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "Item-UoM", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "Item-Prices", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "ApprovedSupplier", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "GTIN", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "LocationUpdate", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "Item-InventBatch", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "Product", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "UOMSeqID", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "CrossReference", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "TradeAgreement", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)

ExportTradeAgreementUpload

PostProcessOriginalExport Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx"

DoCmd.RunSQL "UPDATE SysInfo Set SequenceNumber = SequenceNumber + 1"
MsgBox ("Export Complete")
Forms!main.Refresh

DoCmd.SetWarnings True

Else

'MsgBox ("TEST4")
DoCmd.SetWarnings False

'MsgBox (Me.Filename)
DoCmd.TransferSpreadsheet acExport, 10, "Item-Main", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "Item-MainExist", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "Item-OrderSettings", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "Item-InventModule", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "Item-UoM", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "Item-Prices", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "ApprovedSupplier", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "GTIN", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "LocationUpdate", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "Item-InventBatch", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "Product", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "UOMSeqID", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "CrossReference", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)
DoCmd.TransferSpreadsheet acExport, 10, "TradeAgreement", Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx", True
If Err.Description <> "" Then MsgBox (Err.Description)

ExportTradeAgreementUpload

PostProcessOriginalExport Me.OutLoc & Me.Filename & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx"

DoCmd.RunSQL "Update SysInfo set LastExport = Now()"
DoCmd.RunSQL "UPDATE SysInfo Set SequenceNumber = SequenceNumber + 1"
MsgBox ("Export Complete")
Forms!main.Refresh

End If
End If
End If
End If
'MsgBox ("TEST5")
DoCmd.SetWarnings True

End Sub

Private Sub ExportTradeAgreementUpload()

Dim CostGroupCount As Long
Dim dataAreaID As String
Dim TradeAgreementPath As String
Dim TradeAgreementFile As String

' Only create the additional Trade Agreement file when every Item-Main record
' has a Cost Group of PK or RM.
CostGroupCount = DCount("*", "Item-Main")

If CostGroupCount = 0 Then Exit Sub

If DCount("*", "Item-Main", "A_CostGroupId NOT IN ('PK','RM') OR A_CostGroupId Is Null") > 0 Then
    Exit Sub
End If

' CrossReference will contain one dataAreaId value for the export.
dataAreaID = Nz(DLookup("dataAreaId", "CrossReference"), "")

' Only the agreed dataAreaId values are valid destinations.
If dataAreaID <> "26" And dataAreaID <> "27" And dataAreaID <> "29" And dataAreaID <> "30" And dataAreaID <> "33" Then
    MsgBox "Invalid dataAreaId for Trade Agreement export: " & dataAreaID, vbExclamation, "Trade Agreement Export"
    Exit Sub
End If

TradeAgreementPath = "\\Sbseaxlfs01\" & dataAreaID & "\Trade_Agreement_Upload\"
TradeAgreementFile = TradeAgreementPath & "TradeAgreement_" & Format(Now(), "yyyy-mmm-dd_") & Format(Me.SeqID, "000") & ".xlsx"

If Dir(TradeAgreementPath, vbDirectory) = "" Then
    MsgBox "Trade Agreement folder does not exist:" & vbCrLf & TradeAgreementPath, vbExclamation, "Trade Agreement Export"
    Exit Sub
End If

DoCmd.TransferSpreadsheet acExport, 10, "TradeAgreement", TradeAgreementFile, True

If Err.Description <> "" Then
    MsgBox Err.Description, vbExclamation, "Trade Agreement Export"
End If

End Sub

Private Sub PostProcessOriginalExport(ByVal ExportFile As String)

Dim xlApp As Object
Dim xlBook As Object
Dim wsMain As Object
Dim wsCross As Object
Dim wsOrder As Object
Dim wsLocation As Object
Dim wsSupplier As Object
Dim lastRow As Long
Dim r As Long
Dim newRow As Long
Dim itemID As String
Dim costGroup As String
Dim defaultSite As String
Dim orderSiteId As String
Dim locationId As String
Dim crossSiteId As String
Dim dataAreaID As String
Dim allowedSites As String
Dim doCopy As Boolean
Dim itemKey As Variant
Dim costGroups As Object
Dim itemSites As Object
Dim mixWipItems As Object
Dim itemData As Object
Dim itemAliases As Object
Dim itemDataAreas As Object
Dim ruleDB As Object
Dim ruleRS As Object
Dim aliasColumn As Long
Dim ruleCustomer As String
Dim RuleDataArea As String
Dim RuleSite As String
Dim ruleLocation As String
Dim ruleWarehouse As String
Dim locationExists As Boolean
Dim siteMatches As Boolean
Dim lastItemRow As Long
Dim lastOrderRow As Long
Dim lastCrossRow As Long
Dim lastLocationRow As Long
Dim lastSupplierRow As Long
Dim ErrorNumber As Long
Dim ErrorDescription As String
Dim wsError As String
Dim rName As String
Dim rNameColumn As Long

On Error GoTo PostProcessError

' ============================================================
' OPEN THE ORIGINAL EXCEL WORKBOOK
' ============================================================
Set xlApp = CreateObject("Excel.Application")
xlApp.DisplayAlerts = False
Set xlBook = xlApp.Workbooks.Open(ExportFile)

Set wsMain = xlBook.Worksheets("Item_Main")
Set wsCross = xlBook.Worksheets("CrossReference")
Set wsOrder = xlBook.Worksheets("Item_OrderSettings")
Set wsLocation = xlBook.Worksheets("LocationUpdate")
Set wsSupplier = xlBook.Worksheets("ApprovedSupplier")

' ============================================================
' BUILD ITEM COST-GROUP / SITE LOOKUPS FROM THE EXPORTED FILE
' ============================================================
' This deliberately reads the workbook itself rather than Access.
' That means the post-export rules are applied to exactly what was
' actually exported.
Set costGroups = CreateObject("Scripting.Dictionary")
Set itemSites = CreateObject("Scripting.Dictionary")
Set mixWipItems = CreateObject("Scripting.Dictionary")
Set itemData = CreateObject("Scripting.Dictionary")
Set itemAliases = CreateObject("Scripting.Dictionary")
Set itemDataAreas = CreateObject("Scripting.Dictionary")

' ------------------------------------------------------------
' FIND R_Name COLUMN IN Item_Main
' ------------------------------------------------------------
' R_Name is located by its header rather than relying on a fixed
' column number. Header search is performed across row 1.
rNameColumn = 0

For r = 1 To wsMain.UsedRange.Columns.Count
    If Trim(UCase(CStr(wsMain.Cells(1, r).value))) = "R_NAME" Then
        rNameColumn = r
        Exit For
    End If
Next r

If rNameColumn = 0 Then
    Err.Raise vbObjectError + 1001, "PostProcessOriginalExport", _
              "Could not find the R_Name column in Item_Main."
End If

' ------------------------------------------------------------
' FIND A_NameAlias COLUMN IN Item_Main FOR WAREHOUSE RULE MATCHING
' ------------------------------------------------------------
aliasColumn = 0

For r = 1 To wsMain.UsedRange.Columns.Count
    If Trim(UCase(CStr(wsMain.Cells(1, r).value))) = "A_NAMEALIAS" Then
        aliasColumn = r
        Exit For
    End If
Next r

If aliasColumn = 0 Then
    Err.Raise vbObjectError + 1002, "PostProcessOriginalExport", _
              "Could not find the A_NameAlias column in Item_Main."
End If

' ------------------------------------------------------------
' Item_Main
' A_ItemId = column A
' R_Name   = located by header
' A_CostGroupId = column J
' ------------------------------------------------------------
lastItemRow = wsMain.Cells(wsMain.Rows.Count, "A").End(-4162).Row

For r = 2 To lastItemRow

    itemID = Trim(CStr(wsMain.Cells(r, 1).value))
    rName = Trim(CStr(wsMain.Cells(r, rNameColumn).value))
    costGroup = Trim(UCase(CStr(wsMain.Cells(r, 10).value)))

    ' --------------------------------------------------------
    ' MIX / WIP COST-GROUP OVERRIDES
    ' --------------------------------------------------------
    ' R_Name starts with:
    '   MIX -> Cost Group MX
    '   WIP -> Cost Group IW
    '
    ' Mix/WIP items are manufactured items, not purchased RM/PK
    ' items, so they are tracked separately from normal cost groups.
    '
    ' The source system does not naturally generate MX, so MIX
    ' must be explicitly converted to MX.
    '
    ' WIP is explicitly converted to IW.
If UCase(Left(rName, 3)) = "MIX" Then

    costGroup = "MX"

    If itemID <> "" Then
        mixWipItems(itemID) = "MIX"
    End If

ElseIf UCase(Left(rName, 3)) = "WIP" Then

    costGroup = "IW"

    If itemID <> "" Then
        mixWipItems(itemID) = "WIP"
    End If

End If

' Write the calculated Cost Group back to Item_Main.
wsMain.Cells(r, 10).value = costGroup

If itemID <> "" Then
    costGroups(itemID) = costGroup
    itemAliases(itemID) = Trim(CStr(wsMain.Cells(r, aliasColumn).value))
End If

Next r

' ------------------------------------------------------------
' Item_OrderSettings
' A_DefaultSite = column A
' D_ItemId      = column C
' E_InventSiteId = column D
' LocationId    = column E
' ------------------------------------------------------------
lastOrderRow = wsOrder.Cells(wsOrder.Rows.Count, "C").End(-4162).Row

For r = 2 To lastOrderRow

    itemID = Trim(CStr(wsOrder.Cells(r, 3).value))
    orderSiteId = Trim(UCase(CStr(wsOrder.Cells(r, 4).value)))

    If itemID <> "" And orderSiteId <> "" Then

        If Not itemSites.exists(itemID) Then
            itemSites.Add itemID, "|" & orderSiteId & "|"
        ElseIf InStr(1, itemSites(itemID), "|" & orderSiteId & "|", vbTextCompare) = 0 Then
            itemSites(itemID) = itemSites(itemID) & orderSiteId & "|"
        End If

    End If

Next r

' ============================================================
' APPROVED SUPPLIER CLEANUP
' ============================================================
' Mix/WIP items are manufactured products and should not have
' ApprovedSupplier records.
'
' ApprovedSupplier:
'   Column B = ItemId
'
' Match:
'   ApprovedSupplier.ItemId = Item_Main.A_ItemId
'
' Only ApprovedSupplier rows belonging to Mix/WIP items are deleted.
' All other ApprovedSupplier records are retained.
'
' Header row is retained.
wsError = "ApprovedSupplier cleanup"

lastSupplierRow = wsSupplier.Cells(wsSupplier.Rows.Count, "B").End(-4162).Row

For r = lastSupplierRow To 2 Step -1

    itemID = Trim(CStr(wsSupplier.Cells(r, 2).value))

    If itemID <> "" Then

        If mixWipItems.exists(itemID) Then
            wsSupplier.Rows(r).Delete
        End If

    End If

Next r

' ============================================================
' CROSSREFERENCE FILTERING, PER ITEM
' ============================================================
' RM / PK + BAG:
'   Keep BAG and BMR only.
'
' RM / PK + MLT:
'   Keep MLT only.
'
' FG / SL / AL + BAG:
'   Keep BAG, BMR and MAN only.
'
' MIX / WIP + BAG:
'   Keep BAG and BMR only.
'
' Other cost groups/sites:
'   Leave CrossReference unchanged.
wsError = "CrossReference filtering"

lastCrossRow = wsCross.Cells(wsCross.Rows.Count, "A").End(-4162).Row

For r = lastCrossRow To 2 Step -1

    itemID = Trim(CStr(wsCross.Cells(r, 1).value))
    crossSiteId = Trim(UCase(CStr(wsCross.Cells(r, 4).value)))
    allowedSites = ""

    If costGroups.exists(itemID) Then

        costGroup = costGroups(itemID)

        If itemSites.exists(itemID) Then

            If InStr(1, itemSites(itemID), "|BAG|", vbTextCompare) > 0 Then

                If mixWipItems.exists(itemID) Then

                    ' MIX / WIP:
                    ' Keep BAG and BMR only.
                    allowedSites = "|BAG|BMR|"

                ElseIf costGroup = "RM" Or costGroup = "PK" Then

                    ' RM / PK:
                    ' Keep BAG and BMR only.
                    allowedSites = "|BAG|BMR|"

                ElseIf costGroup = "FG" Or costGroup = "SL" Or costGroup = "AL" Then

                    ' FG / SL / AL:
                    ' Keep BAG, BMR and MAN.
                    allowedSites = "|BAG|BMR|MAN|"

                End If

            ElseIf InStr(1, itemSites(itemID), "|MLT|", vbTextCompare) > 0 Then

                If costGroup = "RM" Or costGroup = "PK" Then
                    allowedSites = "|MLT|"
                End If

            End If

        End If

    End If

    If allowedSites <> "" Then

        If InStr(1, allowedSites, "|" & crossSiteId & "|", vbTextCompare) = 0 Then
            wsCross.Rows(r).Delete
        End If

    End If

Next r

' ============================================================
' BAG -> BMR ITEM_ORDERSETTINGS DUPLICATION
' ============================================================
' RM / PK + BAG:
'   Copy every BAG row, including DEFAULT rows.
'   BAG becomes BMR.
'   WHAG becomes WHMR.
'   Blank locations remain blank.
'
' MIX / WIP + BAG:
'   Copy every BAG row, including DEFAULT rows.
'   BAG becomes BMR.
'   WHAG becomes WHMR.
'   Blank locations remain blank.
'
' FG / SL / AL + BAG:
'   Copy only SITE / BAG / WHAG rows.
'   BAG becomes BMR.
'   WHAG becomes WHMR.
'
' Original rows are retained.
wsError = "Item_OrderSettings duplication"

' Work backwards so newly inserted rows cannot be processed again.
lastOrderRow = wsOrder.Cells(wsOrder.Rows.Count, "A").End(-4162).Row

For r = lastOrderRow To 2 Step -1

    defaultSite = Trim(UCase(CStr(wsOrder.Cells(r, 1).value)))
    itemID = Trim(CStr(wsOrder.Cells(r, 3).value))
    orderSiteId = Trim(UCase(CStr(wsOrder.Cells(r, 4).value)))
    locationId = Trim(UCase(CStr(wsOrder.Cells(r, 5).value)))
    doCopy = False

    If costGroups.exists(itemID) And itemSites.exists(itemID) Then

        costGroup = costGroups(itemID)

        If InStr(1, itemSites(itemID), "|BAG|", vbTextCompare) > 0 Then

            If mixWipItems.exists(itemID) Then

                ' MIX / WIP:
                ' Copy every BAG row, including DEFAULT.
                If orderSiteId = "BAG" Then
                    doCopy = True
                End If

            ElseIf costGroup = "RM" Or costGroup = "PK" Then

                ' RM / PK:
                ' Copy every BAG row, including DEFAULT.
                If orderSiteId = "BAG" Then
                    doCopy = True
                End If

            ElseIf costGroup = "FG" Or costGroup = "SL" Or costGroup = "AL" Then

                ' FG / SL / AL:
                ' Only SITE + BAG + WHAG.
                If defaultSite = "SITE" And orderSiteId = "BAG" And locationId = "WHAG" Then
                    doCopy = True
                End If

            End If

        End If

    End If

    If doCopy Then

        newRow = wsOrder.Cells(wsOrder.Rows.Count, "A").End(-4162).Row + 1

        wsOrder.Rows(r).Copy
        wsOrder.Rows(newRow).Insert
        xlApp.CutCopyMode = False

        wsOrder.Cells(newRow, 4).value = "BMR"

        If locationId = "WHAG" Then
            wsOrder.Cells(newRow, 5).value = "WHMR"
        End If

    End If

Next r

' ============================================================
' FG / SL / AL + BAG -> ADD RR / WHRR AND F1 / WHF1 TO LOCATIONUPDATE
' ============================================================
' Add two rows per unique FG / SL / AL ItemId where that item has BAG in
' Item_OrderSettings.
'
' DataAreaId is taken from the first CrossReference dataAreaId.
wsError = "LocationUpdate additions"

dataAreaID = ""
lastCrossRow = wsCross.Cells(wsCross.Rows.Count, "A").End(-4162).Row

For r = 2 To lastCrossRow

    dataAreaID = Trim(CStr(wsCross.Cells(r, 3).value))

    If dataAreaID <> "" Then Exit For

Next r

For Each itemKey In costGroups.Keys

    itemID = CStr(itemKey)
    costGroup = costGroups(itemID)

    If costGroup = "FG" Or costGroup = "SL" Or costGroup = "AL" Then

        If itemSites.exists(itemID) Then

            If InStr(1, itemSites(itemID), "|BAG|", vbTextCompare) > 0 Then

                newRow = wsLocation.Cells(wsLocation.Rows.Count, "A").End(-4162).Row + 1

                If newRow < 2 Then newRow = 2

                wsLocation.Cells(newRow, 1).value = itemID
                wsLocation.Cells(newRow, 2).value = dataAreaID
                wsLocation.Cells(newRow, 3).value = "RR"
                wsLocation.Cells(newRow, 4).value = "WHRR"

                newRow = wsLocation.Cells(wsLocation.Rows.Count, "A").End(-4162).Row + 1

                wsLocation.Cells(newRow, 1).value = itemID
                wsLocation.Cells(newRow, 2).value = dataAreaID
                wsLocation.Cells(newRow, 3).value = "F1"
                wsLocation.Cells(newRow, 4).value = "WHF1"

            End If

        End If

    End If

Next itemKey

' ============================================================
' TABLE-DRIVEN EXTRA WAREHOUSE ADDITIONS
' ============================================================
' Access table required: WarehouseAdditions
' Fields: Customer, DataAreaId, AXSite, Location, Warehouse, Active
' Matching uses the item's AXSite and retailer/range abbreviations.
'
' Customer matching uses the meaningful words in the Customer rule after
' removing the leading company/site label (Bradgate/Melton) and trailing
' word Items. Every remaining word must occur in A_NameAlias.
'
' Existing LocationUpdate records are left untouched. Only missing exact
' ItemId + DataAreaId + Location + Warehouse combinations are appended.
wsError = "WarehouseAdditions table processing"

If Not WarehouseAdditionsTableExists() Then
    Err.Raise vbObjectError + 1003, "PostProcessOriginalExport", _
              "The Access table 'WarehouseAdditions' does not exist. Create it before running the BIS export."
End If

' Build item -> DataAreaId lookup from CrossReference (ItemId in column A,
' DataAreaId in column C). Keep the first nonblank company for each item.
lastCrossRow = wsCross.Cells(wsCross.Rows.Count, "A").End(-4162).Row
For r = 2 To lastCrossRow
    itemID = Trim(CStr(wsCross.Cells(r, 1).value))
    RuleDataArea = Trim(CStr(wsCross.Cells(r, 3).value))
    If itemID <> "" And RuleDataArea <> "" Then
        If Not itemDataAreas.exists(itemID) Then itemDataAreas.Add itemID, RuleDataArea
    End If
Next r

Set ruleDB = CurrentDb
Set ruleRS = ruleDB.OpenRecordset( _
    "SELECT [Customer], [DataAreaId], [AXSite], [Location], [Warehouse] " & _
    "FROM [WarehouseAdditions] WHERE [Active] = True", dbOpenSnapshot)

Do While Not ruleRS.EOF
    ruleCustomer = Trim(Nz(ruleRS!customer, ""))
    RuleDataArea = Trim(Nz(ruleRS!dataAreaID, ""))
    RuleSite = Trim(Nz(ruleRS!AXSite, ""))
    ruleLocation = Trim(Nz(ruleRS!Location, ""))
    ruleWarehouse = Trim(Nz(ruleRS!Warehouse, ""))

    If ruleCustomer <> "" And RuleDataArea <> "" And ruleLocation <> "" And ruleWarehouse <> "" Then
        For Each itemKey In itemAliases.Keys
            itemID = CStr(itemKey)

            If itemDataAreas.exists(itemID) And itemDataAreas(itemID) = RuleDataArea Then
                If costGroups.exists(itemID) And itemSites.exists(itemID) Then
                    ' This table is currently for Finished Goods cost groups only.
                    If costGroups(itemID) = "FG" Or costGroups(itemID) = "SL" Or costGroups(itemID) = "AL" Then
                        siteMatches = WarehouseRuleSiteMatches(RuleSite, CStr(itemSites(itemID)))

                        If siteMatches And WarehouseCustomerMatchesAlias(ruleCustomer, CStr(itemAliases(itemID)), RuleSite) Then
                            If WarehouseRuleHasMoreSpecificMatch(ruleCustomer, RuleDataArea, RuleSite, CStr(itemAliases(itemID))) Then
                                GoTo NextWarehouseItem
                            End If
                            locationExists = False
                            lastLocationRow = wsLocation.Cells(wsLocation.Rows.Count, "A").End(-4162).Row

                            For r = 2 To lastLocationRow
                                If Trim(CStr(wsLocation.Cells(r, 1).value)) = itemID _
                                   And Trim(CStr(wsLocation.Cells(r, 2).value)) = RuleDataArea _
                                   And StrComp(Trim(CStr(wsLocation.Cells(r, 3).value)), ruleLocation, vbTextCompare) = 0 _
                                   And StrComp(Trim(CStr(wsLocation.Cells(r, 4).value)), ruleWarehouse, vbTextCompare) = 0 Then
                                    locationExists = True
                                    Exit For
                                End If
                            Next r

                            If Not locationExists Then
                                newRow = wsLocation.Cells(wsLocation.Rows.Count, "A").End(-4162).Row + 1
                                If newRow < 2 Then newRow = 2
                                wsLocation.Cells(newRow, 1).value = itemID
                                wsLocation.Cells(newRow, 2).value = RuleDataArea
                                wsLocation.Cells(newRow, 3).value = ruleLocation
                                wsLocation.Cells(newRow, 4).value = ruleWarehouse
                            End If
                        End If
                    End If
                End If
            End If
NextWarehouseItem:
        Next itemKey
    End If

    ruleRS.MoveNext
Loop

ruleRS.Close
Set ruleRS = Nothing
Set ruleDB = Nothing

' ============================================================
' EDI DATA SHEET
' ============================================================
' Build the additional EDIData worksheet from the imported source fields
' and the Product / Item_Main data already exported to this workbook.
wsError = "EDIData generation"
BuildEDIDataSheet xlBook, costGroups

' ============================================================
' SAVE AND CLOSE
' ============================================================
wsError = "Saving workbook"

xlBook.Save
xlBook.Close False
xlApp.Quit

Set itemData = Nothing
Set itemAliases = Nothing
Set itemDataAreas = Nothing
Set mixWipItems = Nothing
Set costGroups = Nothing
Set itemSites = Nothing
Set wsMain = Nothing
Set wsCross = Nothing
Set wsOrder = Nothing
Set wsLocation = Nothing
Set wsSupplier = Nothing
Set xlBook = Nothing
Set xlApp = Nothing

Exit Sub

PostProcessError:

' Capture the actual error BEFORE cleanup can reset Err.
ErrorNumber = Err.Number
ErrorDescription = Err.Description

On Error Resume Next

If Not ruleRS Is Nothing Then ruleRS.Close
Set ruleRS = Nothing
Set ruleDB = Nothing
If Not xlBook Is Nothing Then xlBook.Close False
If Not xlApp Is Nothing Then xlApp.Quit

Set itemData = Nothing
Set itemAliases = Nothing
Set itemDataAreas = Nothing
Set mixWipItems = Nothing
Set costGroups = Nothing
Set itemSites = Nothing
Set wsMain = Nothing
Set wsCross = Nothing
Set wsOrder = Nothing
Set wsLocation = Nothing
Set wsSupplier = Nothing
Set xlBook = Nothing
Set xlApp = Nothing

MsgBox "The original export was created, but the post-export Excel changes could not be completed." & vbCrLf & vbCrLf & _
       "Failed during: " & wsError & vbCrLf & _
       "Error " & ErrorNumber & ": " & ErrorDescription, vbExclamation, "Post-Export Processing"

End Sub

Private Sub BuildEDIDataSheet(ByVal xlBook As Object, ByVal costGroups As Object)

    Dim wsProduct As Object
    Dim wsEDI As Object
    Dim rsEDI As DAO.Recordset
    Dim productNames As Object
    Dim emittedItems As Object
    Dim lastProductRow As Long
    Dim productNumberColumn As Long
    Dim productNameColumn As Long
    Dim r As Long
    Dim outputRow As Long
    Dim itemID As String
    Dim companyCode As String
    Dim sys21Code As String
    Dim custGroup As String
    Dim customerItemCode As String
    Dim gtin As String
    Dim caseSize As String
    Dim ediFlag As String
    Dim costGroup As String
    Dim productName As String
    Dim outputKey As String
    Dim headers As Variant
    Dim c As Long
    Dim savedErrNumber As Long
    Dim savedErrDescription As String

    On Error GoTo BuildEDIDataError

    ' Product supplies the item description used to derive CustItemCode.
    Set wsProduct = xlBook.Worksheets("Product")
    productNumberColumn = FindEDIHeaderColumn(wsProduct, "Product number")
    productNameColumn = FindEDIHeaderColumn(wsProduct, "Product Name")

    Set productNames = CreateObject("Scripting.Dictionary")
    productNames.CompareMode = vbTextCompare

    lastProductRow = wsProduct.Cells(wsProduct.Rows.Count, productNumberColumn).End(-4162).Row

    For r = 2 To lastProductRow
        itemID = GetEDIExcelCellText(wsProduct.Cells(r, productNumberColumn))
        productName = GetEDIExcelCellText(wsProduct.Cells(r, productNameColumn))

        If Len(itemID) > 0 Then
            If Not productNames.Exists(itemID) Then
                productNames.Add itemID, productName
            End If
        End If
    Next r

    ' Reuse the EDIData worksheet if one is already present. Otherwise create it.
    Set wsEDI = Nothing
    On Error Resume Next
    Set wsEDI = xlBook.Worksheets("EDIData")
    Err.Clear
    On Error GoTo BuildEDIDataError

    If wsEDI Is Nothing Then
        Set wsEDI = xlBook.Worksheets.Add
        wsEDI.Name = "EDIData"
    Else
        wsEDI.Cells.Clear
    End If

    ' Remove only completely empty default Excel sheets (Sheet2, Sheet3, etc.).
    ' Any worksheet containing data, or any deliberately named worksheet, is kept.
    For r = xlBook.Worksheets.Count To 1 Step -1
        If xlBook.Worksheets(r).Name <> "EDIData" Then
            If Left$(xlBook.Worksheets(r).Name, 5) = "Sheet" Then
                If IsNumeric(Mid$(xlBook.Worksheets(r).Name, 6)) Then
                    If xlBook.Application.WorksheetFunction.CountA(xlBook.Worksheets(r).UsedRange) = 0 Then
                        xlBook.Worksheets(r).Delete
                    End If
                End If
            End If
        End If
    Next r

    ' Keep EDIData as the final worksheet in the exported workbook.
    If wsEDI.Index <> xlBook.Worksheets.Count Then
        wsEDI.Move After:=xlBook.Worksheets(xlBook.Worksheets.Count)
    End If

    headers = Array("DataAreaId", "ItemId", "ItemS21", "CustGroup", "CustItemCode", "GTIN", "CaseSize")

    For c = LBound(headers) To UBound(headers)
        wsEDI.Cells(1, c + 1).Value = CStr(headers(c))
    Next c

    ' All output columns are text. This is particularly important for GTINs,
    ' which can begin with zero and must not be converted to scientific notation.
    wsEDI.Columns("A:G").NumberFormat = "@"
    wsEDI.Range("A1:G1").Font.Bold = True

    Set emittedItems = CreateObject("Scripting.Dictionary")
    emittedItems.CompareMode = vbTextCompare

    ' These values already exist in SourceDocumentData from the SharePoint import.
    ' AX Item ID is the generated NewAXItemCode, or ExistingAXItemCode for releases.
    Set rsEDI = CurrentDb.OpenRecordset( _
        "SELECT [Company], [ExistingAXItemCode], [NewAXItemCode], [Sys21ItemCode], " & _
        "[CustGroup], [GTIN], [ItemPackSize], [EDIItem] " & _
        "FROM [SourceDocumentData]", dbOpenSnapshot)

    outputRow = 2

    Do While Not rsEDI.EOF

        ediFlag = GetEDIRecordText(rsEDI, "EDIItem")

        If StrComp(ediFlag, "Yes", vbTextCompare) = 0 Then

            companyCode = GetEDIRecordText(rsEDI, "Company")
            itemID = GetEDIRecordText(rsEDI, "NewAXItemCode")

            If Len(itemID) = 0 Then
                itemID = GetEDIRecordText(rsEDI, "ExistingAXItemCode")
            End If

            If Len(companyCode) = 0 Then
                Err.Raise vbObjectError + 1201, "BuildEDIDataSheet", _
                          "An EDI-enabled SourceDocumentData row has no Company / DataAreaId."
            End If

            If Len(itemID) = 0 Then
                Err.Raise vbObjectError + 1202, "BuildEDIDataSheet", _
                          "An EDI-enabled SourceDocumentData row has no NewAXItemCode or ExistingAXItemCode."
            End If

            If Not costGroups.Exists(itemID) Then
                Err.Raise vbObjectError + 1203, "BuildEDIDataSheet", _
                          "EDI item " & itemID & " was not found in Item_Main, so its A_CostGroupId cannot be checked."
            End If

            If Not productNames.Exists(itemID) Then
                Err.Raise vbObjectError + 1204, "BuildEDIDataSheet", _
                          "EDI item " & itemID & " was not found in the Product worksheet."
            End If

            outputKey = companyCode & "|" & itemID

            ' One EDIData row per DataAreaId + ItemId, even if an input query returns duplicates.
            If Not emittedItems.Exists(outputKey) Then

                costGroup = UCase$(Trim$(CStr(costGroups(itemID))))
                productName = CStr(productNames(itemID))

                customerItemCode = vbNullString
                If costGroup = "FG" Then
                    ' Only standard Finished Goods carry CustItemCode.
                    ' The first token of Product Name is the customer item code.
                    customerItemCode = FirstTokenForEDI(productName)
                End If

                sys21Code = GetEDIRecordText(rsEDI, "Sys21ItemCode")
                custGroup = GetEDIRecordText(rsEDI, "CustGroup")
                gtin = GetEDIRecordText(rsEDI, "GTIN")
                caseSize = GetEDIRecordText(rsEDI, "ItemPackSize")

                wsEDI.Cells(outputRow, 1).Value = companyCode
                wsEDI.Cells(outputRow, 2).Value = itemID
                wsEDI.Cells(outputRow, 3).Value = sys21Code
                wsEDI.Cells(outputRow, 4).Value = custGroup
                wsEDI.Cells(outputRow, 5).Value = customerItemCode
                wsEDI.Cells(outputRow, 6).Value = gtin
                wsEDI.Cells(outputRow, 7).Value = caseSize

                emittedItems.Add outputKey, True
                outputRow = outputRow + 1

            End If

        End If

        rsEDI.MoveNext
    Loop

    wsEDI.Columns("A:G").AutoFit

    rsEDI.Close
    Set rsEDI = Nothing
    Set emittedItems = Nothing
    Set productNames = Nothing
    Set wsEDI = Nothing
    Set wsProduct = Nothing

    Exit Sub

BuildEDIDataError:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description

    On Error Resume Next
    If Not rsEDI Is Nothing Then rsEDI.Close
    Set rsEDI = Nothing
    Set emittedItems = Nothing
    Set productNames = Nothing
    Set wsEDI = Nothing
    Set wsProduct = Nothing

    On Error GoTo 0
    Err.Raise savedErrNumber, "BuildEDIDataSheet", savedErrDescription

End Sub


Private Function FindEDIHeaderColumn(ByVal ws As Object, ByVal headerName As String) As Long

    Dim lastColumn As Long
    Dim c As Long
    Dim headerValue As Variant

    lastColumn = ws.Cells(1, ws.Columns.Count).End(-4159).Column

    For c = 1 To lastColumn
        headerValue = ws.Cells(1, c).Value

        If Not IsError(headerValue) Then
            If StrComp(Trim$(CStr(headerValue)), headerName, vbTextCompare) = 0 Then
                FindEDIHeaderColumn = c
                Exit Function
            End If
        End If
    Next c

    Err.Raise vbObjectError + 1210, "FindEDIHeaderColumn", _
              "Could not find the '" & headerName & "' column in worksheet '" & ws.Name & "'."

End Function


Private Function GetEDIExcelCellText(ByVal targetCell As Object) As String

    Dim cellValue As Variant

    cellValue = targetCell.Value

    If IsError(cellValue) Or IsNull(cellValue) Or IsEmpty(cellValue) Then Exit Function

    GetEDIExcelCellText = Trim$(CStr(cellValue))

End Function


Private Function GetEDIRecordText(ByVal rs As DAO.Recordset, ByVal fieldName As String) As String

    Dim fieldValue As Variant

    fieldValue = rs.Fields(fieldName).Value

    If IsNull(fieldValue) Or IsEmpty(fieldValue) Then Exit Function

    GetEDIRecordText = Trim$(CStr(fieldValue))

End Function


Private Function FirstTokenForEDI(ByVal sourceText As String) As String

    sourceText = Replace(sourceText, ChrW(160), " ")
    sourceText = Replace(sourceText, vbTab, " ")
    sourceText = Trim$(sourceText)

    Do While InStr(1, sourceText, "  ", vbBinaryCompare) > 0
        sourceText = Replace(sourceText, "  ", " ")
    Loop

    If Len(sourceText) = 0 Then Exit Function

    If InStr(1, sourceText, " ", vbBinaryCompare) > 0 Then
        FirstTokenForEDI = Left$(sourceText, InStr(1, sourceText, " ", vbBinaryCompare) - 1)
    Else
        FirstTokenForEDI = sourceText
    End If

End Function


' Match the AXSite value stored in WarehouseAdditions against the item's site list.
Private Function WarehouseRuleSiteMatches(ByVal RuleSite As String, ByVal ItemSiteList As String) As Boolean
    RuleSite = UCase(Trim(RuleSite))
    If RuleSite = "" Or ItemSiteList = "" Then Exit Function
    WarehouseRuleSiteMatches = (InStr(1, ItemSiteList, "|" & RuleSite & "|", vbTextCompare) > 0)
End Function

Private Function WarehouseAdditionsTableExists() As Boolean
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Set db = CurrentDb
    For Each tdf In db.TableDefs
        If StrComp(tdf.Name, "WarehouseAdditions", vbTextCompare) = 0 Then
            WarehouseAdditionsTableExists = True
            Exit For
        End If
    Next tdf
    Set db = Nothing
End Function

' Returns the rule's meaningful customer/range phrase, normalized to the
' abbreviations used in A_NameAlias. Longer phrases are converted first.
Private Function WarehouseNormalizeRuleCustomer(ByVal CustomerRule As String) As String
    Dim s As String

    s = LCase(Trim(CustomerRule))

    ' Remove the site/division label at the start of the rule.
    If Left(s, 9) = "bradgate " Then s = Mid(s, 10)
    If Left(s, 7) = "melton " Then s = Mid(s, 8)
    If Left(s, 12) = "manton wood " Then s = Mid(s, 13)
    If Left(s, 10) = "charnwood " Then s = Mid(s, 11)
    If Left(s, 4) = "tcb " Then s = Mid(s, 5)
    If Left(s, 11) = "saladworks " Then s = Mid(s, 12)
    If Left(s, 8) = "kettleby " Then s = Mid(s, 9)
    If Left(s, 10) = "real wrap " Then s = Mid(s, 11)

    ' Remove the generic suffix.
    If Right(s, 6) = " items" Then s = Left(s, Len(s) - 6)

    ' Normalize specific retailer/range phrases before individual names.
    s = Replace(s, "tesco petrol stations", "tes pfs")
    s = Replace(s, "tesco party food", "tes pf")
    s = Replace(s, "tesco finest", "tes fin")
    s = Replace(s, "tesco express", "tes exp")
    s = Replace(s, "tesco xmas fto", "tes xmas fto")
    s = Replace(s, "tesco pfs", "tes pfs")
    s = Replace(s, "tesco pf", "tes pf")
    s = Replace(s, "tesco", "tes")
    s = Replace(s, "old el paso", "oep")
    s = Replace(s, "sainsburys", "js")
    s = Replace(s, "sainsbury", "js")
    s = Replace(s, "waitrose", "wai")
    s = Replace(s, "morrisons", "mor")
    s = Replace(s, "mcdonalds", "mcd")
    s = Replace(s, "aldi", "ald")
    s = Replace(s, "asda", "asd")
    s = Replace(s, "blueberry foods", "bbf")
    s = Replace(s, "coop", "cop")
    s = Replace(s, "co-op", "cop")
    s = Replace(s, "costco", "cco")
    s = Replace(s, "cost co", "cco")
    s = Replace(s, "costa", "cos")
    s = Replace(s, "farmers boy", "fb")
    s = Replace(s, "greencore", "gc")
    s = Replace(s, "ginsters", "gin")
    s = Replace(s, "gressingham", "grs")
    s = Replace(s, "honest crust", "hc")
    s = Replace(s, "independants", "ind")
    s = Replace(s, "intercompany", "int")
    s = Replace(s, "inter company", "int")
    s = Replace(s, "jamie oliver", "jo")
    s = Replace(s, "lidl", "lidl")
    s = Replace(s, "m&s", "m&s")
    s = Replace(s, "marston's brewery", "mar")
    s = Replace(s, "one stop", "os")
    s = Replace(s, "ocado", "oca")
    s = Replace(s, "starbucks", "stb")
    s = Replace(s, "urban eat", "ue")
    s = Replace(s, "vanco", "van")
    s = Replace(s, "walker & son", "w&s")
    s = Replace(s, "2 sisters food group", "2sfg")
    s = Replace(s, "grove turkey", "gt")
    s = Replace(s, "traditional norfolk poultry", "tnp")
    s = Replace(s, "jd wetherspoons", "jdw")

    ' Brand/range aliases supplied by the user.
    s = Replace(s, "little dish", "ld")
    s = Replace(s, "culina", "ld")

    s = Replace(s, "(", " ")
    s = Replace(s, ")", " ")
    s = Replace(s, "-", " ")
    s = Replace(s, "/", " ")
    s = Replace(s, ".", " ")
    s = Replace(s, ",", " ")
    s = Replace(s, "'", " ")
    s = Trim(s)

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    WarehouseNormalizeRuleCustomer = s
End Function

' Checks each normalized rule token as a whole token in A_NameAlias.
' This prevents TES matching an unrelated longer word and allows aliases
' such as TES PF to match a rule named Tesco Party Food.
Private Function WarehouseCustomerMatchesAlias(ByVal CustomerRule As String, ByVal ItemAlias As String, ByVal RuleSite As String) As Boolean
    Dim ruleText As String
    Dim aliasText As String
    Dim words() As String
    Dim word As Variant

    ruleText = WarehouseNormalizeRuleCustomer(CustomerRule)
    aliasText = LCase(Trim(ItemAlias))
    aliasText = Replace(aliasText, "(", " ")
    aliasText = Replace(aliasText, ")", " ")
    aliasText = Replace(aliasText, "-", " ")
    aliasText = Replace(aliasText, "/", " ")
    aliasText = Replace(aliasText, ".", " ")
    aliasText = Replace(aliasText, ",", " ")
    aliasText = Replace(aliasText, "'", " ")
    Do While InStr(aliasText, "  ") > 0
        aliasText = Replace(aliasText, "  ", " ")
    Loop

    If ruleText = "" Or aliasText = "" Then Exit Function

    ' OS/Greencore is an either/or label in the rule name.
    If ruleText = "os gc" Then
        WarehouseCustomerMatchesAlias = _
            (InStr(1, " " & aliasText & " ", " os ", vbTextCompare) > 0) Or _
            (InStr(1, " " & aliasText & " ", " gc ", vbTextCompare) > 0) Or _
            (InStr(1, " " & aliasText & " ", " greencore ", vbTextCompare) > 0)
        Exit Function
    End If

    words = Split(ruleText, " ")
    For Each word In words
        If Trim(CStr(word)) <> "" Then
            If InStr(1, " " & aliasText & " ", " " & Trim(CStr(word)) & " ", vbTextCompare) = 0 Then
                WarehouseCustomerMatchesAlias = False
                Exit Function
            End If
        End If
    Next word

    WarehouseCustomerMatchesAlias = True
End Function

' A general rule is suppressed only when a more-specific active rule for the
' same company and AX site also matches this item's alias. Rules with the same
' normalized customer phrase are not considered more specific, so multiple
' warehouses/locations belonging to one rule can still all be added.
Private Function WarehouseRuleHasMoreSpecificMatch(ByVal CustomerRule As String, ByVal RuleDataArea As String, ByVal RuleSite As String, ByVal ItemAlias As String) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim basePhrase As String
    Dim candidatePhrase As String
    Dim candidateCustomer As String

    basePhrase = WarehouseNormalizeRuleCustomer(CustomerRule)
    If basePhrase = "" Then Exit Function

    ' Intentionally include inactive rules here: if a specific rule exists
    ' but is inactive, the general rule must not be used as a fallback.
    sql = "SELECT [Customer] FROM [WarehouseAdditions] " & _
          "WHERE [DataAreaId] = '" & Replace(RuleDataArea, "'", "''") & "' " & _
          "AND [AXSite] = '" & Replace(RuleSite, "'", "''") & "'"

    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    Do While Not rs.EOF
        candidateCustomer = Trim(Nz(rs!customer, ""))
        candidatePhrase = WarehouseNormalizeRuleCustomer(candidateCustomer)

        If candidatePhrase <> "" And candidatePhrase <> basePhrase Then
            If Left(candidatePhrase, Len(basePhrase) + 1) = basePhrase & " " Then
                If WarehouseCustomerMatchesAlias(candidateCustomer, ItemAlias, RuleSite) Then
                    WarehouseRuleHasMoreSpecificMatch = True
                    Exit Do
                End If
            End If
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

