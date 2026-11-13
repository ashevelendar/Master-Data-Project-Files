Attribute VB_Name = "modSUPP_VATFormBuilder"
Option Compare Database
Option Explicit

'============================================================
' SUPPLIER CHECKING TOOL - HMRC VAT FORM BUILDER
'
' Creates:
'   Form: SUPP_frmVATCheck
'   Main form button: SUPP_cmdOpenVATCheck
'
' Uses the existing CheckHMRCVAT function in Module9.
' Uses archive fields NewVATNo / CurrentVATNo.
' Does not alter the existing Companies House button or form.
'
' Run once:
'   SUPP_BuildVATCheckForm
'============================================================

Private Const SUPP_VAT_FORM As String = "SUPP_frmVATCheck"
Private Const SUPP_MAIN_FORM As String = "Main"
Private Const SUPP_ARCHIVE_TABLE As String = "SUPP_tblSupplierArchive"

Private Const SUPP_MARGIN As Long = 360
Private Const SUPP_GAP As Long = 180
Private Const SUPP_LABEL_H As Long = 240
Private Const SUPP_TEXT_H As Long = 390
Private Const SUPP_BUTTON_H As Long = 480

Public Sub SUPP_BuildVATCheckForm()

    On Error GoTo BuildError

    Dim frm As Access.Form
    Dim oldFormName As String
    Dim mainWasOpen As Boolean
    Dim stage As String

    stage = "Checking archive table"
    If Not SUPP_VAT_TableExists(SUPP_ARCHIVE_TABLE) Then
        Err.Raise vbObjectError + 7100, , _
            "The table " & SUPP_ARCHIVE_TABLE & " was not found."
    End If

    stage = "Ensuring VAT audit fields"
    SUPP_VAT_EnsureAuditFields

    stage = "Closing any existing VAT form"
    If CurrentProject.AllForms(SUPP_VAT_FORM).IsLoaded Then
        DoCmd.Close acForm, SUPP_VAT_FORM, acSaveYes
    End If

    stage = "Backing up the Main form"
    mainWasOpen = CurrentProject.AllForms(SUPP_MAIN_FORM).IsLoaded
    If mainWasOpen Then DoCmd.Close acForm, SUPP_MAIN_FORM, acSaveYes
    SUPP_VAT_BackupForm SUPP_MAIN_FORM

    stage = "Creating the VAT form"
    Set frm = CreateForm
    oldFormName = frm.Name

    SUPP_VAT_ConfigureForm frm

    stage = "Saving and naming the VAT form"
    DoCmd.Save acForm, oldFormName
    DoCmd.Close acForm, oldFormName, acSaveYes

    If StrComp(oldFormName, SUPP_VAT_FORM, vbTextCompare) <> 0 Then
        If SUPP_VAT_FormExists(SUPP_VAT_FORM) Then
            DoCmd.DeleteObject acForm, SUPP_VAT_FORM
        End If
        DoCmd.Rename SUPP_VAT_FORM, acForm, oldFormName
    End If

    stage = "Adding the VAT button to Main"
    DoCmd.OpenForm SUPP_MAIN_FORM, acDesign, , , , acHidden
    SUPP_VAT_AddMainButton Forms(SUPP_MAIN_FORM)
    DoCmd.Close acForm, SUPP_MAIN_FORM, acSaveYes

    MsgBox "HMRC VAT form created successfully." & vbCrLf & vbCrLf & _
           "Form: " & SUPP_VAT_FORM & vbCrLf & _
           "Button added to: " & SUPP_MAIN_FORM & vbCrLf & vbCrLf & _
           "The existing Companies House validation button has not been changed.", _
           vbInformation, "Supplier VAT Form Builder"

    Exit Sub

BuildError:
    Dim errNo As Long
    Dim errText As String
    errNo = Err.Number
    errText = Err.Description

    On Error Resume Next
    If Not frm Is Nothing Then
        If CurrentProject.AllForms(oldFormName).IsLoaded Then _
            DoCmd.Close acForm, oldFormName, acSaveNo
    End If
    If CurrentProject.AllForms(SUPP_MAIN_FORM).IsLoaded Then _
        DoCmd.Close acForm, SUPP_MAIN_FORM, acSaveNo

    MsgBox "The VAT form builder stopped." & vbCrLf & vbCrLf & _
           "Stage: " & stage & vbCrLf & _
           "Error " & errNo & ": " & errText, _
           vbCritical, "Supplier VAT Form Builder"
End Sub

Private Sub SUPP_VAT_ConfigureForm(ByVal frm As Access.Form)

    Dim y As Long
    Dim formWidth As Long
    Dim fullWidth As Long
    Dim halfWidth As Long
    Dim rightX As Long
    Dim ctl As Object

    formWidth = 15000
    fullWidth = formWidth - (SUPP_MARGIN * 2)
    halfWidth = (fullWidth - SUPP_GAP) \ 2
    rightX = SUPP_MARGIN + halfWidth + SUPP_GAP

    frm.Caption = "Supplier VAT Verification"
    frm.RecordSource = vbNullString
    frm.DefaultView = acDefViewSingle
    frm.Width = formWidth
    frm.PopUp = True
    frm.Modal = True
    frm.AutoCenter = True
    frm.AutoResize = True
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.DividingLines = False
    frm.ScrollBars = 2
    frm.OnLoad = "=SUPP_VATFormLoad()"

    y = 360
    SUPP_VAT_AddLabel frm, "SUPP_lblTitle", _
        "HMRC VAT VERIFICATION", SUPP_MARGIN, y, fullWidth, 420, True
    y = y + 600

    SUPP_VAT_AddLabel frm, "SUPP_lblRequestRef", _
        "SUPPLIER REQUEST", SUPP_MARGIN, y, 2100, SUPP_LABEL_H, False
    Set ctl = SUPP_VAT_AddTextBox(frm, "SUPP_txtRequestRef", _
        SUPP_MARGIN + 2250, y - 45, fullWidth - 2250, SUPP_TEXT_H, False)
    SUPP_VAT_StyleTextBox ctl, True
    y = y + 600

    SUPP_VAT_AddLabel frm, "SUPP_lblSupplierName", _
        "SUPPLIER NAME", SUPP_MARGIN, y, halfWidth, SUPP_LABEL_H, False
    SUPP_VAT_AddLabel frm, "SUPP_lblVATNumber", _
        "VAT NUMBER", rightX, y, halfWidth, SUPP_LABEL_H, False
    y = y + 270

    Set ctl = SUPP_VAT_AddTextBox(frm, "SUPP_txtSupplierName", _
        SUPP_MARGIN, y, halfWidth, SUPP_TEXT_H, False)
    SUPP_VAT_StyleTextBox ctl, True
    Set ctl = SUPP_VAT_AddTextBox(frm, "SUPP_txtVATNumber", _
        rightX, y, halfWidth, SUPP_TEXT_H, False)
    SUPP_VAT_StyleTextBox ctl, False
    y = y + 650

    SUPP_VAT_AddLabel frm, "SUPP_lblSupplierAddress", _
        "ADDRESS SUBMITTED BY SUPPLIER", SUPP_MARGIN, y, halfWidth, SUPP_LABEL_H, False
    SUPP_VAT_AddLabel frm, "SUPP_lblRegisteredAddress", _
        "HMRC REGISTERED ADDRESS", rightX, y, halfWidth, SUPP_LABEL_H, False
    y = y + 270

    Set ctl = SUPP_VAT_AddTextBox(frm, "SUPP_txtSupplierAddress", _
        SUPP_MARGIN, y, halfWidth, 1050, True)
    SUPP_VAT_StyleTextBox ctl, True
    Set ctl = SUPP_VAT_AddTextBox(frm, "SUPP_txtRegisteredAddress", _
        rightX, y, halfWidth, 1050, True)
    SUPP_VAT_StyleTextBox ctl, True
    y = y + 1350

    SUPP_VAT_AddLabel frm, "SUPP_lblHMRCName", _
        "REGISTERED BUSINESS NAME", SUPP_MARGIN, y, fullWidth, SUPP_LABEL_H, False
    y = y + 270
    Set ctl = SUPP_VAT_AddTextBox(frm, "SUPP_txtRegisteredName", _
        SUPP_MARGIN, y, fullWidth, SUPP_TEXT_H, False)
    SUPP_VAT_StyleTextBox ctl, True
    y = y + 650

    SUPP_VAT_AddLabel frm, "SUPP_lblNameMatch", _
        "NAME COMPARISON", SUPP_MARGIN, y, fullWidth, SUPP_LABEL_H, False
    y = y + 270
    Set ctl = SUPP_VAT_AddTextBox(frm, "SUPP_txtNameComparison", _
        SUPP_MARGIN, y, fullWidth, SUPP_TEXT_H, False)
    SUPP_VAT_StyleTextBox ctl, True
    y = y + 650

    SUPP_VAT_AddLabel frm, "SUPP_lblStatus", _
        "HMRC RESULT", SUPP_MARGIN, y, fullWidth, SUPP_LABEL_H, True
    y = y + 270
    Set ctl = SUPP_VAT_AddTextBox(frm, "SUPP_txtVATStatus", _
        SUPP_MARGIN, y, fullWidth, 570, False)
    SUPP_VAT_StyleTextBox ctl, True
    ctl.FontSize = 14
    ctl.FontBold = True
    y = y + 780

    SUPP_VAT_AddLabel frm, "SUPP_lblDetails", _
        "CHECK DETAILS", SUPP_MARGIN, y, fullWidth, SUPP_LABEL_H, False
    y = y + 270
    Set ctl = SUPP_VAT_AddTextBox(frm, "SUPP_txtHMRCDetails", _
        SUPP_MARGIN, y, fullWidth, 1050, True)
    SUPP_VAT_StyleTextBox ctl, True
    y = y + 1350

    SUPP_VAT_AddButton frm, "SUPP_cmdRunVATCheck", _
        "CHECK VAT WITH HMRC", SUPP_MARGIN, y, 3300, _
        "=SUPP_RunVATCheck()"
    SUPP_VAT_AddButton frm, "SUPP_cmdCloseVATForm", _
        "CLOSE", SUPP_MARGIN + 3600, y, 1800, _
        "=SUPP_CloseVATForm()"

    frm.Section(acDetail).Height = y + SUPP_BUTTON_H + 450

End Sub

Private Sub SUPP_VAT_AddMainButton(ByVal frm As Access.Form)

    Dim ctl As Object
    Dim referenceButton As Object
    Dim buttonLeft As Long
    Dim buttonTop As Long

    Set ctl = SUPP_VAT_GetControl(frm, "SUPP_cmdOpenVATCheck")

    If ctl Is Nothing Then
        Set referenceButton = SUPP_VAT_GetControl(frm, "ButtonVal")

        If Not referenceButton Is Nothing Then
            buttonLeft = referenceButton.Left + referenceButton.Width + 240
            buttonTop = referenceButton.Top
        Else
            buttonLeft = SUPP_MARGIN
            buttonTop = frm.Section(acDetail).Height + 120
            frm.Section(acDetail).Height = buttonTop + SUPP_BUTTON_H + 240
        End If

        Set ctl = Application.CreateControl(frm.Name, acCommandButton, _
            acDetail, , "SUPP_cmdOpenVATCheck", buttonLeft, buttonTop, 2700, SUPP_BUTTON_H)
    End If

    ctl.Caption = "CHECK VAT"
    ctl.OnClick = "=SUPP_OpenVATCheck()"
    ctl.Visible = True
    ctl.FontName = "Calibri"
    ctl.FontSize = 9
    ctl.FontBold = True

End Sub

Private Function SUPP_VAT_AddTextBox( _
    ByVal frm As Access.Form, ByVal controlName As String, _
    ByVal leftPos As Long, ByVal topPos As Long, _
    ByVal controlWidth As Long, ByVal controlHeight As Long, _
    ByVal multiLine As Boolean) As Object

    Dim ctl As Object
    Set ctl = Application.CreateControl(frm.Name, acTextBox, acDetail, , _
        controlName, leftPos, topPos, controlWidth, controlHeight)

    ctl.Name = controlName
    ctl.ControlSource = vbNullString
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = controlWidth
    ctl.Height = controlHeight
    ctl.EnterKeyBehavior = IIf(multiLine, 1, 0)
    ctl.ScrollBars = IIf(multiLine, 2, 0)
    ctl.TabStop = False

    Set SUPP_VAT_AddTextBox = ctl
End Function

Private Sub SUPP_VAT_AddLabel( _
    ByVal frm As Access.Form, ByVal controlName As String, _
    ByVal captionText As String, ByVal leftPos As Long, _
    ByVal topPos As Long, ByVal controlWidth As Long, _
    ByVal controlHeight As Long, ByVal heading As Boolean)

    Dim ctl As Object
    Set ctl = Application.CreateControl(frm.Name, acLabel, acDetail, , _
        controlName, leftPos, topPos, controlWidth, controlHeight)

    ctl.Name = controlName
    ctl.Caption = captionText
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = controlWidth
    ctl.Height = controlHeight
    ctl.FontName = "Calibri"
    ctl.FontSize = IIf(heading, 11, 9)
    ctl.FontBold = heading
    ctl.ForeColor = RGB(55, 70, 85)
    ctl.BackStyle = 0
End Sub

Private Sub SUPP_VAT_AddButton( _
    ByVal frm As Access.Form, ByVal controlName As String, _
    ByVal captionText As String, ByVal leftPos As Long, _
    ByVal topPos As Long, ByVal controlWidth As Long, _
    ByVal onClickExpression As String)

    Dim ctl As Object
    Set ctl = Application.CreateControl(frm.Name, acCommandButton, acDetail, , _
        controlName, leftPos, topPos, controlWidth, SUPP_BUTTON_H)

    ctl.Name = controlName
    ctl.Caption = captionText
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = controlWidth
    ctl.Height = SUPP_BUTTON_H
    ctl.OnClick = onClickExpression
    ctl.FontName = "Calibri"
    ctl.FontSize = 9
    ctl.FontBold = True
End Sub

Private Sub SUPP_VAT_StyleTextBox(ByVal ctl As Object, ByVal locked As Boolean)
    ctl.FontName = "Calibri"
    ctl.FontSize = 10
    ctl.BackColor = RGB(248, 250, 252)
    ctl.ForeColor = RGB(35, 48, 62)
    ctl.BorderStyle = 1
    ctl.SpecialEffect = 0
    ctl.Locked = locked
    ctl.Enabled = True
    ctl.TabStop = Not locked
End Sub

Private Function SUPP_VAT_GetControl(ByVal frm As Access.Form, _
                                     ByVal controlName As String) As Object
    On Error Resume Next
    Set SUPP_VAT_GetControl = frm.Controls(controlName)
    On Error GoTo 0
End Function

Private Function SUPP_VAT_FormExists(ByVal formName As String) As Boolean
    On Error Resume Next
    SUPP_VAT_FormExists = (CurrentProject.AllForms(formName).Name = formName)
    On Error GoTo 0
End Function

Private Function SUPP_VAT_TableExists(ByVal tableName As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error Resume Next
    Set tdf = CurrentDb.TableDefs(tableName)
    SUPP_VAT_TableExists = (Err.Number = 0 And Not tdf Is Nothing)
    Err.Clear
    Set tdf = Nothing
    On Error GoTo 0
End Function

Private Sub SUPP_VAT_EnsureAuditFields()

    Dim tdf As DAO.TableDef
    Set tdf = CurrentDb.TableDefs(SUPP_ARCHIVE_TABLE)

    SUPP_VAT_EnsureField tdf, "SUPP_VATStatus", dbText, 30
    SUPP_VAT_EnsureField tdf, "SUPP_VATRegisteredName", dbText, 255
    SUPP_VAT_EnsureField tdf, "SUPP_VATRegisteredAddress", dbMemo
    SUPP_VAT_EnsureField tdf, "SUPP_VATCheckDetails", dbMemo
    SUPP_VAT_EnsureField tdf, "SUPP_VATCheckedDate", dbDate
End Sub

Private Sub SUPP_VAT_EnsureField(ByVal tdf As DAO.TableDef, _
                                 ByVal fieldName As String, _
                                 ByVal fieldType As DAO.DataTypeEnum, _
                                 Optional ByVal fieldSize As Long = 0)

    Dim fld As DAO.Field
    For Each fld In tdf.Fields
        If StrComp(fld.Name, fieldName, vbTextCompare) = 0 Then Exit Sub
    Next fld

    If fieldSize > 0 Then
        tdf.Fields.Append tdf.CreateField(fieldName, fieldType, fieldSize)
    Else
        tdf.Fields.Append tdf.CreateField(fieldName, fieldType)
    End If
End Sub

Private Sub SUPP_VAT_BackupForm(ByVal formName As String)

    Dim backupPath As String
    backupPath = CurrentProject.Path & "\SUPP_FormBackup_" & formName & "_" & _
                 Format$(Now(), "yyyymmdd_hhnnss") & ".txt"
    Application.SaveAsText acForm, formName, backupPath
End Sub

'============================================================
' PUBLIC FORM ACTIONS
' These are expression-based event handlers. No form class
' module or Trust Center VBA-project access is required.
'============================================================

Public Function SUPP_OpenVATCheck() As Boolean

    On Error GoTo OpenError

    Dim archiveID As Long

    If DCount("*", SUPP_ARCHIVE_TABLE) = 0 Then
        MsgBox "There are no imported supplier requests to check.", _
               vbInformation, "Supplier VAT Check"
        Exit Function
    End If

    archiveID = CLng(DMax("ID", SUPP_ARCHIVE_TABLE))

    If CurrentProject.AllForms(SUPP_VAT_FORM).IsLoaded Then
        DoCmd.Close acForm, SUPP_VAT_FORM, acSaveNo
    End If

    DoCmd.OpenForm SUPP_VAT_FORM, acNormal, , , acFormReadOnly, _
                   acWindowNormal, CStr(archiveID)

    SUPP_OpenVATCheck = True
    Exit Function

OpenError:
    MsgBox "Could not open the VAT verification form." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Supplier VAT Check"
End Function

Public Function SUPP_VATFormLoad() As Boolean

    On Error GoTo LoadError

    Dim frm As Access.Form
    Dim archiveID As Long
    Dim rs As DAO.Recordset
    Dim vatNumber As String
    Dim supplierName As String
    Dim supplierAddress As String

    Set frm = Forms(SUPP_VAT_FORM)

    If Not IsNumeric(Nz(frm.OpenArgs, vbNullString)) Then
        Err.Raise vbObjectError + 7110, , "Supplier archive ID was not supplied."
    End If

    archiveID = CLng(frm.OpenArgs)

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT * FROM [" & SUPP_ARCHIVE_TABLE & "] WHERE [ID]=" & archiveID & ";", _
        dbOpenSnapshot)

    If rs.EOF Then Err.Raise vbObjectError + 7111, , _
        "Supplier archive record " & archiveID & " was not found."

    frm.Controls("SUPP_txtRequestRef").Value = Nz(rs.Fields("Title").Value, "")
    supplierName = SUPP_VAT_ReadFirstField(rs, _
        Array("SupplierKnownAsName", "AXSupplierName", "CompanyName"))
    vatNumber = SUPP_VAT_ReadFirstField(rs, Array("NewVATNo", "CurrentVATNo"))
    supplierAddress = SUPP_GetSubmittedAddress(archiveID)

    frm.Controls("SUPP_txtSupplierName").Value = supplierName
    frm.Controls("SUPP_txtVATNumber").Value = vatNumber
    frm.Controls("SUPP_txtSupplierAddress").Value = supplierAddress
    frm.Controls("SUPP_txtRegisteredName").Value = Null
    frm.Controls("SUPP_txtRegisteredAddress").Value = Null
    frm.Controls("SUPP_txtNameComparison").Value = "NOT CHECKED"
    frm.Controls("SUPP_txtVATStatus").Value = "NOT CHECKED"
    frm.Controls("SUPP_txtHMRCDetails").Value = _
        "Review the supplier details, then select CHECK VAT WITH HMRC."

    If Len(Trim$(vatNumber)) = 0 Then
        frm.Controls("SUPP_txtVATStatus").Value = "VAT NUMBER MISSING"
        frm.Controls("SUPP_txtHMRCDetails").Value = _
            "No VAT number was found in NewVATNo or CurrentVATNo."
    End If

    SUPP_VATFormLoad = True

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set frm = Nothing
    Exit Function

LoadError:
    MsgBox "Could not load supplier details for the VAT check." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Supplier VAT Check"
    Resume CleanExit
End Function

Public Function SUPP_RunVATCheck() As Boolean

    On Error GoTo CheckError

    Dim frm As Access.Form
    Dim archiveID As Long
    Dim vatNumber As String
    Dim registeredName As String
    Dim registeredAddress As String
    Dim details As String
    Dim result As String
    Dim submittedName As String
    Dim nameComparison As String
    Dim rs As DAO.Recordset

    Set frm = Forms(SUPP_VAT_FORM)
    archiveID = CLng(frm.OpenArgs)
    vatNumber = Nz(frm.Controls("SUPP_txtVATNumber").Value, "")
    submittedName = Nz(frm.Controls("SUPP_txtSupplierName").Value, "")

    frm.Controls("SUPP_txtVATStatus").Value = "CHECK IN PROGRESS"
    frm.Controls("SUPP_txtHMRCDetails").Value = "Contacting HMRC..."
    frm.Repaint

    result = CheckHMRCVAT(vatNumber, registeredName, registeredAddress, details)

    If result = "VALID" Then
        If SUPP_VAT_NormaliseName(submittedName) = _
           SUPP_VAT_NormaliseName(registeredName) Then
            nameComparison = "MATCH"
        Else
            nameComparison = "REVIEW REQUIRED"
        End If
    Else
        nameComparison = "NOT CHECKED"
    End If

    frm.Controls("SUPP_txtVATStatus").Value = result
    frm.Controls("SUPP_txtRegisteredName").Value = registeredName
    frm.Controls("SUPP_txtRegisteredAddress").Value = registeredAddress
    frm.Controls("SUPP_txtNameComparison").Value = nameComparison
    frm.Controls("SUPP_txtHMRCDetails").Value = details

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT * FROM [" & SUPP_ARCHIVE_TABLE & "] WHERE [ID]=" & archiveID & ";", _
        dbOpenDynaset)

    If Not rs.EOF Then
        rs.Edit
        rs.Fields("SUPP_VATStatus").Value = result
        rs.Fields("SUPP_VATRegisteredName").Value = SUPP_VAT_NullIfBlank(registeredName)
        rs.Fields("SUPP_VATRegisteredAddress").Value = SUPP_VAT_NullIfBlank(registeredAddress)
        rs.Fields("SUPP_VATCheckDetails").Value = SUPP_VAT_NullIfBlank(details)
        rs.Fields("SUPP_VATCheckedDate").Value = Now()
        rs.Update
    End If

    SUPP_RunVATCheck = True

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set frm = Nothing
    Exit Function

CheckError:
    On Error Resume Next
    If Not frm Is Nothing Then
        frm.Controls("SUPP_txtVATStatus").Value = "UNABLE TO VERIFY"
        frm.Controls("SUPP_txtHMRCDetails").Value = _
            "VAT check could not be completed. Error " & Err.Number & ": " & Err.Description
    End If
    MsgBox "The HMRC VAT check could not be completed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Supplier VAT Check"
    Resume CleanExit
End Function

Public Function SUPP_CloseVATForm() As Boolean
    On Error Resume Next
    DoCmd.Close acForm, SUPP_VAT_FORM, acSaveNo
    SUPP_CloseVATForm = True
End Function

Private Function SUPP_VAT_ReadFirstField(ByVal rs As DAO.Recordset, _
                                         ByVal fieldNames As Variant) As String
    Dim i As Long
    For i = LBound(fieldNames) To UBound(fieldNames)
        On Error Resume Next
        SUPP_VAT_ReadFirstField = Nz(rs.Fields(CStr(fieldNames(i))).Value, "")
        If Err.Number = 0 And Len(Trim$(SUPP_VAT_ReadFirstField)) > 0 Then
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    Next i
End Function

Private Function SUPP_VAT_NormaliseName(ByVal value As String) As String
    value = UCase$(Trim$(value))
    value = Replace(value, "&", " AND ")
    value = Replace(value, ".", "")
    value = Replace(value, ",", "")
    value = Replace(value, "'", "")
    value = Replace(value, "-", " ")
    Do While InStr(value, "  ") > 0
        value = Replace(value, "  ", " ")
    Loop
    SUPP_VAT_NormaliseName = Trim$(value)
End Function

Private Function SUPP_VAT_NullIfBlank(ByVal value As String) As Variant
    If Len(Trim$(value)) = 0 Then
        SUPP_VAT_NullIfBlank = Null
    Else
        SUPP_VAT_NullIfBlank = value
    End If
End Function
