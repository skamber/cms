Attribute VB_Name = "ModInvoice"
Public mInvoiceItemKey As String
Public mInvoiceItemMode As String

Public Function CheckInvoiceSecurity(sFunction As String) As Boolean
Dim privilege_ctr As Byte
CheckInvoiceSecurity = False
   privilege_ctr = GetPrivileges("Invoice")
    If privilege_ctr = 0 Then
    MsgBox "Invalid access - Action not available for current user access level.", vbExclamation
    Exit Function
    End If
    'Userprivilege.Bookmark = PrivilegeBookMark
    Select Case sFunction
        
        
        Case RECORD_NEW:
        
            If Permissions(privilege_ctr, 2) = "N" Then
                MsgBox "Invalid access - Function not available for current user access level.", vbExclamation
                Exit Function
             ElseIf Permissions(privilege_ctr, 2) = "Y" Then CheckInvoiceSecurity = True
           End If
        Case RECORD_EDIT
            If Permissions(privilege_ctr, 3) = "N" Then
                MsgBox "Invalid access - Function not available for current user access level.", vbExclamation
                Exit Function
             ElseIf Permissions(privilege_ctr, 3) = "Y" Then CheckInvoiceSecurity = True
           End If
        Case RECORD_DELETE
            If Permissions(privilege_ctr, 4) = "N" Then
                MsgBox "Invalid access - Function not available for current user access level.", vbExclamation
                Exit Function
             ElseIf Permissions(privilege_ctr, 4) = "Y" Then CheckInvoiceSecurity = True
           End If
        End Select
End Function

Public Function InitialiseInvoice()
On Error GoTo ErrorHandler

    Dim ctrl As Control


    For Each ctrl In frmInvoice.Controls
        
        If TypeOf ctrl Is TextBox Then ctrl.Text = ""
        If TypeOf ctrl Is ComboBox Then ctrl.ListIndex = -1
        If TypeOf ctrl Is MaskEdBox Then ctrl.Text = ""
        Set ctrl = Nothing
        
        frmInvoice.InvoiceItemList.ListItems.Clear
        
    Next ctrl
    DoEvents
    

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMember", "InitialiseMember", True)

End Function
Public Function SaveInvoice() As Boolean
On Error GoTo ErrorHandler

    Dim objInvoice As CMSInvoice.clsInvoice
    Dim objInvoice_s As CMSInvoice.clsInvoice_s

    SaveInvoice = False
                            
    'Member record
    Set objInvoice = New CMSInvoice.clsInvoice
    Set objInvoice_s = New CMSInvoice.clsInvoice_s
    Set objInvoice_s.DatabaseConnection = objConnection


    PopulateInvoiceObject objInvoice

    'Insert or Update record
    If gRecordMode = RECORD_NEW Then
    
        objInvoice_s.InsertInvoice objInvoice
        gInvoiceId = objInvoice_s.NewInvoiceId
        frmInvoice.txtInvoiceNo = objInvoice_s.NewInvoiceNo
        frmInvoice.txtRef = objInvoice_s.NewRefNo
        
    ElseIf (gRecordMode = RECORD_EDIT) Or (gRecordMode = RECORD_READ) Then
        
        objInvoice.Id = gInvoiceId
        objInvoice_s.UpdateInvoice objInvoice
        
    
  '  ElseIf gRecordMode = RECORD_DELETE Then
        
   '     objMember_s.Deletemember gmemberId
        
    End If
    
    SaveInvoice = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modInvoice", "SaveInvoice", True)
    
End Function

Public Function ValidateInvoice() As Boolean

On Error GoTo ErrorHandler

    ValidateInvoice = False

    With frmInvoice

           
           If Trim(.txtName.Text) = "" Then
                MsgBox "Required field missing - Name  must be entered.", vbExclamation
                .txtName.SetFocus
                Exit Function
            End If
                
            If .cmbTerms.Text = "" Then
                MsgBox "Required field missing - Terms must be selected.", vbExclamation
                .cmbTerms.SetFocus
                Exit Function
            End If
        
    End With

    ValidateInvoice = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modInvoice", "ValidateInvoice", True)

End Function

Public Function LoadInvoiceDefualtValue()

With frmInvoice
       .dteInvoiceDate.Text = Format(Now(), DATE_FORMAT)
End With

End Function

Public Function LoadInvoiceComboBox()
On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rslocal As ADODB.Recordset
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection
    
    'Payment
    Set rslocal = objOrganisation_s.getTerms

    With frmInvoice
            
            .cmbTerms.Clear
            If Not rslocal Is Nothing Then
                Do Until rslocal.EOF
                    .cmbTerms.AddItem rslocal!TermName
                    .cmbTerms.ItemData(.cmbTerms.NewIndex) = rslocal!Id
                    rslocal.MoveNext
                Loop
                Set rslocal = Nothing
            End If
     Set rslocal = objOrganisation_s.getInvoiceItemType
            .CmbItemDescription.Clear
            If Not rslocal Is Nothing Then
                Do Until rslocal.EOF
                    .CmbItemDescription.AddItem rslocal!Descriprtion
                    .CmbItemDescription.ItemData(.CmbItemDescription.NewIndex) = rslocal!Id
                    rslocal.MoveNext
                Loop
                Set rslocal = Nothing
            End If
    End With

    
    Set objOrganisation_s = Nothing



Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modInvoice", "LoadInvoiceComboBox", True)

End Function

Public Function PopulateInvoiceObject(objInvoice As CMSInvoice.clsInvoice)
On Error GoTo ErrorHandler
'save data from form to newly created object for insert or update

   


    With frmInvoice
        If Trim(.txtInvoiceNo.Text) <> "" Then objInvoice.Invoice_no = .txtInvoiceNo.Text
        If Trim(.txtRef.Text) <> "" Then objInvoice.Ref = .txtRef.Text
        If Trim(.txtName.Text) <> "" Then objInvoice.Name1 = .txtName.Text
        If Trim(.txtSurname.Text) <> "" Then objInvoice.name2 = .txtSurname.Text
        If Trim(.txtAddress1.Text) <> "" Then objInvoice.address1 = .txtAddress1.Text
        If Trim(.txtAddress2.Text) <> "" Then objInvoice.address2 = .txtAddress2.Text
        If Trim(.txtPostcode.Text) <> "" Then objInvoice.Address3 = .txtPostcode.Text
        If Trim(.txtMobile.Text) <> "" Then objInvoice.Mobile = .txtMobile.Text
        If Trim(.txtBalance.Text) <> "" Then objInvoice.balance = .txtBalance.Text
        If .dteInvoiceDate.Text <> "" Then objInvoice.dateofInvoice = .dteInvoiceDate.FormattedText
        If .dteOverDueDate.Text <> "" Then objInvoice.over_due_date = .dteOverDueDate.FormattedText
        If .dtePhone.Text <> "" Then objInvoice.Phone = .dtePhone.FormattedText
        If Trim(.txtState) <> "" Then objInvoice.State = .txtState
        If Trim(.txtTotalAmount.Text) <> "" Then objInvoice.Total_Amount = .txtTotalAmount.Text
        If Trim(.cmbTerms.Text) <> "" Then objInvoice.Terms = .cmbTerms.Text
        objInvoice.ChurchId = gChurchId
        
    End With
    
    
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMember", "PopulateInvoiceObject", True)

End Function

Public Sub GenerateInvoiceList(sql As String)
 On Error GoTo ErrorHandler
 
    ' Dim objFollowup_s As PACMSFollowUP.clsFollowup_s
     Dim rslocal As ADODB.Recordset
     Dim strId  As String
     Dim itmx As ListItem
     
    ' Dim lngTotalProspect As Long
    ' Dim lngTotalPlanning As Long
    ' Dim lngTotalCoaching As Long
 
 
     Screen.MousePointer = vbHourglass
 
     With frmInvoiceSearch
             .ListInvoiceView.ListItems.Clear
     
         '==============================================================================
         'get Prospect list
            Set rslocal = New ADODB.Recordset
            rslocal.Open sql, objConnection, adOpenForwardOnly, adLockReadOnly
             If Not rslocal.EOF Then
                 Do While Not rslocal.EOF
                     strId = CStr(rslocal!Id)
                     Set itmx = .ListInvoiceView.ListItems.Add()
                                     itmx.Key = "#" & strId
                                     itmx.Text = CStr(rslocal!Id)
                                     If Not IsNull(rslocal!Invoice_no) Then itmx.SubItems(1) = CStr(rslocal!Invoice_no)
                                     If Not IsNull(rslocal!Ref) Then itmx.SubItems(2) = CStr(rslocal!Ref)
                                     If Not IsNull(rslocal!Name1) Then itmx.SubItems(3) = CStr(rslocal!Name1)
                                     If Not IsNull(rslocal!name2) Then itmx.SubItems(4) = CStr(rslocal!name2)
                                     If Not IsNull(rslocal!Created_date) Then itmx.SubItems(5) = CStr(rslocal!Created_date)
                                     If Not IsNull(rslocal!over_due_date) Then itmx.SubItems(6) = CStr(rslocal!over_due_date)
                                     If Not IsNull(rslocal!Total_Amount) Then itmx.SubItems(7) = CStr(rslocal!Total_Amount)
                                     
                     Set itmx = Nothing
                     rslocal.MoveNext
                     
                 Loop
                 InvoiceSelected = True
                 Set rslocal = Nothing
             Else
               InvoiceSelected = False
             End If
     End With
     
     Screen.MousePointer = vbDefault
 
 
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modInvoice", "GenerateInvoiceList", True)
 
 End Sub

Public Sub SetupInvoiceItemList()
On Error GoTo ErrorHandler

    'Planning/Coaching
    With frmInvoice.InvoiceItemList
    
            .View = lvwReport
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Item Number", 1400, lvwColumnLeft
            .ColumnHeaders.Add , , "Invoice Number", 1400, lvwColumnLeft
            .ColumnHeaders.Add , , "Description", 4000, lvwColumnLeft
            .ColumnHeaders.Add , , "Amount", 1000, lvwColumnLeft
            .ColumnHeaders.Add , , "Gst Amount", 1300, lvwColumnLeft
            .ColumnHeaders.Add , , "total Amount", 1400, lvwColumnLeft
            .ListItems.Clear
            
    End With
    DoEvents


Exit Sub
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modInvoice", "SetupInvoiceItemList", True)

End Sub


Public Function InitialiseInvoiceItem()
On Error GoTo ErrorHandler

    With frmInvoice
           
            .CmbItemDescription.ListIndex = -1
            .txtItemAmount.Text = ""
            .txtItemGST.Text = ""
            .txtTotalItemAmount = ""
    End With
    DoEvents

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "ModInvoice", "InitialiseInvoiceItem", True)
    
End Function

Public Sub GetGstAndTotalAmount()
 On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rslocal As ADODB.Recordset
    Dim num1, num2 As Double
    Dim Found_item  As Boolean
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection
    Found_item = False
    'Payment
     With frmInvoice
     Set rslocal = objOrganisation_s.getInvoiceItemType
            If Not rslocal Is Nothing Then
                Do Until rslocal.EOF
                    If rslocal!Descriprtion = .CmbItemDescription.Text Then
                    'If rslocal!Id = .CmbItemDescription.ItemData(.CmbItemDescription.NewIndex) Then
                    num1 = rslocal!gst * .txtItemAmount.Text
                    Found_item = True
                    .txtItemGST.Text = num1
                    Call ValidNumericEntry(.txtItemGST)
                    num2 = .txtItemAmount.Text
                    num2 = num2 + num1
                    .txtTotalItemAmount.Text = num2
                    Call ValidNumericEntry(.txtItemAmount)
                    Set rslocal = Nothing
                    Exit Sub
                    
                    End If
                    rslocal.MoveNext
                Loop
                Set rslocal = Nothing
            End If
    If Found_item = False Then
    MsgBox "Invalid Item - Invoice Item must be define .", vbExclamation
    .CmbItemDescription.SetFocus
    Set objOrganisation_s = Nothing
    Exit Sub
    End If
    
    End With
    
    Set objOrganisation_s = Nothing

Exit Sub
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modInvoice", "LoadInvoiceComboBox", True)


End Sub

Public Function SaveInvoiceItem() As Boolean
On Error GoTo ErrorHandler

    Dim objInvoiceItem As CMSInvoiceItem.clsInvoiceItem
    Dim objInvoiceItem_s As CMSInvoiceItem.clsInvoiceItem_s

    SaveInvoiceItem = False
                            
    'Member record
    Set objInvoiceItem = New CMSInvoiceItem.clsInvoiceItem
    Set objInvoiceItem_s = New CMSInvoiceItem.clsInvoiceItem_s
    Set objInvoiceItem_s.DatabaseConnection = objConnection


    PopulateInvoiceItemObject objInvoiceItem

    'Insert or Update record
    If mInvoiceItemMode = RECORD_NEW Then
    
        objInvoiceItem_s.InsertInvoiceItem objInvoiceItem
        gInvoiceItemId = objInvoiceItem_s.NewInvoiceItemId
        
    ElseIf gRecordMode = RECORD_EDIT Then
        
        objInvoiceItem.Item_Id = gInvoiceItemId
        objInvoiceItem_s.UpdateInvoiceItem objInvoiceItem
        
    
  '  ElseIf gRecordMode = RECORD_DELETE Then
        
   '     objMember_s.Deletemember gmemberId
        
    End If
    
    SaveInvoiceItem = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modInvoice", "SaveInvoiceitem", True)
    
End Function

Public Function PopulateInvoiceItemObject(objInvoiceItem As CMSInvoiceItem.clsInvoiceItem)
On Error GoTo ErrorHandler
'save data from form to newly created object for insert or update

   


    With frmInvoice
        objInvoiceItem.Invoice_Id = gInvoiceId
        If Trim(.txtItemAmount.Text) <> "" Then objInvoiceItem.Amount = .txtItemAmount.Text
        If Trim(.CmbItemDescription.Text) <> "" Then objInvoiceItem.Description = .CmbItemDescription.Text
        If Trim(.txtItemGST.Text) <> "" Then objInvoiceItem.GstAmount = .txtItemGST.Text
        If Trim(.txtTotalItemAmount.Text) <> "" Then objInvoiceItem.Total_Amount = .txtTotalItemAmount.Text
        
    End With
    
    
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modInvoice", "PopulateInvoiceItemObject", True)

End Function

Sub displayInvoiceItemList()
 On Error GoTo ErrorHandler
 
    ' Dim objFollowup_s As PACMSFollowUP.clsFollowup_s
     Dim objInvoiceItem_s As CMSInvoiceItem.clsInvoiceItem_s
     Dim rslocal As ADODB.Recordset
     Dim strId  As String
     Dim itmx As ListItem
     
     Dim TotalItemsAmount, BalanceAmount As Double
    Set objInvoiceItem_s = New CMSInvoiceItem.clsInvoiceItem_s
    Set objInvoiceItem_s.DatabaseConnection = objConnection
 
    ' Dim lngTotalProspect As Long
    ' Dim lngTotalPlanning As Long
    ' Dim lngTotalCoaching As Long
 Set rslocal = objInvoiceItem_s.getByInvoiceId(gInvoiceId)
     TotalItemsAmount = 0
     Screen.MousePointer = vbHourglass
 
     With frmInvoice
             .InvoiceItemList.ListItems.Clear
     If .txtBalance.Text <> "" Then
       BalanceAmount = .txtBalance.Text
     Else
       BalanceAmount = 0
     End If
         '==============================================================================
         'get Prospect list
             If Not rslocal Is Nothing Then
                 Do While Not rslocal.EOF
                     strId = CStr(rslocal!Item_Id)
                     Set itmx = .InvoiceItemList.ListItems.Add()
                                     itmx.Key = "#" & strId
                                     itmx.Text = CStr(rslocal!Item_Id)
                                     If Not IsNull(rslocal!Invoice_Id) Then itmx.SubItems(1) = CStr(rslocal!Invoice_Id)
                                     If Not IsNull(rslocal!Description) Then itmx.SubItems(2) = CStr(rslocal!Description)
                                     If Not IsNull(rslocal!Amount) Then itmx.SubItems(3) = CStr(rslocal!Amount)
                                     If Not IsNull(rslocal!gst) Then itmx.SubItems(4) = CStr(rslocal!gst)
                                     If Not IsNull(rslocal!Total_Amount) Then
                                       itmx.SubItems(5) = CStr(rslocal!Total_Amount)
                                       TotalItemsAmount = TotalItemsAmount + rslocal!Total_Amount
                                     End If
                     Set itmx = Nothing
                     rslocal.MoveNext
                     
                 Loop
                 InvoiceItemSelected = True
                 Set rslocal = Nothing
                 .txtTotalAmount.Text = TotalItemsAmount
                 Call ValidNumericEntry(.txtTotalAmount)
                .txtBalance.Text = TotalItemsAmount - getReceiptAmount(gInvoiceId)
                 Call ValidNumericEntry(.txtBalance)
                 If .txtBalance.Text <> BalanceAmount Then SaveInvoice
              Else
                 InvoiceItemSelected = False
             End If
     End With
     
     Screen.MousePointer = vbDefault
 
 
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modInvoice", "GenerateInvoiceItemList", True)
 
End Sub
Sub displayInvoiceSearchItemList()
 On Error GoTo ErrorHandler
 
    ' Dim objFollowup_s As PACMSFollowUP.clsFollowup_s
     Dim objInvoiceItem_s As CMSInvoiceItem.clsInvoiceItem_s
     Dim rslocal As ADODB.Recordset
     Dim strId  As String
     Dim itmx As ListItem
     
    Set objInvoiceItem_s = New CMSInvoiceItem.clsInvoiceItem_s
    Set objInvoiceItem_s.DatabaseConnection = objConnection
 
    ' Dim lngTotalProspect As Long
    ' Dim lngTotalPlanning As Long
    ' Dim lngTotalCoaching As Long
 Set rslocal = objInvoiceItem_s.getByInvoiceId(gInvoiceId)
    
     Screen.MousePointer = vbHourglass
 
     With frmInvoiceSearch
             .InvoiceItemList.ListItems.Clear
     
         '==============================================================================
         'get Prospect list
         If rslocal Is Nothing Then
         Screen.MousePointer = vbDefault
         Exit Sub
         End If
             If Not rslocal.EOF Then
                 Do While Not rslocal.EOF
                     strId = CStr(rslocal!Item_Id)
                     Set itmx = .InvoiceItemList.ListItems.Add()
                                     itmx.Key = "#" & strId
                                     itmx.Text = CStr(rslocal!Item_Id)
                                     If Not IsNull(rslocal!Invoice_Id) Then itmx.SubItems(1) = CStr(rslocal!Invoice_Id)
                                     If Not IsNull(rslocal!Description) Then itmx.SubItems(2) = CStr(rslocal!Description)
                                     If Not IsNull(rslocal!Amount) Then itmx.SubItems(3) = CStr(rslocal!Amount)
                                     If Not IsNull(rslocal!gst) Then itmx.SubItems(4) = CStr(rslocal!gst)
                                     If Not IsNull(rslocal!Total_Amount) Then
                                       itmx.SubItems(5) = CStr(rslocal!Total_Amount)
                                       TotalItemsAmount = TotalItemsAmount + rslocal!Total_Amount
                                     End If
                     Set itmx = Nothing
                     rslocal.MoveNext
                     
                 Loop
                 
                 Set rslocal = Nothing
                 
              
                 
             End If
     End With
     
     Screen.MousePointer = vbDefault
 
 
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modInvoice", "GenerateInvoiceItemList", True)
 
End Sub

Public Function DisplayInvoiceItem()

    Dim objInvoiceItem_s As CMSInvoiceItem.clsInvoiceItem_s
    
    Dim rslocal As ADODB.Recordset
   ' Dim intCnt As Integer
   ' Dim strDatabaseValue As String
    
    On Error GoTo ErrorHandler
    
    Call InitialiseInvoiceItem
    
    'Retrieve Prospect record and display on form
    Set objInvoiceItem_s = New CMSInvoiceItem.clsInvoiceItem_s
    Set objInvoiceItem_s.DatabaseConnection = objConnection
    Set rslocal = objInvoiceItem_s.getByInvoiceItemId(gInvoiceItemId)

    If rslocal Is Nothing Then
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    With frmInvoice
        
        .CmbItemDescription.Text = ConvertNull(rslocal!Description)
        .txtItemAmount.Text = ConvertNull(rslocal!Amount)
        .txtItemGST.Text = ConvertNull(rslocal!gst)
        .txtTotalItemAmount.Text = ConvertNull(rslocal!Total_Amount)
        
    End With
   Set objInvoiceItem_s = Nothing
   Screen.MousePointer = vbDefault
Exit Function

ErrorHandler:

    
    
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modInvoice", "DisplayInvoiceItem", True)

End Function



Public Function DisplayInvoice()

    Dim objInvoice_s As CMSInvoice.clsInvoice_s
    
    Dim rslocal As ADODB.Recordset
   ' Dim intCnt As Integer
   ' Dim strDatabaseValue As String
    
    On Error GoTo ErrorHandler
    
    Call InitialiseInvoiceItem
    
    'Retrieve Prospect record and display on form
    Set objInvoice_s = New CMSInvoice.clsInvoice_s
    Set objInvoice_s.DatabaseConnection = objConnection
    Set rslocal = objInvoice_s.getByInvoiceId(gInvoiceId)

    If rslocal Is Nothing Then
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    With frmInvoice
        
        .txtName.Text = ConvertNull(rslocal!Name1)
        .txtSurname.Text = ConvertNull(rslocal!name2)
        .txtAddress1.Text = ConvertNull(rslocal!address1)
        .txtAddress2.Text = ConvertNull(rslocal!address2)
        .txtPostcode.Text = ConvertNull(rslocal!Address3)
        .txtMobile.Text = ConvertNull(rslocal!Mobile)
        .txtBalance.Text = ConvertNull(rslocal!balance)
        .dteInvoiceDate.Text = Format(rslocal!Created_date, DATE_FORMAT)
        .dtePhone.Text = ConvertNull(rslocal!Phone)
        .txtState.Text = ConvertNull(rslocal!State)
        .txtTotalAmount.Text = ConvertNull(rslocal!Total_Amount)
        .cmbTerms.Text = ConvertNull(rslocal!Terms)
        .txtInvoiceNo.Text = ConvertNull(rslocal!Invoice_no)
        .txtRef.Text = ConvertNull(rslocal!Ref)
        .dteOverDueDate.Text = Format(rslocal!over_due_date, DATE_FORMAT)
        
    End With
   Set objInvoice_s = Nothing
   Screen.MousePointer = vbDefault
Exit Function

ErrorHandler:
    
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modInvoice", "DisplayInvoice", True)

End Function


