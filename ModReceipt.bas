Attribute VB_Name = "ModReceipt"
Option Explicit

Public mReceiptMode As String
Public mReceiptAmountAlocated As Double
Public Function CheckReceiptSecurity(sFunction As String) As Boolean
Dim privilege_ctr As Byte
    CheckReceiptSecurity = False
    privilege_ctr = GetPrivileges("RECEIPT")
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
             ElseIf Permissions(privilege_ctr, 2) = "Y" Then CheckReceiptSecurity = True
           End If
        Case RECORD_EDIT
            If Permissions(privilege_ctr, 3) = "N" Then
                MsgBox "Invalid access - Function not available for current user access level.", vbExclamation
                Exit Function
             ElseIf Permissions(privilege_ctr, 3) = "Y" Then CheckReceiptSecurity = True
           End If
        Case RECORD_DELETE
            If Permissions(privilege_ctr, 4) = "N" Then
                MsgBox "Invalid access - Function not available for current user access level.", vbExclamation
                Exit Function
             ElseIf Permissions(privilege_ctr, 4) = "Y" Then CheckReceiptSecurity = True
           End If
        End Select
                
End Function

Public Function ValidateReceipt() As Boolean

On Error GoTo ErrorHandler

    ValidateReceipt = False

    With frmReceipt

           
           If Trim(.txtInvoiceNo.Text) = "" Then
                MsgBox "Required field missing - Invoice_number  must be entered.", vbExclamation
                .txtInvoiceNo.SetFocus
                Exit Function
            End If
        If Trim(.txtAmountToPay.Text) = "" Then
                MsgBox "Required field missing - Amount  must be entered.", vbExclamation
                .txtAmountToPay.SetFocus
                Exit Function
            End If
        If mReceiptAmountAlocated < .txtAmountToPay.Text Then
                MsgBox "You can not Alocate this amount.", vbExclamation
                .txtAmountToPay.SetFocus
                Exit Function
            End If
        
    End With

    ValidateReceipt = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modReceipt", "ValidateReceipt", True)

End Function

Public Function InitialiseReceipt()
On Error GoTo ErrorHandler

    Dim ctrl As Control


    For Each ctrl In frmReceipt.Controls
        
        If TypeOf ctrl Is TextBox Then ctrl.Text = ""
        If TypeOf ctrl Is ComboBox Then ctrl.ListIndex = -1
        If TypeOf ctrl Is MaskEdBox Then ctrl.Text = ""
        Set ctrl = Nothing
        
    Next ctrl
    DoEvents
    

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modReceipt", "InitialiseMember", True)

End Function
Public Function SaveReceipt() As Boolean
On Error GoTo ErrorHandler

    Dim objReceipt As CMSreceipt.clsReceipt
    
    Dim objreceipt_s As CMSreceipt.clsReceipt_s
    


    SaveReceipt = False
                            
    'Receipt record
    Set objReceipt = New CMSreceipt.clsReceipt
    Set objreceipt_s = New CMSreceipt.clsReceipt_s
    Set objreceipt_s.DatabaseConnection = objConnection


    PopulateReceiptObject objReceipt

    'Insert or Update record
    If gRecordMode = RECORD_NEW Then
    
        objreceipt_s.InsertReceipt objReceipt
        gReceiptId = objreceipt_s.NewreceiptID
        frmReceipt.txtReceiptId = gReceiptId
        UpdateInvoiceBalance (frmReceipt.txtInvoiceID)
    ElseIf gRecordMode = RECORD_EDIT Then
        
        objReceipt.Id = gReceiptId
        objreceipt_s.Updatereceipt objReceipt
        UpdateInvoiceBalance (frmReceipt.txtInvoiceID)
    
'    ElseIf gRecordMode = RECORD_DELETE Then
'
'        objMember_s.Deletemember gmemberId
        
    End If
    
    SaveReceipt = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMember", "SaveMember", True)
    
End Function

Public Function PopulateReceiptObject(objReceipt As CMSreceipt.clsReceipt)
On Error GoTo ErrorHandler
'save data from form to newly created object for insert or update

   


    With frmReceipt

            'Push recordset results to form fields
            
            If .dteDate.Text <> "" Then objReceipt.DateOfReceipt = .dteDate.FormattedText
            objReceipt.INV_NO = .txtInvoiceNo.Text
            objReceipt.Amount = .txtAmountToPay.Text
            objReceipt.Accountname = .txtUser.Text
            objReceipt.INV_ID = .txtInvoiceID.Text
            objReceipt.ChequeNumber = .txtChequeNo.Text
            objReceipt.ChurchId = gChurchId
            
            
    End With
    
    
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modReceipt", "PopulateReceiptObject", True)

End Function

Public Function LoadReceiptDefualtValue()

With frmReceipt
       .dteDate.Text = Format(Now(), DATE_FORMAT)
       .txtUser.Text = UserName
       
End With

End Function

Sub GetInvoiceInfo()
 On Error GoTo ErrorHandler
 
    ' Dim objFollowup_s As PACMSFollowUP.clsFollowup_s
     Dim objInvoice_s As CMSinvoice.clsInvoice_s
     Dim rslocal As ADODB.Recordset
     
    Set objInvoice_s = New CMSinvoice.clsInvoice_s
    Set objInvoice_s.DatabaseConnection = objConnection
 
   
 Set rslocal = objInvoice_s.getByInvoiceNo(frmReceipt.txtInvoiceNo)
    
     
 
     With frmReceipt
         'get Prospect list
         If rslocal Is Nothing Then
         MsgBox "Invalid Invoice Number.", vbExclamation
         Screen.MousePointer = vbDefault
         Exit Sub
         Else
          .txtName.Text = rslocal!Name1
          .txtSurname.Text = rslocal!name2
          .txtInvoiceID.Text = rslocal!Id
          .txtAmountToPay.Text = rslocal!balance
          mReceiptAmountAlocated = rslocal!balance
          Call ValidNumericEntry(.txtAmountToPay)
         End If
     End With
     
     Screen.MousePointer = vbDefault
 
 
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modReceipt", "GetInvoiceInfo", True)
 
End Sub

Public Function getReceiptAmount(InvoiceId As Long) As Double
  On Error GoTo ErrorHandler
 
    ' Dim objFollowup_s As PACMSFollowUP.clsFollowup_s
     Dim objreceipt_s As CMSreceipt.clsReceipt_s
     Dim rslocal As ADODB.Recordset
     Dim TotalAmount As Double
    Set objreceipt_s = New CMSreceipt.clsReceipt_s
    Set objreceipt_s.DatabaseConnection = objConnection
 
    ' Dim lngTotalProspect As Long
    ' Dim lngTotalPlanning As Long
    ' Dim lngTotalCoaching As Long
 Set rslocal = objreceipt_s.getByInvoiceId(InvoiceId)
     TotalAmount = 0
     Screen.MousePointer = vbHourglass
             If Not rslocal Is Nothing Then
                 Do While Not rslocal.EOF
                     TotalAmount = TotalAmount + rslocal!Amount
                     rslocal.MoveNext
             Loop
             End If
     getReceiptAmount = TotalAmount
     Screen.MousePointer = vbDefault
 Exit Function
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modReceipt", "getReceiptAmount", True)
 
End Function

Public Sub GenerateReceiptList(sql As String)
 On Error GoTo ErrorHandler
 
    ' Dim objFollowup_s As PACMSFollowUP.clsFollowup_s
     Dim rslocal As ADODB.Recordset
     Dim strId  As String
     Dim itmx As ListItem
     
    ' Dim lngTotalProspect As Long
    ' Dim lngTotalPlanning As Long
    ' Dim lngTotalCoaching As Long
 
 
     Screen.MousePointer = vbHourglass
 
     With frmReceiptSearch
             .ListReceiptView.ListItems.Clear
     
         '==============================================================================
         'get Prospect list
            Set rslocal = New ADODB.Recordset
            rslocal.Open sql, objConnection, adOpenForwardOnly, adLockReadOnly
             If Not rslocal.EOF Then
                 Do While Not rslocal.EOF
                     strId = CStr(rslocal!Id)
                     Set itmx = .ListReceiptView.ListItems.Add()
                                     itmx.Key = "#" & strId
                                     itmx.Text = CStr(rslocal!Id)
                                     If Not IsNull(rslocal!INV_ID) Then itmx.SubItems(1) = CStr(rslocal!INV_ID)
                                     If Not IsNull(rslocal!INV_NO) Then itmx.SubItems(2) = CStr(rslocal!INV_NO)
                                     If Not IsNull(rslocal!Amount) Then itmx.SubItems(3) = CStr(rslocal!Amount)
                                     If Not IsNull(rslocal!Date_Of_Receipt) Then itmx.SubItems(4) = CStr(rslocal!Date_Of_Receipt)
                     Set itmx = Nothing
                     rslocal.MoveNext
                     
                 Loop
                 ReceiptSelected = True
                 Set rslocal = Nothing
             Else
               ReceiptSelected = False
             End If
     End With
     
     Screen.MousePointer = vbDefault
 
 
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMember", "GenerateMemberList", True)
 
 End Sub
Public Sub UpdateInvoiceBalance(InvoiceId As Long)
  On Error GoTo ErrorHandler
    Dim TotalAmount As Double
    Dim TotalReceiptAmount As Double
    Dim AmountBalance As Double
    Dim objInvoice_s As CMSinvoice.clsInvoice_s
    Dim rslocal As ADODB.Recordset
    On Error GoTo ErrorHandler
   
    Set objInvoice_s = New CMSinvoice.clsInvoice_s
    Set objInvoice_s.DatabaseConnection = objConnection
    Set rslocal = objInvoice_s.getByInvoiceId(InvoiceId)
    If rslocal Is Nothing Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    TotalAmount = rslocal!Total_Amount
    Set objInvoice_s = Nothing
    TotalReceiptAmount = getReceiptAmount(InvoiceId)
    AmountBalance = TotalAmount - TotalReceiptAmount
    
    Set objInvoice_s = New CMSinvoice.clsInvoice_s
    Set objInvoice_s.DatabaseConnection = objConnection
    objInvoice_s.UpdateBalance InvoiceId, AmountBalance
    Set objInvoice_s = Nothing
    Screen.MousePointer = vbDefault
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modReceipt", "UpdateInvoiceBalance", True)
 
End Sub
