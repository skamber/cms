Attribute VB_Name = "ModPayment"
Option Explicit

Public mPaymentMode As String
Public Function CheckPaymentSecurity(sFunction As String) As Boolean
Dim privilege_ctr As Byte
CheckPaymentSecurity = False
privilege_ctr = GetPrivileges("PAYMENT")
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
             ElseIf Permissions(privilege_ctr, 2) = "Y" Then CheckPaymentSecurity = True
           End If
        Case RECORD_EDIT
            If Permissions(privilege_ctr, 3) = "N" Then
                MsgBox "Invalid access - Function not available for current user access level.", vbExclamation
                Exit Function
             ElseIf Permissions(privilege_ctr, 3) = "Y" Then CheckPaymentSecurity = True
           End If
        Case RECORD_DELETE
            If Permissions(privilege_ctr, 4) = "N" Then
                MsgBox "Invalid access - Function not available for current user access level.", vbExclamation
                Exit Function
             ElseIf Permissions(privilege_ctr, 4) = "Y" Then CheckPaymentSecurity = True
           End If
        End Select
                
End Function

Public Function InitialisePayment()
On Error GoTo ErrorHandler

    Dim ctrl As Control
    Dim cbobx As ComboBox

    For Each ctrl In frmPayment.Controls
        
        
        If TypeOf ctrl Is TextBox Then ctrl.Text = ""
        If TypeOf ctrl Is ComboBox Then
            ctrl.ListIndex = -1
 
            If ctrl.Name = "cmbAmountInWord" Then
                ctrl.Text = ""
            End If
        End If
        If TypeOf ctrl Is MaskEdBox Then ctrl.Text = ""
        Set ctrl = Nothing
        
    Next ctrl
    DoEvents
    

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMember", "InitialiseMember", True)

End Function

Public Function SavePayment() As Boolean
On Error GoTo ErrorHandler

    Dim objPayment As CMSPayment.clsPayment
    Dim objpayment_s As CMSPayment.clsPayment_s


    SavePayment = False
                            
    'Member record
    Set objPayment = New CMSPayment.clsPayment
    Set objpayment_s = New CMSPayment.clsPayment_s
    Set objpayment_s.DatabaseConnection = objConnection


    PopulatePaymentObject objPayment

    'Insert or Update record
    If gRecordMode = RECORD_NEW Then
    
        objpayment_s.Insertpayment objPayment
        gPaymentId = objpayment_s.NewPaymentID
        frmPayment.txtReceiptNo = gPaymentId
        
    ElseIf gRecordMode = RECORD_EDIT Then
        
        objPayment.ReceiptNo = gPaymentId
        objpayment_s.UpdatePayment objPayment
        
    
'    ElseIf gRecordMode = RECORD_DELETE Then
'
'        objPayment_s.Deletepayment gPaymentId
'
    End If
    
    SavePayment = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMember", "SaveMember", True)
    
End Function

Public Function PopulatePaymentObject(objPayment As CMSPayment.clsPayment)
On Error GoTo ErrorHandler
'save data from form to newly created object for insert or update

   


    With frmPayment

            'Push recordset results to form fields
            objPayment.MNo = .txtmemberNo.Text
            objPayment.Accountname = gUserFullName
            objPayment.Amount = .txtAmount.Text
            objPayment.AmountinWords = .cmbAmountInWord.Text
            objPayment.Comments = .txtComment.Text
            objPayment.Payment = .cmbPaymentKind.Text
            objPayment.PaymentType = .cmbPaymentType.Text
            objPayment.DonationType = .cmbDonationType.Text
            'objPayment.ReceiptNo = .txtReceiptNo.Text
            objPayment.churchId = gChurchId
            
            If .dteDateofPayment.Text <> "" Then objPayment.DateofPayment = .dteDateofPayment.FormattedText
            If .dteEfectiveDate.Text <> "" Then objPayment.MemberEffective = .dteEfectiveDate.FormattedText
            If .dteExpiryDate.Text <> "" Then objPayment.MemberExpiary = .dteExpiryDate.FormattedText

    End With
    
    
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modPayment", "PopulatePaymentObject", True)

End Function

Public Function GetMemberInfo(MemberNo As Long) As Boolean
On Error GoTo ErrorHandler
 Dim objMember_s As CMSMember.clsMember_s
    
    Dim rslocal As ADODB.Recordset
    On Error GoTo ErrorHandler
    frmPayment.txtGivenName = ""
    frmPayment.txtSurname = ""
    frmPayment.cmbStatus.ListIndex = -1
    frmPayment.dteMemberExpiry.Text = ""
    
    'Retrieve Member record and display on form
    Set objMember_s = New CMSMember.clsMember_s
    Set objMember_s.DatabaseConnection = objConnection
    Set rslocal = objMember_s.getByMemberId(MemberNo, gCityId)
    

    If rslocal Is Nothing Then
          MsgBox "Invalid access - No such Member.", vbExclamation
                GetMemberInfo = False
                Exit Function
    End If

    Screen.MousePointer = vbHourglass

    With frmPayment
        
        .txtGivenName.Text = ConvertNull(rslocal!Given_name)
        .txtSurname.Text = ConvertNull(rslocal!surname)
        .cmbStatus.Text = ConvertNull(rslocal!Status)
        .dteMemberExpiry.Text = Format(rslocal!Membership_Expiary, DATE_FORMAT)
    End With
   Set objMember_s = Nothing
   Screen.MousePointer = vbDefault
   GetMemberInfo = True
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modPayment", "GetMemberInfo", True)
End Function

Public Function LoadComboBox()
On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rslocal As ADODB.Recordset
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection
    
    'Payment
    Set rslocal = objOrganisation_s.getPayment

    With frmPayment
            
            .cmbPaymentKind.Clear
            If Not rslocal Is Nothing Then
                Do Until rslocal.EOF
                    .cmbPaymentKind.AddItem rslocal!Payment_type
                    .cmbPaymentKind.ItemData(.cmbPaymentKind.NewIndex) = rslocal!id
                    rslocal.MoveNext
                Loop
                Set rslocal = Nothing
            End If
            
    End With

    
    'Payment type
    Set rslocal = objOrganisation_s.getPaymentType

    With frmPayment
            
            .cmbPaymentType.Clear
            If Not rslocal Is Nothing Then
                Do Until rslocal.EOF
                    .cmbPaymentType.AddItem rslocal!Payment
                    .cmbPaymentType.ItemData(.cmbPaymentType.NewIndex) = rslocal!id
                    rslocal.MoveNext
                Loop
                Set rslocal = Nothing
            End If
            
    End With
     
    Set rslocal = objOrganisation_s.getDonationType

    With frmPayment
            
            .cmbDonationType.Clear
            If Not rslocal Is Nothing Then
                Do Until rslocal.EOF
                    .cmbDonationType.AddItem rslocal!Donation
                    .cmbDonationType.ItemData(.cmbDonationType.NewIndex) = rslocal!id
                    rslocal.MoveNext
                Loop
                Set rslocal = Nothing
            End If
            
    End With
    

    'Amount in Word
    Set rslocal = objOrganisation_s.getAmountType

    With frmPayment
            
            .cmbAmountInWord.Clear
            If Not rslocal Is Nothing Then
                Do Until rslocal.EOF
                    .cmbAmountInWord.AddItem rslocal!AmountInWord
                    .cmbAmountInWord.ItemData(.cmbAmountInWord.NewIndex) = rslocal!id
                    rslocal.MoveNext
                Loop
                Set rslocal = Nothing
            End If
            
    End With

    Set objOrganisation_s = Nothing



Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modPayment", "LoadComboBox", True)

End Function

Public Function LoadPaymnetComboBox()
On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rslocal As ADODB.Recordset
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection
    
    'Payment
    Set rslocal = objOrganisation_s.getPaymentType

    With frmPaymentSearch
            
            .cmbType.Clear
            .cmbType.AddItem ("All")
            .cmbType.ItemData(.cmbType.NewIndex) = 0
            If Not rslocal Is Nothing Then
                Do Until rslocal.EOF
                    .cmbType.AddItem rslocal!Payment
                    .cmbType.ItemData(.cmbType.NewIndex) = rslocal!id
                    rslocal.MoveNext
                Loop
                Set rslocal = Nothing
            End If
                   
    End With

    
    Set objOrganisation_s = Nothing



Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modPayment", "LoadComboBox", True)

End Function

Public Function ValidatePayment() As Boolean

On Error GoTo ErrorHandler

    ValidatePayment = False

    With frmPayment

            If Trim(.txtmemberNo.Text) = "" Then
                MsgBox "Required field missing - Member Number must be entered.", vbExclamation
                .txtmemberNo.SetFocus
                Exit Function
            End If
        
            If Trim(.txtAmount.Text) = "" Then
                MsgBox "Required field missing - Amount must be entered.", vbExclamation
                .txtAmount.SetFocus
                Exit Function
            End If
        
            If .cmbAmountInWord.Text = "" Then
                MsgBox "Required field missing - Amount In Word must be selected.", vbExclamation
                .cmbAmountInWord.SetFocus
                Exit Function
            End If
        
            If .cmbPaymentKind.Text = "" Then
                MsgBox "Required field missing - Payment type must be selected.", vbExclamation
                .cmbPaymentKind.SetFocus
                Exit Function
            End If
        
            If .cmbPaymentType.Text = "" Then
                MsgBox "Required field missing - Payment Being must be selected.", vbExclamation
                .cmbPaymentType.SetFocus
                Exit Function
            End If
            
            If .cmbPaymentType.Text = "Membership" Then
                If .dteEfectiveDate.Text = "" Then
                 MsgBox "Required field missing - Efective Date must be selected.", vbExclamation
                 .dteEfectiveDate.SetFocus
                 Exit Function
                End If
                
                If .dteExpiryDate.Text = "" Then
                 MsgBox "Required field missing - Expiry Date must be selected.", vbExclamation
                 .dteExpiryDate.SetFocus
                 Exit Function
                End If
            End If
        
    End With

    ValidatePayment = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modPayment", "ValidatePayment", True)

End Function

Public Function LoadPaymentDefualtValue()

With frmPayment
       .dteDateofPayment.Text = Format(Now(), DATE_FORMAT)
       .txtUser.Text = gUserFullName
End With

End Function

Public Sub GeneratePaymentList(ByVal memberNumber As Long)
 On Error GoTo ErrorHandler

     Dim rslocal As ADODB.Recordset
     Dim strId  As String
     Dim itmx As ListItem
     Dim sql As String
 
     Screen.MousePointer = vbHourglass
 
     With frmPaymentSearch
             .ListPaymentView.ListItems.Clear
     
         '==============================================================================
        
            Set rslocal = New ADODB.Recordset
            sql = "SELECT payment.* FROM payment, church  WHERE  payment.CHURCH_ID = church.Id AND MNo =" & memberNumber & " AND church.CityId = " & gCityId
            If frmPaymentSearch.cmbType.Text <> "All" And frmPaymentSearch.cmbType.Text <> "" Then
               sql = sql & " AND PAYMENT ='" & frmPaymentSearch.cmbType.Text & "'"
            End If
            rslocal.Open sql, objConnection, adOpenForwardOnly, adLockReadOnly
             If Not rslocal.EOF Then
                 Do While Not rslocal.EOF
                     
                     Set itmx = .ListPaymentView.ListItems.Add(, , CStr(rslocal!MNo))
                                    
                                     If Not IsNull(rslocal!Payment) Then itmx.SubItems(1) = CStr(rslocal!Payment)
                                     If Not IsNull(rslocal!Member_Effective) Then itmx.SubItems(2) = CStr(rslocal!Member_Effective)
                                     If Not IsNull(rslocal!Member_Expiary) Then itmx.SubItems(3) = CStr(rslocal!Member_Expiary)
                                     If Not IsNull(rslocal!Date_Of_Payment) Then itmx.SubItems(4) = CStr(rslocal!Date_Of_Payment)
                                     If Not IsNull(rslocal!Amount) Then itmx.SubItems(5) = CStr(rslocal!Amount)
                                     If Not IsNull(rslocal!Receipt_No) Then itmx.SubItems(6) = CStr(rslocal!Receipt_No)
                     Set itmx = Nothing
                     rslocal.MoveNext
                 Loop
                 PaymentSelected = True
                 Set rslocal = Nothing
             Else
               PaymentSelected = False
             End If
     End With
     
     Screen.MousePointer = vbDefault
 
 
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modPayment", "GeneratePaymentList", True)
 
 End Sub

Public Function GenerateMemberInfo(MemberNo As Long)
On Error GoTo ErrorHandler
 Dim objMember_s As CMSMember.clsMember_s
    
    Dim rslocal As ADODB.Recordset
    On Error GoTo ErrorHandler
    frmPaymentSearch.txtGivenName = ""
    frmPaymentSearch.txtSurname = ""
    GenerateMemberInfo = False
    'Retrieve Member record and display on form
    Set objMember_s = New CMSMember.clsMember_s
    Set objMember_s.DatabaseConnection = objConnection
    Set rslocal = objMember_s.getByMemberId(MemberNo, gCityId)
    

    If rslocal Is Nothing Then
          MsgBox "Invalid access - No such Member.", vbExclamation
                Exit Function
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    With frmPaymentSearch
        
        .txtGivenName.Text = ConvertNull(rslocal!Given_name)
        .txtSurname.Text = ConvertNull(rslocal!surname)
    End With
   Set objMember_s = Nothing
   Screen.MousePointer = vbDefault
   GenerateMemberInfo = True
Exit Function


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modPayment", "GenerateMemberInfo", True)

End Function

Public Function DisplayPayment()

    Dim objpayment_s As CMSPayment.clsPayment_s
    
    Dim rslocal As ADODB.Recordset
   ' Dim intCnt As Integer
   ' Dim strDatabaseValue As String
    
    On Error GoTo ErrorHandler
    
    Call InitialisePayment
    
    'Retrieve Prospect record and display on form
    Set objpayment_s = New CMSPayment.clsPayment_s
    Set objpayment_s.DatabaseConnection = objConnection
    Set rslocal = objpayment_s.getByPaymentId(gPaymentId)

    If rslocal Is Nothing Then
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    With frmPayment
        
        .txtAmount.Text = Format(rslocal!Amount, NUMERIC_FORMAT)
        .txtComment.Text = "" & rslocal!Comments
        .txtmemberNo.Text = ConvertNull(rslocal!MNo)
        .txtReceiptNo.Text = ConvertNull(rslocal!Receipt_No)
        .txtUser.Text = ConvertNull(rslocal!USER_NAME)
        .cmbAmountInWord.Text = ConvertNull(rslocal!Amount_in_Words)
        .cmbPaymentKind.Text = ConvertNull(rslocal!Type)
        .cmbPaymentType.Text = ConvertNull(rslocal!Payment)
        .cmbDonationType.Text = ConvertNull(rslocal!Donation_type)
        .dteDateofPayment.Text = Format(rslocal!Date_Of_Payment, DATE_FORMAT)
        .dteEfectiveDate.Text = Format(rslocal!Member_Effective, DATE_FORMAT)
        .dteExpiryDate.Text = Format(rslocal!Member_Expiary, DATE_FORMAT)
        
    End With
    frmPayment.txtmemberNo_LostFocus
   Set objpayment_s = Nothing
   Screen.MousePointer = vbDefault
Exit Function

ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modPayment", "DisplayPayment", True)

End Function


