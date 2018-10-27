Attribute VB_Name = "ModMember"
Option Explicit

Public Function CheckMemberSecurity(sFunction As String) As Boolean
'New function

    Dim privilege_ctr As Byte
    CheckMemberSecurity = False
    privilege_ctr = GetPrivileges("MEMBER")
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
             ElseIf Permissions(privilege_ctr, 2) = "Y" Then CheckMemberSecurity = True
           End If
        Case RECORD_EDIT
            If Permissions(privilege_ctr, 3) = "N" Then
                MsgBox "Invalid access - Function not available for current user access level.", vbExclamation
                Exit Function
             ElseIf Permissions(privilege_ctr, 3) = "Y" Then CheckMemberSecurity = True
           End If
        Case RECORD_DELETE
            If Permissions(privilege_ctr, 4) = "N" Then
                MsgBox "Invalid access - Function not available for current user access level.", vbExclamation
                Exit Function
             ElseIf Permissions(privilege_ctr, 4) = "Y" Then CheckMemberSecurity = True
           End If
        End Select
    
                
End Function

Public Function InitialiseMember()
On Error GoTo ErrorHandler

    Dim ctrl As Control


    For Each ctrl In frmMember.Controls
        
        If TypeOf ctrl Is TextBox Then ctrl.Text = ""
        If TypeOf ctrl Is ComboBox Then ctrl.ListIndex = -1
        If TypeOf ctrl Is MaskEdBox Then ctrl.Text = ""
        Set ctrl = Nothing
        
    Next ctrl
    DoEvents
    

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMember", "InitialiseMember", True)

End Function

Public Function ValidateMemeber() As Boolean

On Error GoTo ErrorHandler

    ValidateMemeber = False

    With frmMember

            If Trim(.txtGivenName.Text) = "" Then
                MsgBox "Required field missing - Member Name must be entered.", vbExclamation
                .txtGivenName.SetFocus
                Exit Function
            End If
        
            If Trim(.txtSurname.Text) = "" Then
                MsgBox "Required field missing - Surname must be selected.", vbExclamation
                .txtSurname.SetFocus
                Exit Function
            End If
        
            If Trim(.txtAddress1.Text) = "" Then
                MsgBox "Required field missing - Address1 must be selected.", vbExclamation
                .txtAddress1.SetFocus
                Exit Function
            End If
        
            If Trim(.txtAddress2.Text) = "" Then
                MsgBox "Required field missing - Sales Address2 must be selected.", vbExclamation
                .txtAddress2.SetFocus
                Exit Function
            End If
        
            If .cmbStatus.Text = "" Then
                MsgBox "Required field missing - Status must be selected.", vbExclamation
                .cmbStatus.SetFocus
                Exit Function
            End If
        
    End With

    ValidateMemeber = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMember", "ValidateMember", True)

End Function

Public Function SaveMember() As Boolean
On Error GoTo ErrorHandler

    Dim objMember As CMSMember.clsmember
    Dim objMember_s As CMSMember.clsMember_s


    SaveMember = False
                            
    'Member record
    Set objMember = New CMSMember.clsmember
    Set objMember_s = New CMSMember.clsMember_s
    Set objMember_s.DatabaseConnection = objConnection


    PopulateMemberObject objMember

    'Insert or Update record
    If gRecordMode = RECORD_NEW Then
    
        objMember_s.InsertMember objMember
        gmemberId = objMember_s.NewMemberID
        frmMember.txtMno = gmemberId
        
    ElseIf gRecordMode = RECORD_EDIT Then
        
        objMember.MNo = gmemberId
        objMember_s.Updatemember objMember
        
    
  '  ElseIf gRecordMode = RECORD_DELETE Then
        
   ''     objMember_s.Deletemember gmemberId
        
    End If
    
    SaveMember = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMember", "SaveMember", True)
    
End Function

Public Function PopulateMemberObject(objMember As CMSMember.clsmember)
On Error GoTo ErrorHandler
'save data from form to newly created object for insert or update

   


    With frmMember

            'Push recordset results to form fields
            'objMember.MNo = .txtMno.Text
            objMember.GivenName = .txtGivenName.Text
            objMember.surname = .txtSurname.Text
            objMember.FullName = Trim(.txtGivenName.Text) & Trim(.txtSurname.Text)
            objMember.Initial = .txtStatus.Text
            objMember.address1 = .txtAddress1.Text
            objMember.address2 = .txtAddress2.Text
            objMember.Comments = .txtMemo.Text
            
            'objMember.Phone = .txtMemberPhone.Text
            objMember.postCode = .txtPostcode.Text
            objMember.SpouseName = .txtSpouse.Text
            objMember.State = .txtState.Text
            objMember.Status = .cmbStatus.Text
            objMember.Email = .txtEmail.Text
              
            If .dteDateOfBirth.Text <> "" Then objMember.DateOfBirth = .dteDateOfBirth.FormattedText
            If .dteMemberPhone.Text <> "" Then objMember.Phone = .dteMemberPhone.FormattedText
            If .dteExpiryDate.Text <> "" Then objMember.MembershipExpiary = .dteExpiryDate.FormattedText
            If .dteCreateDate.Text <> "" Then objMember.Created_date = .dteCreateDate.FormattedText

    End With
    
    
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMember", "PopulateMemberObject", True)

End Function

Public Sub GenerateMemberList(sql As String)
 On Error GoTo ErrorHandler
 
    ' Dim objFollowup_s As PACMSFollowUP.clsFollowup_s
     Dim rslocal As ADODB.Recordset
     Dim strId  As String
     Dim itmx As ListItem
     
    ' Dim lngTotalProspect As Long
    ' Dim lngTotalPlanning As Long
    ' Dim lngTotalCoaching As Long
 
 
     Screen.MousePointer = vbHourglass
 
     With frmMemberSearch
             .ListMemberView.ListItems.Clear
     
         '==============================================================================
         'get Prospect list
            Set rslocal = New ADODB.Recordset
            rslocal.Open sql, objConnection, adOpenForwardOnly, adLockReadOnly
             If Not rslocal.EOF Then
                 Do While Not rslocal.EOF
                     strId = CStr(rslocal!MNo)
                     Set itmx = .ListMemberView.ListItems.Add()
                                     itmx.Key = "#" & strId
                                     itmx.Text = CStr(rslocal!MNo)
                                     If Not IsNull(rslocal!Given_name) Then itmx.SubItems(1) = CStr(rslocal!Given_name)
                                     If Not IsNull(rslocal!surname) Then itmx.SubItems(2) = CStr(rslocal!surname)
                                     If Not IsNull(rslocal!address1) Then itmx.SubItems(3) = CStr(rslocal!address1)
                                     If Not IsNull(rslocal!address2) Then itmx.SubItems(4) = CStr(rslocal!address2)
                                     If Not IsNull(rslocal!Membership_Expiary) Then itmx.SubItems(5) = CStr(rslocal!Membership_Expiary)
                                     If Not IsNull(rslocal!address2) Then itmx.SubItems(6) = CStr(rslocal!Status)
                                     If Not IsNull(rslocal!Phone) Then itmx.SubItems(7) = CStr(rslocal!Phone)
                     Set itmx = Nothing
                     rslocal.MoveNext
                     
                 Loop
                 MemberSelected = True
                 Set rslocal = Nothing
             Else
               MemberSelected = False
             End If
     End With
     
     Screen.MousePointer = vbDefault
 
 
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMember", "GenerateMemberList", True)
 
 End Sub

Public Function DisplayMember()

    Dim objMember_s As CMSMember.clsMember_s
    
    Dim rslocal As ADODB.Recordset
   ' Dim intCnt As Integer
   ' Dim strDatabaseValue As String
    
    On Error GoTo ErrorHandler
    
    Call InitialiseMember
    
    'Retrieve Prospect record and display on form
    Set objMember_s = New CMSMember.clsMember_s
    Set objMember_s.DatabaseConnection = objConnection
    Set rslocal = objMember_s.getByMemberId(gmemberId)

    If rslocal Is Nothing Then
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    With frmMember
        
        .txtAddress1.Text = ConvertNull(rslocal!address1)
        .txtAddress2.Text = ConvertNull(rslocal!address2)
        .txtGivenName.Text = ConvertNull(rslocal!Given_name)
        .dteMemberPhone.Text = ConvertNull(rslocal!Phone)
        .txtMemo.Text = "" & rslocal!Comments
        .txtMno.Text = ConvertNull(rslocal!MNo)
        .txtPostcode.Text = ConvertNull(rslocal!postCode)
        .txtSpouse.Text = ConvertNull(rslocal!spouse_name)
        .txtState.Text = ConvertNull(rslocal!State)
        .txtStatus.Text = ConvertNull(rslocal!Mr)
        .txtSurname.Text = ConvertNull(rslocal!surname)
        .dteExpiryDate.Text = Format(rslocal!Membership_Expiary, DATE_FORMAT)
        .dteDateOfBirth.Text = Format(rslocal!DATE_OF_BIRTH, DATE_FORMAT)
        .cmbStatus.Text = ConvertNull(rslocal!Status)
        .dteCreateDate.Text = Format(rslocal!joining_date, DATE_FORMAT)
        .txtEmail.Text = ConvertNull(rslocal!Email)
        
    End With
   Set objMember_s = Nothing
   Screen.MousePointer = vbDefault
Exit Function

ErrorHandler:

    
    
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modProspect", "DisplayProspect", True)

End Function
Public Sub UpdateMemberExparyDate()

On Error GoTo ErrorHandler

    
    Dim objMember_s As CMSMember.clsMember_s
                            
    'Member record
    If frmPayment.cmbPaymentType = "Membership" Then
    Set objMember_s = New CMSMember.clsMember_s
    Set objMember_s.DatabaseConnection = objConnection
    objMember_s.UpdateExparydate frmPayment.txtmemberNo.Text, frmPayment.dteExpiryDate.FormattedText
    Set objMember_s = Nothing
 End If
    

Exit Sub
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMember", "UpdateMemberExparyDate", True)
    

End Sub

Public Sub UpdateMemberStatus()

On Error GoTo ErrorHandler

    
    Dim objMember_s As CMSMember.clsMember_s
                            
    'Member record
   
    Set objMember_s = New CMSMember.clsMember_s
    Set objMember_s.DatabaseConnection = objConnection
    objMember_s.UpdateStatus frmPayment.txtmemberNo.Text, frmPayment.cmbStatus.Text
    Set objMember_s = Nothing
 
    

Exit Sub
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMember", "UpdateMemberExparyDate", True)
    

End Sub

Public Function LoadMemberDefualtValue()

With frmMember
       .dteCreateDate.Text = Format(Now(), DATE_FORMAT)
End With

End Function
