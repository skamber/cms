Attribute VB_Name = "ModChildren"
Option Explicit

Public mChiledMode As String

Public Function InitialiseChild()
On Error GoTo ErrorHandler

    Dim ctrl As Control


    For Each ctrl In frmChildren.Controls
        
        If TypeOf ctrl Is TextBox Then ctrl.Text = ""
        If TypeOf ctrl Is ComboBox Then ctrl.ListIndex = -1
        If TypeOf ctrl Is MaskEdBox Then ctrl.Text = ""
        Set ctrl = Nothing
        
    Next ctrl
    DoEvents
    

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modChildren", "InitialiseChild", True)

End Function

Public Function SaveChild() As Boolean
On Error GoTo ErrorHandler

    Dim objChild As CMSChildren.clsChildren
    Dim objChild_s As CMSChildren.clsChildren_s


    SaveChild = False
                            
    'Member record
    Set objChild = New CMSChildren.clsChildren
    Set objChild_s = New CMSChildren.clsChildren_s
    Set objChild_s.DatabaseConnection = objConnection


    PopulateChildObject objChild

    'Insert or Update record
    If gRecordMode = RECORD_NEW Then
    
        objChild_s.InsertChild objChild
        gchildId = objChild_s.NewChildId
        frmChildren.txtChildMno = gchildId
        
    ElseIf gRecordMode = RECORD_EDIT Then
        
        objChild.childno = gchildId
        objChild_s.UpdateChild objChild
        
    
  '  ElseIf gRecordMode = RECORD_DELETE Then
        
   ''     objMember_s.Deletemember gmemberId
        
    End If
    
    SaveChild = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modChildren", "SaveChild", True)
    
End Function

Public Function PopulateChildObject(objChild As CMSChildren.clsChildren)
On Error GoTo ErrorHandler
'save data from form to newly created object for insert or update

   


    With frmChildren

            
            objChild.MNo = .txtMemberMno.Text
            objChild.FirstName = .txtGivenName
            objChild.Genda = .cmbGenda.Text
            objChild.MEMBER = .cmbMemberStatus.Text
            objChild.surname = .txtSurname.Text
            objChild.Email = .txtEmail.Text
            objChild.Memo = .txtMemo.Text
            objChild.cityId = gCityId
            
            
            
            
            If .dteBirthDate.Text <> "" Then objChild.BirthDate = .dteBirthDate.FormattedText
    End With
    
    
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modChildren", "PopulateChildObject", True)

End Function

Public Function ValidateChild() As Boolean

On Error GoTo ErrorHandler

    ValidateChild = False

    With frmChildren

            If Trim(.txtMemberMno.Text) = "" Then
                MsgBox "Required field missing - Member Mno must be entered.", vbExclamation
                .txtMemberMno.SetFocus
                Exit Function
            End If
        
            If Trim(.txtSurname.Text) = "" Then
                MsgBox "Required field missing - Surname must be selected.", vbExclamation
                .txtSurname.SetFocus
                Exit Function
            End If
        
            If Trim(.txtGivenName.Text) = "" Then
                MsgBox "Required field missing - Given Name must be selected.", vbExclamation
                .txtGivenName.SetFocus
                Exit Function
            End If
        
            If .cmbGenda.Text = "" Then
                MsgBox "Required field missing - Genda must be selected.", vbExclamation
                .cmbGenda.SetFocus
                Exit Function
            End If
            
            
    End With

    ValidateChild = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modChildren", "ValidateChild", True)

End Function

Public Sub GenerateChildrenList(memberNumber As Long)
 On Error GoTo ErrorHandler
 
    ' Dim objFollowup_s As PACMSFollowUP.clsFollowup_s
     Dim rslocal As ADODB.Recordset
     Dim strId  As String
     Dim itmx As ListItem
     Dim sql As String
    ' Dim lngTotalProspect As Long
    ' Dim lngTotalPlanning As Long
    ' Dim lngTotalCoaching As Long
 
 
     Screen.MousePointer = vbHourglass
 
     With frmChildSearch
             .ListChildrenView.ListItems.Clear
     
         '==============================================================================
        
            Set rslocal = New ADODB.Recordset
            sql = "SELECT * FROM children WHERE MNo =" & memberNumber
            rslocal.Open sql, objConnection, adOpenForwardOnly, adLockReadOnly
             If Not rslocal.EOF Then
                 Do While Not rslocal.EOF
                     
                     Set itmx = .ListChildrenView.ListItems.Add(, , CStr(rslocal!Id))
                                    
                                     If Not IsNull(rslocal!childno) Then itmx.SubItems(1) = CStr(rslocal!MNo)
                                     If Not IsNull(rslocal!first_name) Then itmx.SubItems(2) = CStr(rslocal!first_name)
                                     If Not IsNull(rslocal!surname) Then itmx.SubItems(3) = CStr(rslocal!surname)
                                     If Not IsNull(rslocal!Genda) Then itmx.SubItems(4) = CStr(rslocal!Genda)
                                     If Not IsNull(rslocal!birth_date) Then itmx.SubItems(5) = CStr(rslocal!birth_date)
                                     If Not IsNull(rslocal!MEMBER) Then itmx.SubItems(6) = CStr(rslocal!MEMBER)
                     Set itmx = Nothing
                     rslocal.MoveNext
                     
                 Loop
                 ChildSelected = True
                 Set rslocal = Nothing
              Else
                ChildSelected = False
             End If
     End With
     
     Screen.MousePointer = vbDefault
 
 
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modChildren", "GenerateChildrenList", True)
 
 End Sub

Public Function DisplayChild()

    Dim objChild_s As CMSChildren.clsChildren_s
    
    Dim rslocal As ADODB.Recordset
   ' Dim intCnt As Integer
   ' Dim strDatabaseValue As String
    
    On Error GoTo ErrorHandler
    
    Call InitialiseChild
    
    'Retrieve Prospect record and display on form
    Set objChild_s = New CMSChildren.clsChildren_s
    Set objChild_s.DatabaseConnection = objConnection
    Set rslocal = objChild_s.getByChildId(gchildId)

    If rslocal Is Nothing Then
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    With frmChildren
        
        .txtChildMno.Text = ConvertNull(rslocal!Id)
        .txtGivenName.Text = ConvertNull(rslocal!first_name)
        .txtSurname.Text = ConvertNull(rslocal!surname)
        .cmbGenda.Text = ConvertNull(rslocal!Genda)
        .cmbMemberStatus.Text = ConvertNull(rslocal!MEMBER)
        .txtMemberMno.Text = ConvertNull(rslocal!MNo)
        .txtEmail.Text = ConvertNull(rslocal!Email)
        .dteBirthDate.Text = Format(rslocal!birth_date, DATE_FORMAT)
        .txtMemo = "" & rslocal!Comments
        GetMemberName (.txtMemberMno.Text)
        
        
    End With
   Set objChild_s = Nothing
   Screen.MousePointer = vbDefault
Exit Function

ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modchildren", "DisplayChild", True)

End Function



Public Sub GetMemberName(MemberNo As Long)
On Error GoTo ErrorHandler
 Dim objMember_s As CMSMember.clsMember_s
    
    Dim rslocal As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    frmChildren.txtMemberName = ""
    
    'Retrieve Member record and display on form
    Set objMember_s = New CMSMember.clsMember_s
    Set objMember_s.DatabaseConnection = objConnection
    Set rslocal = objMember_s.getByMemberId(MemberNo)
    

    If rslocal Is Nothing Then
          MsgBox "Invalid access - No such Member.", vbExclamation
                Exit Sub
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass

    With frmChildren
        
        .txtMemberName.Text = ConvertNull(rslocal!Given_name) & " " & ConvertNull(rslocal!surname)
        
    End With
   Set objMember_s = Nothing
   Screen.MousePointer = vbDefault
Exit Sub


Exit Sub
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modChildren", "GenerateMemberInfo", True)

End Sub

