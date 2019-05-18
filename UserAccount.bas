Attribute VB_Name = "modUserAccount"
' (c) 2002-2002 Marsh Ltd.                                           }
' All Rights Reserved.                                                   }
 
 
 
' Update History
 
 
'Date        By     Comment

'03-May-02   SAM    Create the unit
'13-Jun-02   SAM    Make sure that office been selected . Make sure upcase of user logon
'27-Aug-02   SAM    Add new Privilege called Practice Administrator
'<!<CHECKOUT>!>

'Backed up to 4529 on 16-May-02 by SAM
'Backed up to 4558 on 13-Jun-02 by SAM
'Backed up to 4614 on 27-Aug-02 by SAM
'<!<PREVIOUS_VERSIONS>!>
'}
Option Explicit

Public gRecordUserType As String
Public gUserAccountId As Long
Public gUserAccountName As String


Public Function InitialiseUserAccount()
On Error GoTo ErrorHandler
    
    Dim ctrl As Control
    Dim i As Integer
    Dim Ctr As Byte
    For Each ctrl In frmUserAccount.Controls
        
        If TypeOf ctrl Is TextBox Then ctrl.Text = ""
        If TypeOf ctrl Is ComboBox Then
            ctrl.ListIndex = -1
            'ctrl.Text = ""
        End If
        If TypeOf ctrl Is MaskEdBox Then ctrl.Text = ""
        Set ctrl = Nothing
        
    Next ctrl
    
    DoEvents
    For Ctr = 1 To frmUserAccount.listActions.ListCount
    LoadPermissions(Ctr, 1) = "N"
    LoadPermissions(Ctr, 2) = "N"
    LoadPermissions(Ctr, 3) = "N"
    LoadPermissions(Ctr, 4) = "N"
    Next Ctr
    frmUserAccount.chkDelete.value = 0
    frmUserAccount.chkInsert.value = 0
    frmUserAccount.chkRead.value = 0
    frmUserAccount.chkUpdate.value = 0
    
    
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modUserAccount", "InitialiseUserAccount", True)

End Function

Public Function LoadUserAccountComboBox()
On Error GoTo ErrorHandler

 Dim objuser_s As CMSUser.clsUser_s
 Dim rslocal As Recordset
 Dim objOrganisation_s As CMSOrganisation.clsOrganisation
 

'==============================================================================
    Set objuser_s = New CMSUser.clsUser_s
    Set objuser_s.DatabaseConnection = objConnection
    
    
    Set rslocal = objuser_s.getAllUserName

    With frmUserAccount
            
            .CboUserName.Clear
            If Not rslocal Is Nothing Then
                Do Until rslocal.EOF
                    .CboUserName.AddItem rslocal!Full_Name
                    .CboUserName.ItemData(.CboUserName.NewIndex) = rslocal!id
                    rslocal.MoveNext
                Loop
                Set rslocal = Nothing
            End If
            
    End With

     

    Set objuser_s = Nothing
    
 Set objOrganisation_s = New CMSOrganisation.clsOrganisation
 Set objOrganisation_s.DatabaseConnection = objConnection
    Set rslocal = objOrganisation_s.getActions

    With frmUserAccount
            
            .listActions.Clear
            If Not rslocal Is Nothing Then
                Do Until rslocal.EOF
                    .listActions.AddItem rslocal!Action_Name
                    .listActions.ItemData(.listActions.NewIndex) = rslocal!Action_id
                    rslocal.MoveNext
                Loop
                Set rslocal = Nothing
            End If
            
    End With

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modUserAccount", "LoadUserAccountComboBox", True)

End Function

Public Function SetToolbarUserAccountControl()

    On Error GoTo ErrorHandler

    With frmUserAccount
        Select Case gRecordUserType
            Case RECORD_READ, RECORD_CANCEL
                    .Toolbar1.Buttons.Item(1).Enabled = True
                    .Toolbar1.Buttons.Item(2).Enabled = True
                    .Toolbar1.Buttons.Item(3).Enabled = True
                    .Toolbar1.Buttons.Item(5).Enabled = False
                    .Toolbar1.Buttons.Item(6).Enabled = False
        
            Case RECORD_NEW, RECORD_EDIT, RECORD_DELETE
                    .Toolbar1.Buttons.Item(1).Enabled = False
                    .Toolbar1.Buttons.Item(2).Enabled = False
                    .Toolbar1.Buttons.Item(3).Enabled = False
                    .Toolbar1.Buttons.Item(5).Enabled = True
                    .Toolbar1.Buttons.Item(6).Enabled = True
                    
        End Select
    End With
    
    Exit Function
    
ErrorHandler:

    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMain", "SetToolbarControl", True)

End Function

Public Function DisplayUserAccount()

   
    Dim objuser_s As CMSUser.clsUser_s
    Dim rslocal As ADODB.Recordset
    Dim intCnt As Integer
    'Dim strDatabaseValue As String
    Dim checkSystemManager As String
    Dim checkReportView As String
    On Error GoTo ErrorHandler
   ' Call InitialiseUserAccount
    Screen.MousePointer = vbHourglass
        
        Set objuser_s = New CMSUser.clsUser_s
        Set objuser_s.DatabaseConnection = objConnection
        
        
        Set rslocal = objuser_s.getUserAccountRecord(gUserAccountId)
        If rslocal Is Nothing Then
        Exit Function
        End If
      With frmUserAccount
        .txtUserId = rslocal!id
        .txtUserName = rslocal!Full_Name
        .txtLogonId = rslocal!Logon_Id
        Dim cityId As Long
        Dim churchId As Long
        Dim index As Long
        cityId = ConvertNull(rslocal!City_Id)
        index = FindCBIndexById(.cmbCity, cityId)
        .cmbCity.Text = .cmbCity.List(index)
        LoadChurchComboBoxWithAll (cityId)
        churchId = ConvertNull(rslocal!Church_Id)
        index = FindCBIndexById(.cmbChurch, churchId)
        .cmbChurch.Text = .cmbChurch.List(index)
        
        .txtPassword = rslocal!Logon_Password
        checkSystemManager = rslocal!SYSTEM_MANAGER
        checkReportView = rslocal!Report_View
        If checkSystemManager = "Y" Then
            .chkSystemManager.value = 1
         Else
         .chkSystemManager.value = 0
        End If
        
        If checkReportView = "Y" Then
            .chkReportView.value = 1
         Else
         .chkReportView.value = 0
        End If
        
        .dtePasswordLastUpdate = Format(rslocal!Password_Last_Update, DATE_FORMAT) & ""
        
        getLoadedUserPriveleges (gUserAccountId)
        
       
      End With
    Set objuser_s = Nothing
    Set rslocal = Nothing
  '   DisplayPracticesForUser
     
    Screen.MousePointer = vbDefault

Exit Function

ErrorHandler:

    Screen.MousePointer = vbDefault
    
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modUserAccount", "DisplayUserAccount", True)

End Function


Public Function ProcessUserAccountRecord(ByVal sFunction As String)
Dim i As Integer
Dim iAnswer As Integer
On Error GoTo ErrorHandler
If sFunction = "New" Or sFunction = "Cancel" Or sFunction = "Save" Then
        'Let through
    Else
        'Check to see if the prospect is selected
        If frmUserAccount.CboUserName.ListIndex = -1 Then
            MsgBox "Please select a User before continuing", vbInformation, "CMS"
            Exit Function
        End If
End If
With frmUserAccount
Select Case sFunction

Case RECORD_NEW
                InitialiseUserAccount
                
                .CboUserName.Enabled = False
                .txtPassword.Text = "WELCOME"
                .dtePasswordLastUpdate.Text = Format(Now, "ddmmyyyy")
                gRecordUserType = sFunction
                gUserAccountId = -1
                UserAccountActivate
                .txtUserName.SetFocus
                SetToolbarUserAccountControl
               
Case RECORD_EDIT
               
               ' .CboUserName.Enabled = False
                SetUserRecordMode
                gRecordUserType = sFunction
                UserAccountActivate
                SetToolbarUserAccountControl

Case RECORD_DELETE
                
                .CboUserName.Enabled = False
                iAnswer = MsgBox("Are you sure you want to delete this User?", vbExclamation + vbYesNo)
                If iAnswer = vbNo Then
                .CboUserName.Enabled = True
                    Exit Function
                Else
                    gRecordUserType = sFunction
                    SaveUserAccount
                    InitialiseUserAccount
                End If
                .CboUserName.Enabled = True

Case RECORD_SAVE
                
               objConnection.BeginTrans
                If SaveUserAccount Then
                   objConnection.CommitTrans
                Else
                    objConnection.RollbackTrans
                    MsgBox "An error has occurred saving record to database - changes have not been applied.", vbExclamation
                End If
                gRecordUserType = RECORD_READ
                'frmUserAccount.CboUserName.Text = frmUserAccount.txtUserName.Text
                SetToolbarUserAccountControl
                SetUserRecordMode
                UserAccountActivate
Case RECORD_CANCEL
                
                iAnswer = MsgBox("Do you want to cancel the changes made to the record?", vbExclamation + vbYesNo)
                If iAnswer = vbNo Then Exit Function
                
                If gRecordUserType = RECORD_NEW Then
                            gRecordUserType = RECORD_CANCEL
                            InitialiseUserAccount
                            SetToolbarUserAccountControl
                ElseIf gRecordUserType = RECORD_EDIT Then
                    gRecordUserType = RECORD_CANCEL
                    SetToolbarUserAccountControl
                    .CboUserName_Click
                End If
                gRecordUserType = RECORD_READ
                SetUserRecordMode
                UserAccountActivate
End Select
End With
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modUserAccount", "ProcessUserAccountRecord", True)

End Function
 
Public Sub UserAccountActivate()

    If gRecordUserType = RECORD_NEW Or gRecordUserType = RECORD_EDIT Then
        With frmUserAccount
        .txtLogonId.Enabled = True
        .CboUserName.Enabled = False
        .txtUserName.Enabled = True
        .txtPassword.Enabled = True
        .fram1.Enabled = True
        .listActions.Enabled = True
        .chkSystemManager.Enabled = True
        .chkReportView.Enabled = True
        .cmbCity.Enabled = True
        .cmbChurch.Enabled = True
        
        
'        .cboOffice.Enabled = True
'        .lstPractice.ListIndex = -1
'        .lstPrivilege.ListIndex = -1
'        .lstPractice.Enabled = True
'        .lstPrivilege.Enabled = True
        .dtePasswordLastUpdate.Enabled = True
        End With
        
    Else
        With frmUserAccount
        .txtLogonId.Enabled = False
        .txtUserName.Enabled = False
        .txtPassword.Enabled = False
        .fram1.Enabled = False
        .chkSystemManager.Enabled = False
        .chkReportView.Enabled = False
        .cmbCity.Enabled = False
        .cmbChurch.Enabled = False
        .listActions.Enabled = True
'        .cboOffice.Enabled = False
'        .lstPractice.ListIndex = -1
'        .lstPrivilege.ListIndex = -1
'        .lstPractice.Enabled = False
'        .lstPrivilege.Enabled = False
        .dtePasswordLastUpdate.Enabled = False
       End With
    End If
    
End Sub

Public Function SaveUserAccount() As Boolean
On Error GoTo ErrorHandler

    Dim objUser As CMSUser.clsUser
    Dim objuser_s As CMSUser.clsUser_s
    Dim objUserPermission As CMSUser.clsUserPermissions
    Dim objUserPermission_s As CMSUser.clsUserPermissions_s
    Dim Ctr As Integer
    SaveUserAccount = False
                            
    'UserAccount record
    Set objUser = New CMSUser.clsUser
    Set objuser_s = New CMSUser.clsUser_s
    Set objuser_s.DatabaseConnection = objConnection
    Set objUserPermission = New CMSUser.clsUserPermissions
    Set objUserPermission_s = New CMSUser.clsUserPermissions_s
    Set objUserPermission_s.DatabaseConnection = objConnection

    If Not PopulateUserAccountObject(objUser) Then
     MsgBox "Invalid data entry.", vbExclamation
     Exit Function
    End If
    
    'Insert or Update record
    Select Case gRecordUserType
Case RECORD_NEW:

        objuser_s.insertUser objUser
        objUser.UserId = objuser_s.NewUserId
        gUserAccountId = objUser.UserId
        gUserAccountName = objUser.userName
        frmUserAccount.CboUserName.AddItem gUserAccountName
        frmUserAccount.CboUserName.ItemData(frmUserAccount.CboUserName.NewIndex) = gUserAccountId
        frmUserAccount.txtUserId = objuser_s.NewUserId
        For Ctr = 0 To frmUserAccount.listActions.ListCount - 1
        PopulateUserPermissionObject objUserPermission, frmUserAccount.listActions.ItemData(Ctr), objuser_s.NewUserId
        
        objUserPermission_s.InsertpermissionForUser objUserPermission
        Next Ctr
Case RECORD_EDIT:
        objUser.UserId = frmUserAccount.txtUserId
        objuser_s.UpdateUser objUser
        For Ctr = 0 To frmUserAccount.listActions.ListCount - 1
        PopulateUserPermissionObject objUserPermission, frmUserAccount.listActions.ItemData(Ctr), objUser.UserId
        objUserPermission_s.UpdatepermissionForUser objUserPermission
        Next Ctr
 End Select
    SaveUserAccount = True

Set objUser = Nothing
Set objuser_s = Nothing
Set objUserPermission = Nothing
Set objUserPermission_s = Nothing

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modUserAccount", "SaveUserAccount", True)
    
End Function



Public Function PopulateUserAccountObject(objUser As CMSUser.clsUser)

On Error GoTo ErrorHandler
PopulateUserAccountObject = False
    With frmUserAccount
    
            If Trim(.txtUserName.Text) <> "" Then objUser.userName = UCase(.txtUserName.Text)
            If Trim(.cmbCity.Text) <> "" Then
            .cmbCity.ListIndex = FindCBIndexByName(.cmbCity, .cmbCity.Text)
              objUser.cityId = .cmbCity.ItemData(.cmbCity.ListIndex)
            Else
            MsgBox "Invalid city selection.", vbExclamation
            Exit Function
            End If
            
            If Trim(.cmbChurch.Text) <> "" Then
            .cmbChurch.ListIndex = FindCBIndexByName(.cmbChurch, .cmbChurch.Text)
              objUser.churchId = .cmbChurch.ItemData(.cmbChurch.ListIndex)
            Else
            MsgBox "Invalid Church selection.", vbExclamation
            Exit Function
            End If
            
            If Trim(.txtLogonId.Text) <> "" Then objUser.LogonId = UCase(.txtLogonId.Text)
            If Trim(.txtPassword.Text) <> "" Then objUser.LogonPassword = .txtPassword.Text
            If .chkSystemManager.value = 1 Then
                objUser.systemManager = "Y"
            Else
               objUser.systemManager = "N"
            End If
            If .chkReportView.value = 1 Then
                objUser.ReportView = "Y"
            Else
               objUser.ReportView = "N"
            End If
            If .dtePasswordLastUpdate.Text <> "" Then objUser.PasswordLastChange = .dtePasswordLastUpdate.FormattedText
            PopulateUserAccountObject = True
     End With
    
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modUserAccount", "PopulateUserAccount", True)

End Function

Public Function PopulateUserPermissionObject(objpermission As CMSUser.clsUserPermissions, Action_id As Long, User_id As Long)

On Error GoTo ErrorHandler
       If LoadPermissions(Action_id, 1) = "Y" Then
         objpermission.ReadPermissions = "Y"
       Else
         objpermission.ReadPermissions = "N"
       End If
       If LoadPermissions(Action_id, 2) = "Y" Then
         objpermission.WritePermissions = "Y"
       Else
         objpermission.WritePermissions = "N"
       End If
         
      If LoadPermissions(Action_id, 3) = "Y" Then
         objpermission.UpdatePermissions = "Y"
       Else
         objpermission.UpdatePermissions = "N"
      End If
         If LoadPermissions(Action_id, 4) = "Y" Then
         objpermission.DeletePermissions = "Y"
       Else
         objpermission.DeletePermissions = "N"
    End If
    objpermission.ActionId = Action_id
    objpermission.UserId = User_id
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modUserAccount", "PopulateUserPermissionObject", True)

End Function
Public Function SetUserRecordMode()

    On Error GoTo ErrorHandler
        If gRecordUserType = RECORD_NEW Or gRecordUserType = RECORD_EDIT Then
            If gRecordUserType = RECORD_NEW Then
                frmUserAccount.CboUserName.ListIndex = -1
            End If
        Else
            If gRecordUserType = RECORD_SAVE Or gRecordUserType = RECORD_CANCEL Or gRecordUserType = RECORD_READ Then
                frmUserAccount.CboUserName.Enabled = True
            End If
        End If
        
  Exit Function
    
ErrorHandler:

    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modUserAccount", "SetUserRecordMode", True)
End Function



Public Function getUserPriveleges(ByVal UserId As Long) As Boolean
On Error GoTo ErrorHandler

    Dim sql As String
    Dim Count As Byte
    Count = 1
    sql = "SELECT * FROM privilege WHERE User_Id = " & UserId
    sql = sql & " Order by Action_Id"
    Set Userprivilege = New ADODB.Recordset
        Userprivilege.Open sql, objConnection, adOpenForwardOnly, adLockReadOnly
    
    If Userprivilege.EOF = True Then
        MsgBox "No Privilege found for this user.", vbExclamation
        getUserPriveleges = False
    Else
        getUserPriveleges = True
  
        Do While Not Userprivilege.EOF
            Permissions(Count, 1) = (Userprivilege!read_data)
            Permissions(Count, 2) = (Userprivilege!Create_data)
            Permissions(Count, 3) = (Userprivilege!Edit_data)
            Permissions(Count, 4) = (Userprivilege!Delete_data)
            Count = Count + 1
            Userprivilege.MoveNext
        Loop
    End If
   Set Userprivilege = Nothing
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "moduser", "getUserPriveleges", True)

End Function

Public Function getLoadedUserPriveleges(ByVal UserId As Long) As Boolean
On Error GoTo ErrorHandler

    Dim sql As String
    Dim Count As Byte
    Count = 1
    sql = "SELECT * FROM privilege WHERE User_Id = " & UserId
    sql = sql & " Order by Action_Id"
    Set Userprivilege = New ADODB.Recordset
        Userprivilege.Open sql, objConnection, adOpenForwardOnly, adLockReadOnly
    
    If Userprivilege.EOF = True Then
        MsgBox "No Privilege found for this user.", vbExclamation
        getLoadedUserPriveleges = False
    Else
        getLoadedUserPriveleges = True
  
        Do While Not Userprivilege.EOF
            LoadPermissions(Count, 1) = (Userprivilege!read_data)
            LoadPermissions(Count, 2) = (Userprivilege!Create_data)
            LoadPermissions(Count, 3) = (Userprivilege!Edit_data)
            LoadPermissions(Count, 4) = (Userprivilege!Delete_data)
            Count = Count + 1
            Userprivilege.MoveNext
        Loop
    End If
   Set Userprivilege = Nothing
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "moduser", "getUserPriveleges", True)

End Function



Public Sub do_permissioms(Item)
On Error GoTo ErrorHandler

With frmUserAccount
If frmUserAccount.listActions.ListIndex = -1 Then Exit Sub
Select Case Item

Case 1:
        If .chkRead.value = 1 Then
        LoadPermissions(frmUserAccount.listActions.ItemData(frmUserAccount.listActions.ListIndex), 1) = "Y"
        Else
        LoadPermissions(frmUserAccount.listActions.ItemData(frmUserAccount.listActions.ListIndex), 1) = "N"
        End If
Case 2:
        If .chkInsert.value = 1 Then
        LoadPermissions(frmUserAccount.listActions.ItemData(frmUserAccount.listActions.ListIndex), 2) = "Y"
        Else
        LoadPermissions(frmUserAccount.listActions.ItemData(frmUserAccount.listActions.ListIndex), 2) = "N"
        End If
Case 3:
        If .chkUpdate.value = 1 Then
        LoadPermissions(frmUserAccount.listActions.ItemData(frmUserAccount.listActions.ListIndex), 3) = "Y"
        Else
        LoadPermissions(frmUserAccount.listActions.ItemData(frmUserAccount.listActions.ListIndex), 3) = "N"
        End If
Case 4:
        If .chkDelete.value = 1 Then
        LoadPermissions(frmUserAccount.listActions.ItemData(frmUserAccount.listActions.ListIndex), 4) = "Y"
        Else
        LoadPermissions(frmUserAccount.listActions.ItemData(frmUserAccount.listActions.ListIndex), 4) = "N"
        End If
End Select

End With

Exit Sub
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "moduser", "getUserPriveleges", True)

End Sub


Public Sub loadActionPermisions(Item)
On Error GoTo ErrorHandler

With frmUserAccount


If LoadPermissions(Item, 1) = "Y" Then
        .chkRead.value = 1
Else
.chkRead.value = 0
End If

If LoadPermissions(Item, 2) = "Y" Then
        .chkInsert.value = 1
Else
.chkInsert.value = 0
End If
If LoadPermissions(Item, 3) = "Y" Then
        .chkUpdate.value = 1
Else
.chkUpdate.value = 0
End If
If LoadPermissions(Item, 4) = "Y" Then
        .chkDelete.value = 1
Else
.chkDelete.value = 0
End If
End With

Exit Sub
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "moduser", "getUserPriveleges", True)

End Sub
Public Sub doActionPermisions()

'Select Case gRecordUserType

loadActionPermisions (frmUserAccount.listActions.ItemData(frmUserAccount.listActions.ListIndex))
'Case RECORD_NEW:
'Case RECORD_EDIT: loadActionPermisions (frmUserAccount.listActions.ItemData(frmUserAccount.listActions.ListIndex))
'End Select


End Sub
Public Function InitialisePermission()

On Error GoTo ErrorHandler
    Dim Ctr As Byte
    
    DoEvents
    For Ctr = 1 To frmUserAccount.listActions.ListCount
    LoadPermissions(Ctr, 1) = "N"
    LoadPermissions(Ctr, 2) = "N"
    LoadPermissions(Ctr, 3) = "N"
    LoadPermissions(Ctr, 4) = "N"
    Next Ctr
    frmUserAccount.chkDelete.value = 0
    frmUserAccount.chkInsert.value = 0
    frmUserAccount.chkRead.value = 0
    frmUserAccount.chkUpdate.value = 0
    
    
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modUserAccount", "InitialiseUserAccount", True)

End Function

Public Function LoadCityNames()
On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rslocal As ADODB.Recordset
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection

    Set rslocal = objOrganisation_s.getCities()

    With frmUserAccount
            
            .cmbCity.Clear
            If Not rslocal Is Nothing Then
                Do Until rslocal.EOF
                    .cmbCity.AddItem rslocal!cityName
                    .cmbCity.ItemData(.cmbCity.NewIndex) = rslocal!id
                    rslocal.MoveNext
                Loop
                Set rslocal = Nothing
            End If
    End With
    Set objOrganisation_s = Nothing

Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modlogon", "LoadChurchComboBox", True)

End Function




Public Function LoadChurchComboBoxWithAll(ByVal sCity As Long)
On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rslocal As ADODB.Recordset
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection

    Set rslocal = objOrganisation_s.getChurchName(sCity)

    With frmUserAccount
            
            .cmbChurch.Clear
            .cmbChurch.AddItem "ALL"
            .cmbChurch.ItemData(.cmbChurch.NewIndex) = 0
            If Not rslocal Is Nothing Then
                Do Until rslocal.EOF
                    .cmbChurch.AddItem rslocal!Name
                    .cmbChurch.ItemData(.cmbChurch.NewIndex) = rslocal!id
                    rslocal.MoveNext
                Loop
                Set rslocal = Nothing
            End If
    End With
    Set objOrganisation_s = Nothing

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modUserAccount", "LoadChurchComboBoxWithAll", True)

End Function
