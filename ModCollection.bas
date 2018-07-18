Attribute VB_Name = "ModCollection"
Option Explicit

Public mCollectionMode As String

Public Function CheckCollectionSecurity(sFunction As String) As Boolean
Dim privilege_ctr As Byte
CheckCollectionSecurity = False
     privilege_ctr = GetPrivileges("COLLECTION")
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
             ElseIf Permissions(privilege_ctr, 2) = "Y" Then CheckCollectionSecurity = True
           End If
        Case RECORD_EDIT
            If Permissions(privilege_ctr, 3) = "N" Then
                MsgBox "Invalid access - Function not available for current user access level.", vbExclamation
                Exit Function
             ElseIf Permissions(privilege_ctr, 3) = "Y" Then CheckCollectionSecurity = True
           End If
        Case RECORD_DELETE
            If Permissions(privilege_ctr, 4) = "N" Then
                MsgBox "Invalid access - Function not available for current user access level.", vbExclamation
                Exit Function
             ElseIf Permissions(privilege_ctr, 4) = "Y" Then CheckCollectionSecurity = True
           End If
        End Select
End Function

Public Function InitialiseCollection()
On Error GoTo ErrorHandler

    Dim ctrl As Control


    For Each ctrl In frmCollection.Controls
        
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
Public Function SaveCollection() As Boolean
On Error GoTo ErrorHandler

    Dim objcollection As CMSCollection.clsCollection
    Dim objCollection_s As CMSCollection.clsCollection_s


    SaveCollection = False
                            
    'Member record
    Set objcollection = New CMSCollection.clsCollection
    Set objCollection_s = New CMSCollection.clsCollection_s
    Set objCollection_s.DatabaseConnection = objConnection


    PopulateCollectionObject objcollection

    'Insert or Update record
    If gRecordMode = RECORD_NEW Then
    
        objCollection_s.InsertCollection objcollection
        gCollectionId = objCollection_s.NewCollectionId
        frmCollection.txtCollectionNo = gCollectionId
        
    ElseIf gRecordMode = RECORD_EDIT Then
        
        objcollection.COL_ID = gCollectionId
        objCollection_s.UpdateCollection objcollection
        
    
'    ElseIf gRecordMode = RECORD_DELETE Then
'
'        objMember_s.Deletemember gmemberId
'
    End If
    
    SaveCollection = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCollection", "SaveCollection", True)
    
End Function

Public Function PopulateCollectionObject(objcollection As CMSCollection.clsCollection)
On Error GoTo ErrorHandler
'save data from form to newly created object for insert or update

   


    With frmCollection
            'Push recordset results to form fields
            objcollection.Amount = .txtAmount.Text
            objcollection.Comments = .txtComments.Text
            objcollection.Payment = .cmbType.Text
            objcollection.types = .cmbTypes.Text
            objcollection.UserName = .txtUserName.Text
            objcollection.ChurchId = gChurchId
            
            If .dteDateofPayment <> "" Then objcollection.Dateofcollection = .dteDateofPayment.FormattedText

    End With
    
    
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCllection", "PopulateCollectionobject", True)

End Function


Public Function LoadCollectionComboBox()
On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rslocal As ADODB.Recordset
     Dim rslocal1 As ADODB.Recordset
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection
    
    'Collection type
    Set rslocal = objOrganisation_s.getCollectionType
    Set rslocal1 = objOrganisation_s.getPayment

    With frmCollection
            
            .cmbType.Clear
            If Not rslocal Is Nothing Then
                Do Until rslocal.EOF
                    .cmbType.AddItem rslocal!Collection_Type
                    .cmbType.ItemData(.cmbType.NewIndex) = rslocal!Id
                    rslocal.MoveNext
                Loop
                Set rslocal = Nothing
            End If
                   
            .cmbTypes.Clear
            If Not rslocal1 Is Nothing Then
                Do Until rslocal1.EOF
                    .cmbTypes.AddItem rslocal1!Payment_type
                    .cmbTypes.ItemData(.cmbTypes.NewIndex) = rslocal1!Id
                    rslocal1.MoveNext
                Loop
                Set rslocal1 = Nothing
            End If
    End With

    
    Set objOrganisation_s = Nothing



Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modcollection", "LoadcollectionComboBox", True)

End Function

Public Function LoadcollectionDefualtValue()

With frmCollection
       .dteDateofPayment.Text = Format(Now(), DATE_FORMAT)
       .txtUserName.Text = UserName
End With

End Function

Public Function ValidateCollection() As Boolean

On Error GoTo ErrorHandler

    ValidateCollection = False

    With frmCollection

           
        
            If .cmbType.Text = "" Then
                MsgBox "Required field missing - Type of Collection must be selected.", vbExclamation
                .cmbType.SetFocus
                Exit Function
            End If
            If .cmbTypes.Text = "" Then
                MsgBox "Required field missing - Type of payment must be selected.", vbExclamation
                .cmbTypes.SetFocus
                Exit Function
            End If
             If Trim(.txtAmount.Text) = "" Then
                MsgBox "Required field missing - Amount  must be entered.", vbExclamation
                .txtAmount.SetFocus
                Exit Function
            End If
        
    End With

    ValidateCollection = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modPayment", "ValidatePayment", True)

End Function

Public Sub GenerateCollectionList(sql As String)
 On Error GoTo ErrorHandler
 
    
     Dim rslocal As ADODB.Recordset
     Dim strId  As String
     Dim itmx As ListItem
     
    
 
 
     Screen.MousePointer = vbHourglass
 
     With frmCollectionSearch
             .ListCollectionView.ListItems.Clear
     
         '==============================================================================
         'get Prospect list
            Set rslocal = New ADODB.Recordset
            rslocal.Open sql, objConnection, adOpenForwardOnly, adLockReadOnly
             If Not rslocal.EOF Then
                 Do While Not rslocal.EOF
                     strId = CStr(rslocal!COL_ID)
                     Set itmx = .ListCollectionView.ListItems.Add(, , CStr(rslocal!COL_ID))
                                     If Not IsNull(rslocal!Payment) Then itmx.SubItems(1) = CStr(rslocal!Payment)
                                     If Not IsNull(rslocal!DATE_OF_COLLECTION) Then itmx.SubItems(2) = CStr(rslocal!DATE_OF_COLLECTION)
                                     If Not IsNull(rslocal!Amount) Then itmx.SubItems(3) = CStr(rslocal!Amount)
                                     If Not IsNull(rslocal!Comments) Then itmx.SubItems(4) = CStr(rslocal!Comments)
                     Set itmx = Nothing
                     rslocal.MoveNext
                     
                 Loop
                 CollectionSelected = True
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

Public Function DisplayCollection()

    Dim objCollection_s As CMSCollection.clsCollection_s
    
    Dim rslocal As ADODB.Recordset
   ' Dim intCnt As Integer
   ' Dim strDatabaseValue As String
    
    On Error GoTo ErrorHandler
    
    Call InitialiseCollection
    
    'Retrieve Prospect record and display on form
    Set objCollection_s = New CMSCollection.clsCollection_s
    Set objCollection_s.DatabaseConnection = objConnection
    Set rslocal = objCollection_s.getByCollectionId(gCollectionId)

    If rslocal Is Nothing Then
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    With frmCollection
        
        
        .dteDateofPayment.Text = Format(rslocal!DATE_OF_COLLECTION, DATE_FORMAT)
        .txtCollectionNo.Text = ConvertNull(rslocal!COL_ID)
        .cmbType.Text = ConvertNull(rslocal!Payment)
        .cmbTypes.Text = ConvertNull(rslocal!Type)
        .txtAmount.Text = ConvertNull(rslocal!Amount)
        .txtComments.Text = ConvertNull(rslocal!Comments)
        .txtUserName.Text = ConvertNull(rslocal!USER_NAME)
        
    End With
   Set objCollection_s = Nothing
   Screen.MousePointer = vbDefault
Exit Function

ErrorHandler:

    
    
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modProspect", "DisplayProspect", True)

End Function

