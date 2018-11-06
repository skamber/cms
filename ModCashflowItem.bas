Attribute VB_Name = "ModCashflowItem"
Public Function CheckCashflowItemSecurity(sFunction As String) As Boolean
Dim privilege_ctr As Byte

CheckCashflowItemSecurity = False
    privilege_ctr = GetPrivileges("Cash Flow Item")
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
             ElseIf Permissions(privilege_ctr, 2) = "Y" Then CheckCashflowItemSecurity = True
           End If
        Case RECORD_EDIT
            If Permissions(privilege_ctr, 3) = "N" Then
                MsgBox "Invalid access - Function not available for current user access level.", vbExclamation
                Exit Function
             ElseIf Permissions(privilege_ctr, 3) = "Y" Then CheckCashflowItemSecurity = True
           End If
        Case RECORD_DELETE
            If Permissions(privilege_ctr, 4) = "N" Then
                MsgBox "Invalid access - Function not available for current user access level.", vbExclamation
                Exit Function
             ElseIf Permissions(privilege_ctr, 4) = "Y" Then CheckCashflowItemSecurity = True
           End If
        End Select
End Function

Public Function InitialiseCashflowItem()
On Error GoTo ErrorHandler

    Dim ctrl As Control


    For Each ctrl In frmCashFlowItem.Controls
        
        If TypeOf ctrl Is TextBox Then ctrl.Text = ""
        If TypeOf ctrl Is ComboBox Then ctrl.ListIndex = -1
        If TypeOf ctrl Is MaskEdBox Then ctrl.Text = ""
        Set ctrl = Nothing
        
    Next ctrl
    DoEvents
    

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "CashflowItem", "InitialiseCashflowItem", True)

End Function
Public Function ValidateCashflowItem() As Boolean

On Error GoTo ErrorHandler

    ValidateCashflowItem = False

    With frmCashFlowItem

            If Trim(.txtItemName.Text) = "" Then
                MsgBox "Required field missing - Account Name must be entered.", vbExclamation
                .txtItemName.SetFocus
                Exit Function
            End If
        
            If Trim(.txtItemCode.Text) = "" Then
                MsgBox "Required field missing - Item Code must be entered.", vbExclamation
                .txtItemCode.SetFocus
                Exit Function
            End If
        
            If Trim(.txtGSTRate.Text) = "" Then
                MsgBox "Required field missing - GST must be entered.", vbExclamation
                .txtGSTRate.SetFocus
                Exit Function
            End If
        
            If .cmbItemType.Text = "" Then
                MsgBox "Required field missing - Item type must be selected.", vbExclamation
                .cmbItemType.SetFocus
                Exit Function
            End If
            
                        
    End With

    ValidateCashflowItem = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashIn", "ValidateCashIn", True)

End Function

Public Function SaveCashflowItem() As Boolean
On Error GoTo ErrorHandler

    Dim objCashflowItem As CMSCashFlowItem.clsCashflowItem
    Dim objCashflowItem_s As CMSCashFlowItem.clsCashflowItem_s


    SaveCashflowItem = False
                            
    'Member record
    Set objCashflowItem = New CMSCashFlowItem.clsCashflowItem
    Set objCashflowItem_s = New CMSCashFlowItem.clsCashflowItem_s
    Set objCashflowItem_s.DatabaseConnection = objConnection


    PopulateCashflowItemObject objCashflowItem

    'Insert or Update record
    If gRecordMode = RECORD_NEW Then
    
        objCashflowItem_s.InsertCashflowItem objCashflowItem
        gCashflowItemId = objCashflowItem_s.NewCashFlowItemId
        frmCashFlowItem.txtId = gCashflowItemId
        
    ElseIf gRecordMode = RECORD_EDIT Then
        
        objCashflowItem.Id = gCashflowItemId
        objCashflowItem_s.UpdateCashflowItem objCashflowItem
        
    
'    ElseIf gRecordMode = RECORD_DELETE Then
'
'        objMember_s.Deletemember gmemberId
'
    End If
    
    SaveCashflowItem = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashflowItem", "SaveCashflowitem", True)
    
End Function

Public Function PopulateCashflowItemObject(objCashflowItem As CMSCashFlowItem.clsCashflowItem)
On Error GoTo ErrorHandler
'save data from form to newly created object for insert or update

   


    With frmCashFlowItem
            'Push recordset results to form fields
            objCashflowItem.ITEMCODE = .txtItemCode.Text
            objCashflowItem.ITEMNAME = .txtItemName.Text
            objCashflowItem.gst = .txtGSTRate.Text
            objCashflowItem.ItemType = .cmbItemType.Text
            
    End With
    
    
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashflowitem", "PopulateCashflowItemObject", True)

End Function
Public Function DispayCashflowItem()

    Dim objCashflowItem_s As CMSCashFlowItem.clsCashflowItem_s
    
    Dim rslocal As ADODB.Recordset
   
    On Error GoTo ErrorHandler
    
    Call InitialiseCashflowItem
    
    'Retrieve Prospect record and display on form
    Set objCashflowItem_s = New CMSCashFlowItem.clsCashflowItem_s
    Set objCashflowItem_s.DatabaseConnection = objConnection
    Set rslocal = objCashflowItem_s.getByCashflowItemId(gCashflowItemId)

    If rslocal Is Nothing Then
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    With frmCashFlowItem
        
        .txtGSTRate.Text = Format(rslocal!gst, NUMERIC_FORMAT)
        .txtId.Text = ConvertNull(rslocal!Id)
        .txtItemCode.Text = ConvertNull(rslocal!ITEMCODE)
        .txtItemName.Text = ConvertNull(rslocal!ITEMNAME)
        .cmbItemType.Text = ConvertNull(rslocal!Type)
        
    End With
    
   Set objCashflowItem_s = Nothing
   Screen.MousePointer = vbDefault
Exit Function

ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashflowItem", "DisplayCashflowItem", True)

End Function

Public Sub LoadCashflowItemList()

 On Error GoTo ErrorHandler
 
    ' Dim objFollowup_s As PACMSFollowUP.clsFollowup_s
     Dim rslocal As ADODB.Recordset
     Dim strId  As String
     Dim itmx As ListItem
     Dim sql As String

 
     Screen.MousePointer = vbHourglass
 
     With frmCashFlowItem
             .ListCashflowItemView.ListItems.Clear
     
         '==============================================================================
        
            Set rslocal = New ADODB.Recordset
            sql = "SELECT * FROM cashflowitem"
            rslocal.Open sql, objConnection, adOpenForwardOnly, adLockReadOnly
             If Not rslocal Is Nothing Then
                 Do While Not rslocal.EOF
                     
                     Set itmx = .ListCashflowItemView.ListItems.Add(, , CStr(rslocal!Id))
                                    
                                     If Not IsNull(rslocal!ITEMNAME) Then itmx.SubItems(1) = CStr(rslocal!ITEMNAME)
                                     If Not IsNull(rslocal!ITEMCODE) Then itmx.SubItems(2) = CStr(rslocal!ITEMCODE)
                                     If Not IsNull(rslocal!gst) Then itmx.SubItems(3) = CStr(Format(rslocal!gst, NUMERIC_FORMAT))
                                     If Not IsNull(rslocal!Type) Then itmx.SubItems(4) = CStr(rslocal!Type)
                     Set itmx = Nothing
                     rslocal.MoveNext
                 Loop
             
                 Set rslocal = Nothing
             End If
     End With
     
     Screen.MousePointer = vbDefault
     CashflowItemSelected = True
     
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "FrmCashflow", "LoadCashInList", True)

End Sub


