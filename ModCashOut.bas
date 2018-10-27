Attribute VB_Name = "ModCashOut"
Option Explicit



Public Function CheckCashflowSecurity(sFunction As String) As Boolean
Dim privilege_ctr As Byte
CheckCashflowSecurity = False
privilege_ctr = GetPrivileges("Cash Flow")
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
             ElseIf Permissions(privilege_ctr, 2) = "Y" Then CheckCashflowSecurity = True
           End If
        Case RECORD_EDIT
            If Permissions(privilege_ctr, 3) = "N" Then
                MsgBox "Invalid access - Function not available for current user access level.", vbExclamation
                Exit Function
             ElseIf Permissions(privilege_ctr, 3) = "Y" Then CheckCashflowSecurity = True
           End If
        Case RECORD_DELETE
            If Permissions(privilege_ctr, 4) = "N" Then
                MsgBox "Invalid access - Function not available for current user access level.", vbExclamation
                Exit Function
             ElseIf Permissions(privilege_ctr, 4) = "Y" Then CheckCashflowSecurity = True
           End If
        End Select
End Function

Public Function InitialiseCashout()
On Error GoTo ErrorHandler

    Dim ctrl As Control


    For Each ctrl In frmCashOut.Controls
        
        If TypeOf ctrl Is TextBox Then ctrl.Text = ""
        If TypeOf ctrl Is ComboBox Then ctrl.ListIndex = -1
        If TypeOf ctrl Is MaskEdBox Then ctrl.Text = ""
        Set ctrl = Nothing
        
    Next ctrl
    DoEvents
    

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "Cashout", "InitialiseCashout", True)

End Function
Public Function SaveCashOut() As Boolean
On Error GoTo ErrorHandler

    Dim objCashOut As CMScashout.clsCashOut
    Dim objCashOut_s As CMScashout.clsCashout_s


    SaveCashOut = False
                            
    'Member record
    Set objCashOut = New CMScashout.clsCashOut
    Set objCashOut_s = New CMScashout.clsCashout_s
    Set objCashOut_s.DatabaseConnection = objConnection


    PopulateCashoutObject objCashOut

    'Insert or Update record
    If gRecordMode = RECORD_NEW Then
    
        objCashOut_s.InsertCashOut objCashOut
        gCashoutId = objCashOut_s.NewCashOutId
        frmCashOut.txtCashoutId = gCashoutId
        
    ElseIf gRecordMode = RECORD_EDIT Then
        
        objCashOut.Id = gCashoutId
        objCashOut_s.UpdateCashOut objCashOut
        
    
'    ElseIf gRecordMode = RECORD_DELETE Then
'
'        objMember_s.Deletemember gmemberId
'
    End If
    
    SaveCashOut = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashOut", "SaveCashOut", True)
    
End Function

Public Function PopulateCashoutObject(objCashOut As CMScashout.clsCashOut)
On Error GoTo ErrorHandler
'save data from form to newly created object for insert or update

   


    With frmCashOut
            'Push recordset results to form fields
            objCashOut.Amount = .txtAmount.Text
            objCashOut.Comment = .txtCommand.Text
            objCashOut.gst = .txtGST.Text
            objCashOut.Item = .cmbItem.Text
            objCashOut.Item_code = .txtItemCode.Text
            objCashOut.Total_Amount = .txtNetAmount.Text
            If .dteDateofCashOut <> "" Then objCashOut.DateOfCashOut = .dteDateofCashOut.FormattedText

    End With
    
    
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashOut", "PopulateCashOutObject", True)

End Function

Public Function ValidateCashOut() As Boolean

On Error GoTo ErrorHandler

    ValidateCashOut = False

    With frmCashOut

            If Trim(.txtNetAmount.Text) = "" Then
                MsgBox "Required field missing - Amount must be entered.", vbExclamation
                .txtNetAmount.SetFocus
                Exit Function
            End If
        
            If Trim(.txtGST.Text) = "" Then
                MsgBox "Required field missing - GST must be entered.", vbExclamation
                .txtGST.SetFocus
                Exit Function
            End If
        
            If Trim(.txtAmount.Text) = "" Then
                MsgBox "Required field missing - NetAmount must be entered.", vbExclamation
                .txtNetAmount.SetFocus
                Exit Function
            End If
        
            If .cmbItem.Text = "" Then
                MsgBox "Required field missing - Item must be selected.", vbExclamation
                .cmbItem.SetFocus
                Exit Function
            End If
            
                        
    End With

    ValidateCashOut = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashIn", "ValidateCashIn", True)

End Function

Function LoadCashOutComboBox()
On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rslocal As ADODB.Recordset
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection
    
    'Payment
    Set rslocal = objOrganisation_s.getCashflowItem("CASHOUT")

    With frmCashOut
            
            .cmbItem.Clear
            If Not rslocal Is Nothing Then
                Do Until rslocal.EOF
                    .cmbItem.AddItem rslocal!ITEMNAME
                    .cmbItem.ItemData(.cmbItem.NewIndex) = rslocal!Id
                    rslocal.MoveNext
                Loop
                Set rslocal = Nothing
            End If
                   
    End With

    
    Set objOrganisation_s = Nothing



Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashout", "LoadCashoutComboBox", True)

End Function

Public Function LoadCashOutDefualtValue()

With frmCashOut
       .dteDateofCashOut.Text = Format(Now(), DATE_FORMAT)
End With

End Function
Public Function loadItemCodeAndGstAmountforCashOut()
  On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rslocal As ADODB.Recordset
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection
    
    'Payment
    Set rslocal = objOrganisation_s.getCashflowItemcodeandGst(frmCashOut.cmbItem.Text)

    With frmCashOut
            
            
            If Not rslocal Is Nothing Then
                      .txtItemCode.Text = rslocal!ITEMCODE
                      .txtGST.Text = rslocal!gst
                      Call ValidNumericEntry(.txtGST)
                Set rslocal = Nothing
            End If
                   
    End With
    Set objOrganisation_s = Nothing
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashOut", "loadItemCodeAndGstAmountforCashOut", True)

End Function

Public Function loadItemCodeForCashOut()
  On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rslocal As ADODB.Recordset
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection
    
    'Payment
    Set rslocal = objOrganisation_s.getCashflowItemcodeandGst(frmCashOut.cmbItem.Text)
    

    With frmCashOut
            
            
            If Not rslocal Is Nothing Then
                      .txtItemCode.Text = rslocal!ITEMCODE
                      Call ValidNumericEntry(.txtGST)
                Set rslocal = Nothing
            End If
                   
    End With
    Set objOrganisation_s = Nothing
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashOut", "loadItemCodeAndGstAmountforCashOut", True)

End Function

Public Function DisplayCashOut()

    Dim objCashOut_s As CMScashout.clsCashout_s
    
    Dim rslocal As ADODB.Recordset
   ' Dim intCnt As Integer
   ' Dim strDatabaseValue As String
    
    On Error GoTo ErrorHandler
    
    Call InitialiseCashout
    
    'Retrieve Prospect record and display on form
    Set objCashOut_s = New CMScashout.clsCashout_s
    Set objCashOut_s.DatabaseConnection = objConnection
    Set rslocal = objCashOut_s.getByCashoutId(gCashoutId)

    If rslocal Is Nothing Then
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    With frmCashOut
        
        .txtAmount.Text = Format(rslocal!Amount, NUMERIC_FORMAT)
        .txtCommand.Text = ConvertNull(rslocal!Comment)
        .txtCashoutId.Text = ConvertNull(rslocal!Id)
        .txtGST.Text = Format(rslocal!gst, NUMERIC_FORMAT)
        .txtNetAmount.Text = Format(rslocal!net_amount, NUMERIC_FORMAT)
        .cmbItem.Text = ConvertNull(rslocal!Item)
        .dteDateofCashOut.Text = Format(rslocal!DateOfCashOut, DATE_FORMAT)
        
    End With
    
   Set objCashOut_s = Nothing
   loadItemCodeForCashOut
   Screen.MousePointer = vbDefault
Exit Function

ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashOut", "DisplayCashout", True)

End Function
