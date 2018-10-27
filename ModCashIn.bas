Attribute VB_Name = "ModCashIn"
Option Explicit





Public Function InitialiseCashIn()
On Error GoTo ErrorHandler

    Dim ctrl As Control


    For Each ctrl In frmCashIn.Controls
        
        If TypeOf ctrl Is TextBox Then ctrl.Text = ""
        If TypeOf ctrl Is ComboBox Then ctrl.ListIndex = -1
        If TypeOf ctrl Is MaskEdBox Then ctrl.Text = ""
        Set ctrl = Nothing
        
    Next ctrl
    DoEvents
    

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashIn", "InitialiseCashIn", True)

End Function
Public Function SaveCashIn() As Boolean
On Error GoTo ErrorHandler

    Dim objCashIn As CMScashin.clsCashIn
    Dim objCashIn_s As CMScashin.clsCashIn_s


    SaveCashIn = False
                            
    'Member record
    Set objCashIn = New CMScashin.clsCashIn
    Set objCashIn_s = New CMScashin.clsCashIn_s
    Set objCashIn_s.DatabaseConnection = objConnection


    PopulateCashInObject objCashIn

    'Insert or Update record
    If gRecordMode = RECORD_NEW Then
    
        objCashIn_s.InsertCashIn objCashIn
        gCashInId = objCashIn_s.NewCashInId
        frmCashIn.txtCashInId = gCashInId
        
    ElseIf gRecordMode = RECORD_EDIT Then
        
        objCashIn.id = gCashInId
        objCashIn_s.UpdateCashIn objCashIn
        
    
'    ElseIf gRecordMode = RECORD_DELETE Then
'
'        objMember_s.Deletemember gmemberId
'
    End If
    
    SaveCashIn = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashin", "SaveCashIn", True)
    
End Function

Public Function PopulateCashInObject(objCashIn As CMScashin.clsCashIn)
On Error GoTo ErrorHandler
'save data from form to newly created object for insert or update

   


    With frmCashIn
            'Push recordset results to form fields
            
            objCashIn.Amount = .txtAmount.Text
            objCashIn.Comment = .txtComment.Text
            objCashIn.gst = .txtGST.Text
            objCashIn.Item = .cmbItem.Text
            objCashIn.Total_Amount = .txtNetAmount.Text
            If .dteDateOfCashIn <> "" Then objCashIn.DateOfCashin = .dteDateOfCashIn.FormattedText

    End With
    
    
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashIn", "PopulateCashInObject", True)

End Function

Public Function ValidateCashIn() As Boolean

On Error GoTo ErrorHandler

    ValidateCashIn = False

    With frmCashIn

            If Trim(.txtAmount.Text) = "" Then
                MsgBox "Required field missing - Amount must be entered.", vbExclamation
                .txtNetAmount.SetFocus
                Exit Function
            End If
        
            If Trim(.txtGST.Text) = "" Then
                MsgBox "Required field missing - GST must be entered.", vbExclamation
                .txtGST.SetFocus
                Exit Function
            End If
        
            If Trim(.txtNetAmount.Text) = "" Then
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

    ValidateCashIn = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashIn", "ValidateCashIn", True)

End Function

Function LoadCashInComboBox()
On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rsLocal As ADODB.Recordset
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection
    
    'Payment
    Set rsLocal = objOrganisation_s.getCashflowItem("CASHIN")

    With frmCashIn
            
            .cmbItem.Clear
            If Not rsLocal Is Nothing Then
                Do Until rsLocal.EOF
                    .cmbItem.AddItem rsLocal!ITEMNAME
                    .cmbItem.ItemData(.cmbItem.NewIndex) = rsLocal!id
                    rsLocal.MoveNext
                Loop
                Set rsLocal = Nothing
            End If
                   
    End With

    
    Set objOrganisation_s = Nothing



Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashIn", "LoadCashInComboBox", True)

End Function

Public Function LoadCashInDefualtValue()

With frmCashIn
       .dteDateOfCashIn.Text = Format(Now(), DATE_FORMAT)
End With

End Function

Public Function loadItemCodeAndGstAmountforCashIn()
  On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rsLocal As ADODB.Recordset
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection
    
    'Payment
    Set rsLocal = objOrganisation_s.getCashflowItemcodeandGst(frmCashIn.cmbItem.Text)

    With frmCashIn
            
            
            If Not rsLocal Is Nothing Then
                      .txtItemCode.Text = rsLocal!ITEMCODE
                      .txtGST.Text = rsLocal!gst
                      Call ValidNumericEntry(.txtGST)
                Set rsLocal = Nothing
            End If
                   
    End With
    Set objOrganisation_s = Nothing
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashIn", "LoadCashInComboBox", True)

End Function

Public Function DisplayCashIn()

    Dim objCashIn_s As CMScashin.clsCashIn_s
    
    Dim rsLocal As ADODB.Recordset
   ' Dim intCnt As Integer
   ' Dim strDatabaseValue As String
    
    On Error GoTo ErrorHandler
    
    Call InitialiseCashIn
    
    'Retrieve Prospect record and display on form
    Set objCashIn_s = New CMScashin.clsCashIn_s
    Set objCashIn_s.DatabaseConnection = objConnection
    Set rsLocal = objCashIn_s.getByCashInId(gCashInId)

    If rsLocal Is Nothing Then
        Exit Function
    End If

    Screen.MousePointer = vbHourglass

    With frmCashIn
        
        .txtAmount.Text = Format(rsLocal!Amount, NUMERIC_FORMAT)
        .txtComment.Text = ConvertNull(rsLocal!Comment)
        .txtCashInId.Text = ConvertNull(rsLocal!id)
        .txtGST.Text = Format(rsLocal!gst, NUMERIC_FORMAT)
        .txtNetAmount.Text = Format(rsLocal!net_amount, NUMERIC_FORMAT)
        .cmbItem.Text = ConvertNull(rsLocal!Item)
        .dteDateOfCashIn.Text = Format(rsLocal!DateOfCashin, DATE_FORMAT)
        
    End With
    
   Set objCashIn_s = Nothing
   loadItemCodeforCashIn
   Screen.MousePointer = vbDefault
Exit Function

ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashIn", "DisplayCashin", True)

End Function


Public Function loadItemCodeforCashIn()
  On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rsLocal As ADODB.Recordset
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection
    
    'Payment
    
    Set rsLocal = objOrganisation_s.getCashflowItemcodeandGst(frmCashIn.cmbItem.Text)

    With frmCashIn
            
            
            If Not rsLocal Is Nothing Then
                      .txtItemCode.Text = rsLocal!ITEMCODE
                      Call ValidNumericEntry(.txtGST)
                Set rsLocal = Nothing
            End If
                   
    End With
    Set objOrganisation_s = Nothing
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashIn", "LoadCashInComboBox", True)

End Function
