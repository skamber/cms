Attribute VB_Name = "modMain"
Option Explicit
Public Const INI_FILE_NAME = "CMS.INI"
Public Const gDatabasePassword = "WingedBull"
Public gmemberId As Long
Public gchildId As Long
Public gPaymentId As Long
Public gReceiptId As Long
Public gCollectionId As Long
Public gCashInId As Long
Public gInvoiceId As Long
Public gInvoiceItemId As Long
Public gCashoutId As Long
Public gCashflowItemId As Long
Public gChurchId As Long
Public gCityId As Long
Public gCityName As String
Public PrivilegeBookMark As Variant
Public Userprivilege As ADODB.Recordset
Public objConnection As ADODB.Connection
Public objError As New CMSErrorHandler.clsErrorHandler
Public NumLogIN As Long
Public MemberSelected As Boolean
Public ChildSelected As Boolean
Public PaymentSelected As Boolean
Public CollectionSelected As Boolean
Public InvoiceSelected As Boolean
Public ReceiptSelected As Boolean
Public InvoiceItemSelected As Boolean
Public CashInSelected As Boolean
Public CashOutSelected As Boolean
Public CashflowItemSelected As Boolean
Public systemManager As Boolean
Public ReportView As Boolean
Public CompulsoryChangePassword As Boolean
Public UserName As String
Public Permissions(1 To 20, 1 To 4) As String
Public LoadPermissions(1 To 20, 1 To 4) As String
Public Cities As New COLLECTION
Public UserId As Long
Public Const DATE_FORMAT = "dd/mm/yyyy"
Public Const DATE_TIME_FORMAT = "yyyy/mm/dd hh:mm:ss"
Public Const NUMERIC_FORMAT = "###,###,###,##0.00"
Public Const MEMBER = "Member"
'Public Const CASHFLOW = "CashFlow"
Public Const CASHIN = "CashIn"
Public Const CASHFLOW_VIEW = "cashflow"
Public Const CASHOUT = "CashOut"
Public Const CASHFLOWITEM = "CashFlowItem"
Public Const Payment = "payment"
Public Const COLLECTION = "Collection"
Public Const INVOICE = "Invoice"
Public Const RECEIPT = "Receipt"
Public Const REPORTING = "Reporting"
Public Const Member_information = "MemberMenu"
Public Const Member_Search = "SearchMembers"
Public Const Child_Search = "SearchChildren"
Public Const Children_information = "ChildrenMenu"
Public Const Payment_information = "PaymentMenu"
Public Const Payment_search = "PaymentSearch"
Public Const Collection_information = "CollectionMenu"
Public Const Invoice_information = "InvoiceMenu"
Public Const Invoice_search = "InvoiceSearch"
Public Const Receipt_information = "ReceiptMenu"
Public Const receipt_search = "ReceiptSearch"
Public Const collection_search = "CollectionSearch"
Public Const Report = "Report"
Public gRecordType As String    'indicate form/record type
Public gRecordMode As String    'regulates record mode
Public Const RECORD_NEW = "New"
Public Const RECORD_EDIT = "Edit"
Public Const RECORD_DELETE = "Delete"
Public Const RECORD_SAVE = "Save"
Public Const RECORD_CANCEL = "Cancel"
Public Const RECORD_READ = "Read"
Public Const RECORD_ADD = "Add"
Public Const RECORD_CHANGE = "Change"
Public Const RECORD_REMOVE = "Remove"
Public Const RECORD_COPY = "Copy"
Public Const RECORD_PRINT = "Print"
Public Const RECORD_PRIVIEW = "Preview"


Public Enum eSortOrder
    sortAscending = 0
    sortDescending = 1
End Enum
Public Enum eSortType
    sortAlpha = 0
    sortNumeric = 1
    sortDate = 2
End Enum


Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub CentreForm(frm As Form, intPosition As Integer)

    Select Case intPosition
                Case 1: frm.Move (MDIFrm.ScaleWidth - frm.Width) / 2, (MDIFrm.ScaleHeight - frm.Height) / 6
                Case 2: frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 6
                Case 3: frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - (frm.Height * 2))
    End Select

End Sub

Public Function ConnectDatabase()

    Dim sAppName As String * 100

    Dim sHost As String
    Dim sDatabaseName As String
    Dim sUserName As String
    Dim sPassword As String
    Dim sCertPath As String
    
    Dim sDefault As String * 100
    Dim sRet As String * 100
    Dim lSize As Long
    Dim sFileName As String * 255
    Dim lRet As Long
    
    On Error GoTo ErrorHandler
    
    ' connect to local database
    sAppName = "Connection"
    

    sDefault = ""
    sRet = ""
    lSize = 100
    sFileName = App.Path & "\" & INI_FILE_NAME
    
    lRet = GetPrivateProfileString(sAppName, "Host", sDefault, sRet, lSize, sFileName)
    sHost = Left(sRet, InStrB(1, sRet, Chr(0)) / 2)
    
    lRet = GetPrivateProfileString(sAppName, "Database", sDefault, sRet, lSize, sFileName)
    sDatabaseName = Left(sRet, InStrB(1, sRet, Chr(0)) / 2)
    
    lRet = GetPrivateProfileString(sAppName, "UserName", sDefault, sRet, lSize, sFileName)
    sUserName = Left(sRet, InStrB(1, sRet, Chr(0)) / 2)
    
    lRet = GetPrivateProfileString(sAppName, "UserPassword", sDefault, sRet, lSize, sFileName)
    sPassword = Left(sRet, InStrB(1, sRet, Chr(0)) / 2)
    sPassword = DecryptPassword(sPassword)

    
    lRet = GetPrivateProfileString(sAppName, "CertPath", sDefault, sRet, lSize, sFileName)
    sCertPath = Left(sRet, InStrB(1, sRet, Chr(0)) / 2)
    
            
    Set objConnection = New ADODB.Connection
    On Error Resume Next
    
    With objConnection
            .Open "Driver={MySQL ODBC 5.3 ANSI Driver};Server=" & sHost & ";Database=" & _
            sDatabaseName & "; User=" & sUserName & ";Password=" & sPassword & _
            ";sslca=" & sCertPath & "; sslverify=1; Option=3;"
            
            
    End With
                            

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMain", "ConnectDatabase", True)

End Function
Public Function SetToolbarControl()

    On Error GoTo ErrorHandler

    With MDIFrm
        Select Case gRecordType
            Case REPORTING, Member_Search, Child_Search, Payment_search, CASHFLOW_VIEW, Invoice_search
                .Toolbar1.Buttons.Item(1).Enabled = False
                .Toolbar1.Buttons.Item(2).Enabled = False
                .Toolbar1.Buttons.Item(3).Enabled = False
                .Toolbar1.Buttons.Item(4).Enabled = False
                .Toolbar1.Buttons.Item(5).Enabled = False
                .Toolbar1.Buttons.Item(6).Enabled = False
                .Toolbar1.Buttons.Item(7).Enabled = False
            Case MEMBER, Payment, COLLECTION, INVOICE, RECEIPT, Children_information, CASHIN, CASHOUT, CASHFLOWITEM
                Select Case gRecordMode
                    Case RECORD_READ, RECORD_CANCEL
                            .Toolbar1.Buttons.Item(1).Enabled = True
                            .Toolbar1.Buttons.Item(2).Enabled = True
                            .Toolbar1.Buttons.Item(3).Enabled = True
                            .Toolbar1.Buttons.Item(4).Enabled = False
                            .Toolbar1.Buttons.Item(5).Enabled = False
                            .Toolbar1.Buttons.Item(6).Enabled = True
                            .Toolbar1.Buttons.Item(7).Enabled = True
                
                    Case RECORD_NEW, RECORD_EDIT, RECORD_DELETE
                            .Toolbar1.Buttons.Item(1).Enabled = False
                            .Toolbar1.Buttons.Item(2).Enabled = False
                            .Toolbar1.Buttons.Item(3).Enabled = False
                            .Toolbar1.Buttons.Item(4).Enabled = True
                            .Toolbar1.Buttons.Item(5).Enabled = True
                            .Toolbar1.Buttons.Item(6).Enabled = True
                            .Toolbar1.Buttons.Item(7).Enabled = True
                            
                End Select
        End Select
    End With

    Exit Function
    
ErrorHandler:

    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMain", "SetToolbarControl", True)

End Function


Public Function ProcessRecord(ByVal sFunction As String)

    Dim iAnswer As Integer
    On Error GoTo ErrorHandler
    

   
    
    If ModChildren.mChiledMode = RECORD_ADD Or ModChildren.mChiledMode = RECORD_CHANGE Then
        MsgBox "Please Update or Cancel the Chiled Memeber before continuing", vbInformation, "CMS - Error Continuing"
        Exit Function
    End If
    
    
    If ModCollection.mCollectionMode = RECORD_ADD Or ModCollection.mCollectionMode = RECORD_CHANGE Then
        MsgBox "Please Update or Cancel the Collection before continuing", vbInformation, "CMS - Error Continuing"
        Exit Function
    End If
    If ModPayment.mPaymentMode = RECORD_ADD Or ModPayment.mPaymentMode = RECORD_CHANGE Then
        MsgBox "Please Update or Cancel the Payment before continuing", vbInformation, "CMS - Error Continuing"
        Exit Function
    End If
    If ModReceipt.mReceiptMode = RECORD_ADD Or ModReceipt.mReceiptMode = RECORD_CHANGE Then
        MsgBox "Please Update or Cancel the Strategy Red Flag before continuing", vbInformation, "CMS - Error Continuing"
        Exit Function
    End If
    
        
    Select Case sFunction
        
              
        Case RECORD_NEW
                'MDIFrm.pnlStatusBar.Panels(2).Text = ""
                                                
                Select Case gRecordType
                            
                    Case MEMBER:
                        If Not CheckMemberSecurity(sFunction) Then Exit Function
                        gRecordMode = sFunction
                        InitialiseMember
                        MDIFrm.ActiveForm.Form_Activate
                        frmMember.ZOrder 0
                        DoEvents
                        LoadMemberDefualtValue
                        frmMember.txtStatus.SetFocus
                        SetToolbarControl
                        
                    Case Children_information:
                        If Not CheckMemberSecurity(sFunction) Then Exit Function
                        gRecordMode = sFunction
                        InitialiseChild
                        MDIFrm.ActiveForm.Form_Activate
                        frmChildren.ZOrder 0
                        DoEvents
                        frmChildren.txtMemberMno.SetFocus
                        SetToolbarControl
                                
                    Case Payment:
                        If Not CheckPaymentSecurity(sFunction) Then Exit Function
                        gRecordMode = sFunction
                        InitialisePayment
                        MDIFrm.ActiveForm.Form_Activate
                        frmPayment.ZOrder 0
                        DoEvents
                        LoadPaymentDefualtValue
                        frmPayment.txtmemberNo.SetFocus
                        SetToolbarControl
                        
                    Case COLLECTION:
                    If Not CheckCollectionSecurity(sFunction) Then Exit Function
                        gRecordMode = sFunction
                        InitialiseCollection
                        MDIFrm.ActiveForm.Form_Activate
                        frmCollection.ZOrder 0
                        DoEvents
                        LoadcollectionDefualtValue
                        frmCollection.cmbType.SetFocus
                        SetToolbarControl
                        
                    Case INVOICE:
                    If Not CheckInvoiceSecurity(sFunction) Then Exit Function
                        gRecordMode = sFunction
                        InitialiseInvoice
                        MDIFrm.ActiveForm.Form_Activate
                        frmInvoice.ZOrder 0
                        DoEvents
                        LoadInvoiceDefualtValue
                        frmInvoice.txtName.SetFocus
                        frmInvoice.fraInvoiceItem.Visible = False
                        frmInvoice.fraInvoiceItem.Enabled = False
                        SetToolbarControl
                        
                    Case RECEIPT:
                        If Not CheckReceiptSecurity(sFunction) Then Exit Function
                        gRecordMode = sFunction
                        InitialiseReceipt
                        MDIFrm.ActiveForm.Form_Activate
                        frmReceipt.ZOrder 0
                        DoEvents
                        LoadReceiptDefualtValue
                        frmReceipt.txtInvoiceNo.SetFocus
                        SetToolbarControl
                    
                    Case CASHIN:
                        If Not CheckCashflowSecurity(sFunction) Then Exit Function
                        gRecordMode = sFunction
                        InitialiseCashIn
                        MDIFrm.ActiveForm.Form_Activate
                        frmCashIn.ZOrder 0
                        DoEvents
                        LoadCashInDefualtValue
                        frmCashIn.cmbItem.SetFocus
                        SetToolbarControl
                    Case CASHOUT:
                        If Not CheckCashflowSecurity(sFunction) Then Exit Function
                        gRecordMode = sFunction
                        InitialiseCashout
                        MDIFrm.ActiveForm.Form_Activate
                        frmCashOut.ZOrder 0
                        DoEvents
                        LoadCashOutDefualtValue
                        frmCashOut.cmbItem.SetFocus
                        SetToolbarControl
                    
                    Case CASHFLOWITEM:
                        If Not CheckCashflowItemSecurity(sFunction) Then Exit Function
                        gRecordMode = sFunction
                        InitialiseCashflowItem
                        MDIFrm.ActiveForm.Form_Activate
                        frmCashFlowItem.ZOrder 0
                        DoEvents
                        frmCashFlowItem.txtItemName.SetFocus
                        SetToolbarControl
                End Select
                
                
        Case RECORD_EDIT
                Select Case gRecordType
                            
                    Case MEMBER:
                        If Not CheckMemberSecurity(sFunction) Then Exit Function
                        If MemberSelected Then
                        gRecordMode = sFunction
                        MDIFrm.ActiveForm.Form_Activate
                        SetToolbarControl
                        End If
                    Case Children_information:
                        If Not CheckMemberSecurity(sFunction) Then Exit Function
                        If ChildSelected Then
                        gRecordMode = sFunction
                        MDIFrm.ActiveForm.Form_Activate
                        SetToolbarControl
                        End If
                    Case Payment:
                         If Not CheckPaymentSecurity(sFunction) Then Exit Function
                        If PaymentSelected Then
                        gRecordMode = sFunction
                        MDIFrm.ActiveForm.Form_Activate
                        SetToolbarControl
                        End If
                    
                    Case COLLECTION:
                    If Not CheckCollectionSecurity(sFunction) Then Exit Function
                        If CollectionSelected Then
                        gRecordMode = sFunction
                        MDIFrm.ActiveForm.Form_Activate
                        SetToolbarControl
                        End If
                    Case INVOICE:
                    If Not CheckInvoiceSecurity(sFunction) Then Exit Function
                        If frmInvoice.txtInvoiceNo.Text <> "" Then
                        gRecordMode = sFunction
                        MDIFrm.ActiveForm.Form_Activate
                        SetToolbarControl
                        frmInvoice.fraInvoiceItem.Visible = True
                        frmInvoice.fraInvoiceItem.Enabled = True
                        End If
                    Case RECEIPT:
                        If Not CheckReceiptSecurity(sFunction) Then Exit Function
                        gRecordMode = sFunction
                        MDIFrm.ActiveForm.Form_Activate
                        SetToolbarControl
                        
                    Case CASHIN:
                         If Not CheckCashflowSecurity(sFunction) Then Exit Function
                         If CashInSelected Then
                        gRecordMode = sFunction
                        MDIFrm.ActiveForm.Form_Activate
                        SetToolbarControl
                        End If
                    Case CASHOUT:
                         If Not CheckCashflowSecurity(sFunction) Then Exit Function
                         If CashOutSelected Then
                        gRecordMode = sFunction
                        MDIFrm.ActiveForm.Form_Activate
                        SetToolbarControl
                        End If
                    Case CASHFLOWITEM:
                         If Not CheckCashflowItemSecurity(sFunction) Then Exit Function
                         If CashflowItemSelected Then
                        gRecordMode = sFunction
                        MDIFrm.ActiveForm.Form_Activate
                        SetToolbarControl
                        End If
                End Select
        
        
        Case RECORD_DELETE
                    ' If User.Manager = "N" Then
                    '    MsgBox "Delete function restricted to Manager access level.", vbExclamation
                    '    Exit Function
                    ' End If
                    Select Case gRecordType
                    
                    Case MEMBER:
                    
                        If frmMember.txtGivenName.Text = "" Then
                           MsgBox "No Member has been selected .", vbExclamation
                           'frmMember.txtMno.SetFocus
                           Exit Function
                        End If
                         
                         iAnswer = MsgBox("Are you sure you want to delete this Member?", vbExclamation + vbYesNo)
                        If iAnswer = vbNo Then
                            Exit Function
                        Else
                            gRecordMode = sFunction
                            SaveMember
                            Call InitialiseMember
                        End If
                        
                    Case Payment:
                        If frmPayment.txtGivenName.Text = "" Then
                          MsgBox "No Payment has been selected .", vbExclamation
                          'frmPayment.txtmemberNo.SetFocus
                          Exit Function
                        End If
                        
                        
                        iAnswer = MsgBox("Are you sure you want to delete this Payment?", vbExclamation + vbYesNo)
                        If iAnswer = vbNo Then
                            
                            Exit Function
                        Else
                            gRecordMode = sFunction
                            SavePayment
                            Call InitialisePayment
                        End If
                        'Delete the
                    Case COLLECTION:
                        If frmCollection.txtAmount = "" Then
                            MsgBox "Collection must be selected prior to deleting record.", vbExclamation
                            'frmCollection.txtCollectionNo.SetFocus
                            Exit Function
                        End If
                        
                        iAnswer = MsgBox("Are you sure you want to delete this Collection record?", vbExclamation + vbYesNo)
                        If iAnswer = vbNo Then
                            Exit Function
                        Else
                            gRecordMode = sFunction
                            SaveCollection
                            Call InitialiseCollection
                        End If
                    Case INVOICE:
                        If frmInvoice.txtName = "" Then
                            MsgBox "Invoice must be selected prior to deleting record.", vbExclamation
                            'frmInvoice.txtInvoiceNo.SetFocus
                            Exit Function
                        End If
                        
                        iAnswer = MsgBox("Are you sure you want to delete this Invoice record?", vbExclamation + vbYesNo)
                        If iAnswer = vbNo Then
                            Exit Function
                        Else
                            gRecordMode = sFunction
                            SaveInvoice
                            Call InitialiseInvoice
                        End If
                    Case RECEIPT:
                        If frmReceipt.txtAmountToPay = "" Then
                            MsgBox "Collection must be selected prior to deleting record.", vbExclamation
                            'frmCollection.txtCollectionNo.SetFocus
                            Exit Function
                        End If
                        
                        iAnswer = MsgBox("Are you sure you want to delete this Receipt record?", vbExclamation + vbYesNo)
                        If iAnswer = vbNo Then
                            Exit Function
                        Else
                            gRecordMode = sFunction
                            SaveReceipt
                            Call InitialiseReceipt
                        End If
                    Case CASHIN:
                    
                    Case CASHOUT:
                    Case CASHFLOWITEM:
                End Select
        
        Case RECORD_SAVE
        
                Select Case gRecordType
                            
                    Case MEMBER:
                            If Not ValidateMemeber Then Exit Function
                            objConnection.BeginTrans
                            If SaveMember Then
                            objConnection.CommitTrans
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Record successfully saved to database."
                            Else
                            objConnection.RollbackTrans
                            MsgBox "An error has occurred saving record to database - changes have not been applied.", vbExclamation
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Error saving record to database, changes not applied."
                            End If
                            gRecordMode = RECORD_READ
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                    
                    Case Children_information:
                            If Not ValidateChild Then Exit Function
                            objConnection.BeginTrans
                            If SaveChild Then
                            objConnection.CommitTrans
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Record successfully saved to database."
                            Else
                            objConnection.RollbackTrans
                            MsgBox "An error has occurred saving record to database - changes have not been applied.", vbExclamation
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Error saving record to database, changes not applied."
                            End If
                            gRecordMode = RECORD_READ
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                            
                    Case Payment:
                            If Not ValidatePayment Then Exit Function
                            objConnection.BeginTrans
                            If SavePayment Then
                            objConnection.CommitTrans
                            PaymentSelected = True
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Record successfully saved to database."
                            Else
                            objConnection.RollbackTrans
                            MsgBox "An error has occurred saving record to database - changes have not been applied.", vbExclamation
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Error saving record to database, changes not applied."
                            End If
                            UpdateMemberExparyDate
                            gRecordMode = RECORD_READ
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                         
                    Case COLLECTION:
                            If Not ValidateCollection Then Exit Function
                            objConnection.BeginTrans
                            If SaveCollection Then
                            objConnection.CommitTrans
                            CollectionSelected = True
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Record successfully saved to database."
                            Else
                            objConnection.RollbackTrans
                            MsgBox "An error has occurred saving record to database - changes have not been applied.", vbExclamation
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Error saving record to database, changes not applied."
                            
                            End If
                            gRecordMode = RECORD_READ
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                        
                    Case INVOICE:
                            If mInvoiceItemMode <> RECORD_READ Then
                            MsgBox "Please Save Invoice Item First.", vbExclamation
                            Exit Function
                            End If
                            If Not ValidateInvoice Then Exit Function
                              objConnection.BeginTrans
                              If SaveInvoice Then
                              objConnection.CommitTrans
                              MDIFrm.pnlStatusBar.Panels(1).Text = "Record successfully saved to database."
                            Else
                            objConnection.RollbackTrans
                            MsgBox "An error has occurred saving record to database - changes have not been applied.", vbExclamation
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Error saving record to database, changes not applied."
                            
                            End If
                            gRecordMode = RECORD_READ
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                            frmInvoice.fraInvoiceItem.Visible = True
                            frmInvoice.fraInvoiceItem.Enabled = True
                    Case RECEIPT:
                            If Not ValidateReceipt Then Exit Function
                            objConnection.BeginTrans
                            If SaveReceipt Then
                            objConnection.CommitTrans
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Record successfully saved to database."
                            Else
                            objConnection.RollbackTrans
                            MsgBox "An error has occurred saving record to database - changes have not been applied.", vbExclamation
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Error saving record to database, changes not applied."
                            
                            End If
                            gRecordMode = RECORD_READ
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                    Case CASHIN:
                            If Not ValidateCashIn Then Exit Function
                            objConnection.BeginTrans
                            If SaveCashIn Then
                            objConnection.CommitTrans
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Record successfully saved to database."
                            Else
                            objConnection.RollbackTrans
                            MsgBox "An error has occurred saving record to database - changes have not been applied.", vbExclamation
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Error saving record to database, changes not applied."
                            End If
                            gRecordMode = RECORD_READ
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                    Case CASHOUT:
                            If Not ValidateCashOut Then Exit Function
                            objConnection.BeginTrans
                            If SaveCashOut Then
                            objConnection.CommitTrans
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Record successfully saved to database."
                            Else
                            objConnection.RollbackTrans
                            MsgBox "An error has occurred saving record to database - changes have not been applied.", vbExclamation
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Error saving record to database, changes not applied."
                            End If
                            gRecordMode = RECORD_READ
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                     Case CASHFLOWITEM:
                            If Not ValidateCashflowItem Then Exit Function
                            objConnection.BeginTrans
                            If SaveCashflowItem Then
                            objConnection.CommitTrans
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Record successfully saved to database."
                            Else
                            objConnection.RollbackTrans
                            MsgBox "An error has occurred saving record to database - changes have not been applied.", vbExclamation
                            MDIFrm.pnlStatusBar.Panels(1).Text = "Error saving record to database, changes not applied."
                            End If
                            Call LoadCashflowItemList
                            gRecordMode = RECORD_READ
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                End Select
               
        
        Case RECORD_CANCEL
                
                If gRecordMode = RECORD_READ Then Exit Function
                
                iAnswer = MsgBox("Do you want to cancel the changes made to the record?", vbExclamation + vbYesNo)
                If iAnswer = vbNo Then Exit Function
        
                Select Case gRecordType
                            
                    Case MEMBER:
                                                                
                        If gRecordMode = RECORD_NEW Then
                            gRecordMode = RECORD_CANCEL
                            InitialiseMember
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                            
                        ElseIf gRecordMode = RECORD_EDIT Then
                            gRecordMode = RECORD_CANCEL
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                        End If
                        gRecordMode = RECORD_READ
                     
                     Case Children_information:
                        If gRecordMode = RECORD_NEW Then
                            gRecordMode = RECORD_CANCEL
                            InitialiseChild
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                            
                        ElseIf gRecordMode = RECORD_EDIT Then
                            gRecordMode = RECORD_CANCEL
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                        End If
                        gRecordMode = RECORD_READ
                        
                    Case Payment:

                        If gRecordMode = RECORD_NEW Then
                            gRecordMode = RECORD_CANCEL
                            InitialisePayment
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                        ElseIf gRecordMode = RECORD_EDIT Then
                            gRecordMode = RECORD_CANCEL
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                        End If
                        gRecordMode = RECORD_READ
                    Case COLLECTION:
                            If gRecordMode = RECORD_NEW Then
                            gRecordMode = RECORD_CANCEL
                            InitialiseCollection
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                        ElseIf gRecordMode = RECORD_EDIT Then
                            gRecordMode = RECORD_CANCEL
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                        End If
                            gRecordMode = RECORD_READ
                    Case INVOICE:
                            If gRecordMode = RECORD_NEW Then
                                gRecordMode = RECORD_CANCEL
                                InitialiseInvoice
                                MDIFrm.ActiveForm.Form_Activate
                                SetToolbarControl
                            ElseIf gRecordMode = RECORD_EDIT Then
                                gRecordMode = RECORD_CANCEL
                                MDIFrm.ActiveForm.Form_Activate
                                SetToolbarControl
                            End If
                            gRecordMode = RECORD_READ
                    Case RECEIPT:
                            If gRecordMode = RECORD_NEW Then
                                gRecordMode = RECORD_CANCEL
                                InitialiseReceipt
                                MDIFrm.ActiveForm.Form_Activate
                                SetToolbarControl
                            ElseIf gRecordMode = RECORD_EDIT Then
                                gRecordMode = RECORD_CANCEL
                                MDIFrm.ActiveForm.Form_Activate
                                SetToolbarControl
                            End If
                            gRecordMode = RECORD_READ
                    
                    Case CASHIN:
                            If gRecordMode = RECORD_NEW Then
                            gRecordMode = RECORD_CANCEL
                            InitialiseCashIn
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                            
                        ElseIf gRecordMode = RECORD_EDIT Then
                            gRecordMode = RECORD_CANCEL
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                        End If
                        gRecordMode = RECORD_READ
                    Case CASHOUT:
                            If gRecordMode = RECORD_NEW Then
                            gRecordMode = RECORD_CANCEL
                            InitialiseCashout
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                            
                        ElseIf gRecordMode = RECORD_EDIT Then
                            gRecordMode = RECORD_CANCEL
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                        End If
                        gRecordMode = RECORD_READ
                   
                   Case CASHFLOWITEM:
                            If gRecordMode = RECORD_NEW Then
                            gRecordMode = RECORD_CANCEL
                            InitialiseCashflowItem
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                            
                        ElseIf gRecordMode = RECORD_EDIT Then
                            gRecordMode = RECORD_CANCEL
                            MDIFrm.ActiveForm.Form_Activate
                            SetToolbarControl
                        End If
                        gRecordMode = RECORD_READ
                End Select

                gRecordMode = RECORD_READ
       Case RECORD_PRINT
                  Select Case gRecordType
                            
                    Case MEMBER:
                     
                    Case Children_information:
                        
                    Case Payment:
                                If frmPayment.txtReceiptNo.Text = "" Then
                                  MsgBox "Please Save the payment before print or Priview", vbInformation, "CMS - Error Continuing"
                                Else
                                  If frmPayment.cmbPaymentType.Text = "Membership" Then
                                    Call GenerateReport(9, "Print", objConnection)
                                  Else
                                    Call GenerateReport(10, "Print", objConnection)
                                  End If
                                 End If
                    Case COLLECTION:
                                If frmCollection.txtCollectionNo.Text = "" Then
                                  MsgBox "Please Save the collection before print or Priview", vbInformation, "CMS - Error Continuing"
                                Else
                                  Call GenerateReport(11, "Print", objConnection)
                                 End If
                          
                    Case INVOICE:
                               If frmInvoice.txtInvoiceNo.Text = "" Then
                                  MsgBox "Please Save the Invoice before print or Priview", vbInformation, "CMS - Error Continuing"
                                Else
                                  Call GenerateReport(7, "Print", objConnection)
                                  
                               End If
                    Case RECEIPT:
                               If frmReceipt.txtReceiptId.Text = "" Then
                               MsgBox "Please Save the Receipt before print or Priview", vbInformation, "CMS - Error Continuing"
                                Else
                                  Call GenerateReport(6, "Print", objConnection)
                                  
                               End If
                    Case CASHFLOW_VIEW:
                                
                                  Call GenerateReport(5, "Print", objConnection)
                End Select
       Case RECORD_PRIVIEW
                 
                 Select Case gRecordType
                            
                   Case MEMBER:
                     
                   Case Children_information:
                       
                   Case Payment:
                        If frmPayment.txtReceiptNo.Text = "" Then
                                  MsgBox "Please Save the payment before print or Priview", vbInformation, "CMS - Error Continuing"
                        Else
                           If frmPayment.cmbPaymentType.Text = "Membership" Then
                           Call GenerateReport(9, "View", objConnection)
                           Else
                            Call GenerateReport(10, "View", objConnection)
                           End If
                        End If
                   Case COLLECTION:
                        If frmCollection.txtCollectionNo.Text = "" Then
                            MsgBox "Please Save the collection before print or Priview", vbInformation, "CMS - Error Continuing"
                        Else
                            Call GenerateReport(11, "View", objConnection)
                        End If
                          
                   Case INVOICE:
                           If frmInvoice.txtInvoiceNo.Text = "" Then
                                  MsgBox "Please Save the Invoice before print or Priview", vbInformation, "CMS - Error Continuing"
                                Else
                                  Call GenerateReport(7, "View", objConnection)
                                  
                               End If
                   Case RECEIPT:
                           If frmReceipt.txtReceiptId.Text = "" Then
                                  MsgBox "Please Save the Receipt before print or Priview", vbInformation, "CMS - Error Continuing"
                                Else
                                  Call GenerateReport(6, "View", objConnection)
                                  
                               End If
                   Case CASHFLOW_VIEW:
                                
                                  Call GenerateReport(5, "View", objConnection)
                End Select
    End Select

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMain", "ProcessRecord", True)

End Function

Public Function GetPrivileges(ActionName As String) As Byte

Dim objOrganisation As CMSOrganisation.clsOrganisation
Dim ActionId As Long
Dim strFind As String
On Error GoTo ErrorHandler


Set objOrganisation = New CMSOrganisation.clsOrganisation
Set objOrganisation.DatabaseConnection = objConnection
ActionId = objOrganisation.getActionId(ActionName)

If ActionId <> 0 Then
    GetPrivileges = ActionId
Else
  GetPrivileges = 0
 
End If
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMain", "GetPrivilege", True)

End Function

Public Function ValidDateEntry(ctrl As MaskEdBox)
On Error GoTo ErrorHandler
'==============================================================================
'Synopsis:  Checks for Valid Date entered into MaskEdBox.
'               Manual checking preferred over VB Date checking as VB calls Windows
'               settings - this can give unexpected results.
'               Force date validation for Australian Date - independent of
'               Windows settings.
'Range:      Date check range 1900-2050 (sufficient life expectancy of Application)
'Note:       MaskEdit control must set properties as "PromptInclude = False"
'==============================================================================

        Dim bDateOk As Boolean
        Dim iDay, iMonth, iYear As Integer
        
        
        If Len(ctrl) = 0 Then Exit Function
        
        bDateOk = True
        If Len(ctrl) < 8 Then
            bDateOk = False
            GoTo exitProcedure
        End If
        
        iDay = CInt(Mid(ctrl, 1, 2))
        iMonth = CInt(Mid(ctrl, 3, 2))
        iYear = CInt(Mid(ctrl, 5, 4))
        
        If (iDay > 31) Or (iMonth > 12) Or (iYear < 1900 Or iYear > 2050) Then
            bDateOk = False
            GoTo exitProcedure
        End If
        
        Select Case iDay
                    
                    Case Is = 31:
                                       Select Case iMonth
                                                    Case 1, 3, 5, 7, 8, 10, 12
                                                    Case Else: bDateOk = False: GoTo exitProcedure
                                        End Select
                    
                    Case Is = 30:
                                        Select Case iMonth
                                                    Case 2: bDateOk = False: GoTo exitProcedure
                                        End Select
                    
                    Case Is = 29:
                                        If iMonth = 2 Then
                                            Select Case iYear
                                                        Case 1900, 1904, 1908, 1912, 1916, 1920, 1924, _
                                                               1928, 1932, 1936, 1940, 1944, 1948, _
                                                               1952, 1956, 1960, 1964, 1968, 1972, _
                                                               1976, 1980, 1984, 1988, 1992, 1996, _
                                                               2000, 2004, 2008, 2012, 2016, 2020, _
                                                               2024, 2028, 2032, 2036, 2040, 2044, 2050
                                                        Case Else: bDateOk = False: GoTo exitProcedure
                                            End Select
                                        End If
        
        End Select


exitProcedure:
        If bDateOk = False Then
            MsgBox "Date entered is invalid or out of range.  Date must be entered in format dd/mm/yyyy.", vbExclamation
            ctrl.Text = ""
            ctrl.SetFocus
        End If
        Exit Function

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMain", "ValidDateEntry", True)

End Function

Public Sub HighlightText(frm As Form)
On Error Resume Next
    'incase form is not loaded
    frm.ActiveControl.SelStart = 0
    If TypeOf frm.ActiveControl Is MaskEdBox Then
        frm.ActiveControl.SelLength = Len(frm.ActiveControl.FormattedText)
    Else
        frm.ActiveControl.SelLength = Len(frm.ActiveControl.Text)
    End If

End Sub


Public Sub SortListView(lvwCtrl As ListView, ColumnHeader As ColumnHeader)
On Error GoTo ErrorHandler

    If (ColumnHeader.index - 1 <> lvwCtrl.SortKey) Then
        lvwCtrl.SortKey = ColumnHeader.index - 1
    Else
        If (lvwCtrl.SortOrder = lvwAscending) Then
            lvwCtrl.SortOrder = lvwDescending
        Else
            lvwCtrl.SortOrder = lvwAscending
        End If
    End If
    lvwCtrl.Sorted = True
  
Exit Sub
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMain", "SortListView", True)

End Sub

Private Sub SortList(ByVal LV As MSComctlLib.ListView, SortOrder As eSortOrder, SortKey As Integer)
    If SortOrder = sortAscending Then
        LV.SortOrder = lvwAscending
    ElseIf SortOrder = sortDescending Then
        LV.SortOrder = lvwDescending
    End If
    LV.SortKey = SortKey
    LV.Sorted = True
End Sub

Public Function SortColumn(ByVal LV As ListView, ColumnIndex As Integer, SortOrder As eSortOrder, SortType As eSortType) As Boolean
    
On Error GoTo EH:
    
    Dim X       As Integer
    Dim Y       As Integer
    Dim strMax  As String
    Dim strNew  As String
    
    strMax = "0"
    Select Case SortType
        Case eSortType.sortAlpha
            'TEXT SORT
            Call SortList(LV, SortOrder, ColumnIndex - 1)
        Case eSortType.sortNumeric
            'NUMERIC SORT
            
            'GET THE LONGEST NUMBER STRING LENGTH
            If ColumnIndex > 1 Then
                For X = 1 To LV.ListItems.Count
                    If Len(LV.ListItems(X).ListSubItems(ColumnIndex - 1)) <> 0 Then 'DONT BOTHER WITH 0 LENGTH STRINGS
                        If Len(CStr(Int(LV.ListItems(X).ListSubItems(ColumnIndex - 1)))) > Len(strMax) Then
                       
                        '  If Int(LV.ListItems(X).ListSubItems(ColumnIndex - 1)) > Int(strMax) Then
                            strMax = CStr(Int(LV.ListItems(X).SubItems(ColumnIndex - 1)))
                        End If
                    End If
                Next
            Else
                For X = 1 To LV.ListItems.Count
                    If Len(LV.ListItems(X)) <> 0 Then
                        If Len(CStr(Int(LV.ListItems(X)))) > Len(strMax) Then
                            strMax = CStr(Int(LV.ListItems(X)))
                        End If
                    End If
                Next
            End If
            
            'MAKE TO CONTROL INVISIBLE TO INCREASE PERFORMANCE
            LV.Visible = False
            
            If ColumnIndex > 1 Then
                For X = 1 To LV.ListItems.Count
                    If Len(LV.ListItems(X).ListSubItems(ColumnIndex - 1)) = 0 Then
                        LV.ListItems(X).ListSubItems(ColumnIndex - 1) = "0" 'IF 0 LENGTH STRING, MAKE IT 0 SO THE SORT WONT FAIL
                    ElseIf Len(CStr(Int(LV.ListItems(X).ListSubItems(ColumnIndex - 1)))) < Len(strMax) Then
                        'PAD NUMBER WITH 0s THIS ALLOWS A ROUND ABOUT NUMERIC SORT
                        'THEN WE CAN REMOVE THE 0s LATER.
                        strNew = LV.ListItems(X).ListSubItems(ColumnIndex - 1)
                        For Y = 1 To Len(strMax) - Len(CStr(Int(LV.ListItems(X).ListSubItems(ColumnIndex - 1))))
                            strNew = "0" & strNew
                        Next
                        LV.ListItems(X).ListSubItems(ColumnIndex - 1) = strNew
                    End If
                Next
            Else
                For X = 1 To LV.ListItems.Count
                    If Len(LV.ListItems(X).Text) = 0 Then
                        LV.ListItems(X).Text = "0" 'make 0 length strings = To "0"
                    ElseIf Len(CStr(Int(LV.ListItems(X)))) < Len(strMax) Then
                        'PAD NUMBER WITH 0s THIS ALLOWS A ROUND ABOUT NUMERIC SORT
                        'THEN WE CAN REMOVE THE 0s LATER.
                        strNew = LV.ListItems(X).Text
                        For Y = 1 To Len(strMax) - Len(CStr(Int(LV.ListItems(X))))
                            strNew = "0" & strNew
                        Next
                        LV.ListItems(X).Text = strNew
                    End If
                Next
            End If
            
            'SORT THE LIST
            Call SortList(LV, SortOrder, ColumnIndex - 1)
            
            'GET RID OF PADDED 0s
            If ColumnIndex > 1 Then
                For X = 1 To LV.ListItems.Count
                    LV.ListItems(X).ListSubItems(ColumnIndex - 1) = CDbl(LV.ListItems(X).ListSubItems(ColumnIndex - 1))
                   If LV.ListItems(X).ListSubItems(ColumnIndex - 1) = 0 Then LV.ListItems(X).ListSubItems(ColumnIndex - 1) = ""
                Next
            Else
                For X = 1 To LV.ListItems.Count
                    LV.ListItems(X).Text = CDbl(LV.ListItems(X).Text)
                    If LV.ListItems(X).Text = 0 Then LV.ListItems(X).Text = ""
                Next
            End If
            LV.Visible = True
        Case eSortType.sortDate
            'DATE SORT
                'MAKE TO CONTROL INVISIBLE TO INCREASE PERFORMANCE
                LV.Visible = False
                If ColumnIndex > 1 Then
                    'CHANGE DATE TO A FORMAT THAT CAN BE SORTED LIKE TEXT
                    For X = 1 To LV.ListItems.Count
                        LV.ListItems(X).ListSubItems(ColumnIndex - 1) = Format(LV.ListItems(X).ListSubItems(ColumnIndex - 1), "YYYY MM DD hh:mm:ss")
                    Next

                    Call SortList(LV, SortOrder, ColumnIndex - 1)
                    'CONVERT BACK TO FORMAT YOU WANT TO DISPLAY
                        For X = 1 To LV.ListItems.Count
                            LV.ListItems(X).ListSubItems(ColumnIndex - 1) = Format(LV.ListItems(X).ListSubItems(ColumnIndex - 1), DATE_FORMAT)
                        Next
                Else
                    'CHANGE DATE TO A FORMAT THAT CAN BE SORTED LIKE TEXT
                    For X = 1 To LV.ListItems.Count
                        LV.ListItems(X).Text = Format(LV.ListItems(X).Text, "YYYY MM DD hh:mm:ss")
                    Next
                        Call SortList(LV, SortOrder, ColumnIndex - 1)
                        For X = 1 To LV.ListItems.Count
                            LV.ListItems(X).Text = Format(LV.ListItems(X).Text, DATE_FORMAT)
                        Next
                End If
                LV.Visible = True
        End Select
        'RETURN TRUE
        SortColumn = True
        Exit Function
EH:
    'IF ERROR OCCURED RETURN FALSE
    SortColumn = False
    MsgBox "ERROR IN SORT ROUTINE:" & vbCrLf & Err.Number & vbTab & Err.Description
End Function



Public Function ConvertNull(rvarValue As Variant)
    
    If IsNull(rvarValue) Then
        ConvertNull = ""
    Else
        ConvertNull = rvarValue
    End If
    
End Function

Public Function ValidNumericEntry(ctrl As TextBox)
On Error GoTo ErrorHandler
    
    If ctrl.Text <> "" Then
        If Not IsNumeric(ctrl.Text) Then
            MsgBox "Invalid character entered.  Numeric values only valid.", vbExclamation
            ctrl.SetFocus
            Exit Function
        End If
        ctrl.Text = Format(ctrl.Text, NUMERIC_FORMAT)
    End If
            
    ValidNumericEntry = True


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modMain", "ValidNumericEntry", True)

End Function

Public Sub clearAllSelection()
MemberSelected = False
'ChildSelected = False
PaymentSelected = False
CashInSelected = False
CashOutSelected = False
CollectionSelected = False
InvoiceSelected = False
CashflowItemSelected = False
End Sub

Public Function get_MoreThanAndLessThan(ActionName As String) As String
Dim s, Start As String
Dim AsciiNumber As Long

s = Mid(ActionName, Len(ActionName), 1)
Start = Mid(ActionName, 1, Len(ActionName) - 1)
AsciiNumber = Asc(s)
If (AsciiNumber = 90) Or (AsciiNumber = 122) Then
  get_MoreThanAndLessThan = ActionName & "z"
Else
  AsciiNumber = AsciiNumber + 1
  s = Chr(AsciiNumber)
  get_MoreThanAndLessThan = Start & s
End If
End Function


Public Function FindCBIndexByName(ByRef cbComboBox As ComboBox, ByRef strSearchValue As String) As Integer
    Dim n As Integer
    For n = 0 To cbComboBox.ListCount - 1
        If cbComboBox.List(n) = strSearchValue Then
          ' // Return the found index
            FindCBIndexByName = n
          ' // and exit
            Exit Function
        End If
    Next
  ' // Set not found value
    FindCBIndexByName = -1
End Function


Public Function FindCBIndexById(ByRef cbComboBox As ComboBox, ByRef id As Long) As Integer
    Dim n As Integer
    For n = 0 To cbComboBox.ListCount - 1
        If cbComboBox.ItemData(n) = id Then
          ' // Return the found index
            FindCBIndexById = n
          ' // and exit
            Exit Function
        End If
    Next
  ' // Set not found value
    FindCBIndexById = -1
End Function

