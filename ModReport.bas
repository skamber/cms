Attribute VB_Name = "ModReport"
 Option Explicit

 Private Const REPORTS = 4
 
 Public Function LoadComboBox_PaymentType(cboComboBox As ComboBox)
On Error GoTo ErrorHandler

    Dim PaymentType As CMSOrganisation.clsOrganisation
    Dim rslocal As ADODB.Recordset
    Dim rslocal1 As ADODB.Recordset

    'get Funnel Positions associated to Strategy records
    Set PaymentType = New CMSOrganisation.clsOrganisation
    Set PaymentType.DatabaseConnection = objConnection
    Set rslocal = PaymentType.getPaymentType
    Set rslocal1 = PaymentType.getCollectionType
    
                            
    cboComboBox.Clear
    cboComboBox.AddItem "All"
    If Not rslocal Is Nothing Then
    
        Do Until rslocal.EOF
            cboComboBox.AddItem rslocal!Payment
            rslocal.MoveNext
        Loop
        Set rslocal = Nothing
    
    End If
    
    If Not rslocal1 Is Nothing Then
    
        Do Until rslocal1.EOF
            cboComboBox.AddItem rslocal1!Collection_Type
            rslocal1.MoveNext
        Loop
        Set rslocal1 = Nothing
    
    End If
    
    


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modReport", "LoadComboBox_PostCode", True)

End Function
Public Function LoadComboBox_Type(cboComboBox As ComboBox)
On Error GoTo ErrorHandler

    Dim PaymentType As CMSOrganisation.clsOrganisation
    Dim rslocal As ADODB.Recordset
    Dim rslocal1 As ADODB.Recordset

    'get Funnel Positions associated to Strategy records
    Set PaymentType = New CMSOrganisation.clsOrganisation
    Set PaymentType.DatabaseConnection = objConnection
    Set rslocal = PaymentType.getPayment
    
    
                            
    cboComboBox.Clear
    cboComboBox.AddItem "All"
    If Not rslocal Is Nothing Then
    
        Do Until rslocal.EOF
            cboComboBox.AddItem rslocal!Payment_type
            rslocal.MoveNext
        Loop
        Set rslocal = Nothing
    
    End If
     

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modReport", "LoadComboBox_PostCode", True)

End Function

Public Function LoadComboBox_PostCode(cboComboBox As ComboBox)
On Error GoTo ErrorHandler

    Dim objPostCode As CMSMember.clsMember_s
    Dim rslocal As ADODB.Recordset


    'get Funnel Positions associated to Strategy records
    Set objPostCode = New CMSMember.clsMember_s
    Set objPostCode.DatabaseConnection = objConnection
    Set rslocal = objPostCode.GetPostCodes
                            
    cboComboBox.Clear
    cboComboBox.AddItem "All"
    If Not rslocal Is Nothing Then
    
        Do Until rslocal.EOF
          If IsNull(rslocal!postCode) Then
            cboComboBox.AddItem ("")
          Else
            cboComboBox.AddItem rslocal!postCode
          End If
            rslocal.MoveNext
        Loop
        Set rslocal = Nothing
    
    End If

    


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modReport", "LoadComboBox_PostCode", True)

End Function

Public Function LoadComboBox_UserName(cboComboBox As ComboBox)
On Error GoTo ErrorHandler

    Dim objUserName As CMSUser.clsUser_s
    Dim rslocal As ADODB.Recordset


    'get Funnel Positions associated to Strategy records
    Set objUserName = New CMSUser.clsUser_s
    Set objUserName.DatabaseConnection = objConnection
    Set rslocal = objUserName.getAllUserName
    
                            
    cboComboBox.Clear
    cboComboBox.AddItem "All"
    If Not rslocal Is Nothing Then
    
        Do Until rslocal.EOF
            cboComboBox.AddItem rslocal!Full_Name
            rslocal.MoveNext
        Loop
        Set rslocal = Nothing
    
    End If

    


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modReport", "LoadComboBox_UserName", True)

End Function

Public Sub GenerateReport(lngReportId As Long, strDestination As String, objReportConnection As ADODB.Connection)
On Error GoTo ErrorHandler

    Dim strReportId As String
    Dim strReportSQL As String
    Dim strReportSQLSubReport As String
    Dim strStartDate As String
    Dim strEndDate As String
    Dim intCnt As Integer
    Dim intNoOfTables As Integer
    Dim NewDate As String
    On Error GoTo ErrorHandler

    Screen.MousePointer = vbHourglass
    Select Case lngReportId
    Case 5: 'CashflowView
    
       With frmCashFlowView.Report
        
        frmCashFlowView.Report.WindowWidth = Screen.Width
        frmCashFlowView.Report.WindowHeight = Screen.Height
        For intCnt = 0 To 30
            .Formulas(intCnt) = ""
            .DataFiles(intCnt) = ""
        Next
       .SelectionFormula = ""
        .GroupSelectionFormula = ""
        
        
        .DiscardSavedData = True
      End With
    
    Case 6: 'Receipt
    With frmReceipt.Report
        frmReceipt.Report.WindowWidth = Screen.Width
        frmReceipt.Report.WindowHeight = Screen.Height
        For intCnt = 0 To 30
            .Formulas(intCnt) = ""
            .DataFiles(intCnt) = ""
        Next
       .SelectionFormula = ""
        .GroupSelectionFormula = ""
        .DiscardSavedData = True
    End With
    Case 7: 'Invoice
    With frmInvoice.Report
        frmInvoice.Report.WindowWidth = Screen.Width
        frmInvoice.Report.WindowHeight = Screen.Height
        For intCnt = 0 To 30
            .Formulas(intCnt) = ""
            .DataFiles(intCnt) = ""
        Next
       .SelectionFormula = ""
        .GroupSelectionFormula = ""
        .DiscardSavedData = True
    End With
    Case 9, 10: 'MemberShip Invoice
    With frmPayment.Report
        frmPayment.Report.WindowWidth = Screen.Width
        frmPayment.Report.WindowHeight = Screen.Height
        For intCnt = 0 To 30
            .Formulas(intCnt) = ""
            .DataFiles(intCnt) = ""
        Next
       .SelectionFormula = ""
        .GroupSelectionFormula = ""
        .DiscardSavedData = True
    End With
    Case 0, 1, 2, 3, 4: 'Reporting
    With frmReport.Report
        frmReport.Report.WindowWidth = Screen.Width
        frmReport.Report.WindowHeight = Screen.Height
        For intCnt = 0 To 30
            .Formulas(intCnt) = ""
            .DataFiles(intCnt) = ""
        Next
       .SelectionFormula = ""
        .GroupSelectionFormula = ""
        .SortFields(0) = ""
        .DiscardSavedData = True
    End With
    
    
    End Select
    Select Case lngReportId
        Case 5: 'sashflowview
                With frmCashFlowView
                strReportId = "CashIn and Cashout Repoert.rpt"
                strStartDate = "Date(" & Format(.dteFromDate.FormattedText, "yyyy,mm,dd") & ")"
                strEndDate = "Date(" & Format(.dteToDate.FormattedText, "yyyy,mm,dd") & ")"
                If .dteFromDate.Text = "" Then strStartDate = "Date(1970,01,01)"
                If .dteToDate.Text = "" Then strEndDate = "Date(2070,01,01)"
                'used to display "for the period..." on report
                .Report.Formulas(0) = "STARTDATE= " & strStartDate
                .Report.Formulas(1) = "ENDDATE= " & strEndDate
                If strReportSQL <> "" Then
                strReportSQL = strReportSQL & "AND" & "{cashin.dateofcashin} >= " & strStartDate _
                & " AND {cashin.dateofcashin} <= " & strEndDate
                Else
                strReportSQL = "{cashin.dateofcashin} >= " & strStartDate _
                & " AND {cashin.dateofcashin} <= " & strEndDate
                End If
                
                If strReportSQLSubReport <> "" Then
                strReportSQLSubReport = strReportSQLSubReport & "AND" & "{cashout.dateofcashout} >= " & strStartDate _
                & " AND {cashout.dateofcashout} <= " & strEndDate
                Else
                strReportSQLSubReport = "{cashout.dateofcashout} >= " & strStartDate _
                & " AND {cashout.dateofcashout} <= " & strEndDate
                End If
                
                End With
                intNoOfTables = 1
            
            
        Case 6: 'Receipt
                
                strReportId = "InvoiceReceipt.rpt"
                strReportSQL = "{Receipt.ID} = " & frmReceipt.txtReceiptId.Text
            
            intNoOfTables = 2
            
        Case 7: 'Invoice
                
                strReportId = "Invoice.rpt"
                strReportSQL = "{invoice.Invoice_no} = '" & frmInvoice.txtInvoiceNo.Text & "'"
            
            intNoOfTables = 1
        Case 9: 'MemberShip Invoice
                
                strReportId = "MemberShip.rpt"
                strReportSQL = "{Payment.Receipt_No} = " & frmPayment.txtReceiptNo.Text
            
            intNoOfTables = 3
        Case 10: 'Payment
            
                strReportId = "PaymentReceipt.rpt"
                strReportSQL = "{Payment.Receipt_No} = " & frmPayment.txtReceiptNo.Text
        intNoOfTables = 3
        
        Case 0:
                strReportId = "MemberReport.rpt"
                With frmReport
                If .cboReportCriteria1.Text <> "All" Then
                   strReportSQL = "{Member.PostCode} = '" & .cboReportCriteria1.Text & "'"
                End If
                
                If .cboReportCriteria2.Text <> "All" Then
                   If strReportSQL <> "" Then
                     strReportSQL = strReportSQL & "AND" & "{Member.status} = '" & .cboReportCriteria2.Text & "'"
                   Else
                    strReportSQL = "{Member.status} = '" & .cboReportCriteria2.Text & "'"
                   End If
                End If
                
                Select Case .cboReportCriteria3.Text
                
                Case "SURNAME": frmReport.Report.SortFields(0) = "+{Member.Surname}"
                                
                Case "GIVEN_NAME": frmReport.Report.SortFields(0) = "+{Member.GIVEN_NAME}"
                                   
                Case "POSTCODE": frmReport.Report.SortFields(0) = "+{Member.POSTCODE}"
                Case "MEMBERSHIP_EXPIARY": frmReport.Report.SortFields(0) = "+{Member.MEMBERSHIP_EXPIARY}"
                Case "STATUS": frmReport.Report.SortFields(0) = "+{Member.STATUS}"
                Case "MNO": frmReport.Report.SortFields(0) = "+{Member.MNO}"
               End Select
                
                
                strStartDate = "Date(" & Format(.dteStartDate.FormattedText, "yyyy,mm,dd") & ")"
                strEndDate = "Date(" & Format(.dteEndDate.FormattedText, "yyyy,mm,dd") & ")"
                If .dteStartDate.Text = "" Then strStartDate = "Date(1970,01,01)"
                If .dteEndDate.Text = "" Then strEndDate = "Date(2070,01,01)"
                If .dteStartDate.Text <> "" And .dteEndDate.Text <> "" Then
                'used to display "for the period..." on report
                .Report.Formulas(0) = "STARTDATE= " & strStartDate
                .Report.Formulas(1) = "ENDDATE= " & strEndDate
                  If strReportSQL <> "" Then
                     strReportSQL = strReportSQL & "AND" & "{Member.Membership_Expiary} >= " & strStartDate _
                     & " AND {Member.Membership_Expiary} <= " & strEndDate
                  Else
                    strReportSQL = "{Member.Membership_Expiary} >= " & strStartDate _
                    & " AND {Member.Membership_Expiary} <= " & strEndDate
                  End If
                End If
                End With
                intNoOfTables = 1
        Case 1:
               strReportId = "PaymentReport.rpt"
                With frmReport
                If .cboReportCriteria1.Text <> "All" Then
                   strReportSQL = "{AllIncome.User_Name} = '" & .cboReportCriteria1.Text & "'"
                End If

                If .cboReportCriteria2.Text <> "All" Then
                   If strReportSQL <> "" Then
                     strReportSQL = strReportSQL & " AND " & "{AllIncome.Payment} = '" & .cboReportCriteria2.Text & "'"
                   Else
                    strReportSQL = "{AllIncome.Payment} = '" & .cboReportCriteria2.Text & "'"
                   End If
                End If
                
                If .cboReportCriteria3.Text <> "All" Then
                   If strReportSQL <> "" Then
                     strReportSQL = strReportSQL & " AND " & "{AllIncome.type} = '" & .cboReportCriteria3.Text & "'"
                   Else
                    strReportSQL = "{AllIncome.type} = '" & .cboReportCriteria3.Text & "'"
                   End If
                End If
                
                If .txtMemberRno.Text <> "" Then
                    If strReportSQL <> "" Then
                     strReportSQL = strReportSQL & " AND " & "{AllIncome.Mno} = " & .txtMemberRno.Text
                   Else
                    strReportSQL = "{AllIncome.Mno} = " & .txtMemberRno.Text
                   End If
                 End If
                

                strStartDate = "Date(" & Format(.dteStartDate.FormattedText, "yyyy,mm,dd") & ")"
                strEndDate = "Date(" & Format(.dteEndDate.FormattedText, "yyyy,mm,dd") & ")"
                If .dteStartDate.Text = "" Then strStartDate = "Date(1970,01,01)"
                If .dteEndDate.Text = "" Then strEndDate = "Date(2070,01,01)"
                'used to display "for the period..." on report
                .Report.Formulas(0) = "STARTDATE= " & strStartDate
                .Report.Formulas(1) = "ENDDATE= " & strEndDate
                If strReportSQL <> "" Then
                strReportSQL = strReportSQL & "AND" & "{AllIncome.Date_Of_Payment} >= " & strStartDate _
                & " AND {AllIncome.Date_Of_Payment} <= " & strEndDate
                Else
                strReportSQL = "{AllIncome.Date_Of_Payment} >= " & strStartDate _
                & " AND {AllIncome.Date_Of_Payment} <= " & strEndDate
                End If

                
                End With
                intNoOfTables = 1
        Case 2:
                strReportId = "Children Over 18.rpt"
                With frmReport
                NewDate = DateAdd("yyyy", -18, .dteStartDate.FormattedText)

                strStartDate = "Date(" & Format(NewDate, "yyyy,mm,dd") & ")"
                'used to display "for the period..." on report
                strReportSQL = "{Children.Birth_Date} <= " & strStartDate
                If .cboReportCriteria2.Text <> "All" Then
                    If .cboReportCriteria2.Text = "Member" Then
                       strReportSQL = strReportSQL & " AND {Children.MEMBER} ='Y'"
                    Else
                       strReportSQL = strReportSQL & " AND {Children.MEMBER} ='N'"
                    End If
                End If
                End With
                intNoOfTables = 1
        Case 3:
               strReportId = "InvoiceReport.rpt"
                strStartDate = "Date(" & Format(Date, "yyyy,mm,dd") & ")"
                'used to display "for the period..." on report
                strReportSQL = "{invoice.over_due_date} <= " & strStartDate _
                & " AND {Invoice.Balance} <> " & 0
                intNoOfTables = 1
        Case 4:
               strReportId = "ReceiptReport.rpt"
               With frmReport
               strStartDate = "Date(" & Format(.dteStartDate.FormattedText, "yyyy,mm,dd") & ")"
                strEndDate = "Date(" & Format(.dteEndDate.FormattedText, "yyyy,mm,dd") & ")"
                If .dteStartDate.Text = "" Then strStartDate = "Date(1970,01,01)"
                If .dteEndDate.Text = "" Then strEndDate = "Date(2070,01,01)"
                'used to display "for the period..." on report
                .Report.Formulas(0) = "STARTDATE= " & strStartDate
                .Report.Formulas(1) = "ENDDATE= " & strEndDate
                If strReportSQL <> "" Then
                strReportSQL = strReportSQL & " AND " & "{Receipt.date_of_Receipt} >= " & strStartDate _
                & " AND {Receipt.date_of_Receipt} <= " & strEndDate
                Else
                strReportSQL = "{Receipt.date_of_Receipt} >= " & strStartDate _
                & " AND {Receipt.date_of_Receipt} <= " & strEndDate
                End If
                End With
                intNoOfTables = 2
                
                
                
    End Select
    Select Case lngReportId
    Case 5: 'CashflowView
         With frmCashFlowView
        .Report.ReportFileName = App.Path & "\Reports\" & strReportId
        .Report.SelectionFormula = strReportSQL
        
        If strDestination = "View" Then
            .Report.Destination = crptToWindow
        ElseIf strDestination = "Print" Then
            .Report.Destination = crptToPrinter
        End If
        
            For intCnt = 0 To (intNoOfTables - 1)
                .Report.DataFiles(intCnt) = gLocalDBPath
            Next
        'Temporary setup for reporting
        'Note: DSN must be setup and report attached to exact version.
        'Further investigation warranted into OLE DB interface ....
        
        .Report.Connect = objReportConnection
        .Report.SubreportToChange = .Report.GetNthSubreportName(0)
        .Report.SelectionFormula = strReportSQLSubReport
        .Report.Action = 1
       .Report.SubreportToChange = ""



    End With
    Case 6: 'Receipt
    
    With frmReceipt
        .Report.ReportFileName = App.Path & "\Reports\" & strReportId
        .Report.SelectionFormula = strReportSQL
        
        If strDestination = "View" Then
            .Report.Destination = crptToWindow
        ElseIf strDestination = "Print" Then
            .Report.Destination = crptToPrinter
        End If
        
            For intCnt = 0 To (intNoOfTables - 1)
                .Report.DataFiles(intCnt) = gLocalDBPath
            Next
        'Temporary setup for reporting
        'Note: DSN must be setup and report attached to exact version.
        'Further investigation warranted into OLE DB interface ....
        
        .Report.Connect = objReportConnection
        .Report.Action = 1
    End With
    
    Case 7: 'Invoice
    
    With frmInvoice
        .Report.ReportFileName = App.Path & "\Reports\" & strReportId
        .Report.SelectionFormula = strReportSQL
        
        If strDestination = "View" Then
            .Report.Destination = crptToWindow
        ElseIf strDestination = "Print" Then
            .Report.Destination = crptToPrinter
        End If
        
            For intCnt = 0 To (intNoOfTables - 1)
                .Report.DataFiles(intCnt) = gLocalDBPath
            Next
        'Temporary setup for reporting
        'Note: DSN must be setup and report attached to exact version.
        'Further investigation warranted into OLE DB interface ....
        
        .Report.Connect = objReportConnection
        .Report.Action = 1
        
    End With
    Case 9, 10: 'MemberShip Invoice
    Dim result As String
    With frmPayment
        .Report.ReportFileName = App.Path & "\Reports\" & strReportId
        .Report.SelectionFormula = strReportSQL
        
        If strDestination = "View" Then
            .Report.Destination = crptToWindow
        ElseIf strDestination = "Print" Then
            .Report.Destination = crptToPrinter
        End If
        
            For intCnt = 0 To (intNoOfTables - 1)
                .Report.DataFiles(intCnt) = gLocalDBPath
            Next
        'Temporary setup for reporting
        'Note: DSN must be setup and report attached to exact version.
        'Further investigation warranted into OLE DB interface ....
        
        .Report.Connect = objReportConnection
        .Report.Action = 1
                
        
    End With
    
    Case 0, 1, 2, 3, 4:
      With frmReport
        .Report.ReportFileName = App.Path & "\Reports\" & strReportId
        .Report.SelectionFormula = strReportSQL
        
        If strDestination = "View" Then
            .Report.Destination = crptToWindow
        ElseIf strDestination = "Print" Then
            .Report.Destination = crptToPrinter
        End If
        
            For intCnt = 0 To (intNoOfTables - 1)
                .Report.DataFiles(intCnt) = gLocalDBPath
            Next
        'Temporary setup for reporting
        'Note: DSN must be setup and report attached to exact version.
        'Further investigation warranted into OLE DB interface ....
        
        .Report.Connect = objReportConnection
        .Report.Action = 1
    End With
    End Select
    DoEvents
    
    Screen.MousePointer = vbDefault

    Exit Sub
    
ErrorHandler:
    
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "frmReport", "GenerateReport", True)

End Sub

Public Function InitialiseReport()
On Error GoTo ErrorHandler

    Dim ctrl As Control
    
    For Each ctrl In frmReport.Controls
        
        If TypeOf ctrl Is TextBox Then ctrl.Text = ""
        If TypeOf ctrl Is ComboBox Then ctrl.ListIndex = -1
        If TypeOf ctrl Is MaskEdBox Then ctrl.Text = ""
        Set ctrl = Nothing
        
    Next ctrl
    DoEvents
    

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modReport", "InitialiseReport", True)

End Function
Public Sub ReportCriteria_Reset()
On Error GoTo ErrorHandler

    With frmReport
    
            .fraCriteria.Enabled = False
            .CmdView.Enabled = True
            .CmdPrint.Enabled = True
            
            'Hide all Label and ComboBox controls
            .lblSelectionLabel1.Visible = False
            .lblSelectionLabel2.Visible = False
            
            
            .lblSelectionLabel5.Visible = False
            .lblSelectionLabel6.Visible = False
            .lblSelectionLabel7.Visible = False
            .cboReportCriteria1.Visible = False
            .cboReportCriteria2.Visible = False
            .txtMemberRno.Visible = False
            
            .dteStartDate.Visible = False
            .dteEndDate.Visible = False
            
            .fraCriteria.Enabled = True
    
    End With
    InitialiseReport

Exit Sub
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modReport", "ValidateReportCriteria", True)

End Sub

Public Function CheckReportSecurity() As Boolean

CheckReportSecurity = False
If Not ReportView Then
         MsgBox "Invalid access - Reporting not available for current user access level.", vbExclamation
     Exit Function
Else
   CheckReportSecurity = True
End If
                
End Function
Public Sub Load_Default_Value(lngReportId As Long)
With frmReport
 Select Case lngReportId
   Case 0:
          .cboReportCriteria1.Text = "All"
          .cboReportCriteria2.Text = "All"
          .cboReportCriteria3.Text = "SURNAME"
   Case 1:
          .cboReportCriteria1.Text = "All"
          .cboReportCriteria2.Text = "All"
          .cboReportCriteria3.Text = "All"
          .dteStartDate.Text = Format(Now(), DATE_FORMAT)
          .dteEndDate.Text = Format(Now(), DATE_FORMAT)
   Case 2:
          .dteStartDate.Text = Format(Now(), DATE_FORMAT)
           .cboReportCriteria2.Text = "All"
  End Select
 End With
End Sub



Public Function ReportCriteria_1(lngReportId As Long)
On Error GoTo ErrorHandler
  ReportCriteria_Reset
    With frmReport

            .lblSelectionLabel1.Visible = False
            .lblSelectionLabel2.Visible = False
            .lblSelectionLabel3.Visible = False
           
            .lblSelectionLabel5.Visible = False
            .lblSelectionLabel6.Visible = False
            .lblSelectionLabel7.Visible = False
            .cboReportCriteria1.Visible = False
            .cboReportCriteria2.Visible = False
            .cboReportCriteria3.Visible = False
            .txtMemberRno.Visible = False
            .dteStartDate.Visible = False
            .dteEndDate.Visible = False


            Select Case lngReportId
        
        
                            Case 0:
                                        'Individual Prospect Profile
                                        .lblSelectionLabel1.Caption = "Post Code:"
                                        .lblSelectionLabel1.Top = 400
                                        .lblSelectionLabel1.Width = 2000
                                        .lblSelectionLabel1.Left = 3800
                                        .lblSelectionLabel1.Visible = True
                                        
                                        .cboReportCriteria1.Top = 400
                                        .cboReportCriteria1.Left = 6000
                                        .cboReportCriteria1.Visible = True
                                        
                                        LoadComboBox_PostCode .cboReportCriteria1
                                        
                                        .lblSelectionLabel2.Caption = "Status:"
                                        .lblSelectionLabel2.Top = 800
                                        .lblSelectionLabel2.Width = 2000
                                        .lblSelectionLabel2.Left = 3800
                                        .lblSelectionLabel2.Visible = True
                                                            
                                        .cboReportCriteria2.Top = 800
                                        .cboReportCriteria2.Left = 6000
                                        .cboReportCriteria2.Visible = True
                                        
                                        .cboReportCriteria2.Clear
                                        
                                        .cboReportCriteria2.AddItem "All"
                                        .cboReportCriteria2.AddItem "ACTIVE"
                                        .cboReportCriteria2.AddItem "NOT ACTIVE"
                                        .cboReportCriteria2.AddItem "DECEASED"
                                        
                                        .lblSelectionLabel3.Caption = "Sort By:"
                                        .lblSelectionLabel3.Top = 1200
                                        .lblSelectionLabel3.Width = 2000
                                        .lblSelectionLabel3.Left = 3800
                                        .lblSelectionLabel3.Visible = True
                                                            
                                        .cboReportCriteria3.Top = 1200
                                        .cboReportCriteria3.Left = 6000
                                        .cboReportCriteria3.Visible = True
                                        
                                        .cboReportCriteria3.Clear
                                        
                                        .cboReportCriteria3.AddItem "SURNAME"
                                        .cboReportCriteria3.AddItem "GIVEN_NAME"
                                        .cboReportCriteria3.AddItem "POSTCODE"
                                        .cboReportCriteria3.AddItem "MEMBERSHIP_EXPIARY"
                                        .cboReportCriteria3.AddItem "STATUS"
                                        .cboReportCriteria3.AddItem "MNO"
                                        
                                        
                                        .lblSelectionLabel5.Top = 1800
                                        .lblSelectionLabel5.Width = 2500
                                        .lblSelectionLabel5.Left = 4300
                                        .lblSelectionLabel5.Caption = " Expiry Date From"
                                        .lblSelectionLabel5.Visible = True
                                        
                                        .dteStartDate.Top = 1800
                                        .dteStartDate.Left = 6000
                                        .dteStartDate.Visible = True
                                        
                                        .lblSelectionLabel6.Top = 1800
                                        .lblSelectionLabel6.Width = 2500
                                        .lblSelectionLabel6.Left = 7300
                                        .lblSelectionLabel6.Caption = " Expiry Date To"
                                        .lblSelectionLabel6.Visible = True
                                        
                                        .dteEndDate.Top = 1800
                                        .dteEndDate.Left = 8900
                                        .dteEndDate.Visible = True
                                        Load_Default_Value (lngReportId)
        
                            Case 1:
                                        'New Business Activity
                                        .lblSelectionLabel1.Caption = "User Name:"
                                        
                                        .lblSelectionLabel1.Top = 400
                                        .lblSelectionLabel1.Width = 2000
                                        .lblSelectionLabel1.Left = 3800
                                        .lblSelectionLabel1.Visible = True
                                        
                                        .cboReportCriteria1.Top = 400
                                        .cboReportCriteria1.Left = 6000
                                        .cboReportCriteria1.Visible = True
                                        LoadComboBox_UserName .cboReportCriteria1
        
                                        .lblSelectionLabel2.Caption = "Payment Type:"
                                        .lblSelectionLabel2.Top = 800
                                        .lblSelectionLabel2.Width = 2000
                                        .lblSelectionLabel2.Left = 3800
                                        .lblSelectionLabel2.Visible = True
                                                            
                                        .cboReportCriteria2.Top = 800
                                        .cboReportCriteria2.Left = 6000
                                        .cboReportCriteria2.Visible = True
                                        LoadComboBox_PaymentType .cboReportCriteria2
                                        
                                        .lblSelectionLabel3.Caption = "Type:"
                                        .lblSelectionLabel3.Top = 1200
                                        .lblSelectionLabel3.Width = 2000
                                        .lblSelectionLabel3.Left = 3800
                                        .lblSelectionLabel3.Visible = True
                                                            
                                        .cboReportCriteria3.Top = 1200
                                        .cboReportCriteria3.Left = 6000
                                        .cboReportCriteria3.Visible = True
                                        
                                        LoadComboBox_Type .cboReportCriteria3
                                        
                                        
                                        
                                        
                                        .lblSelectionLabel5.Top = 1800
                                        .lblSelectionLabel5.Width = 2000
                                        .lblSelectionLabel5.Left = 4900
                                        .lblSelectionLabel5.Caption = " Star Date"
                                        .lblSelectionLabel5.Visible = True
                                        
                                        .dteStartDate.Top = 1800
                                        .dteStartDate.Left = 6000
                                        .dteStartDate.Visible = True
                                        
                                        .lblSelectionLabel6.Top = 1800
                                        .lblSelectionLabel6.Width = 2000
                                        .lblSelectionLabel6.Left = 7800
                                        .lblSelectionLabel6.Caption = " End Date"
                                        .lblSelectionLabel6.Visible = True
                                        
                                        .dteEndDate.Top = 1800
                                        .dteEndDate.Left = 8900
                                        .dteEndDate.Visible = True
                                        
                                        .lblSelectionLabel7.Top = 2100
                                        .lblSelectionLabel7.Width = 2000
                                        .lblSelectionLabel7.Left = 3800
                                        
                                        .lblSelectionLabel7.Visible = True
                                        .txtMemberRno.Top = 2100
                                        .txtMemberRno.Left = 6000
                                        
                                        .txtMemberRno.Visible = True
                                        Load_Default_Value (lngReportId)
                            Case 2:
                                        'Sales Executive Follow up
                                        .lblSelectionLabel5.Caption = "Today Date:"
                                        
                                        .lblSelectionLabel5.Top = 400
                                        .lblSelectionLabel5.Width = 2000
                                        .lblSelectionLabel5.Left = 3800
                                        .lblSelectionLabel5.Visible = True
                                        
                                        .dteStartDate.Top = 400
                                        .dteStartDate.Left = 6000
                                        .dteStartDate.Visible = True
                                        
                                        .lblSelectionLabel2.Caption = "Status:"
                                        .lblSelectionLabel2.Top = 800
                                        .lblSelectionLabel2.Width = 2000
                                        .lblSelectionLabel2.Left = 3800
                                        .lblSelectionLabel2.Visible = True
                                                            
                                        .cboReportCriteria2.Top = 800
                                        .cboReportCriteria2.Left = 6000
                                        .cboReportCriteria2.Visible = True
                                        
                                        .cboReportCriteria2.Clear
                                        
                                        .cboReportCriteria2.AddItem "All"
                                        .cboReportCriteria2.AddItem "Member"
                                        .cboReportCriteria2.AddItem "Not Member"
                                        
                                        
                                        Load_Default_Value (lngReportId)
                            Case 3:
                            Case 4:
                            .lblSelectionLabel5.Caption = "Today Date:"
                                        
                                        .lblSelectionLabel5.Top = 1800
                                        .lblSelectionLabel5.Width = 2000
                                        .lblSelectionLabel5.Left = 4900
                                        .lblSelectionLabel5.Caption = " Star Date"
                                        .lblSelectionLabel5.Visible = True
                                        
                                        .dteStartDate.Top = 1800
                                        .dteStartDate.Left = 6000
                                        .dteStartDate.Visible = True
                                        
                                        .lblSelectionLabel6.Top = 1800
                                        .lblSelectionLabel6.Width = 2000
                                        .lblSelectionLabel6.Left = 7800
                                        .lblSelectionLabel6.Caption = " End Date"
                                        .lblSelectionLabel6.Visible = True
                                        
                                        .dteEndDate.Top = 1800
                                        .dteEndDate.Left = 8900
                                        .dteEndDate.Visible = True
        
            End Select
            
            .fraCriteria.Enabled = True

    End With
    

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modReport", "SetupCriteria_1", True)

End Function

Public Function ValidateReportCriteria(lngReportId As Long) As Boolean
On Error GoTo ErrorHandler

    Dim i As Integer
    Dim bReportSelected As Boolean
    
    
    ValidateReportCriteria = False

    With frmReport
                        
            bReportSelected = False
            For i = 0 To REPORTS
                If .optReport(i).Value = True Then
                    bReportSelected = True
                End If
            Next i
        
            If bReportSelected = False Then
                MsgBox "Report type must be selected prior to generating report.", vbExclamation
                .optReport(0).SetFocus
                Exit Function
            End If


            'Note: Multiple CASE statements due to validation for multiple report types
            
            '1st Select CASE validation
            Select Case lngReportId

                        Case 0:

                                If Trim(.cboReportCriteria1.Text) = "" Then
                                    MsgBox "Invalid report critieria.  Selection criteria must be selected prior to generating report.", vbExclamation
                                    .cboReportCriteria1.SetFocus
                                    Exit Function
                                End If
                                If Trim(.cboReportCriteria2.Text) = "" Then
                                    MsgBox "Invalid report critieria.  Selection criteria must be selected prior to generating report.", vbExclamation
                                    .cboReportCriteria1.SetFocus
                                    Exit Function
                                End If
                                If Trim(.cboReportCriteria3.Text) = "" Then
                                    MsgBox "Invalid report critieria.  Selection criteria must be selected prior to generating report.", vbExclamation
                                    .cboReportCriteria1.SetFocus
                                    Exit Function
                                End If
                                
                        Case 1:

                                If Trim(.cboReportCriteria1.Text) = "" Then
                                    MsgBox "Invalid report critieria.  Selection criteria must be selected prior to generating report.", vbExclamation
                                    .cboReportCriteria1.SetFocus
                                    Exit Function
                                End If
                                If Trim(.cboReportCriteria2.Text) = "" Then
                                    MsgBox "Invalid report critieria.  Selection criteria must be selected prior to generating report.", vbExclamation
                                    .cboReportCriteria1.SetFocus
                                    Exit Function
                                End If
                                
                         Case 2:
                                    If Trim(.dteStartDate) = "" Then
                                        MsgBox "Invalid report critieria.  Selection criteria must be selected prior to generating report.", vbExclamation, "Today Date"
                                        Exit Function
                                    End If
                                    If Trim(.cboReportCriteria2.Text) = "" Then
                                        MsgBox "Invalid report critieria.  Selection criteria must be selected prior to generating report.", vbExclamation
                                        .dteStartDate.SetFocus
                                    Exit Function
                                
                                    End If
                                
                        
            End Select
            
            '2nd Select CASE validation
            Select Case lngReportId

                        Case 1:

                                If Trim(.dteStartDate) = "" Then
                                    MsgBox "Invalid report critieria.  Selection criteria must be selected prior to generating report.", vbExclamation, "Enter Start date"
                                    Exit Function
                                End If
                                If Trim(.dteEndDate) = "" Then
                                    MsgBox "Invalid report critieria.  Selection criteria must be selected prior to generating report.", vbExclamation, "Enter End date"
                                    Exit Function
                                End If
                                
                                    If Trim(.txtMemberRno.Text) = "" Then
                                    
                                    Else
                                      If Trim(.txtMemberRno.Text) <> "All" Then
                                      If Val(Trim(.txtMemberRno.Text)) = 0 Then
                                      MsgBox "Invalid report critieria.  Selection criteria must be selected prior to generating report.", vbExclamation
                                      .txtMemberRno.SetFocus
                                      Exit Function
                                      End If
                                      End If
                                     End If
                         Case 4:
                              If Trim(.dteStartDate) = "" Then
                                    MsgBox "Invalid report critieria.  Selection criteria must be selected prior to generating report.", vbExclamation, "Enter Start date"
                                    Exit Function
                                End If
                                If Trim(.dteEndDate) = "" Then
                                    MsgBox "Invalid report critieria.  Selection criteria must be selected prior to generating report.", vbExclamation, "Enter End date"
                                    Exit Function
                                End If
                                     
                        
            End Select
          
    End With

    ValidateReportCriteria = True

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modReport", "ValidateReportCriteria", True)

End Function


