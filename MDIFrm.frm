VERSION 5.00
Object = "{15138B51-7EB6-11D0-9BB7-0000C0F04C96}#1.0#0"; "SSLstBar.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.MDIForm MDIFrm 
   BackColor       =   &H8000000C&
   Caption         =   "CMS"
   ClientHeight    =   7815
   ClientLeft      =   315
   ClientTop       =   2085
   ClientWidth     =   10875
   Icon            =   "MDIFrm.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin Listbar.SSListBar SSListBar1 
      Align           =   3  'Align Left
      Height          =   6840
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   12065
      _Version        =   65537
      PictureBackgroundStyle=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureBackground=   "MDIFrm.frx":030A
      OLEDragMode     =   1
      OLEDropMode     =   2
      GroupCount      =   7
      IconsLargeCount =   13
      Image(1).Index  =   1
      Image(1).Picture=   "MDIFrm.frx":355C
      Image(1).Key    =   "Child"
      Image(2).Index  =   2
      Image(2).Picture=   "MDIFrm.frx":3878
      Image(2).Key    =   "Sms"
      Image(3).Index  =   3
      Image(3).Picture=   "MDIFrm.frx":44CA
      Image(3).Key    =   "Item"
      Image(4).Index  =   4
      Image(4).Picture=   "MDIFrm.frx":47E4
      Image(4).Key    =   "Cashout"
      Image(5).Index  =   5
      Image(5).Picture=   "MDIFrm.frx":4AFE
      Image(5).Key    =   "Cashin"
      Image(6).Index  =   6
      Image(6).Picture=   "MDIFrm.frx":4E18
      Image(6).Key    =   "Receipt"
      Image(7).Index  =   7
      Image(7).Picture=   "MDIFrm.frx":5134
      Image(7).Key    =   "Reprt"
      Image(8).Index  =   8
      Image(8).Picture=   "MDIFrm.frx":5450
      Image(8).Key    =   "Search"
      Image(9).Index  =   9
      Image(9).Picture=   "MDIFrm.frx":576C
      Image(9).Key    =   "Member"
      Image(10).Index =   10
      Image(10).Picture=   "MDIFrm.frx":5A88
      Image(10).Key   =   "Invoice"
      Image(11).Index =   11
      Image(11).Picture=   "MDIFrm.frx":5DA4
      Image(11).Key   =   "Collection"
      Image(12).Index =   12
      Image(12).Picture=   "MDIFrm.frx":60C0
      Image(12).Key   =   "Search1"
      Image(13).Index =   13
      Image(13).Picture=   "MDIFrm.frx":63DC
      Image(13).Key   =   "Payment"
      Groups(1).ItemCount=   5
      Groups(1).PictureBackgroundStyle=   1
      Groups(1).CurrentGroup=   -1  'True
      Groups(1).Caption=   "Member"
      Groups(1).ListItems(1).Text=   "Member"
      Groups(1).ListItems(1).Key=   "MemberMenu"
      Groups(1).ListItems(1).IconLarge=   "Member"
      Groups(1).ListItems(2).Index=   2
      Groups(1).ListItems(2).Text=   "Search Members"
      Groups(1).ListItems(2).Key=   "SearchMembers"
      Groups(1).ListItems(2).IconLarge=   "Search"
      Groups(1).ListItems(3).Index=   3
      Groups(1).ListItems(3).Text=   "Children"
      Groups(1).ListItems(3).Key=   "ChildrenMenu"
      Groups(1).ListItems(3).IconLarge=   "Child"
      Groups(1).ListItems(4).Index=   4
      Groups(1).ListItems(4).Text=   "Search Children"
      Groups(1).ListItems(4).Key=   "SearchChildren"
      Groups(1).ListItems(4).IconLarge=   "Search"
      Groups(1).ListItems(5).Index=   5
      Groups(1).ListItems(5).Text=   "Send SMS"
      Groups(1).ListItems(5).Key=   "SendSms"
      Groups(1).ListItems(5).IconLarge=   "Sms"
      Groups(2).Index =   2
      Groups(2).ItemCount=   2
      Groups(2).PictureBackgroundStyle=   1
      Groups(2).Caption=   "Payment"
      Groups(2).ListItems(1).Text=   "Payment"
      Groups(2).ListItems(1).Key=   "PaymentMenu"
      Groups(2).ListItems(1).IconLarge=   "Payment"
      Groups(2).ListItems(2).Index=   2
      Groups(2).ListItems(2).Text=   "Payment Search"
      Groups(2).ListItems(2).Key=   "PaymentSearch"
      Groups(2).ListItems(2).IconLarge=   "Search1"
      Groups(3).Index =   3
      Groups(3).ItemCount=   2
      Groups(3).PictureBackgroundStyle=   1
      Groups(3).Caption=   "Collection"
      Groups(3).ListItems(1).Text=   "Collection"
      Groups(3).ListItems(1).Key=   "CollectionMenu"
      Groups(3).ListItems(1).IconLarge=   "Collection"
      Groups(3).ListItems(2).Index=   2
      Groups(3).ListItems(2).Text=   "Collection Search"
      Groups(3).ListItems(2).Key=   "CollectionSearch"
      Groups(3).ListItems(2).IconLarge=   "Search"
      Groups(4).Index =   4
      Groups(4).ItemCount=   2
      Groups(4).PictureBackgroundStyle=   1
      Groups(4).Caption=   "Invoice"
      Groups(4).ListItems(1).Text=   "Invoice "
      Groups(4).ListItems(1).Key=   "InvoiceMenu"
      Groups(4).ListItems(1).IconLarge=   "Invoice"
      Groups(4).ListItems(2).Index=   2
      Groups(4).ListItems(2).Text=   "Invoice Search"
      Groups(4).ListItems(2).Key=   "InvoiceSearch"
      Groups(4).ListItems(2).IconLarge=   "Search"
      Groups(5).Index =   5
      Groups(5).ItemCount=   2
      Groups(5).PictureBackgroundStyle=   1
      Groups(5).Caption=   "Receipt"
      Groups(5).ListItems(1).Text=   "Receipt"
      Groups(5).ListItems(1).Key=   "ReceiptMenu"
      Groups(5).ListItems(1).IconLarge=   "Receipt"
      Groups(5).ListItems(2).Index=   2
      Groups(5).ListItems(2).Text=   "Receipt Search"
      Groups(5).ListItems(2).Key=   "ReceiptSearch"
      Groups(5).ListItems(2).IconLarge=   "Search"
      Groups(6).Index =   6
      Groups(6).ItemCount=   4
      Groups(6).PictureBackgroundStyle=   1
      Groups(6).Caption=   "Cash Flow"
      Groups(6).ListItems(1).Text=   "Cash flow"
      Groups(6).ListItems(1).Key=   "cashflow"
      Groups(6).ListItems(1).IconLarge=   "Search"
      Groups(6).ListItems(2).Index=   2
      Groups(6).ListItems(2).Text=   "Cash In"
      Groups(6).ListItems(2).Key=   "CashIn"
      Groups(6).ListItems(2).IconLarge=   "Cashin"
      Groups(6).ListItems(3).Index=   3
      Groups(6).ListItems(3).Text=   "Cash Out"
      Groups(6).ListItems(3).Key=   "CashOut"
      Groups(6).ListItems(3).IconLarge=   "Cashout"
      Groups(6).ListItems(4).Index=   4
      Groups(6).ListItems(4).Text=   "Item"
      Groups(6).ListItems(4).Key=   "CashFlowItem"
      Groups(6).ListItems(4).IconLarge=   "Item"
      Groups(7).Index =   7
      Groups(7).ItemCount=   1
      Groups(7).PictureBackgroundStyle=   1
      Groups(7).Caption=   "Report"
      Groups(7).ListItems(1).Text=   "Report"
      Groups(7).ListItems(1).Key=   "Report"
      Groups(7).ListItems(1).IconLarge=   "Reprt"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrm.frx":66F8
            Key             =   "update"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrm.frx":6A14
            Key             =   "insert"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrm.frx":6D30
            Key             =   "save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrm.frx":704C
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrm.frx":7368
            Key             =   "cancle"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrm.frx":7684
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrm.frx":79A0
            Key             =   "Preview"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   1058
      ButtonWidth     =   2223
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar pnlStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7440
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   18699
            MinWidth        =   18699
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8469
            MinWidth        =   8469
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MainMnu 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu ExitMnu 
         Caption         =   "&Exit"
         Index           =   0
      End
      Begin VB.Menu MnuSystemSetting 
         Caption         =   "&System Setting"
      End
      Begin VB.Menu MnuChangePassword 
         Caption         =   "&Change Password"
      End
   End
   Begin VB.Menu MainMnu 
      Caption         =   "&Help"
      Index           =   1
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   0
      End
   End
End
Attribute VB_Name = "MDIFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ExitMnu_Click(index As Integer)
Select Case index
Case 0:     objConnection.Close
            Set objConnection = Nothing
            End
End Select
End Sub

Private Sub MDIForm_Load()
pnlStatusBar.Panels(1).Text = ""
getChruchNameById (gChurchId)
getCityNameById
If Not systemManager Then MDIFrm.MnuSystemSetting.Visible = False
End Sub

Private Sub MnuChangePassword_Click()
frmPassword.Show vbModal
End Sub

Private Sub mnuHelp_Click(index As Integer)
Select Case index
Case 0: frmAbout.Show
End Select

End Sub

Private Sub MnuSystemSetting_Click()
frmUserAccount.Show vbModal
End Sub

Private Sub SSListBar1_GroupClick(ByVal GroupClicked As Listbar.SSGroup, ByVal PreviousGroup As Listbar.SSGroup)

'On Error GoTo ErrorHandler
If gRecordMode = RECORD_NEW Or gRecordMode = RECORD_EDIT Then
        SSListBar1.CurrentGroup = PreviousGroup.index
        MsgBox "Please select Save or Cancel before continuing", vbExclamation + vbOKOnly
        Exit Sub
    End If
Select Case PreviousGroup.Caption
    Case "Member":
                Unload frmMember
                Unload frmChildren
                Unload frmMemberSearch
                Unload frmChildSearch
                Unload frmSms
                Set frmMember = Nothing
                Set frmChildren = Nothing
                Set frmMemberSearch = Nothing
                Set frmSms = Nothing
                MemberSelected = False
                ChildSelected = False
    Case "Payment":
                    Unload frmPayment
                    Unload frmPaymentSearch
                    Set frmPayment = Nothing
                    Set frmPaymentSearch = Nothing
                    PaymentSelected = False
    Case "Collection":
                        Unload frmCollection
                        Unload frmCollectionSearch
                        Set frmCollection = Nothing
                        Set frmCollectionSearch = Nothing
                        CollectionSelected = False
    Case "Invoice":
                        Unload frmInvoice
                        Unload frmInvoiceSearch
                        Set frmInvoice = Nothing
                        Set frmInvoiceSearch = Nothing
                        InvoiceSelected = False
                        InvoiceItemSelected = False
                        
    Case "Receipt":
                        Unload frmReceipt
                        Unload frmReceiptSearch
                        Set frmReceipt = Nothing
                        Set frmReceiptSearch = Nothing
    
    Case "Cash Flow":    Unload frmCashIn
                        Unload frmCashOut
                        Unload frmCashFlowView
                        Unload frmCashFlowItem
                        CashInSelected = False
                        CashOutSelected = False
                        CashflowItemSelected = False
    Case "Report":
                    Unload frmReport
                    Set frmReport = Nothing
    
End Select
Select Case GroupClicked.Caption
    Case "Member": frmMember.Show
    Case "Payment": frmPayment.Show
    Case "Collection": frmCollection.Show
    Case "Invoice": frmInvoice.Show
    Case "Receipt": frmReceipt.Show
    Case "Cash Flow": frmCashFlowView.Show
    Case "Report": frmReport.Show
End Select
DoEvents

Select Case GroupClicked.Caption
 Case "Member": frmMember.ZOrder 0
    Case "Payment": frmPayment.ZOrder 0
    Case "Collection": frmCollection.ZOrder 0
    Case "Invoice": frmInvoice.ZOrder 0
    Case "Receipt": frmReceipt.ZOrder 0
    Case "Report": frmReport.ZOrder 0
     Case "Cash Flow": frmCashFlowView.ZOrder 0
End Select
Exit Sub
ErrorHandler:

    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "MDIFRM", "SSListBar1_GroupClick", True)

End Sub


Private Sub SSListBar1_ListItemClick(ByVal ItemClicked As Listbar.SSListItem)
If gRecordMode = RECORD_NEW Or gRecordMode = RECORD_EDIT Then
        MsgBox "Please select Save or Cancel before continuing", vbExclamation + vbOKOnly
        Exit Sub
    End If
Select Case ItemClicked.Key
    Case Member_information:
                     frmMember.Show
                     DoEvents
                     frmMember.ZOrder 0
                     If MemberSelected Then
                       gmemberId = frmMemberSearch.ListView.SelectedItem
                       DisplayMember
                     End If
                     gRecordType = MEMBER
                     
    Case Member_Search: frmMemberSearch.Show: DoEvents: frmMemberSearch.ZOrder 0
    Case Children_information:
                     frmChildren.Show
                     DoEvents
                     frmChildren.ZOrder 0
                     If ChildSelected Then
                       gchildId = frmChildSearch.ListView.SelectedItem
                       DisplayChild
                     End If
                     clearAllSelection
    Case Child_Search: frmChildSearch.Show: DoEvents: frmChildSearch.ZOrder 0
    Case Send_Sms: frmSms.Show: DoEvents: frmSms.ZOrder 0
    Case Payment_information:
                    frmPayment.Show
                    DoEvents
                    frmPayment.ZOrder 0
                    If PaymentSelected Then
                    gPaymentId = frmPaymentSearch.ListPaymentView.SelectedItem.SubItems(6)
                       DisplayPayment
                     End If
                     
    Case Payment_search: frmPaymentSearch.Show: DoEvents: frmPaymentSearch.ZOrder 0
    Case Collection_information:
                        frmCollection.Show
                        DoEvents
                        frmCollection.ZOrder 0
                        If CollectionSelected Then
                            gCollectionId = frmCollectionSearch.ListCollectionView.SelectedItem
                            DisplayCollection
                        End If
    Case collection_search: frmCollectionSearch.Show: DoEvents: frmCollectionSearch.ZOrder 0
    Case Invoice_information:
                        frmInvoice.Show
                        DoEvents
                        frmInvoice.ZOrder 0
                        If InvoiceSelected Then
                    gInvoiceId = frmInvoiceSearch.ListInvoiceView.SelectedItem
                       DisplayInvoice
                       displayInvoiceItemList
                     End If
    Case Invoice_search: frmInvoiceSearch.Show: DoEvents: frmInvoiceSearch.ZOrder 0
    Case Receipt_information:
                        frmReceipt.Show
                        DoEvents
                      '  frmReceipt.ZOrder 0
                      '  If ReceiptSelected Then
                      '      gReceiptId = frmReceiptSearch.ListReceiptView.SelectedItem
                      '      DisplayReceipt
                      '  End If
    Case receipt_search: frmReceiptSearch.Show: DoEvents: frmReceiptSearch.ZOrder 0
    Case CASHFLOW_VIEW: frmCashFlowView.Show: DoEvents: frmCashFlowView.ZOrder 0
    Case CASHOUT:
               frmCashOut.Show
               DoEvents
               frmCashOut.ZOrder 0
               If CashOutSelected Then
               gCashoutId = frmCashFlowView.ListCashOutView.SelectedItem
                       DisplayCashOut
                     End If
               
    Case CASHIN:
               frmCashIn.Show
               DoEvents
               frmCashIn.ZOrder 0
               If CashInSelected Then
               gCashInId = frmCashFlowView.ListCashInView.SelectedItem
                       DisplayCashIn
                     End If
               
    Case CASHFLOWITEM: frmCashFlowItem.Show
                        DoEvents
                        frmCashFlowItem.ZOrder 0
                        gCashflowItemId = frmCashFlowItem.ListCashflowItemView.SelectedItem
                        Call DispayCashflowItem
                        
    Case Report: frmReport.Show: DoEvents: frmReport.ZOrder 0
End Select
SetToolbarControl

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Call ProcessRecord(Button)
End Sub

Public Function getChruchNameById(ByVal churchId As Long)
  On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rslocal As ADODB.Recordset
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection
    
    'Payment
    Set rslocal = objOrganisation_s.getChurchNameById(churchId)

    With frmCashIn
            If Not rslocal Is Nothing Then
             pnlStatusBar.Panels(2).Text = rslocal!Name
              pnlStatusBar.Panels(2).Width = 5000
              
             Set rslocal = Nothing
            End If
                   
    End With
    Set objOrganisation_s = Nothing
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashIn", "LoadCashInComboBox", True)

End Function

Public Function getCityNameById()
  On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rslocal As ADODB.Recordset
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection
    Set rslocal = objOrganisation_s.getCityNameById(gCityId)
    
    With frmCashIn
            If Not rslocal Is Nothing Then
             pnlStatusBar.Panels(3).Text = rslocal!cityName
             pnlStatusBar.Panels(3).Width = 4000
             Set rslocal = Nothing
            End If
                   
    End With
 
    Set objOrganisation_s = Nothing
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modCashIn", "LoadCashInComboBox", True)

End Function
