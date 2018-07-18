VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmCashFlowView 
   ClientHeight    =   9495
   ClientLeft      =   765
   ClientTop       =   1935
   ClientWidth     =   11970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   11970
   WindowState     =   2  'Maximized
   Begin VB.Frame FraCashFlow 
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   11895
      Begin VB.TextBox txtCashOutAmount 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   3240
         TabIndex        =   20
         Top             =   5400
         Width           =   1300
      End
      Begin VB.TextBox txtCashInAmount 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   3120
         TabIndex        =   19
         Top             =   2760
         Width           =   1300
      End
      Begin VB.TextBox txtCashOutGST 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   4680
         TabIndex        =   18
         Top             =   5400
         Width           =   1300
      End
      Begin VB.TextBox txtCashInGST 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   4560
         TabIndex        =   17
         Top             =   2760
         Width           =   1300
      End
      Begin VB.TextBox txtCashFlowNetAmount 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   6120
         TabIndex        =   11
         Top             =   6120
         Width           =   1300
      End
      Begin VB.TextBox txtCashOutNetAmount 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   6120
         TabIndex        =   10
         Top             =   5400
         Width           =   1300
      End
      Begin VB.TextBox txtCashInNetCashIn 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   6000
         TabIndex        =   9
         Top             =   2760
         Width           =   1300
      End
      Begin MSComctlLib.ListView ListCashOutView 
         Height          =   1935
         Left            =   120
         TabIndex        =   8
         Top             =   3240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   3413
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListCashInView 
         Height          =   1815
         Left            =   0
         TabIndex        =   7
         Top             =   720
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   3201
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSMask.MaskEdBox dteToDate 
         Height          =   315
         Left            =   4680
         TabIndex        =   1
         Top             =   30
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox dteFromDate 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   30
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Crystal.CrystalReport Report 
         Left            =   8400
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowBorderStyle=   1
         WindowControlBox=   -1  'True
         WindowMaxButton =   0   'False
         WindowMinButton =   0   'False
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Label Label12 
         Caption         =   "TOTAL AMOUNT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3240
         TabIndex        =   24
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "GST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4680
         TabIndex        =   23
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "TOTAL AMOUNT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3120
         TabIndex        =   22
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "GST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   21
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "CASHOUT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Width           =   900
      End
      Begin VB.Label Label7 
         Caption         =   "CASHIN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label6 
         Caption         =   "NET CASHFLOW :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6120
         TabIndex        =   14
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "NET CASHOUT "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6120
         TabIndex        =   13
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "NET CASHIN "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6000
         TabIndex        =   12
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "FROM DATE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   6
         Top             =   30
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "TO DATE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3840
         TabIndex        =   5
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00800000&
      FillStyle       =   6  'Cross
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11910
      TabIndex        =   2
      Top             =   0
      Width           =   11970
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "CASH FLOW VIEW"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmCashFlowView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TotalCashIn_GST As Double
Dim TotalCashin_Netamount As Double
Dim TotalCashin_Amount As Double
Dim TotalCashOut_GST As Double
Dim TotalCashOut_Netamount As Double
Dim TotalCashOut_Amount As Double
Dim TotalCashFlow_Netamount As Double

Private Sub CmdPrint_Click()

End Sub



Private Sub cmdView_LostFocus()
Call ValidDateEntry(dteFromDate)
End Sub





Private Sub dteToDate_LostFocus()
Call ValidDateEntry(dteToDate)
LoadCashFlowList
End Sub

Private Sub Form_Activate()
gRecordType = CASHFLOW_VIEW
End Sub

Private Sub Form_Load()
ListCashInView.ListItems.Clear
ListCashInView.ColumnHeaders.Add , , "Cash In ID", ListCashInView.Width / 12
ListCashInView.ColumnHeaders.Add , , "ITEM", ListCashInView.Width / 10
ListCashInView.ColumnHeaders.Add , , "Date Of CashIn", ListCashInView.Width / 7
ListCashInView.ColumnHeaders.Add , , "Amount", ListCashInView.Width / 7
ListCashInView.ColumnHeaders.Add , , "GST", ListCashInView.Width / 7
ListCashInView.ColumnHeaders.Add , , "Net Amount", ListCashInView.Width / 7
ListCashInView.ColumnHeaders.Add , , "Comment", ListCashInView.Width / 4

ListCashInView.View = lvwReport

ListCashOutView.ListItems.Clear
ListCashOutView.ColumnHeaders.Add , , "Cash Out ID", ListCashOutView.Width / 12
ListCashOutView.ColumnHeaders.Add , , "Item", ListCashOutView.Width / 10
ListCashOutView.ColumnHeaders.Add , , "date Of CashOut", ListCashOutView.Width / 7
ListCashOutView.ColumnHeaders.Add , , "Amount", ListCashOutView.Width / 7
ListCashOutView.ColumnHeaders.Add , , "GST", ListCashOutView.Width / 7
ListCashOutView.ColumnHeaders.Add , , "Net Amount", ListCashOutView.Width / 7
ListCashOutView.ColumnHeaders.Add , , "Comment", ListCashOutView.Width / 4
ListCashOutView.View = lvwReport

End Sub



Private Sub LoadCashFlowList()
LoadCashInList
If TotalCashIn_GST <> 0 Then
txtCashInGST.Text = TotalCashIn_GST
Call ValidNumericEntry(txtCashInGST)
End If

If TotalCashin_Netamount <> 0 Then
frmCashFlowView.txtCashInNetCashIn.Text = TotalCashin_Netamount
Call ValidNumericEntry(txtCashInNetCashIn)
End If

If TotalCashin_Amount <> 0 Then
txtCashInAmount.Text = TotalCashin_Amount
Call ValidNumericEntry(txtCashInAmount)
End If

LoadcashOutList
If TotalCashOut_GST <> 0 Then
txtCashOutGST.Text = TotalCashOut_GST
Call ValidNumericEntry(txtCashOutGST)
End If

If TotalCashOut_Netamount <> 0 Then
frmCashFlowView.txtCashOutNetAmount.Text = TotalCashOut_Netamount
Call ValidNumericEntry(txtCashOutNetAmount)
End If

If TotalCashOut_Amount <> 0 Then
txtCashOutAmount.Text = TotalCashOut_Amount
Call ValidNumericEntry(txtCashOutAmount)
End If
If TotalCashin_Netamount - TotalCashOut_Netamount <> 0 Then

txtCashFlowNetAmount.Text = TotalCashin_Netamount - TotalCashOut_Netamount
End If
Call ValidNumericEntry(txtCashFlowNetAmount)
End Sub
Private Sub LoadCashInList()

 On Error GoTo ErrorHandler
 
    ' Dim objFollowup_s As PACMSFollowUP.clsFollowup_s
     Dim rslocal As ADODB.Recordset
     Dim strId  As String
     Dim itmx As ListItem
     Dim sql As String
     
    ' Dim lngTotalProspect As Long
    ' Dim lngTotalPlanning As Long
    ' Dim lngTotalCoaching As Long
 TotalCashIn_GST = 0
 TotalCashin_Netamount = 0
 TotalCashin_Amount = 0
 
 
     Screen.MousePointer = vbHourglass
 
     With frmCashFlowView
             .ListCashInView.ListItems.Clear
     
         '==============================================================================
        
            Set rslocal = New ADODB.Recordset
            sql = "SELECT * FROM CASHIN Where DateOFCashIn between #" & Format(frmCashFlowView.dteFromDate.FormattedText, "dd-mmm-yyyy") & _
            "# AND #" & Format(frmCashFlowView.dteToDate.FormattedText, "dd-mmm-yyyy") & "#"
            rslocal.Open sql, objConnection, adOpenForwardOnly, adLockReadOnly
             If Not rslocal Is Nothing Then
                 Do While Not rslocal.EOF
                     
                     Set itmx = .ListCashInView.ListItems.Add(, , CStr(rslocal!Id))
                                    
                                     If Not IsNull(rslocal!Item) Then itmx.SubItems(1) = CStr(rslocal!Item)
                                     If Not IsNull(rslocal!DateOfCashin) Then itmx.SubItems(2) = CStr(rslocal!DateOfCashin)
                                     If Not IsNull(rslocal!Amount) Then itmx.SubItems(3) = CStr(Format(rslocal!Amount, NUMERIC_FORMAT))
                                     If Not IsNull(rslocal!gst) Then itmx.SubItems(4) = CStr(Format(rslocal!gst, NUMERIC_FORMAT))
                                     If Not IsNull(rslocal!net_amount) Then itmx.SubItems(5) = CStr(Format(rslocal!net_amount, NUMERIC_FORMAT))
                                     If Not IsNull(rslocal!Comment) Then itmx.SubItems(6) = CStr(rslocal!Comment)
                                     TotalCashIn_GST = TotalCashIn_GST + rslocal!gst
                                     TotalCashin_Netamount = TotalCashin_Netamount + rslocal!net_amount
                                     TotalCashin_Amount = TotalCashin_Amount + rslocal!Amount
                     Set itmx = Nothing
                     rslocal.MoveNext
                 Loop
             
                 Set rslocal = Nothing
             End If
     End With
     
     Screen.MousePointer = vbDefault
     CashInSelected = True
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "FrmCashflow", "LoadCashInList", True)

End Sub

Private Sub LoadcashOutList()

 On Error GoTo ErrorHandler
 
    ' Dim objFollowup_s As PACMSFollowUP.clsFollowup_s
     Dim rslocal As ADODB.Recordset
     Dim strId  As String
     Dim itmx As ListItem
     Dim sql As String
     
    ' Dim lngTotalProspect As Long
    ' Dim lngTotalPlanning As Long
    ' Dim lngTotalCoaching As Long
 TotalCashOut_GST = 0
 TotalCashOut_Netamount = 0
 TotalCashOut_Amount = 0
 
 
     Screen.MousePointer = vbHourglass
 
     With frmCashFlowView
             .ListCashOutView.ListItems.Clear
     
         '==============================================================================
        
            Set rslocal = New ADODB.Recordset
            sql = "SELECT * FROM CASHOUT Where DateOFCashOut between #" & Format(frmCashFlowView.dteFromDate.FormattedText, "dd-mmm-yyyy") & _
            "# AND #" & Format(frmCashFlowView.dteToDate.FormattedText, "dd-mmm-yyyy") & "#"
            rslocal.Open sql, objConnection, adOpenForwardOnly, adLockReadOnly
             If Not rslocal Is Nothing Then
                 Do While Not rslocal.EOF
                     
                     Set itmx = .ListCashOutView.ListItems.Add(, , CStr(rslocal!Id))
                                    
                                     If Not IsNull(rslocal!Item) Then itmx.SubItems(1) = CStr(rslocal!Item)
                                     If Not IsNull(rslocal!DateOfCashOut) Then itmx.SubItems(2) = CStr(rslocal!DateOfCashOut)
                                     If Not IsNull(rslocal!Amount) Then itmx.SubItems(3) = CStr(Format(rslocal!Amount, NUMERIC_FORMAT))
                                     If Not IsNull(rslocal!gst) Then itmx.SubItems(4) = CStr(Format(rslocal!gst, NUMERIC_FORMAT))
                                     If Not IsNull(rslocal!net_amount) Then itmx.SubItems(5) = CStr(Format(rslocal!net_amount, NUMERIC_FORMAT))
                                     If Not IsNull(rslocal!Comment) Then itmx.SubItems(6) = CStr(rslocal!Comment)
                                     
                                     TotalCashOut_GST = TotalCashOut_GST + rslocal!gst
                                     TotalCashOut_Netamount = TotalCashOut_Netamount + rslocal!net_amount
                                     TotalCashOut_Amount = TotalCashOut_Amount + rslocal!Amount
                     Set itmx = Nothing
                     rslocal.MoveNext
                 Loop
             
                 Set rslocal = Nothing
             End If
     End With
     
     Screen.MousePointer = vbDefault
     CashOutSelected = True
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "FrmCashflow", "LoadCashInList", True)

End Sub


