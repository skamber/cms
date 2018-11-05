VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmPayment 
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8820
   ScaleMode       =   0  'User
   ScaleWidth      =   12255
   WindowState     =   2  'Maximized
   Begin VB.Frame fraPayment 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   8415
      Left            =   0
      TabIndex        =   11
      Top             =   480
      Width           =   12255
      Begin VB.ComboBox cmbDonationType 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPayment.frx":0000
         Left            =   6960
         List            =   "frmPayment.frx":0002
         TabIndex        =   4
         Text            =   "cmbDonationType"
         Top             =   2280
         Width           =   2535
      End
      Begin MSMask.MaskEdBox dteMemberExpiry 
         Height          =   315
         Left            =   7920
         TabIndex        =   32
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin VB.ComboBox cmbPaymentKind 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPayment.frx":0004
         Left            =   6960
         List            =   "frmPayment.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtReceiptNo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         TabIndex        =   17
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtUser 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         TabIndex        =   16
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtGivenName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtMemberNo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtSurname 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         TabIndex        =   14
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtAmount 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   3120
         Width           =   1455
      End
      Begin VB.ComboBox cmbStatus 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPayment.frx":0008
         Left            =   6240
         List            =   "frmPayment.frx":0015
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox cmbPaymentType 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPayment.frx":0037
         Left            =   4560
         List            =   "frmPayment.frx":0039
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2280
         Width           =   2175
      End
      Begin VB.ComboBox cmbAmountInWord 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   6
         Top             =   3120
         Width           =   4335
      End
      Begin VB.TextBox txtcomment 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   3960
         Width           =   9615
      End
      Begin MSMask.MaskEdBox dteExpiryDate 
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
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
      Begin MSMask.MaskEdBox dteEfectiveDate 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
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
      Begin MSMask.MaskEdBox dteDateOfPayment 
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
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
      Begin VB.Label Label17 
         Caption         =   "Donation Type"
         Enabled         =   0   'False
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
         Left            =   6960
         TabIndex        =   34
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Member Expiry Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   33
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Payment Been"
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
         TabIndex        =   31
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Memo"
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
         Left            =   240
         TabIndex        =   30
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Date Of Payment"
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
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Receipt No"
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
         Left            =   2640
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "User Name"
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
         Left            =   5160
         TabIndex        =   27
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Member No."
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
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Given Name"
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
         Left            =   1800
         TabIndex        =   25
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Surename"
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
         Left            =   3960
         TabIndex        =   24
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Status"
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
         Left            =   6360
         TabIndex        =   23
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Effective From Date"
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
         Left            =   240
         TabIndex        =   22
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Expiry Date"
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
         Left            =   2640
         TabIndex        =   21
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Payment Type"
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
         Left            =   6960
         TabIndex        =   20
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Amount"
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
         Left            =   240
         TabIndex        =   19
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Amount In Word"
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
         Left            =   2280
         TabIndex        =   18
         Top             =   2880
         Width           =   1695
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
      ScaleWidth      =   12195
      TabIndex        =   9
      Top             =   0
      Width           =   12255
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "PAYMENT"
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
         TabIndex        =   10
         Top             =   0
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport report 
      Left            =   8000
      Top             =   5000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowWidth     =   1200
      WindowBorderStyle=   1
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdPreview_Click()
 If cmbPaymentType.Text = "Membership" Then
Call GenerateReport(0, "View", objConnection)
Else
 Call GenerateReport(1, "View", objConnection)
End If
End Sub

Private Sub CmdPrint_Click()
If cmbPaymentType.Text = "Membership" Then
Call GenerateReport(0, "Print", objConnection)
Else
 Call GenerateReport(1, "Print", objConnection)
End If
End Sub

Private Sub cmbPaymentType_LostFocus()
 If cmbPaymentType.Text = "Donation" Then
 cmbDonationType.Enabled = True
 Label17.Enabled = True
 cmbDonationType.SetFocus
 Else
 cmbDonationType.Enabled = False
 cmbDonationType.Text = ""
 Label17.Enabled = False
 End If
End Sub

Private Sub cmbStatus_LostFocus()
Call UpdateMemberStatus
End Sub

Private Sub dteEfectiveDate_LostFocus()
 Call ValidDateEntry(dteEfectiveDate)
End Sub


Private Sub dteExpiryDate_LostFocus()
Call ValidDateEntry(dteExpiryDate)
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
gRecordMode = RECORD_READ
     With MDIFrm
     
            ' .pnlStatusBar.Panels(1).Text = "Loading information for Member..."
             
                  
     End With
     DoEvents
               
     'set record control variables
     gRecordType = Payment
     gRecordMode = RECORD_READ
     SetToolbarControl
     LoadComboBox
     MDIFrm.pnlStatusBar.Panels(1).Text = ""
     DoEvents
 
 
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "frmPayment", "Form_Load", True)



End Sub

Public Sub Form_Activate()

    If gRecordMode = RECORD_NEW Or gRecordMode = RECORD_EDIT Then
        fraPayment.Enabled = True
          
    Else
        fraPayment.Enabled = False
        
    End If
 gRecordType = Payment
End Sub


Private Sub txtAmount_LostFocus()
 Call ValidNumericEntry(txtAmount)
End Sub

Public Sub txtmemberNo_LostFocus()
If gRecordMode = RECORD_NEW Then
    If txtmemberNo.Text = "" Or Not IsNumeric(txtmemberNo.Text) Then
        MsgBox "Please Enter Member Number.", vbExclamation
        txtmemberNo.SetFocus
        Exit Sub
    ElseIf Not GetMemberInfo(txtmemberNo.Text) Then txtmemberNo.SetFocus
    End If
End If
End Sub

