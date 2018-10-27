VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmReceipt 
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   11385
   WindowState     =   2  'Maximized
   Begin VB.Frame fraReceipt 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   8055
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   11415
      Begin VB.TextBox txtChequeNo 
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
         Left            =   4680
         TabIndex        =   18
         Top             =   1800
         Width           =   2000
      End
      Begin VB.TextBox txtInvoiceID 
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
         Left            =   2520
         TabIndex        =   16
         Top             =   1080
         Width           =   1395
      End
      Begin VB.TextBox txtUser 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtReceiptId 
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
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtInvoiceNo 
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
         Left            =   360
         TabIndex        =   0
         Top             =   1080
         Width           =   1515
      End
      Begin VB.TextBox txtName 
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
         Left            =   360
         TabIndex        =   7
         Top             =   1800
         Width           =   1515
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
         Left            =   2460
         TabIndex        =   6
         Top             =   1800
         Width           =   2000
      End
      Begin VB.TextBox txtAmountToPay 
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
         Left            =   360
         TabIndex        =   1
         Top             =   2520
         Width           =   1395
      End
      Begin MSMask.MaskEdBox dteDate 
         Height          =   315
         Left            =   4260
         TabIndex        =   5
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
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
      Begin Crystal.CrystalReport Report 
         Left            =   6480
         Top             =   2520
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
      Begin VB.Label Label8 
         Caption         =   "Cheque Number:"
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
         TabIndex        =   19
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Receipt ID:"
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
         Left            =   2520
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Receipt No:"
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
         Left            =   360
         TabIndex        =   14
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Invoice Number:"
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
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Name:"
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
         Left            =   360
         TabIndex        =   12
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Date:"
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
         Left            =   4260
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Surname:"
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
         Left            =   2460
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Amount To Pay :"
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
         Left            =   360
         TabIndex        =   9
         Top             =   2280
         Width           =   1455
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
      ScaleWidth      =   11325
      TabIndex        =   2
      Top             =   0
      Width           =   11385
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "RECEIPT"
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
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo ErrorHandler

     With MDIFrm
     
            ' .pnlStatusBar.Panels(1).Text = "Loading information for Member..."
             
                  
     End With
     DoEvents
               
     'set record control variables
     gRecordType = RECEIPT
     gRecordMode = RECORD_READ
     SetToolbarControl
     
     MDIFrm.pnlStatusBar.Panels(1).Text = ""
     DoEvents
 
' LoadCollectionComboBox
 
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "frmReceipt", "Form_Load", True)


End Sub

Public Sub Form_Activate()

    If gRecordMode = RECORD_NEW Or gRecordMode = RECORD_EDIT Then
        fraReceipt.Enabled = True
          
    Else
        fraReceipt.Enabled = False
        
    End If
    
End Sub


Private Sub txtAmountToPay_LostFocus()
Call ValidNumericEntry(txtAmountToPay)
End Sub

Private Sub txtInvoiceNo_LostFocus()
GetInvoiceInfo
End Sub

Private Sub txtxSurname_Change()

End Sub
