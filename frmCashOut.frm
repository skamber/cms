VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCashOut 
   ClientHeight    =   9285
   ClientLeft      =   915
   ClientTop       =   2460
   ClientWidth     =   12225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   12225
   WindowState     =   2  'Maximized
   Begin VB.Frame FraCashOut 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   8775
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   12255
      Begin VB.TextBox txtitemCode 
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
         Left            =   3480
         TabIndex        =   18
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox cmbItem 
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
         Left            =   600
         TabIndex        =   0
         Top             =   1320
         Width           =   2415
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
         Left            =   600
         TabIndex        =   2
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtGST 
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
         Left            =   3480
         TabIndex        =   3
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtNetAmount 
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
         Left            =   5520
         TabIndex        =   4
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtCommand 
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
         Left            =   480
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   3360
         Width           =   7455
      End
      Begin VB.TextBox txtCashoutId 
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
         Left            =   600
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
      Begin MSMask.MaskEdBox dteDateofCashOut 
         Height          =   315
         Left            =   5520
         TabIndex        =   1
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label2 
         Caption         =   "Item"
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
         Left            =   600
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Date of Cash out"
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
         Left            =   5520
         TabIndex        =   16
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label4 
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
         Left            =   600
         TabIndex        =   15
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label5 
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
         Left            =   3480
         TabIndex        =   14
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Net Amount"
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
         Left            =   5520
         TabIndex        =   13
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Item Code"
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
         Left            =   3480
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "ID"
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
         Left            =   600
         TabIndex        =   11
         Top             =   240
         Width           =   255
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
      ScaleWidth      =   12165
      TabIndex        =   6
      Top             =   0
      Width           =   12225
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "CASH OUT"
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
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Label Label7 
      Caption         =   "ID"
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
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "frmCashOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbItem_LostFocus()
loadItemCodeAndGstAmountforCashOut
End Sub

Private Sub cmdChangeGST_Click()


End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
gRecordMode = RECORD_READ
     With MDIFrm
     
            ' .pnlStatusBar.Panels(1).Text = "Loading information for Member..."
             
                  
     End With
     DoEvents
               
     'set record control variables
     gRecordType = CASHOUT
     gRecordMode = RECORD_READ
     SetToolbarControl
     LoadCashOutComboBox
     MDIFrm.pnlStatusBar.Panels(1).Text = ""
     DoEvents
 
 
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "frmCashout", "Form_Load", True)



End Sub

Public Sub Form_Activate()

    If gRecordMode = RECORD_NEW Or gRecordMode = RECORD_EDIT Then
        FraCashOut.Enabled = True
          
    Else
        FraCashOut.Enabled = False
        
    End If
 gRecordType = CASHOUT
End Sub



Private Sub Form_Unload(Cancel As Integer)
clearAllSelection
End Sub

Private Sub txtAmount_LostFocus()
Call ValidNumericEntry(txtAmount)
If txtAmount.Text <> "" And txtGST.Text <> "" Then
txtGST.Text = txtGST.Text * txtAmount.Text
Call ValidNumericEntry(txtGST)
txtNetAmount.Text = txtAmount.Text - txtGST.Text
Call ValidNumericEntry(txtNetAmount)
End If
 
End Sub


Private Sub txtGST_LostFocus()
Call ValidNumericEntry(txtGST)
If txtAmount.Text <> "" And txtGST.Text <> "" Then
txtNetAmount.Text = txtAmount.Text - txtGST.Text
End If
Call ValidNumericEntry(txtNetAmount)
End Sub
