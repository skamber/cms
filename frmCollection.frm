VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCollection 
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12255
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtUserName 
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
      TabIndex        =   15
      Top             =   1080
      Width           =   1875
   End
   Begin VB.Frame FraCollection 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   8295
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   12255
      Begin VB.ComboBox cmbTypes 
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
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtCollectionNo 
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
         Left            =   2280
         TabIndex        =   8
         Top             =   600
         Width           =   1395
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
         Left            =   3480
         TabIndex        =   1
         Top             =   1440
         Width           =   1275
      End
      Begin VB.TextBox txtComments 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   2280
         Width           =   6975
      End
      Begin VB.ComboBox cmbType 
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
         ItemData        =   "frmCollection.frx":0000
         Left            =   240
         List            =   "frmCollection.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1440
         Width           =   2655
      End
      Begin MSMask.MaskEdBox dteDateofPayment 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1635
         _ExtentX        =   2884
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
      Begin VB.Label Label10 
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
         Height          =   255
         Left            =   5280
         TabIndex        =   17
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label8 
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
         Left            =   3960
         TabIndex        =   16
         Top             =   360
         Width           =   1575
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
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Collection No"
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
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Type"
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
         TabIndex        =   12
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label5 
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
         Left            =   3360
         TabIndex        =   11
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Comments"
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
         TabIndex        =   9
         Top             =   2040
         Width           =   1095
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
      TabIndex        =   4
      Top             =   0
      Width           =   12255
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "COLLECTION"
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
         TabIndex        =   5
         Top             =   30
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmCollection"
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
     gRecordType = COLLECTION
     gRecordMode = RECORD_READ
     SetToolbarControl
     
     MDIFrm.pnlStatusBar.Panels(1).Text = ""
     DoEvents
 
 LoadCollectionComboBox
 
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "frmCollection", "Form_Load", True)


End Sub

Public Sub Form_Activate()

    If gRecordMode = RECORD_NEW Or gRecordMode = RECORD_EDIT Then
        FraCollection.Enabled = True
          
    Else
        FraCollection.Enabled = False
        
    End If
   gRecordType = COLLECTION
End Sub



Private Sub txtAmount_LostFocus()
Call ValidNumericEntry(txtAmount)
End Sub

