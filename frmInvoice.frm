VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmInvoice 
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   13560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   13560
   WindowState     =   2  'Maximized
   Begin VB.Frame fraInvoice 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   8895
      Left            =   0
      TabIndex        =   13
      Top             =   480
      Width           =   13455
      Begin VB.TextBox txtEmail 
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
         Left            =   9120
         TabIndex        =   4
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Frame fraInvoiceItem 
         Height          =   2415
         Left            =   120
         TabIndex        =   34
         Top             =   5520
         Visible         =   0   'False
         Width           =   13335
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   11760
            TabIndex        =   47
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton CmdChange 
            Caption         =   "Change"
            Enabled         =   0   'False
            Height          =   375
            Left            =   11760
            TabIndex        =   46
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton CmdSave 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   375
            Left            =   11760
            TabIndex        =   45
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton CmdCancel 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   375
            Left            =   11760
            TabIndex        =   44
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Frame framInvoiceItem 
            Enabled         =   0   'False
            Height          =   2055
            Left            =   240
            TabIndex        =   35
            Top             =   240
            Width           =   8655
            Begin VB.ComboBox CmbItemDescription 
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
               Left            =   120
               TabIndex        =   39
               Top             =   600
               Width           =   6015
            End
            Begin VB.TextBox txtTotalItemAmount 
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
               Left            =   4320
               TabIndex        =   38
               Top             =   1560
               Width           =   1935
            End
            Begin VB.TextBox txtItemGST 
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
               Left            =   2520
               TabIndex        =   37
               Top             =   1560
               Width           =   1215
            End
            Begin VB.TextBox txtItemAmount 
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
               Left            =   120
               TabIndex        =   36
               Top             =   1560
               Width           =   1695
            End
            Begin VB.Label Label20 
               Caption         =   "Amount"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label Label19 
               Caption         =   "GST"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2520
               TabIndex        =   42
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label Label18 
               Caption         =   "Total Amount"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4320
               TabIndex        =   41
               Top             =   1320
               Width           =   1455
            End
            Begin VB.Label Label16 
               Caption         =   "Description"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   360
               Width           =   1215
            End
         End
      End
      Begin MSMask.MaskEdBox dteInvoiceDate 
         Height          =   315
         Left            =   5760
         TabIndex        =   30
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSMask.MaskEdBox dtePhone 
         Height          =   315
         Left            =   5040
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "(##)####-####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtState 
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
         Left            =   5400
         TabIndex        =   7
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtInvoiceNo 
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
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtRef 
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
         Left            =   3240
         TabIndex        =   15
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtName 
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
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtSurname 
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
         TabIndex        =   1
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtMobile 
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
         Left            =   7080
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtAddress1 
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
         TabIndex        =   5
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtAddress2 
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
         Left            =   2400
         TabIndex        =   6
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtTotalAmount 
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
         Left            =   4680
         TabIndex        =   10
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox txtPostcode 
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
         Left            =   6960
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtBalance 
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
         Left            =   6720
         TabIndex        =   14
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ComboBox cmbTerms 
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
         Left            =   360
         TabIndex        =   9
         Top             =   3000
         Width           =   1815
      End
      Begin MSComctlLib.ListView InvoiceItemList 
         Height          =   1695
         Left            =   0
         TabIndex        =   31
         Top             =   3720
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   2990
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
      Begin Crystal.CrystalReport Report 
         Left            =   9360
         Top             =   2760
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
      Begin MSMask.MaskEdBox dteOverDueDate 
         Height          =   315
         Left            =   2400
         TabIndex        =   32
         Top             =   3000
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.Label Label22 
         Caption         =   "Email"
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
         Left            =   9120
         TabIndex        =   50
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Invoice No"
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
         TabIndex        =   49
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Over Due Date"
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
         Left            =   2400
         TabIndex        =   33
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Ref"
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
         TabIndex        =   29
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Date"
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
         Left            =   5760
         TabIndex        =   28
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Name or Company Name"
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
         TabIndex        =   27
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Surname"
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
         TabIndex        =   26
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Phone"
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
         Left            =   5040
         TabIndex        =   25
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Mobile"
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
         Left            =   7080
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Address1"
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
         TabIndex        =   23
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Address2"
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
         Left            =   2400
         TabIndex        =   22
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "State"
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
         Left            =   5400
         TabIndex        =   21
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Postcode"
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
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Balance"
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
         Left            =   6720
         TabIndex        =   19
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Terms"
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
         TabIndex        =   18
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Total Amount"
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
         TabIndex        =   17
         Top             =   2760
         Width           =   1335
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
      ScaleWidth      =   13500
      TabIndex        =   11
      Top             =   0
      Width           =   13560
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "INVOICE"
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
         TabIndex        =   12
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Ref"
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
      Left            =   0
      TabIndex        =   48
      Top             =   1000
      Width           =   855
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbTerms_LostFocus()
       Select Case cmbTerms.Text
        Case "7 Days":  dteOverDueDate.Text = Format(Date + 7, DATE_FORMAT)
        Case "14 Days": dteOverDueDate.Text = Format(Date + 14, DATE_FORMAT)
        Case "28 Days": dteOverDueDate.Text = Format(Date + 28, DATE_FORMAT)
       End Select
End Sub

Private Sub cmdAdd_Click()
If txtInvoiceNo.Text = "" Then
MsgBox "Please Save the Invoice before Adding Invoice Items", vbInformation, "CMS - Error Continuing"
Else
    CmdChange.Enabled = False
    cmdAdd.Enabled = False
    
    CmdSave.Enabled = True
    InvoiceItemList.Enabled = False
    InitialiseInvoiceItem
    framInvoiceItem.Enabled = True
    CmdCancel.Enabled = True

    mInvoiceItemMode = RECORD_NEW
    CmbItemDescription.SetFocus
 
End If
End Sub

Private Sub cmdCancel_Click()
 On Error GoTo ErrorHandler

   ' fraStrategyActionList.Enabled = True
    cmdAdd.Enabled = True
    CmdChange.Enabled = False
    
    
    InitialiseInvoiceItem
    
    CmdSave.Enabled = False
    CmdCancel.Enabled = False
    framInvoiceItem.Enabled = False
    mInvoiceItemMode = RECORD_READ
    InvoiceItemList.Enabled = True
   'Call DisplayStrategyActionFromList

    Exit Sub
    
ErrorHandler:

    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "frmStrategyAction", "FormLoad", True)

End Sub

Private Sub CmdChange_Click()
On Error GoTo ErrorHandler

   ' fraStrategyActionDetail.Enabled = True
    CmdSave.Enabled = True
    CmdCancel.Enabled = True
    InvoiceItemList.Enabled = False
     cmdAdd.Enabled = False
    CmdChange.Enabled = False
    
    framInvoiceItem.Enabled = True
    
    mInvoiceItemMode = RECORD_EDIT


Exit Sub
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "frmStrategyAction", "FormLoad", True)

End Sub

Private Sub CmdSave_Click()
SaveInvoiceItem
displayInvoiceItemList
If mInvoiceItemMode = RECORD_NEW Or gRecordMode = RECORD_EDIT Then
  frmInvoice.cmdAdd.Enabled = True
  frmInvoice.CmdCancel.Enabled = False
  frmInvoice.CmdSave.Enabled = False
  framInvoiceItem.Enabled = False
  frmInvoice.CmdChange.Enabled = True
  InvoiceItemList.Enabled = True
  mInvoiceItemMode = RECORD_READ
End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler

     With MDIFrm
     
            ' .pnlStatusBar.Panels(1).Text = "Loading information for Member..."
             
                  
     End With
     DoEvents
               
     'set record control variables
     gRecordType = INVOICE
     gRecordMode = RECORD_READ
     mInvoiceItemMode = RECORD_READ
     SetToolbarControl
     
     MDIFrm.pnlStatusBar.Panels(1).Text = ""
     DoEvents
 LoadInvoiceComboBox
 SetupInvoiceItemList
 
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "frmInvoice", "Form_Load", True)


End Sub

Public Sub Form_Activate()

    If gRecordMode = RECORD_NEW Or gRecordMode = RECORD_EDIT Then
        fraInvoice.Enabled = True
    Else
        fraInvoice.Enabled = False
    End If
 gRecordType = INVOICE
End Sub



Private Sub Text1_Change()

End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub InvoiceItemList_Click()
If InvoiceItemSelected Then
frmInvoice.CmdChange.Enabled = True
gInvoiceItemId = frmInvoice.InvoiceItemList.SelectedItem
DisplayInvoiceItem
End If
End Sub



Private Sub txtItemAmount_LostFocus()
Call ValidNumericEntry(txtItemAmount)
GetGstAndTotalAmount
End Sub
