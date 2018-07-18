VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmChildSearch 
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12510
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   12510
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtmemberNo 
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
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00800000&
      FillStyle       =   6  'Cross
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   12450
      TabIndex        =   0
      Top             =   0
      Width           =   12510
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "CHILD SEARCH"
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
         TabIndex        =   1
         Top             =   0
         Width           =   3015
      End
   End
   Begin MSComctlLib.ListView ListChildrenView 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4683
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
   Begin VB.Label Label1 
      Caption         =   "Member No"
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
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmChildSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
gRecordType = Child_Search
End Sub

Private Sub Form_Load()
ListChildrenView.ListItems.Clear
ListChildrenView.Width = Screen.Width - 5000
ListChildrenView.Height = Screen.Height - 5000

ListChildrenView.ColumnHeaders.Add , , "Child Number", ListChildrenView.Width / 7
ListChildrenView.ColumnHeaders.Add , , "Member Number", ListChildrenView.Width / 7
ListChildrenView.ColumnHeaders.Add , , "First Name", ListChildrenView.Width / 5
ListChildrenView.ColumnHeaders.Add , , "Last Name", ListChildrenView.Width / 5
ListChildrenView.ColumnHeaders.Add , , "Genda", ListChildrenView.Width / 14
ListChildrenView.ColumnHeaders.Add , , "Birth Data", ListChildrenView.Width / 6
ListChildrenView.ColumnHeaders.Add , , "Status", ListChildrenView.Width / 9
ListChildrenView.View = lvwReport
End Sub



Private Sub txtmemberNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtmemberNo.Text <> "" Then GenerateChildrenList (txtmemberNo.Text)

End Sub

Private Sub txtmemberNo_LostFocus()
If txtmemberNo.Text <> "" Then GenerateChildrenList (txtmemberNo.Text)

End Sub
