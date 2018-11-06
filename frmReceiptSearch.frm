VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmReceiptSearch 
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   13275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8203.734
   ScaleMode       =   0  'User
   ScaleWidth      =   13275
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbTypeSearch 
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
      ItemData        =   "frmReceiptSearch.frx":0000
      Left            =   840
      List            =   "frmReceiptSearch.frx":000A
      TabIndex        =   3
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtInputText 
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
      Left            =   840
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00800000&
      FillStyle       =   6  'Cross
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   13215
      TabIndex        =   0
      Top             =   0
      Width           =   13275
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "RECEIPT SEARCH"
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
   Begin MSComctlLib.ListView ListReceiptView 
      Height          =   2775
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   4895
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
End
Attribute VB_Name = "frmReceiptSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
gRecordType = receipt_search
End Sub

Private Sub Form_Load()
ListReceiptView.ListItems.Clear
ListReceiptView.Width = Screen.Width - 5000
ListReceiptView.Height = Screen.Height - 5000
ListReceiptView.ColumnHeaders.Add , , "RECEIPT NUMBER", ListReceiptView.Width / 7
ListReceiptView.ColumnHeaders.Add , , "INVOICE ID", ListReceiptView.Width / 7
ListReceiptView.ColumnHeaders.Add , , "INVOICE NUMBER", ListReceiptView.Width / 4
ListReceiptView.ColumnHeaders.Add , , "AMOUNT", ListReceiptView.Width / 4
ListReceiptView.ColumnHeaders.Add , , "DATE", ListReceiptView.Width / 4
ListReceiptView.View = lvwReport

End Sub

Private Sub ListReceiptView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Call SortListView(ListReceiptView, ColumnHeader)

End Sub



Private Sub txtInputText_KeyPress(KeyAscii As Integer)
    
   If KeyAscii = 13 And txtInputText.Text <> "" Then ShowDetais


End Sub
Private Sub ShowDetais()
    Dim sql As String
    Dim retuns As String
    Dim s As Long
    
    
Select Case cmbTypeSearch.Text
    
    Case "Receipt Number"
        sql = "SELECT * FROM receipt WHERE ID = " & txtInputText.Text
        GenerateReceiptList (sql)
    Case "Invoice Number"
        sql = "SELECT * FROM receipt WHERE INV_NO = '" & txtInputText.Text & "'" & _
        " ORDER BY ID"
        GenerateReceiptList (sql)
    End Select
    
    
End Sub

Private Sub txtInputText_LostFocus()
' If txtInputText.Text <> "" Then ShowDetais
End Sub


