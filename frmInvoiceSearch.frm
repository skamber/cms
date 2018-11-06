VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmInvoiceSearch 
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   13275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   13275
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00800000&
      FillStyle       =   6  'Cross
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   13215
      TabIndex        =   2
      Top             =   0
      Width           =   13275
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "INVOICE SEARCH"
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
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
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
      ItemData        =   "frmInvoiceSearch.frx":0000
      Left            =   840
      List            =   "frmInvoiceSearch.frx":000D
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
   Begin MSComctlLib.ListView ListInvoiceView 
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
   Begin MSComctlLib.ListView InvoiceItemList 
      Height          =   1695
      Left            =   0
      TabIndex        =   5
      Top             =   5280
      Width           =   10695
      _ExtentX        =   18865
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
End
Attribute VB_Name = "frmInvoiceSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
gRecordType = Invoice_search
End Sub

Private Sub Form_Load()
ListInvoiceView.ListItems.Clear
ListInvoiceView.ColumnHeaders.Add , , "INVOICE ID", ListInvoiceView.Width / 9
ListInvoiceView.ColumnHeaders.Add , , "INVOICE NUMBER", ListInvoiceView.Width / 9
ListInvoiceView.ColumnHeaders.Add , , "REFERENCE NUMBER", ListInvoiceView.Width / 9
ListInvoiceView.ColumnHeaders.Add , , "NAME", ListInvoiceView.Width / 6
ListInvoiceView.ColumnHeaders.Add , , "SURNAME", ListInvoiceView.Width / 6
ListInvoiceView.ColumnHeaders.Add , , "CREATED DATE", ListInvoiceView.Width / 5
ListInvoiceView.ColumnHeaders.Add , , "OVER DUE DATE", ListInvoiceView.Width / 6
ListInvoiceView.ColumnHeaders.Add , , "TOTAL AMOUNT", ListInvoiceView.Width / 7
ListInvoiceView.View = lvwReport



InvoiceItemList.ColumnHeaders.Clear
InvoiceItemList.ColumnHeaders.Add , , "ITEM NUMBER", 1400, lvwColumnLeft
InvoiceItemList.ColumnHeaders.Add , , "INVOICE NUMBER", 1400, lvwColumnLeft
InvoiceItemList.ColumnHeaders.Add , , "DESCRIPTION", 4000, lvwColumnLeft
InvoiceItemList.ColumnHeaders.Add , , "AMOUNT", 1000, lvwColumnLeft
InvoiceItemList.ColumnHeaders.Add , , "GST AMOUNT", 1300, lvwColumnLeft
InvoiceItemList.ColumnHeaders.Add , , "TOTAL AMOUNT", 1400, lvwColumnLeft
InvoiceItemList.View = lvwReport


End Sub

Private Sub ListInvoiceView_Click()
If InvoiceSelected Then
    gInvoiceId = frmInvoiceSearch.ListInvoiceView.SelectedItem
    displayInvoiceSearchItemList
End If
End Sub

Private Sub ListInvoiceView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Call SortListView(ListInvoiceView, ColumnHeader)

End Sub



Private Sub txtInputText_KeyPress(KeyAscii As Integer)
    
   If KeyAscii = 13 And txtInputText.Text <> "" Then
   frmInvoiceSearch.ListInvoiceView.ListItems.Clear
   frmInvoiceSearch.InvoiceItemList.ListItems.Clear
   ShowDetais
   End If

End Sub
Private Sub ShowDetais()
    Dim sql As String
    Dim retuns As String
    
    
Select Case cmbTypeSearch.Text
    
    Case "Invoice Number"
        sql = "SELECT * FROM invoice WHERE Invoice_No >= '" & txtInputText.Text & "' AND " & _
              " Invoice_No < '" & get_MoreThanAndLessThan(txtInputText.Text) & "'" & _
              " ORDER BY Invoice_No"
        GenerateInvoiceList (sql)
    Case "Reference Number"
        sql = "SELECT * FROM invoice WHERE Ref >= '" & txtInputText.Text & "' AND " & _
        " Ref < '" & get_MoreThanAndLessThan(txtInputText.Text) & "'" & _
        " ORDER BY Ref"
        GenerateInvoiceList (sql)
    Case "Surname"
        sql = "SELECT * FROM invoice WHERE Name2 >='" & txtInputText.Text & "' AND " & _
        " Name2 < '" & get_MoreThanAndLessThan(txtInputText.Text) & "'" & _
        " ORDER BY Name2"
        GenerateInvoiceList (sql)
    End Select
    
    
End Sub

Private Sub txtInputText_LostFocus()
' If txtInputText.Text <> "" Then ShowDetais
End Sub


