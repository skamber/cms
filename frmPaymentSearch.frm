VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPaymentSearch 
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   12885
   WindowState     =   2  'Maximized
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
      Left            =   4080
      TabIndex        =   1
      Top             =   720
      Width           =   2055
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
      Left            =   1320
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
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
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtMno 
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
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00800000&
      FillStyle       =   6  'Cross
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   12825
      TabIndex        =   2
      Top             =   0
      Width           =   12885
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "PAYMENT SEARCH"
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
         Width           =   2775
      End
   End
   Begin MSComctlLib.ListView ListPaymentView 
      Height          =   3855
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImgSorted"
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
   Begin MSComctlLib.ImageList ImgSorted 
      Left            =   10320
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPaymentSearch.frx":0000
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPaymentSearch.frx":059A
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
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
      Height          =   200
      Left            =   3480
      TabIndex        =   9
      Top             =   780
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
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
      Height          =   200
      Left            =   0
      TabIndex        =   8
      Top             =   1710
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Height          =   200
      Left            =   240
      TabIndex        =   7
      Top             =   1230
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Mno"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   240
      TabIndex        =   6
      Top             =   750
      Width           =   855
   End
End
Attribute VB_Name = "frmPaymentSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmbType_Click()
If txtMno.Text <> "" Then GeneratePaymentList (txtMno.Text)
End Sub

Private Sub Form_Activate()
gRecordType = Payment_search
End Sub

Private Sub Form_Load()
ListPaymentView.ListItems.Clear

ListPaymentView.Width = Screen.Width - 5000
ListPaymentView.Height = Screen.Height - 5000


'ListPaymentView.Width = 11700
ListPaymentView.ColumnHeaders.Add , , "MEMBER NUMBER", ListPaymentView.Width / 10
ListPaymentView.ColumnHeaders.Add , , "PAYMENT", ListPaymentView.Width / 7
ListPaymentView.ColumnHeaders.Add , , "MEMBER EFFECTIVE", ListPaymentView.Width / 6
ListPaymentView.ColumnHeaders.Add , , "MEMBER EXPIARY", ListPaymentView.Width / 6
ListPaymentView.ColumnHeaders.Add , , "DATE OF PAYMENT", ListPaymentView.Width / 6
ListPaymentView.ColumnHeaders.Add , , "AMOUNT", ListPaymentView.Width / 7
ListPaymentView.ColumnHeaders.Add , , "RECEIPT NO", ListPaymentView.Width / 10
ListPaymentView.View = lvwReport
ListPaymentView.ColumnHeaders(1).Icon = "down"
LoadPaymnetComboBox

End Sub

Private Sub ListPaymentView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'Call SortListView(ListPaymentView, ColumnHeader)
    Call ClearHeaderIcons(ColumnHeader.index)
    Select Case ColumnHeader.index
        Case 1
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortDescending, sortNumeric)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortAscending, sortNumeric)
               Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortAscending, sortNumeric)
            End Select
        Case 2
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortDescending, sortAlpha)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortAscending, sortAlpha)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortAscending, sortAlpha)
            End Select
        Case 3
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortDescending, sortDate)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortAscending, sortDate)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortAscending, sortDate)
            End Select
    
        Case 4
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortDescending, sortDate)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortAscending, sortDate)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortAscending, sortDate)
            End Select
        Case 5
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortDescending, sortDate)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortAscending, sortDate)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortAscending, sortDate)
            End Select
        Case 6
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortDescending, sortNumeric)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortAscending, sortNumeric)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortAscending, sortNumeric)
            End Select
        Case 7
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortDescending, sortNumeric)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortAscending, sortNumeric)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(ListPaymentView, ColumnHeader.index, sortAscending, sortNumeric)
            End Select
    End Select

End Sub

Private Sub ClearHeaderIcons(CurrentHeader As Integer)
    Dim i As Integer
    For i = 1 To ListPaymentView.ColumnHeaders.Count
        If ListPaymentView.ColumnHeaders(i).index <> CurrentHeader Then
            ListPaymentView.ColumnHeaders(i).Icon = Empty
        End If
    Next
End Sub


Private Sub txtmno_KeyPress(KeyAscii As Integer)
 
 
    If KeyAscii = 13 And txtMno.Text <> "" Then
      If GenerateMemberInfo(txtMno.Text) Then
        GeneratePaymentList (txtMno.Text)
      End If
    End If
    
End Sub




Private Sub txtMno_LostFocus()
If txtMno.Text <> "" Then
     ' GenerateMemberInfo (txtMno.Text)
     ' GeneratePaymentList (txtMno.Text)
    End If
End Sub
