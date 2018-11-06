VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCollectionSearch 
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   17715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   17715
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00800000&
      FillStyle       =   6  'Cross
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   17655
      TabIndex        =   2
      Top             =   0
      Width           =   17715
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Collection SEARCH"
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
      Top             =   1440
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
      ItemData        =   "frmCollectionSearch.frx":0000
      Left            =   840
      List            =   "frmCollectionSearch.frx":000A
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
   Begin MSComctlLib.ListView ListCollectionView 
      Height          =   5055
      Left            =   0
      TabIndex        =   4
      Top             =   2520
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   8916
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
   Begin MSMask.MaskEdBox dteInputText 
      Height          =   315
      Left            =   840
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
End
Attribute VB_Name = "frmCollectionSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmbTypeSearch_Click()
Select Case cmbTypeSearch.Text
    
    Case "Collection Number"
            txtInputText.Visible = True
            dteInputText.Visible = False
    Case "Collection Date"
            txtInputText.Visible = False
            dteInputText.Visible = True
            dteInputText.Top = 1440
    End Select

End Sub



Private Sub dteInputText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And dteInputText.Text <> "" Then ShowDetais
End Sub

Private Sub Form_Activate()
gRecordType = collection_search
End Sub

Private Sub Form_Load()
ListCollectionView.ListItems.Clear
ListCollectionView.Width = Screen.Width - 5000
ListCollectionView.Height = Screen.Height - 5000

ListCollectionView.ColumnHeaders.Add , , "COLLECTION NUMBER", ListCollectionView.Width / 7
ListCollectionView.ColumnHeaders.Add , , "COLLECTION TYPE", ListCollectionView.Width / 7
ListCollectionView.ColumnHeaders.Add , , "DATE OF COLLECTION", ListCollectionView.Width / 4
ListCollectionView.ColumnHeaders.Add , , "AMOUNT", ListCollectionView.Width / 4
ListCollectionView.ColumnHeaders.Add , , "COMMENTS", ListCollectionView.Width / 4
ListCollectionView.View = lvwReport
txtInputText.Visible = False
dteInputText.Visible = False

End Sub

Private Sub ListCollectionView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Call SortListView(ListCollectionView, ColumnHeader)

End Sub



Private Sub txtInputText_KeyPress(KeyAscii As Integer)
    
   If KeyAscii = 13 And txtInputText.Text <> "" Then ShowDetais


End Sub
Private Sub ShowDetais()
    Dim sql As String
    Dim retuns As String
    Dim s As Long
     Dim strCollectionDate As String
    
Select Case cmbTypeSearch.Text
    
    Case "Collection Number"
        sql = "SELECT * FROM collection WHERE COL_ID = " & txtInputText.Text
        GenerateCollectionList (sql)
    Case "Collection Date"
        strCollectionDate = Format(dteInputText.FormattedText, "yyyy-mm-dd")

        sql = "SELECT * FROM collection WHERE Collection.DATE_OF_COLLECTION = '" & strCollectionDate & _
        "' ORDER BY PAYMENT"
        
        GenerateCollectionList (sql)
    End Select
 
    
End Sub

Private Sub txtInputText_LostFocus()
' If txtInputText.Text <> "" Then ShowDetais
End Sub



