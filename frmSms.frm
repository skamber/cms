VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSms 
   ClientHeight    =   8970
   ClientLeft      =   660
   ClientTop       =   2040
   ClientWidth     =   15300
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
   ScaleHeight     =   8970
   ScaleWidth      =   15300
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton CmdSendSms 
      Caption         =   "Send SMS"
      Height          =   375
      Left            =   11160
      TabIndex        =   2
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox txtSmsMessage 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   7080
      Width           =   10335
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
      ScaleWidth      =   15240
      TabIndex        =   5
      Top             =   0
      Width           =   15300
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "SMS MEMBER"
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
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   3015
      End
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
      ItemData        =   "frmSms.frx":0000
      Left            =   840
      List            =   "frmSms.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   7646
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
   Begin MSMask.MaskEdBox dteExpairyDate 
      Height          =   315
      Left            =   840
      TabIndex        =   10
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
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
   Begin VB.Label Label5 
      Caption         =   "Message To Send"
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
      TabIndex        =   8
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "List To Send"
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
      TabIndex        =   7
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Search Criteria"
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
      Left            =   840
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "frmSms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
cmbTypeSearch.ListIndex = -1
    dteExpairyDate.Text = ""
End Sub

Private Sub cmbTypeSearch_Click()
dteExpairyDate.Text = ""
End Sub

Private Sub cmdSearch_Click()
ShowDetais
End Sub


Private Sub CmdSendSms_Click()

Dim Count, i, j As Integer

'ListView.ListItem(1).Selected = True
'ListView.ListItem.EnsureVisible = True

Count = Me.ListView.ListItems.Count  ' count of items (number of rows in the list view
Dim strFriends(0 To 10) As String
For i = 0 To Count
'For j = 0 To 1  'set j value from 0 to (countof columns-1).  example listview has 2 columns

MsgBox ListView.ListItems(i).Text

Dim value As String
' strFriends(i) = ListView.ListItems(i).SubItems(8).Text
 
'Next j
Next i

End Sub

Private Sub Form_Activate()
gRecordType = Send_Sms
End Sub

Private Sub Form_Load()
ListView.ListItems.Clear
ListView.Width = Screen.Width - 5000
'ListView.Height = Screen.Height - 5000

ListView.ColumnHeaders.Add , , "NUMBER", ListView.Width / 13
ListView.ColumnHeaders.Add , , "NAME", ListView.Width / 13
ListView.ColumnHeaders.Add , , "SURNAME", ListView.Width / 11
ListView.ColumnHeaders.Add , , "ADDRESS1", ListView.Width / 6
ListView.ColumnHeaders.Add , , "ADDRESS2", ListView.Width / 6
ListView.ColumnHeaders.Add , , "MEMBER EXPIARY DATE", ListView.Width / 7
ListView.ColumnHeaders.Add , , "STATUS", ListView.Width / 7
ListView.ColumnHeaders.Add , , "PHONE", ListView.Width / 7
ListView.ColumnHeaders.Add , , "MOBILE", ListView.Width / 7
ListView.ColumnHeaders.Add , , "EMAIL", ListView.Width / 5

ListView.View = lvwReport

End Sub


Private Sub ListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Call SortListView(ListView, ColumnHeader)

End Sub


Private Sub txtInputText_KeyPress(KeyAscii As Integer)
    
   If KeyAscii = 13 And txtInputText.Text <> "" And cmbTypeSearch.Text <> "" Then
    txtInputText1.Text = ""
    cmbTypeSearch1.ListIndex = -1
    ShowDetais
   End If

End Sub
Private Sub ShowDetais()
    Dim sql As String
    Dim retuns As String
    Dim s As Long
    Dim frm As Object
    Set frm = Me
    Dim strMemberExpiartDate As String
    
    Select Case cmbTypeSearch.Text
        Case "All Member"
            sql = "SELECT * FROM member WHERE CITY_ID = " & gCityId
            sql = sql & " AND  Status ='ACTIVE' AND Mobile is not null"
            Call GenerateMemberList(sql, frm)
        Case "Membership Expiry"
            sql = "SELECT * FROM member WHERE CITY_ID = " & gCityId
            sql = sql & " AND  Status ='ACTIVE' AND Mobile is not null"
            strMemberExpiartDate = "Date('" & Format(dteExpairyDate.FormattedText, "yyyy,mm,dd") & "')"
            sql = sql & " AND MEMBERSHIP_EXPIARY <= " & strMemberExpiartDate
            Call GenerateMemberList(sql, frm)
        Case "All Youth"
            sql = "SELECT * FROM children WHERE CITY_ID = " & gCityId
            sql = sql & " AND MEMBER = 'Y' AND Mobile is not null"
        Call GenerateChildrenList(sql, frm)
    End Select

End Sub

