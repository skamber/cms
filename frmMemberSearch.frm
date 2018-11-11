VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMemberSearch 
   ClientHeight    =   8970
   ClientLeft      =   660
   ClientTop       =   2040
   ClientWidth     =   16515
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
   ScaleWidth      =   16515
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cmbTypeSearch1 
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
      ItemData        =   "frmMemberSearch.frx":0000
      Left            =   3840
      List            =   "frmMemberSearch.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtInputText1 
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
      Left            =   3840
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListMemberView 
      Height          =   6495
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   11456
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
      ScaleWidth      =   16455
      TabIndex        =   4
      Top             =   0
      Width           =   16515
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "MEMBER SEARCH"
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
         TabIndex        =   5
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
      Left            =   960
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
      ItemData        =   "frmMemberSearch.frx":003D
      Left            =   960
      List            =   "frmMemberSearch.frx":004D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Second Search Criteria"
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
      Left            =   3840
      TabIndex        =   8
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "First Search Criteria"
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
      Left            =   960
      TabIndex        =   7
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "frmMemberSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

End Sub

Private Sub cmdClear_Click()
 cmbTypeSearch.ListIndex = -1
 cmbTypeSearch1.ListIndex = -1
    'cmbTypeSearch.Text = ""
    'cmbTypeSearch1.Text = ""
    txtInputText.Text = ""
    txtInputText1.Text = ""
End Sub

Private Sub cmdSearch_Click()
SearchDetais
End Sub

Private Sub Form_Activate()
gRecordType = Member_Search
End Sub

Private Sub Form_Load()
ListMemberView.ListItems.Clear
ListMemberView.Width = Screen.Width - 5000
ListMemberView.Height = Screen.Height - 5000

ListMemberView.ColumnHeaders.Add , , "MEMBER NUMBER", ListMemberView.Width / 13
ListMemberView.ColumnHeaders.Add , , "NAME", ListMemberView.Width / 13
ListMemberView.ColumnHeaders.Add , , "SURNAME", ListMemberView.Width / 11
ListMemberView.ColumnHeaders.Add , , "ADDRESS1", ListMemberView.Width / 6
ListMemberView.ColumnHeaders.Add , , "ADDRESS2", ListMemberView.Width / 6
ListMemberView.ColumnHeaders.Add , , "MEMBER EXPIARY DATE", ListMemberView.Width / 7
ListMemberView.ColumnHeaders.Add , , "STATUS", ListMemberView.Width / 7
ListMemberView.ColumnHeaders.Add , , "PHONE", ListMemberView.Width / 7

ListMemberView.View = lvwReport



End Sub

Private Sub ListMemberView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Call SortListView(ListMemberView, ColumnHeader)

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
    
    sql = "SELECT * FROM member WHERE CITY_ID = " & gCityId
    
    Select Case cmbTypeSearch.Text

        Case "Name"
            sql = sql & " AND  Given_Name >= '" & txtInputText.Text & "' AND " & _
                  " Given_Name < '" & get_MoreThanAndLessThan(txtInputText.Text) & "'" & _
                  " ORDER BY Given_name"
    
        Case "Surname"
            sql = sql & " AND Surname >= '" & txtInputText.Text & "' AND " & _
            " Surname < '" & get_MoreThanAndLessThan(txtInputText.Text) & "'" & _
            " ORDER BY Surname"
    
        Case "Member Number"
            s = Val(txtInputText.Text)
            If s = 0 Then
              MsgBox "Member Number has to be Numeric Value.", vbExclamation
              Exit Sub
            End If
            sql = sql & " AND Mno >=" & txtInputText.Text & _
            " ORDER BY mno"
    
        Case "Post Code"
            s = Val(txtInputText.Text)
            If s = 0 Then
              MsgBox "Post Code has to be Numeric Value.", vbExclamation
              Exit Sub
            End If
            sql = sql & " AND postCode = '" & txtInputText.Text & "'"
        
    End Select
    GenerateMemberList (sql)
    
End Sub
Private Sub SearchDetais()
    Dim sql As String
    Dim retuns As String
    Dim s As Long
    Dim check As Boolean
    Dim linkSql As String
        
    
    sql = "SELECT * FROM member WHERE CITY_ID = " & gCityId & " AND "
    check = False
    linkSql = ""
        
    Select Case cmbTypeSearch.Text
             
         Case "Name"
            If txtInputText.Text <> "" Then
            sql = sql & " Given_Name >= '" & txtInputText.Text & "' AND " & _
                  " Given_Name < '" & get_MoreThanAndLessThan(txtInputText.Text) & "'"
            check = True
            End If
         Case "Surname"
            If txtInputText.Text <> "" Then
            sql = sql & " Surname >= '" & txtInputText.Text & "' AND " & _
            " Surname < '" & get_MoreThanAndLessThan(txtInputText.Text) & "'"
            check = True
            End If
         Case "Member Number"
            If txtInputText.Text <> "" Then
                s = Val(txtInputText.Text)
                If s = 0 Then
                    MsgBox "Member Number has to be Numeric Value.", vbExclamation
                    Exit Sub
                End If
                sql = sql & " Mno >=" & txtInputText.Text
                check = True
            End If
         Case "Post Code"
            If txtInputText.Text <> "" Then
            s = Val(txtInputText.Text)
            If s = 0 Then
                MsgBox "Post Code has to be Numeric Value.", vbExclamation
                Exit Sub
            End If
            sql = sql & " postCode = '" & txtInputText.Text & "'"
            check = True
            End If
    End Select
        
    If check = True Then
        linkSql = " and "
    End If
        
    Select Case cmbTypeSearch1.Text
             
        Case "Name"
            If txtInputText1.Text <> "" Then
            sql = sql & linkSql & " Given_Name >= '" & txtInputText1.Text & "' AND " & _
                  " Given_Name < '" & get_MoreThanAndLessThan(txtInputText1.Text) & "'"
            check = True
            End If
        Case "Surname"
            If txtInputText1.Text <> "" Then
            sql = sql & linkSql & " Surname >= '" & txtInputText1.Text & "' AND " & _
            " Surname < '" & get_MoreThanAndLessThan(txtInputText1.Text) & "'"
            check = True
            End If
        Case "Member Number"
            If txtInputText1.Text <> "" Then
            s = Val(txtInputText1.Text)
            If s = 0 Then
              MsgBox "Member Number has to be Numeric Value.", vbExclamation
              Exit Sub
            End If
            sql = sql & linkSql & " Mno >=" & txtInputText1.Text
            check = True
            End If
        Case "Post Code"
            If txtInputText1.Text <> "" Then
            s = Val(txtInputText1.Text)
            If s = 0 Then
                MsgBox "Post Code has to be Numeric Value.", vbExclamation
                Exit Sub
            End If
            sql = sql & linkSql & " postCode = '" & txtInputText1.Text & "'"
            check = True
            End If
    End Select
        
    If check = True Then
       GenerateMemberList (sql)
    End If
        
    
End Sub

Private Sub txtInputText_LostFocus()
' If txtInputText.Text <> "" Then ShowDetais
End Sub




