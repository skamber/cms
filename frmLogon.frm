VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CMS"
   ClientHeight    =   4365
   ClientLeft      =   1590
   ClientTop       =   2370
   ClientWidth     =   6735
   Icon            =   "frmLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6735
   Begin VB.ComboBox cmbCityName 
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
      ItemData        =   "frmLogon.frx":030A
      Left            =   3240
      List            =   "frmLogon.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2520
      Width           =   2655
   End
   Begin VB.ComboBox cmbChurchName 
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
      ItemData        =   "frmLogon.frx":030E
      Left            =   3240
      List            =   "frmLogon.frx":0310
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox txtLogonId 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Church "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Management System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Login ID "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1590
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CMS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   5000
      Left            =   0
      Picture         =   "frmLogon.frx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7000
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbCityName_Click()
   If cmbCityName.ListIndex >= 0 Then
      gCityId = cmbCityName.ItemData(cmbCityName.ListIndex)
      LoadChurchComboBox
   End If
End Sub

Private Sub cmdEnter_Click()
    Dim sLogonID As String
    Dim sLogonPassword As String
    Dim dtePasswordLastUpdate As String
    
    If Not validateApplicationVersion Then
        Exit Sub
    End If
    
    sLogonID = UCase(Trim(txtLogonId.Text))
    sLogonPassword = UCase(Trim(txtPassword.Text))
    sChurchName = UCase(Trim(cmbChurchName.Text))
    If Not ValidateLogon(sLogonID, sLogonPassword, sChurchName) Then
        Exit Sub
    End If

    'gChurchId = cmbChurchName.ListIndex
    gChurchId = cmbChurchName.ItemData(cmbChurchName.ListIndex)
    
    If Not CheckLogonId(sLogonID, sLogonPassword) Then Exit Sub
    If Not CheckPasswordChange Then
        frmPassword.Show vbModal
    End If

    
    Unload frmLogon
    frmMember.Show
    
End Sub


Private Sub cmdExit_Click()
End
End Sub

Private Sub Form_Load()
Call ConnectDatabase

LoadCityNamesComboBox
'LoadChurchComboBox

'If cmbChurchName.ListCount > 0 Then
'    cmbChurchName.
'    cmbChurchName.ListIndex = 0
'End If
Call CentreForm(Me, 2)

End Sub

Private Sub txtLogonId_GotFocus()
HighlightText Me
End Sub

Private Sub txtPassword_GotFocus()
HighlightText Me
End Sub
