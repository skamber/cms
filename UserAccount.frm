VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Account"
   ClientHeight    =   8400
   ClientLeft      =   2295
   ClientTop       =   2595
   ClientWidth     =   7560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   7560
   Begin VB.ComboBox cmbChurch 
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
      ItemData        =   "UserAccount.frx":0000
      Left            =   1560
      List            =   "UserAccount.frx":0002
      TabIndex        =   24
      Top             =   4200
      Width           =   2655
   End
   Begin VB.ComboBox cmbCity 
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
      ItemData        =   "UserAccount.frx":0004
      Left            =   1560
      List            =   "UserAccount.frx":0006
      TabIndex        =   23
      Text            =   "cmbCity"
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CheckBox chkReportView 
      Caption         =   "REPORT VIEW"
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
      Left            =   4680
      TabIndex        =   22
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CheckBox chkSystemManager 
      Caption         =   "SYSTEM MANAGER"
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
      Left            =   4680
      TabIndex        =   21
      Top             =   1200
      Width           =   2295
   End
   Begin VB.ListBox listActions 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   4560
      TabIndex        =   15
      Top             =   4680
      Width           =   2535
   End
   Begin MSMask.MaskEdBox dtePasswordLastUpdate 
      Height          =   315
      Left            =   1560
      TabIndex        =   14
      Top             =   3120
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.Frame fram1 
      Caption         =   "Permissions"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   480
      TabIndex        =   12
      Top             =   5040
      Width           =   3735
      Begin VB.CheckBox chkDelete 
         Caption         =   "Delete Data"
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
         Left            =   240
         TabIndex        =   19
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CheckBox chkInsert 
         Caption         =   "Insert Data"
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
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox chkUpdate 
         Caption         =   "Update Data"
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
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkRead 
         Caption         =   "Read Data"
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
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7680
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
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Top             =   2640
      Width           =   1815
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
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtUserName 
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
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtUserId 
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
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ComboBox CboUserName 
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
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   4455
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   609
      ButtonWidth     =   1561
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Object.ToolTipText     =   "Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel"
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6720
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserAccount.frx":0008
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserAccount.frx":0120
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserAccount.frx":0238
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserAccount.frx":034C
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserAccount.frx":0464
            Key             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label9 
      Caption         =   "Church"
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
      Left            =   120
      TabIndex        =   26
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "City"
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
      Left            =   120
      TabIndex        =   25
      Top             =   3600
      Width           =   1350
   End
   Begin VB.Label Label7 
      Caption         =   "Object Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   20
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Password Last Update"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   13
      Top             =   3060
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "LOGON ID"
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
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "PASSWORD"
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
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "USER NAME"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "USER ID"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "USER NAME"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmUserAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Sub CboUserName_Click()
On Error GoTo ErrorHandler

    If CboUserName.ListIndex = -1 Then Exit Sub
    gUserAccountId = CboUserName.ItemData(CboUserName.ListIndex)
    Call InitialisePermission
    gUserAccountName = CboUserName.Text
    DisplayUserAccount
Exit Sub
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "frmUserAccount", "CboUserName_Click", True)
            
End Sub

Private Sub chkDelete_Click()
    do_permissioms (4)
End Sub

Private Sub chkInsert_Click()
    do_permissioms (2)
End Sub

Private Sub chkRead_Click()
    do_permissioms (1)
End Sub

Private Sub chkUpdate_Click()
    do_permissioms (3)
End Sub



Private Sub cmbCity_GotFocus()
  If cmbCity.ListIndex >= 0 Then
   Dim sCity As Integer
    sCity = cmbCity.ItemData(cmbCity.ListIndex)
    LoadChurchComboBoxWithAll (sCity)
  End If
End Sub

Private Sub cmdOK_Click()
Dim iAnswer As Integer
If gRecordUserType = RECORD_NEW Or gRecordUserType = RECORD_EDIT Then
    iAnswer = MsgBox("Are you sure you want to Exit without saving this record?", vbExclamation + vbYesNo)
        If iAnswer = vbNo Then
            Exit Sub
        Else: Unload frmUserAccount
        End If
Else: Unload frmUserAccount
End If
End Sub



Private Sub Form_Load()
 On Error GoTo ErrorHandler

LoadUserAccountComboBox
LoadCityNames
'InitialiseUserAccount
gRecordUserType = RECORD_READ
SetToolbarUserAccountControl
UserAccountActivate
'frmUserAccount.Enabled = False
Exit Sub
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modUserAccount", "LoadForm", True)
End Sub

Private Sub listActions_Click()
doActionPermisions
End Sub

Private Sub listActions_ItemCheck(Item As Integer)

'doActionPermisions (Item)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Call ProcessUserAccountRecord(Button)
End Sub


