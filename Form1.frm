VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUserAccount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Setting"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "User Setting"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ImageList1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Toolbar1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dtePasswordLastUpdate"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkReportView"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkSystemManager"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "listActions"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "fram1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdok"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtPassword"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtLogonId"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtUserName"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtUserId"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "CboUserName"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "System Setting"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
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
         TabIndex        =   15
         Top             =   1080
         Width           =   4455
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
         TabIndex        =   14
         Top             =   1560
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
         TabIndex        =   13
         Top             =   2040
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
         TabIndex        =   12
         Top             =   2520
         Width           =   1815
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
         TabIndex        =   11
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   3960
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   6960
         Width           =   1095
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
         TabIndex        =   5
         Top             =   4200
         Width           =   3735
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
            TabIndex        =   9
            Top             =   480
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
            TabIndex        =   8
            Top             =   960
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
            TabIndex        =   7
            Top             =   1320
            Width           =   1455
         End
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
            TabIndex        =   6
            Top             =   1800
            Width           =   1575
         End
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
         TabIndex        =   3
         Top             =   3840
         Width           =   2535
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
         TabIndex        =   2
         Top             =   1560
         Width           =   2295
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
         TabIndex        =   1
         Top             =   2040
         Width           =   2295
      End
      Begin MSMask.MaskEdBox dtePasswordLastUpdate 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   3480
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
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   9540
         _ExtentX        =   16828
         _ExtentY        =   556
         ButtonWidth     =   1455
         ButtonHeight    =   503
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
         Top             =   840
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
               Picture         =   "Form1.frx":0038
               Key             =   "New"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0150
               Key             =   "Edit"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0268
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":037C
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0494
               Key             =   "Cancel"
            EndProperty
         EndProperty
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
         TabIndex        =   23
         Top             =   1080
         Width           =   1335
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
         TabIndex        =   22
         Top             =   1560
         Width           =   975
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
         TabIndex        =   21
         Top             =   2040
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
         TabIndex        =   20
         Top             =   3000
         Width           =   1215
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
         TabIndex        =   19
         Top             =   2520
         Width           =   1215
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
         TabIndex        =   18
         Top             =   3420
         Width           =   1335
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
         TabIndex        =   17
         Top             =   3600
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmUserAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub TabStrip1_Click()

End Sub






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



