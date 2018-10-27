VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmReport 
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   13980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   13980
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00800000&
      FillStyle       =   6  'Cross
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   13920
      TabIndex        =   20
      Top             =   0
      Width           =   13980
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   " REPORTING"
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
         TabIndex        =   21
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   13335
      Begin VB.OptionButton optReport 
         Caption         =   "Receipts Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   19
         Top             =   2160
         Width           =   4215
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Invoices Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   18
         Top             =   1800
         Width           =   4215
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Members Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   720
         Width           =   4215
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Payments, Collections and Receipts Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   4935
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Children Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   13
         Top             =   1440
         Width           =   4215
      End
   End
   Begin VB.Frame fraCriteria 
      Caption         =   "Selection Criteria"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3405
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   13335
      Begin VB.ComboBox cboReportCriteria3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmReport.frx":0000
         Left            =   5640
         List            =   "frmReport.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtMemberRno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   5880
         TabIndex        =   17
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print"
         Height          =   375
         Left            =   11775
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "View"
         Height          =   375
         Left            =   10485
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1095
      End
      Begin VB.ComboBox cboReportCriteria2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmReport.frx":0004
         Left            =   5640
         List            =   "frmReport.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.ComboBox cboReportCriteria1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmReport.frx":0008
         Left            =   5640
         List            =   "frmReport.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSMask.MaskEdBox dteStartDate 
         Height          =   315
         Left            =   5880
         TabIndex        =   6
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox dteEndDate 
         Height          =   315
         Left            =   8760
         TabIndex        =   7
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label lblSelectionLabel3 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   22
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblSelectionLabel7 
         Alignment       =   1  'Right Justify
         Caption         =   "Member Rno :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   16
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lblSelectionLabel2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblSelectionLabel6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7440
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblSelectionLabel5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4320
         TabIndex        =   9
         Top             =   2040
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label lblSelectionLabel1 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin Crystal.CrystalReport Report 
      Left            =   2520
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowBorderStyle=   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   0   'False
      WindowMinButton =   0   'False
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngReportId As Long
Public objReportConnection As ADODB.Connection



Private Sub CmdPrint_Click()
If ValidateReportCriteria(lngReportId) Then
        Call GenerateReport(lngReportId, "Print", objReportConnection)
    End If
End Sub

Private Sub CmdView_Click()
Report.WindowWidth = Screen.Width
    Report.WindowHeight = Screen.Width
    
    If ValidateReportCriteria(lngReportId) Then
        Call GenerateReport(lngReportId, "View", objReportConnection)
    End If
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler

    'setup mdi controls
    With MDIFrm
    
            .pnlStatusBar.Panels(1).Text = "Loading information for Reporting..."

        
    End With

    Set objReportConnection = New ADODB.Connection

    'Connect to database for reporting...
   
        'objReportConnection = objConnection & ";PWD=" & gDatabasePassword
        objReportConnection = objConnection
    Report.Connect = objReportConnection
    
    MDIFrm.pnlStatusBar.Panels(1).Text = ""
    
    
Exit Sub
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "frmReport", "Form_Load", True)

End Sub


Private Sub optReport_Click(Index As Integer)
     ReportCriteria_Reset
    lngReportId = Index
    If Not CheckReportSecurity Then Exit Sub
    Call ReportCriteria_1(lngReportId)
    
    DoEvents
End Sub
