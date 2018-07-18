VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStartupStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Application - Update"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmStartupStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5400
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar pbStatus 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   5750
      _ExtentX        =   10134
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblStatus 
      Caption         =   "Preparing Update..."
      Height          =   495
      Left            =   195
      TabIndex        =   2
      Top             =   720
      Width           =   5730
   End
   Begin VB.Label Label1 
      Caption         =   " Update in progress..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmStartupStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.Caption = APPLICATION_NAME & " - Update"
    
End Sub

Private Sub Timer1_Timer()

    Static intCnt As Integer
    
    intCnt = (intCnt + 1)
    
    If intCnt = 11 Then
        Unload Me
    Else
        pbStatus.Value = intCnt
        lblStatus = "Starting in " & Format(11 - intCnt) & " seconds..."
    End If
    
End Sub


