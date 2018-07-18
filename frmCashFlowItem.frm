VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCashFlowItem 
   ClientHeight    =   8760
   ClientLeft      =   615
   ClientTop       =   2235
   ClientWidth     =   12135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   12135
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00800000&
      FillStyle       =   6  'Cross
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   12075
      TabIndex        =   11
      Top             =   0
      Width           =   12135
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "ITEM"
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
         Left            =   240
         TabIndex        =   12
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.Frame fraCashfloeItem 
      BorderStyle     =   0  'None
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   12135
      Begin MSComctlLib.ListView ListCashflowItemView 
         Height          =   2415
         Left            =   480
         TabIndex        =   13
         Top             =   2640
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4260
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
      Begin VB.TextBox txtId 
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
         Left            =   2760
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtGSTRate 
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
         Left            =   5040
         TabIndex        =   4
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtItemCode 
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
         Left            =   3120
         TabIndex        =   3
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtItemName 
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
         Left            =   360
         TabIndex        =   2
         Top             =   1680
         Width           =   2655
      End
      Begin VB.ComboBox cmbItemType 
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
         ItemData        =   "frmCashFlowItem.frx":0000
         Left            =   360
         List            =   "frmCashFlowItem.frx":000A
         TabIndex        =   1
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Item Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "ItemType"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "GST Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5040
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Item Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   3120
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Item Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCashFlowItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text3_Change()

End Sub
Public Sub Form_Activate()

    If gRecordMode = RECORD_NEW Or gRecordMode = RECORD_EDIT Then
        fraCashfloeItem.Enabled = True
          
    Else
        fraCashfloeItem.Enabled = False
        
    End If
 gRecordType = CASHFLOWITEM
End Sub
Private Sub Form_Load()

On Error GoTo ErrorHandler
gRecordMode = RECORD_READ

ListCashflowItemView.ListItems.Clear
ListCashflowItemView.ColumnHeaders.Add , , "ITEM ID", ListCashflowItemView.Width / 9
ListCashflowItemView.ColumnHeaders.Add , , "ITEM NAME", ListCashflowItemView.Width / 3
ListCashflowItemView.ColumnHeaders.Add , , "ITEM CODE", ListCashflowItemView.Width / 5
ListCashflowItemView.ColumnHeaders.Add , , "GST", ListCashflowItemView.Width / 10
ListCashflowItemView.ColumnHeaders.Add , , "ITEM TYPE", ListCashflowItemView.Width / 5
ListCashflowItemView.View = lvwReport
With MDIFrm
Call LoadCashflowItemList

     
            ' .pnlStatusBar.Panels(1).Text = "Loading information for Member..."
             
                  
     End With
     DoEvents
               
     'set record control variables
     gRecordType = CASHFLOWITEM
     gRecordMode = RECORD_READ
     SetToolbarControl
     MDIFrm.pnlStatusBar.Panels(1).Text = ""
     DoEvents
 
 
 Exit Sub
ErrorHandler:
     Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "frmCashflowitem", "Form_Load", True)

End Sub


Private Sub ListCashflowItemView_Click()
gCashflowItemId = frmCashFlowItem.ListCashflowItemView.SelectedItem
Call DispayCashflowItem
End Sub

Private Sub txtGSTRate_LostFocus()
 Call ValidNumericEntry(txtGSTRate)
End Sub
