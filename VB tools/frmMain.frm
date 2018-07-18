VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sort ListView"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton cmdWhatsSel 
      Caption         =   "What Is Selected?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   2295
   End
   Begin MSComctlLib.ListView lvwTest 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImgSorted"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Test"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgSorted 
      Left            =   4440
      Top             =   -120
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
            Picture         =   "frmMain.frx":0000
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":059A
            Key             =   "down"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const NUM_STRING = "One,Two,Three,Four,Five,Six,Seven,Eight,Nine,Ten,Eleven,Twelve,Thirteen,Fourteen,Fifteen,Sixteen,Seventeen,Eighteen,Nineteen"


Private Sub LoadList()
    
    Dim iCtr        As Integer
    Dim strNames()  As String
    Dim lstItem     As ListItem
    
    strNames = Split(NUM_STRING, ",")
       
    For iCtr = 1 To 19
        Set lstItem = lvwTest.ListItems.Add(, "K" & iCtr, iCtr)
        lstItem.SubItems(1) = strNames(iCtr - 1)
        lstItem.SubItems(2) = Format$(DateAdd("d", iCtr - 1, Now), DATE_FORMAT)
    Next iCtr
    lvwTest.ColumnHeaders(1).Icon = "down"
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdWhatsSel_Click()
    With lvwTest.SelectedItem
        MsgBox .Text & vbCrLf & .SubItems(1) & vbCrLf & .SubItems(2)
    End With
End Sub

Private Sub Form_Load()
    LoadList
End Sub


Private Sub lvwTest_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call ClearHeaderIcons(ColumnHeader.Index)
    Select Case ColumnHeader.Index
        Case 1
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(lvwTest, ColumnHeader.Index, sortDescending, sortNumeric)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvwTest, ColumnHeader.Index, sortAscending, sortNumeric)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvwTest, ColumnHeader.Index, sortAscending, sortNumeric)
            End Select
        Case 2
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(lvwTest, ColumnHeader.Index, sortDescending, sortAlpha)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvwTest, ColumnHeader.Index, sortAscending, sortAlpha)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvwTest, ColumnHeader.Index, sortAscending, sortAlpha)
            End Select
        Case 3
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(lvwTest, ColumnHeader.Index, sortDescending, sortDate)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvwTest, ColumnHeader.Index, sortAscending, sortDate)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvwTest, ColumnHeader.Index, sortAscending, sortDate)
            End Select
    End Select
End Sub
Private Sub ClearHeaderIcons(CurrentHeader As Integer)
    Dim i As Integer
    For i = 1 To lvwTest.ColumnHeaders.Count
        If lvwTest.ColumnHeaders(i).Index <> CurrentHeader Then
            lvwTest.ColumnHeaders(i).Icon = Empty
        End If
    Next
End Sub
