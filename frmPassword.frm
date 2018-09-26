VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   3540
   ClientLeft      =   2940
   ClientTop       =   4200
   ClientWidth     =   6765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6765
   Begin VB.Frame fraFrame 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   2760
         TabIndex        =   11
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5160
         TabIndex        =   4
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtNewPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtOldPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox txtConfirmPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2160
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "New Password:"
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
         Index           =   1
         Left            =   315
         TabIndex        =   10
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password:"
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
         Index           =   0
         Left            =   315
         TabIndex        =   9
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm New Password:"
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
         Index           =   2
         Left            =   315
         TabIndex        =   8
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Logon Password expired"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2295
         TabIndex        =   7
         Top             =   240
         Width           =   2595
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Please enter password details then click OK."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   1080
         TabIndex        =   6
         Top             =   600
         Width           =   4635
      End
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    If Not CompulsoryChangePassword Then
        Unload frmPassword
    Else
      MsgBox "you have to change the password.", vbInformation
    txtNewPassword.SetFocus
   ' frmLogon.Show
    End If
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdOK_Click()

    Dim strOldPassword As String
    Dim strNewPassword As String
    Dim strConfirmPassword As String
    Dim sql As String
    Dim encryptedpass As String

    Dim objRecordset As Recordset
    

    strOldPassword = UCase(Trim(txtOldPassword.Text))
    strNewPassword = UCase(Trim(txtNewPassword.Text))
    strConfirmPassword = UCase(Trim(txtConfirmPassword.Text))
    
    If (strOldPassword = "") Or (strNewPassword = "") Or (strConfirmPassword = "") Then
        MsgBox "Required field missing - All Password details must be entered.", vbExclamation
        txtOldPassword.SetFocus
        Exit Sub
    End If
    
        sql = "SELECT * FROM Users WHERE Id = " & UserId
        
        Set objRecordset = New ADODB.Recordset
            objRecordset.Open sql, objConnection, adOpenForwardOnly, adLockOptimistic
    
        If objRecordset.EOF = True Then
            MsgBox "Logon is invalid.  Access has been denied.", vbExclamation
            Exit Sub
        End If
         
        If objRecordset!Logon_Password = strOldPassword Then
         ElseIf DecryptPassword(objRecordset!Logon_Password) <> strOldPassword Then
            MsgBox "Password details entered are invalid.  Password update unsuccessful.", vbInformation
            txtOldPassword.SetFocus
            Exit Sub
        End If
         If strOldPassword = strNewPassword Then
            MsgBox "You have used this password before. Please choose a new one", vbInformation
            txtNewPassword.SetFocus
            Exit Sub
          End If
         If Len(strNewPassword) < 5 Then
         MsgBox "you must supply the minimum number of characters required for a password.", vbInformation
            txtNewPassword.SetFocus
            Exit Sub
         End If
         If Len(strNewPassword) > 12 Then
         MsgBox "you must supply the maximum number of characters required for a password.", vbInformation
            txtNewPassword.SetFocus
            Exit Sub
         End If
         If strNewPassword <> strConfirmPassword Then
            MsgBox "Password details entered are invalid.  Password update unsuccessful.", vbInformation
            txtNewPassword.SetFocus
            Exit Sub
       End If
       encryptedpass = encryptPassword(strNewPassword)
       encryptedpass = Replace(encryptedpass, "'", "''")
        objConnection.BeginTrans
                
                sql = "UPDATE Users SET" _
                            & " Logon_Password =" & "'" & encryptedpass & "'" _
                            & " ,Password_Last_Update =" & "'" & Format(Date, "dd-mmm-yyyy") & "'"
                           
                            
                            'Oracle and Access use different keyword for the sysdate
                           
                            
                sql = sql & " WHERE Id = " & UserId
    
                objConnection.Execute sql

        objConnection.CommitTrans

        MDIFrm.pnlStatusBar.Panels(1).Text = "Password information change successful."
        CompulsoryChangePassword = False
        Unload frmPassword
        

exitProcedure:
    Set objRecordset = Nothing

End Sub

Private Sub Form_Load()
    Call CentreForm(frmPassword, 2)
    
End Sub


