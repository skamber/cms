Attribute VB_Name = "Modlogon"
Option Explicit
Private dtePasswordLastUpdate As Date
Private CheckPasword As String
Public Function ValidateLogon(sLogonID, sLogonPassword, sChurchName) As Boolean
On Error GoTo ErrorHandler

    ValidateLogon = False

    If (sLogonID = "") Or (sLogonPassword = "") Or (sChurchName = "") Then
        MsgBox "Required field missing - Logon ID, Password and Church name must be entered.", vbExclamation
        frmLogon.txtLogonId.SetFocus
        Exit Function
    End If
    
    ValidateLogon = True
    
    
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modLogon", "CheckPasswordChange", True)

End Function

Public Function LoadCityNamesComboBox()
On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rslocal As ADODB.Recordset
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection

    Set rslocal = objOrganisation_s.getCities()

    With frmLogon
            
            .cmbCityName.Clear
            If Not rslocal Is Nothing Then
                Do Until rslocal.EOF
                    .cmbCityName.AddItem rslocal!cityName
                    .cmbCityName.ItemData(.cmbCityName.NewIndex) = rslocal!id
                    Dim cityName As String
                    Dim cityId As String
                    cityName = rslocal!cityName
                    cityId = rslocal!id
                    Cities.Add cityName, cityId
                    
                    rslocal.MoveNext
                Loop
                Set rslocal = Nothing
            End If
    End With
    Set objOrganisation_s = Nothing


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modlogon", "LoadChurchComboBox", True)

End Function

Public Function LoadChurchComboBox()
On Error GoTo ErrorHandler

    Dim objOrganisation_s As CMSOrganisation.clsOrganisation
    Dim rslocal As ADODB.Recordset
    Set objOrganisation_s = New CMSOrganisation.clsOrganisation
    Set objOrganisation_s.DatabaseConnection = objConnection

       
     
    Set rslocal = objOrganisation_s.getChurchName(gCityId)

    With frmLogon
            
            .cmbChurchName.Clear
            If Not rslocal Is Nothing Then
                Do Until rslocal.EOF
                    .cmbChurchName.AddItem rslocal!Name
                    .cmbChurchName.ItemData(.cmbChurchName.NewIndex) = rslocal!id
                    rslocal.MoveNext
                Loop
                Set rslocal = Nothing
            End If
    End With
    Set objOrganisation_s = Nothing

Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modlogon", "LoadChurchComboBox", True)

End Function

Public Function validateApplicationVersion()
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rslocal As ADODB.Recordset

     validateApplicationVersion = False
         
    sql = "SELECT * FROM app WHERE version=  '" & VERSION & "'"
    Set rslocal = New ADODB.Recordset
        rslocal.Open sql, objConnection, adOpenForwardOnly, adLockReadOnly

    If rslocal.EOF = True Then
        MsgBox "Invalid application version number " & VERSION & ".  Please contact the systems administration.", vbExclamation
        frmLogon.txtLogonId.SetFocus
        Exit Function
    End If
    validateApplicationVersion = True
    
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modLogon", "CheckLogonId", True)

End Function

Public Function CheckLogonId(ByVal strLogonID As String, ByVal strLogonPassword As String)
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rslocal As ADODB.Recordset
    Dim checkSystemManager As String
    Dim checkReportView As String
    
    Dim Ctr As Integer
    CheckLogonId = False
         
    sql = "SELECT * FROM users WHERE Logon_ID = '" & strLogonID & "'" & " AND city_Id =" & gCityId
    
    Set rslocal = New ADODB.Recordset
        rslocal.Open sql, objConnection, adOpenForwardOnly, adLockReadOnly

    If rslocal.EOF = True Then
        MsgBox "Logon is invalid.  Access has been denied.", vbExclamation
        frmLogon.txtLogonId.SetFocus
        Exit Function
    End If
    If rslocal!Logon_Password = "LOCKED" Then
    MsgBox "You have been locked. Please contact System Administrator.", vbExclamation
    End
    End If
    If (rslocal!Logon_Password = "WELCOME") And (strLogonPassword = "WELCOME") Then
        
    Else
      If cmdDecrypt(rslocal!Logon_Password) <> strLogonPassword Then
        MsgBox "Logon Password is invalid.  Access has been denied.", vbExclamation
        frmLogon.txtPassword.SetFocus
        NumLogIN = NumLogIN + 1
        If NumLogIN = 3 Then
        LockUser (strLogonID)
        End
        End If
        Exit Function
    End If
    End If
    
    gChurchRestriction = rslocal!Church_Id
    If gChurchRestriction <> 0 And gChurchRestriction <> gChurchId Then
         MsgBox "You don't have access to this church.  Access has been denied.", vbExclamation
         CheckLogonId = False
         Exit Function
    End If
     
    
    If getUserPriveleges(rslocal!id) Then
              CheckLogonId = True
    Else
        CheckLogonId = False
        
    End If
    gUserFullName = rslocal!Full_Name
    UserId = rslocal!id
    dtePasswordLastUpdate = rslocal!Password_Last_Update
    checkSystemManager = rslocal!SYSTEM_MANAGER
    checkReportView = rslocal!Report_View
   
    CheckPasword = rslocal!Logon_Password
    If checkSystemManager = "Y" Then
      systemManager = True
      
    Else
      systemManager = False
    End If
    
    If checkReportView = "Y" Then
      ReportView = True
      
    Else
      ReportView = False
    End If
    Set rslocal = Nothing

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modLogon", "CheckLogonId", True)

End Function

Public Function encryptPassword(s As String) As String
encryptPassword = EDS(s)
End Function

Public Function DecryptPassword(s As String) As String
DecryptPassword = EDS(s)
End Function

Public Function EDS(ByVal s As String) As String
Dim encrypt(1 To 12) As Byte
Dim i As Byte
'encrypt = Array(223, 127, 100, 174, 20, 80, 129, 156, 168, 166, 9, 242)

encrypt(1) = 23
encrypt(2) = 127
encrypt(3) = 8
encrypt(4) = 96
encrypt(5) = 3
encrypt(6) = 16
encrypt(7) = 124
encrypt(8) = 12
encrypt(9) = 15
encrypt(10) = 25
encrypt(11) = 30
encrypt(12) = 94
For i = 1 To Len(s)
Mid(s, i, 1) = Chr(Asc(Mid(s, i, 1)) Xor encrypt(i))
Next i
EDS = s

End Function

Public Function cmdEncrypt(txtText As String)
    ' You can encrypt twice for extra security
    cmdEncrypt = EncryptText((txtText), SALT_PASSWORD)
    cmdEncrypt = EncryptText((cmdEncrypt), SALT_PASSWORD)
End Function

Public Function cmdDecrypt(txtText As String)
    cmdDecrypt = DecryptText((txtText), SALT_PASSWORD)
    cmdDecrypt = DecryptText((cmdDecrypt), SALT_PASSWORD)
End Function

Public Function cmdDecryptLicence(txtText As String)
    cmdDecryptLicence = DecryptText((txtText), SALT_LICENCE)
    cmdDecryptLicence = DecryptText((cmdDecryptLicence), SALT_LICENCE)
End Function

'Encrypt text
Private Function EncryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

#End If

    'Encrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    EncryptText = strBuff
End Function

'Decrypt text encrypted with EncryptText
Private Function DecryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

#End If

    'Decrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    DecryptText = strBuff
End Function
       


Public Function LockUser(strLogonID As String)

On Error GoTo ErrorHandler
    Dim sql As String
            
        objConnection.BeginTrans
                
                sql = "UPDATE Users SET" _
                            & " Logon_Password =" & "'" & "LOCKED" & "'"
                            
                            
                sql = sql & " WHERE Logon_Id = '" & strLogonID & "'"
    
                objConnection.Execute sql

        objConnection.CommitTrans
Exit Function
ErrorHandler:
Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modLogon", "CheckLogonId", True)
End Function

Public Function CheckPasswordChange()
On Error GoTo ErrorHandler

    Dim dteCurrentDate As Date
    Dim dtePasswordDate As Date
    Dim lngNoDays As Long
    
    CheckPasswordChange = False
    
    'Check password last change date
    CompulsoryChangePassword = False
    dtePasswordDate = Format(dtePasswordLastUpdate, DATE_TIME_FORMAT)
    dteCurrentDate = Format(Date, DATE_TIME_FORMAT)
    lngNoDays = DateDiff("d", dtePasswordDate, dteCurrentDate)
    If (lngNoDays < 0) Or (lngNoDays > 30) Or (CheckPasword = "WELCOME") Then
        CompulsoryChangePassword = True
        Exit Function
     Else
    CheckPasswordChange = True
    End If
    
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "modLogon", "CheckPasswordChange", True)

End Function

