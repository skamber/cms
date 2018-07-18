Attribute VB_Name = "modStartupMain"
' (c) 2002-2002 Marsh Ltd.                                           }
' All Rights Reserved.                                                   }
 
  
' Update History
 
 
'Date        By     Comment
'22-May-02   SAM    Set Parameter for ForceRegistration
'<!<CHECKOUT>!>

'Backed up to 4548 on 23-May-02 by SAM
'<!<PREVIOUS_VERSIONS>!>
'}

Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const APPLICATION_NAME = "CMS"
Private Const INI_FILE_NAME = "CMS.INI"
Private Const APPLICATION_EXE = "CMS.exe"

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

' Registry API declarations and constants.
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey As Long)

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_READ = &H20000
Public Const STANDARD_RIGHTS_WRITE = &H20000
Public Const STANDARD_RIGHTS_EXECUTE = &H20000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = (KEY_READ)
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const ERROR_SUCCESS = 0&
Public Const REG_NONE = 0
Public Const REG_SZ = 1
Public Const REG_EXPAND_SZ = 2
Public Const REG_DWORD = 4
Public Const REG_OPENED_EXISTING_KEY = &H2
Public Const REG_OPTION_NON_VOLATILE = 0
Sub Main()

    Dim strVersion As String
    Dim strLocation As String
    Dim strRegEntry As String
    Dim strIniEntry As String
    Dim strCurrVersion As String
    Dim strNeedToRegister As String
    Dim strForceRegistration As String
    Dim lngResult As Long
    Dim intCnt As Integer
    Dim intNoDllsToUnRegister As Integer
    Dim intNoDllsToRegister As Integer
    Dim strIniKey As String
    Dim strShellString As String
    Dim lngRes As Long
    Dim strKeyValue As String
    Dim lngKeyHandle As Long
    Dim lngKeyLength As Long
    Dim lngKeyResult As Long
    Dim udtSecurityAttributes As SECURITY_ATTRIBUTES
    ReDim strDllToUnRegister(1 To 100) As String
    ReDim strDllToRegister(1 To 100) As String
    
    lngRes = RegCreateKeyEx(HKEY_LOCAL_MACHINE, "Software\CHURCH", 0, "REG_SZ", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, udtSecurityAttributes, lngKeyHandle, lngKeyResult)
        
    If lngRes <> ERROR_SUCCESS Then
        ' Has not been installed before.
        strVersion = "0.0.0"
    Else
        strKeyValue = String(255, 0)
        lngKeyLength = 255
        lngRes = RegQueryValueEx(lngKeyHandle, "CMS", 0, REG_SZ, strKeyValue, lngKeyLength)
    
        If lngKeyLength = 0 Or lngKeyLength = 255 Then
            ' Settings are incorrect, reinstall anyway.
            strVersion = "0.0.0"
        Else
            ' Get the key setting.
            strVersion = Trim(Mid(strKeyValue, 1, (lngKeyLength - 1)))
        End If
    End If
    
    On Error GoTo Error_Handler
    
    strLocation = App.Path & "\" & INI_FILE_NAME
    
    If UCase(Trim(Dir(strLocation))) <> UCase(INI_FILE_NAME) Then
        MsgBox "The version control file was not found. Please notify the helpdesk.", vbCritical, App.Title
        
        End
    Else
        ' Read the INI File, find out the current version, and whether
        ' the registry settings need to be re-freshed.
        
        ' Current Version
        strIniEntry = String(255, 0)
        
        lngResult = GetPrivateProfileString(APPLICATION_NAME, "CurrentVersion", "", strIniEntry, 255, strLocation)
        If lngResult = 0 Then
            MsgBox "The Current Version could not be read. Please call the helpdesk.", vbCritical, App.Title
            End
        Else
            strCurrVersion = Mid(strIniEntry, 1, lngResult)
        End If
        
        ' Force Registration ?
        strIniEntry = String(255, 0)
        
        lngResult = GetPrivateProfileString(APPLICATION_NAME, "ForceRegistration", "", strIniEntry, 255, strLocation)
        If lngResult = 0 Then
            MsgBox "The Force Registration Option could not be read. Please call the helpdesk.", vbCritical, App.Title
            End
        Else
            strForceRegistration = Mid(strIniEntry, 1, lngResult)
        End If
        
        If (strVersion <> strCurrVersion) Or (UCase(strForceRegistration) = "Y") Or (LoadCommandLine("/R=") = "Y") Then
            ' Need to Register ?
            strIniEntry = String(255, 0)
            
            lngResult = GetPrivateProfileString(APPLICATION_NAME, "NeedToRegister", "", strIniEntry, 255, strLocation)
            If lngResult = 0 Then
                MsgBox "The Need to Register section of the " & INI_FILE_NAME & " file could not be read. Please call the helpdesk.", vbCritical, App.Title
                End
            Else
                strNeedToRegister = Mid(strIniEntry, 1, lngResult)
            End If
            
            If (UCase(strNeedToRegister) = "Y") Or (UCase(strForceRegistration) = "Y") Then
                intNoDllsToUnRegister = 0
                intNoDllsToRegister = 0
                
                ' Read in all the Dll's to Unregister.
                Do
                    ' File to Un-Register
                    strIniEntry = String(255, 0)
                    strIniKey = "FileToUnRegister" & Format(intNoDllsToUnRegister + 1)
                    lngResult = GetPrivateProfileString(APPLICATION_NAME, strIniKey, "", strIniEntry, 255, strLocation)
                    If lngResult = 0 Then
                        Exit Do
                    Else
                        intNoDllsToUnRegister = (intNoDllsToUnRegister + 1)
                        If intNoDllsToUnRegister = UBound(strDllToUnRegister) Then
                            ReDim Preserve strDllToUnRegister(1 To intNoDllsToUnRegister + 50)
                        End If
                        strDllToUnRegister(intNoDllsToUnRegister) = Mid(strIniEntry, 1, lngResult)
                    End If
                Loop
                
                Do
                    ' Current Version
                    strIniEntry = String(255, 0)
                    strIniKey = "FileToRegister" & Format(intNoDllsToRegister + 1)
                    lngResult = GetPrivateProfileString(APPLICATION_NAME, strIniKey, "", strIniEntry, 255, strLocation)
                    If lngResult = 0 Then
                        Exit Do
                    Else
                        intNoDllsToRegister = (intNoDllsToRegister + 1)
                        If intNoDllsToRegister = UBound(strDllToRegister) Then
                            ReDim Preserve strDllToRegister(1 To intNoDllsToRegister + 50)
                        End If
                        strDllToRegister(intNoDllsToRegister) = Mid(strIniEntry, 1, lngResult)
                    End If
                Loop
    
                frmStartupStatus.pbStatus.Max = (intNoDllsToRegister + intNoDllsToUnRegister)
                frmStartupStatus.pbStatus.Value = 0
                frmStartupStatus.Show
                
                ' Un-Register the dll's.
                For intCnt = 1 To intNoDllsToUnRegister
                    Call RegisterDll(strDllToUnRegister(intCnt), True)
                Next
                
                ' Register the dll's.
                For intCnt = 1 To intNoDllsToRegister
                    Call RegisterDll(strDllToRegister(intCnt), False)
                Next
                
                Unload frmStartupStatus
                frmStartupStatus.pbStatus.Max = 10
                frmStartupStatus.Timer1.Enabled = True
                frmStartupStatus.Show vbModal
            End If
            
            ' Write the new version into the registry file.
            lngRes = RegSetValueEx(lngKeyHandle, "CMS", 0, REG_SZ, strCurrVersion, Len(strCurrVersion))
        End If
    End If
    
    lngRes = RegCloseKey(lngKeyHandle)
    
    strShellString = (App.Path & "\" & APPLICATION_EXE & " " & Chr(34) & Trim(Command) & Chr(34))
    Call Shell(strShellString, vbNormalFocus)
    
    Exit Sub
    
Error_Handler:
    
    'MsgBox "There was a problem when starting the " & APPLICATION_NAME & " application." & Chr(13) & "Please contact the HelpDesk and inform them of the following message." & Chr(13) & "Error: " & Format(Err.Number) & " " & Format(Err.Description) & ".",
    
End Sub
Private Sub RegisterDll(rstrDllToRegister As String, rblnUnRegister As Boolean)

    frmStartupStatus.pbStatus.Value = (frmStartupStatus.pbStatus.Value + 1)
    frmStartupStatus.pbStatus.Refresh
    
    If rblnUnRegister Then
        frmStartupStatus.lblStatus = "Un-Registering " & rstrDllToRegister & "..."
        frmStartupStatus.Refresh
        
        Call Shell("regsvr32 " & Chr(34) & rstrDllToRegister & Chr(34) & " /u /s", vbHide)
    Else
        frmStartupStatus.lblStatus = "Registering " & rstrDllToRegister & "..."
        frmStartupStatus.Refresh
        
        Call Shell("regsvr32 " & Chr(34) & rstrDllToRegister & Chr(34) & " /s", vbHide)
    End If

End Sub




Private Function LoadCommandLine(ParmDef As String) As String


Dim T1 As String
Dim Ptr1, Ptr2 As Long
Dim ParmStr As String

LoadCommandLine = ""

T1 = Trim(UCase(Command$)) ' Save and tidy the command line
For Ptr1 = 1 To Len(T1)
   If Mid(T1, Ptr1, Ptr1 + 2) = ParmDef Then
      ParmStr = ""
      For Ptr2 = (Ptr1 + 3) To Len(T1)
          If Mid(T1, Ptr2, 1) = "/" Then Exit For
          ParmStr = ParmStr & Mid(T1, Ptr2, 1)
      Next Ptr2
      LoadCommandLine = ParmStr
      Exit For
   End If
Next Ptr1
End Function


