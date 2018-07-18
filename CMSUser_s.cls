VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser_s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' (c) 2002-2002 Marsh Ltd.                                           }
' All Rights Reserved.                                                   }
 
 
 
' Update History
 
 
'Date        By     Comment

'03-May-02   SAM    New function GetAllUserName
'27-Aug-02   SAM    Create New Privilege called Practice Administrator
'<!<CHECKOUT>!>
 
'Backed up to 4534 on 16-May-02 by SAM
'Backed up to 4615 on 27-Aug-02 by SAM
'<!<PREVIOUS_VERSIONS>!>
'}
Option Explicit
Private mvarNewUserId As Long
Private dbConnection As ADODB.Connection
Public Property Set DatabaseConnection(vData As ADODB.Connection)
    Set dbConnection = vData
End Property

Public Property Let NewUserId(ByVal vData As Long)
    mvarNewUserId = vData
End Property

Public Property Get NewUserId() As Long
    NewUserId = mvarNewUserId
End Property

Public Function getUserName(lngUserID As Long) As String
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT Id, FULL_NAME FROM Users " _
        & "WHERE Id = " & lngUserID
        
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly
    
    If rsLocal.EOF = True Then
        'MsgBox "No records found for selected User.", vbExclamation
        getUserName = ""
    Else
        getUserName = rsLocal!Full_name & ""
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsUser_s", "getUserName", True)

End Function




Public Function getAllUserName() As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT Id, Full_name FROM Users "
        
        
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly
    
    If rsLocal.EOF = True Then
        MsgBox "No records found for Account Executive.", vbExclamation
        Set getAllUserName = Nothing
    Else
        Set getAllUserName = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsUser_s", "getAllUserName", True)

End Function

Public Function getUserAccountRecord(lngUserID As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT * FROM Users " _
        & "WHERE Id =" & lngUserID
        
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly
    
    If rsLocal.EOF = True Then
        MsgBox "No Account found for This User.", vbExclamation
        Set getUserAccountRecord = Nothing
    Else
        Set getUserAccountRecord = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsUser_s", "getUserAccountRecord", True)

End Function



Public Function UpdateUser(ByVal objUser As CMSUser.clsUser)
On Error GoTo ErrorHandler

    Dim sql As String

    
    sql = "UPDATE USERS SET" _
                & " Full_Name =" & "'" & objUser.Username & "'" _
                & " ,LOGON_ID =" & "'" & objUser.LogonId & "'" _
                & " ,SYSTEM_MANAGER =" & "'" & objUser.SystemManager & "'" _
                & " ,LOGON_PASSWORD =" & "'" & objUser.LogonPassword & "'" _
                & " ,Report_View =" & "'" & objUser.ReportView & "'"
        sql = sql & " ,PASSWORD_LAST_UPDATE =" & "'" & Format(objUser.PasswordLastChange, "dd-mmm-yyyy") & "'"
    
    sql = sql & " WHERE ID = " & objUser.UserId
    
    dbConnection.Execute sql


Exit Function
ErrorHandler:
    '    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsUser_s", "UpdateUser", True)

End Function

Public Function InsertUser(ByVal objUser As CMSUser.clsUser)
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
           
   
        'get next number -- cannot use AutoNumber function
        sql = "SELECT MAX(ID) FROM Users"
        Set rsLocal = New ADODB.Recordset
            rsLocal.Open sql, dbConnection
            
        If IsNull(rsLocal.Fields(0)) Then
            NewUserId = 1
        Else
            NewUserId = rsLocal.Fields(0) + 1
        End If
        rsLocal.Close
        Set rsLocal = Nothing

    
    
    'insert new record
    sql = "INSERT INTO Users" _
                & " (Id, Full_Name, LOGON_ID,LOGON_PASSWORD,SYSTEM_MANAGER,REPORT_VIEW,PASSWORD_LAST_UPDATE)" _
                & " VALUES" _
                & " (" & NewUserId _
                & ",'" & objUser.Username & "'" _
                & ",'" & objUser.LogonId & "'" _
                & ",'" & objUser.LogonPassword & "'" _
                & ",'" & objUser.SystemManager & "'" _
                & ",'" & objUser.ReportView & "'"
                 sql = sql & ",'" & Format(objUser.PasswordLastChange, "dd-mmm-yyyy") & "')"
           
        
    dbConnection.Execute sql


Exit Function
ErrorHandler:
       
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsUser_s", "InsertUser", True)

End Function

Public Function DeleteUser(lngUserID As Long)
On Error GoTo ErrorHandler
    
    Dim sql As String


    sql = "DELETE FROM Users" _
        & " WHERE Id = " & lngUserID
    
    dbConnection.Execute sql


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "clsUser_s", "DeleteUser", True)
    
End Function


