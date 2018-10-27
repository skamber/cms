VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUserPermissions_s"
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

'13-May-02   SAM    Create the unit
'<!<CHECKOUT>!>

'Backed up to 4533 on 16-May-02 by SAM
'<!<PREVIOUS_VERSIONS>!>
'}
Option Explicit
Private dbConnection As ADODB.Connection
'local variable(s) to hold property value(s)

Public Property Set DatabaseConnection(vData As ADODB.Connection)
    Set dbConnection = vData
End Property



Public Function DeletePermissionForUser(lngUserID As Long)
On Error GoTo ErrorHandler
    
    Dim sql As String


    sql = "DELETE FROM User_Practice_link" _
        & " WHERE User_Id = " & lngUserID
    
    dbConnection.Execute sql


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "clsUserPractice_s", "DeletePracticesForUser", True)
    
End Function



Public Function InsertpermissionForUser(ByVal objUserpermission As CMSUser.clsUserPermissions)
On Error GoTo ErrorHandler

    Dim sql As String
    
           
   'insert new record
    sql = "INSERT INTO privilege" _
                & " (User_Id, Action_id,Read_data,Create_data,Edit_data,Delete_Data)" _
                & " VALUES" _
                & " (" & objUserpermission.UserId & "," _
                & objUserpermission.ActionId & "," _
                & "'" & objUserpermission.ReadPermissions & "'" _
                & ",'" & objUserpermission.WritePermissions & "'" _
                & ",'" & objUserpermission.UpdatePermissions & "'" _
                & ",'" & objUserpermission.DeletePermissions & "'" & ")"
                
    dbConnection.Execute sql


Exit Function
ErrorHandler:
    
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsUserPractice_s", "InsertUserPractice", True)

End Function

Public Function UpdatepermissionForUser(ByVal objUserpermission As CMSUser.clsUserPermissions)
On Error GoTo ErrorHandler

    Dim sql As String

    
    sql = "UPDATE privilege SET" _
                & " READ_data =" & "'" & objUserpermission.ReadPermissions & "'" _
                & " ,Create_data =" & "'" & objUserpermission.WritePermissions & "'" _
                & " ,Edit_data =" & "'" & objUserpermission.UpdatePermissions & "'" _
                & " ,Delete_data =" & "'" & objUserpermission.DeletePermissions & "'"
    sql = sql & " WHERE User_Id = " & objUserpermission.UserId & " AND " & "Action_id =" & objUserpermission.ActionId
    
    dbConnection.Execute sql


Exit Function
ErrorHandler:
        'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsUser_s", "UpdateUser", True)

End Function
