VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChildren_s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private dbConnection As ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarNewChildId As Long 'local copy
Public Property Let NewChildId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewChildrenId = 5
    mvarNewChildId = vData
End Property


Public Property Get NewChildId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewChildrenId
    NewChildId = mvarNewChildId
End Property



Public Property Set DatabaseConnection(vData As ADODB.Connection)
    Set dbConnection = vData
End Property

Public Function InsertChild(ByVal objChild As CMSChildren.clsChildren)
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
        
   
        'get next number -- cannot use AutoNumber function
        sql = "SELECT MAX(ID) FROM CHILDREN"
        
        Set rsLocal = New ADODB.Recordset
            rsLocal.Open sql, dbConnection
            
        If IsNull(rsLocal.Fields(0)) Then
            NewChildId = 1
        Else
            NewChildId = rsLocal.Fields(0) + 1
        End If
        rsLocal.Close
        Set rsLocal = Nothing

    
    
    'insert new record
    sql = "INSERT INTO CHILDREN" _
                & " (MNo, ID, First_Name, Surname, Email,COMMENTS,Genda, Member, Birth_Date)" _
                & " VALUES" _
                & " (" & objChild.MNo _
                & "," & NewChildId _
                & ",'" & objChild.FirstName & "','" & objChild.Surname & "'" _
                & ",'" & objChild.Email & "','" & objChild.Memo & "'" _
                & ",'" & objChild.Genda & "','" & objChild.Member & "'"
                
                If objChild.BirthDate = "" Then
                sql = sql & ", NULL)"
                Else
                sql = sql & ",'" & Format(objChild.BirthDate, "dd-mmm-yyyy") & "')"
                End If
    
    dbConnection.Execute sql

Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsmember_s", "Insertmember", True)

End Function

Public Function UpdateChild(ByVal objChild As CMSChildren.clsChildren)
On Error GoTo ErrorHandler

    Dim sql As String

    
    sql = "UPDATE CHILDREN SET" _
                & " MNo =" & objChild.MNo _
                & " ,First_Name =" & "'" & objChild.FirstName & "'" _
                & " ,Surname =" & "'" & objChild.Surname & "'" _
                & " ,Email =" & "'" & objChild.Email & "'" _
                & " ,COMMENTS =" & "'" & objChild.Memo & "'" _
                & " ,Member =" & "'" & objChild.Member & "'" _
                & " ,Genda =" & "'" & objChild.Genda & "'" _

                
    If objChild.BirthDate <> "" Then
        sql = sql & " ,Birth_Date =" & "'" & Format(objChild.BirthDate, "dd-mmm-yyyy") & "'"
    Else
        sql = sql & " ,Birth Date = NULL "
    End If
    
    sql = sql & " WHERE ID = " & objChild.ChildNo
    
    dbConnection.Execute sql


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsMember_s", "UpdateMember", True)

End Function

Public Function getByChildId(ByVal lngChildID As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT * FROM CHILDREN WHERE Id = " & lngChildID
    
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly
    
    If rsLocal.EOF = True Then
        MsgBox "No records found for This Member.", vbExclamation
        Set getByChildId = Nothing
    Else
        Set getByChildId = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsChildren_s", "getByChildId", True)

End Function

