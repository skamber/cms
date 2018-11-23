VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCollection_s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private dbConnection As ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarNewCollectionId As Long 'local copy
Public Property Let NewCollectionId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewCollectionId = 5
    mvarNewCollectionId = vData
End Property


Public Property Get NewCollectionId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewCollectionId
    NewCollectionId = mvarNewCollectionId
End Property



Public Property Set DatabaseConnection(vData As ADODB.Connection)
    Set dbConnection = vData
End Property

Public Function InsertCollection(ByVal objCollection As CMSCollection.clsCollection)
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
   
    
   
        'get next number -- cannot use AutoNumber function
        sql = "SELECT MAX(col_ID) FROM collection"
        Set rsLocal = New ADODB.Recordset
            rsLocal.Open sql, dbConnection
            
        If IsNull(rsLocal.Fields(0)) Then
            NewCollectionId = 1
        Else
            NewCollectionId = rsLocal.Fields(0) + 1
        End If
        rsLocal.Close
        Set rsLocal = Nothing

    
    
    'insert new record
    sql = "INSERT INTO collection" _
                & " (col_ID,payment, Amount, church_id, User_Name,type,comments,Date_of_collection)" _
                & " VALUES" _
                & " (" & NewCollectionId _
                & ",'" & objCollection.payment & "'" _
                & "," & objCollection.Amount _
                & "," & objCollection.ChurchId _
                & ",'" & objCollection.UserName & "'" _
                & ",'" & objCollection.types & "'" _
                
                If objCollection.comments = "" Then
                sql = sql & ", NULL"
                Else
                sql = sql & ",'" & objCollection.comments & "'"
                End If
                
                If objCollection.Dateofcollection = "" Then
                sql = sql & ", NULL)"
                Else
                sql = sql & ",'" & Format(objCollection.Dateofcollection, "yyyy-mm-dd") & "')"
                End If
    
    dbConnection.Execute sql

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsCollection_s", "InsertCollection", True)

End Function

Public Function UpdateCollection(ByVal objCollection As CMSCollection.clsCollection)
On Error GoTo ErrorHandler

    Dim sql As String

    
    sql = "UPDATE collection SET" _
                & " payment =" & "'" & objCollection.payment & "'" _
                & " ,type =" & "'" & objCollection.types & "'" _
                & " ,church_id =" & objCollection.ChurchId _
                & " ,Amount =" & objCollection.Amount _
                & " ,User_Name =" & "'" & objCollection.UserName & "'" _
                & " ,comments =" & "'" & objCollection.comments & "'"
               
                
    If objCollection.Dateofcollection <> "" Then
        sql = sql & " ,Date_of_collection =" & "'" & Format(objCollection.Dateofcollection, "yyyy-mm-dd") & "'"
    Else
        sql = sql & " ,Date_of_collection = NULL "
    End If
    
    sql = sql & " WHERE col_ID = " & objCollection.col_ID
    
    dbConnection.Execute sql


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsCollection_s", "UpdateCollection", True)

End Function


Public Function getByCollectionId(ByVal lngCollectionID As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT * FROM collection where COL_ID = " & lngCollectionID
    
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly
    
    If rsLocal.EOF = True Then
        MsgBox "No records found for This Collection.", vbExclamation
        Set getByCollectionId = Nothing
    Else
        Set getByCollectionId = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsCollection_s", "getByCollectionId", True)

End Function
