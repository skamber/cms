VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCashIn_s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private dbConnection As ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarNewCashInId As Long 'local copy
Public Property Let NewCashInId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewCashInId = 5
    mvarNewCashInId = vData
End Property


Public Property Get NewCashInId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewCashInId
    NewCashInId = mvarNewCashInId
End Property



Public Property Set DatabaseConnection(vData As ADODB.Connection)
    Set dbConnection = vData
End Property

Public Function InsertCashIn(ByVal objCashIn As CMSCashIn.clsCashIn)
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
        
   
        'get next number -- cannot use AutoNumber function
        sql = "SELECT MAX(ID) FROM cashin"
        Set rsLocal = New ADODB.Recordset
            rsLocal.Open sql, dbConnection
            
        If IsNull(rsLocal.Fields(0)) Then
            NewCashInId = 1
        Else
            NewCashInId = rsLocal.Fields(0) + 1
        End If
        rsLocal.Close
        Set rsLocal = Nothing

    
    
    'insert new record
    sql = "INSERT INTO cashin" _
                & " (ID, ITEM, AMOUNT, GST, Net_AMOUNT, COMMENT" _
                & " ,DateOfCashIn )" _
                & " VALUES" _
                & " (" & NewCashInId _
                & ",'" & objCashIn.ITEM & "'" _
                & "," & objCashIn.Amount _
                & "," & objCashIn.GST & "," & objCashIn.Total_Amount _
                & ",'" & objCashIn.Comment & "'"
                
                If objCashIn.DateofCashIn = "" Then
                sql = sql & ", NULL)"
                Else
                sql = sql & ",'" & Format(objCashIn.DateofCashIn, "yyyy-mm-dd") & "')"
                End If
    
    dbConnection.Execute sql

Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsCashIn_s", "InsertCasiIn", True)

End Function

Public Function UpdateCashIn(ByVal objCashIn As CMSCashIn.clsCashIn)
On Error GoTo ErrorHandler

    Dim sql As String

    
    sql = "UPDATE cashin SET" _
                & " ITEM =" & "'" & objCashIn.ITEM & "'" _
                & " ,AMOUNT =" & objCashIn.Amount _
                & " ,GST =" & objCashIn.GST _
                & " ,Net_AMOUNT =" & objCashIn.Total_Amount _
                & " ,COMMENT =" & "'" & objCashIn.Comment & "'"
                
                
    If objCashIn.DateofCashIn <> "" Then
        sql = sql & " ,DateOfCashIn =" & "'" & Format(objCashIn.DateofCashIn, "yyyy-mm-dd") & "'"
    Else
        sql = sql & " ,DateOfCashIn = NULL "
    End If
    
    sql = sql & " WHERE ID = " & objCashIn.ID
    
    dbConnection.Execute sql


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsMember_s", "UpdateMember", True)

End Function

Public Function getByCashInId(ByVal lngCashInID As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT * FROM cashin WHERE ID = " & lngCashInID
    
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly
    
    If rsLocal.EOF = True Then
        MsgBox "No record found for This Payment.", vbExclamation
        Set getByCashInId = Nothing
    Else
        Set getByCashInId = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsPayment_s", "getByPaymentId", True)

End Function





