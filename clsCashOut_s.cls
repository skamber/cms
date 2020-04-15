VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCashOut_s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private dbConnection As ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarNewCashOutId As Long 'local copy
Public Property Let NewCashOutId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewCashInId = 5
    mvarNewCashOutId = vData
End Property


Public Property Get NewCashOutId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewCashInId
    NewCashOutId = mvarNewCashOutId
End Property



Public Property Set DatabaseConnection(vData As ADODB.Connection)
    Set dbConnection = vData
End Property

Public Function InsertCashOut(ByVal objCashOut As CMSCashOut.clsCashOut)
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
        
   
        'get next number -- cannot use AutoNumber function
        sql = "SELECT MAX(ID) FROM CASHOUT"
        Set rsLocal = New ADODB.Recordset
            rsLocal.Open sql, dbConnection
            
        If IsNull(rsLocal.Fields(0)) Then
            NewCashOutId = 1
        Else
            NewCashOutId = rsLocal.Fields(0) + 1
        End If
        rsLocal.Close
        Set rsLocal = Nothing

    
    
    'insert new record
    sql = "INSERT INTO cashout" _
                & " (ID, ITEM, ITEM_CODE, AMOUNT, GST, Net_AMOUNT, COMMENT" _
                & " ,DateOfCashOut )" _
                & " VALUES" _
                & " (" & NewCashOutId _
                & ",'" & objCashOut.ITEM & "'" _
                & ",'" & objCashOut.Item_code & "'," & objCashOut.Amount _
                & "," & objCashOut.GST & "," & objCashOut.Total_Amount _
                & ",'" & objCashOut.Comment & "'"
                
                If objCashOut.DateOfCashOut = "" Then
                sql = sql & ", NULL)"
                Else
                sql = sql & ",'" & Format(objCashOut.DateOfCashOut, "yyyy-mm-dd") & "')"
                End If
    
    dbConnection.Execute sql

Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsmember_s", "Insertmember", True)

End Function

Public Function UpdateCashOut(ByVal objCashOut As CMSCashOut.clsCashOut)
On Error GoTo ErrorHandler

    Dim sql As String

    
    sql = "UPDATE cashout SET" _
                & " ITEM =" & "'" & objCashOut.ITEM & "'" _
                & " ,ITEM_CODE =" & "'" & objCashOut.Item_code & "'" _
                & " ,AMOUNT =" & objCashOut.Amount _
                & " ,GST =" & objCashOut.GST _
                & " ,Net_AMOUNT =" & objCashOut.Total_Amount _
                & " ,COMMENT =" & "'" & objCashOut.Comment & "'"
                
                
    If objCashOut.DateOfCashOut <> "" Then
        sql = sql & " ,DateOfCashOut =" & "'" & Format(objCashOut.DateOfCashOut, "yyyy-mm-dd") & "'"
    Else
        sql = sql & " ,DateOfCashOut = NULL "
    End If
    
    sql = sql & " WHERE ID = " & objCashOut.ID
    
    dbConnection.Execute sql


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsMember_s", "UpdateMember", True)

End Function

Public Function getByCashOutId(ByVal lngCashOutID As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT * FROM cashout WHERE ID = " & lngCashOutID
    
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly
    
    If rsLocal.EOF = True Then
        MsgBox "No record found for This Payment.", vbExclamation
        Set getByCashOutId = Nothing
    Else
        Set getByCashOutId = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsPayment_s", "getByPaymentId", True)

End Function






