VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCashflowItem_s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private dbConnection As ADODB.Connection
Private mvarNewCashflowItemId As Long 'local copy
Public Property Let NewCashFlowItemId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewCashInId = 5
    mvarNewCashflowItemId = vData
End Property


Public Property Get NewCashFlowItemId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewCashInId
    NewCashFlowItemId = mvarNewCashflowItemId
End Property

Public Property Set DatabaseConnection(vData As ADODB.Connection)
    Set dbConnection = vData
End Property
Public Function InsertCashflowItem(ByVal objCashflowItem As CMSCashFlowItem.clsCashflowItem)
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
        
   
        'get next number -- cannot use AutoNumber function
        sql = "SELECT MAX(ID) FROM CASHFLOWITEM"
        Set rsLocal = New ADODB.Recordset
            rsLocal.Open sql, dbConnection
            
        If IsNull(rsLocal.Fields(0)) Then
            NewCashFlowItemId = 1
        Else
            NewCashFlowItemId = rsLocal.Fields(0) + 1
        End If
        rsLocal.Close
        Set rsLocal = Nothing

    
    
    'insert new record
    sql = "INSERT INTO CASHFLOWITEM" _
                & " (ID, ITEMNAME, ITEMCODE, GST, type)" _
                & " VALUES" _
                & " (" & NewCashFlowItemId _
                & ",'" & objCashflowItem.ItemName & "'" _
                & ",'" & objCashflowItem.ItemCode & "'" _
                & "," & objCashflowItem.GST & ",'" & objCashflowItem.ItemType & "')"
    
    dbConnection.Execute sql

Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsCashIn_s", "InsertCasiIn", True)

End Function

Public Function UpdateCashflowItem(ByVal objCashflowItem As CMSCashFlowItem.clsCashflowItem)
On Error GoTo ErrorHandler

    Dim sql As String

    
    sql = "UPDATE CASHFLOWITEM SET" _
                & " ITEMNAME =" & "'" & objCashflowItem.ItemName & "'" _
                & " ,ITEMCODE =" & "'" & objCashflowItem.ItemCode & "'" _
                & " ,GST =" & objCashflowItem.GST _
                & " ,TYPE =" & "'" & objCashflowItem.ItemType & "'"
    sql = sql & " WHERE ID = " & objCashflowItem.ID
    
    dbConnection.Execute sql


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsMember_s", "UpdateMember", True)

End Function

Public Function getByCashflowItemId(ByVal lngCashflowItemID As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT * FROM CASHFLOWITEM WHERE ID = " & lngCashflowItemID
    
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly
    
    If rsLocal.EOF = True Then
        MsgBox "No record found for This Account.", vbExclamation
        Set getByCashflowItemId = Nothing
    Else
        Set getByCashflowItemId = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsPayment_s", "getByPaymentId", True)

End Function






