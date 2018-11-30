VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReceipt_s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private dbConnection As ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarNewReceiptId As Long 'local copy
Public Property Let NewReceiptId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewInvoiceId = 5
    mvarNewReceiptId = vData
End Property


Public Property Get NewReceiptId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewInvoiceId
    NewReceiptId = mvarNewReceiptId
End Property



Public Property Set DatabaseConnection(vData As ADODB.Connection)
    Set dbConnection = vData
End Property

Public Function InsertReceipt(ByVal objReceipt As CMSReceipt.clsReceipt)
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
        
   
        'get next number -- cannot use AutoNumber function
        sql = "SELECT MAX(ID) FROM receipt"
        Set rsLocal = New ADODB.Recordset
            rsLocal.Open sql, dbConnection
            
        If IsNull(rsLocal.Fields(0)) Then
            NewReceiptId = 1
        Else
            NewReceiptId = rsLocal.Fields(0) + 1
        End If
        rsLocal.Close
        Set rsLocal = Nothing

    
    
    'insert new record
    sql = "INSERT INTO receipt" _
                & " (ID, INV_ID,INV_NO, church_id, amount,user_Name, Cheque_number,Date_Of_Receipt,COMMENTS)" _
                & " VALUES" _
                & " (" & NewReceiptId _
                & "," & objReceipt.INV_ID _
                & ",'" & objReceipt.INV_NO & "'" _
                & "," & objReceipt.ChurchId _
                & "," & objReceipt.amount _
                & ",'" & objReceipt.Accountname & "'" _
                & ",'" & objReceipt.ChequeNumber & "'"
                If objReceipt.DateOfReceipt = "" Then
                sql = sql & ", NULL"
                Else
                sql = sql & ",'" & Format(objReceipt.DateOfReceipt, "yyyy-mm-dd") & "'"
                End If
                
                If objReceipt.comments = "" Then
                sql = sql & ", NULL)"
                Else
                sql = sql & ",'" & objReceipt.comments & "')"
                End If
    
    dbConnection.Execute sql

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsreceipt_s", "Insertreceipt", True)

End Function

Public Function UpdateReceipt(ByVal objReceipt As CMSReceipt.clsReceipt)
On Error GoTo ErrorHandler

    Dim sql As String

    
    sql = "UPDATE receipt SET " _
                & "INV_NO =" & "'" & objReceipt.INV_NO & "'" _
                & " ,INV_ID =" & objReceipt.INV_ID _
                & " ,church_id =" & objReceipt.ChurchId _
                & " ,amount =" & objReceipt.amount _
                & " ,User_Name = '" & objReceipt.Accountname & "'" _
                & " ,Cheque_number ='" & objReceipt.ChequeNumber & "'"
    If objReceipt.DateOfReceipt <> "" Then
        sql = sql & " ,Date_Of_Receipt =" & "'" & Format(objReceipt.DateOfReceipt, "yyyy-mm-dd") & "'"
    Else
        sql = sql & " ,Date_Of_Receipt = NULL "
    End If
    
    sql = sql & " WHERE ID = " & objReceipt.ID
    
    dbConnection.Execute sql


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsMember_s", "UpdateMember", True)

End Function

Public Function getByInvoiceId(ByVal lngInvoiceid As String) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT * FROM receipt WHERE INV_ID = " & lngInvoiceid
    
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly
    
    If rsLocal.EOF = True Then
        Set getByInvoiceId = Nothing
    Else
        Set getByInvoiceId = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsReceipt_s", "getByInvoiceId", True)

End Function


Public Function getByReceiptId(ByVal lngReceiptID As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT * FROM receipt where ID = " & lngReceiptID
    
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly
    
    If rsLocal.EOF = True Then
        MsgBox "No records found for This Receipt.", vbExclamation
        Set getByReceiptId = Nothing
    Else
        Set getByReceiptId = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsCollection_s", "getByCollectionId", True)

End Function




