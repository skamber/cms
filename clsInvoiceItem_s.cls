VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInvoiceItem_s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private dbConnection As ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarNewInvoiceItemId As Long 'local copy
Public Property Let NewInvoiceItemId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewInvoiceItemId = 5
    mvarNewInvoiceItemId = vData
End Property


Public Property Get NewInvoiceItemId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewInvoiceItemId
    NewInvoiceItemId = mvarNewInvoiceItemId
End Property



Public Property Set DatabaseConnection(vData As ADODB.Connection)
    Set dbConnection = vData
End Property
Public Function InsertInvoiceItem(ByVal objInvoiceItem As CMSInvoiceItem.clsInvoiceItem)
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
   
    
   
        'get next number -- cannot use AutoNumber function
        sql = "SELECT MAX(Item_Id) FROM InvoiceItem"
        Set rsLocal = New ADODB.Recordset
            rsLocal.Open sql, dbConnection
            
        If IsNull(rsLocal.Fields(0)) Then
            NewInvoiceItemId = 1
        Else
            NewInvoiceItemId = rsLocal.Fields(0) + 1
        End If
        rsLocal.Close
        Set rsLocal = Nothing

    
    
    'insert new record
    sql = "INSERT INTO InvoiceItem" _
                & " (Item_Id, Invoice_id, Description, Amount, Gst, total_amount)" _
                & " VALUES" _
                & " (" & NewInvoiceItemId _
                & "," & objInvoiceItem.Invoice_id _
                & ",'" & objInvoiceItem.Description & "'" _
                & "," & objInvoiceItem.Amount _
                & "," & objInvoiceItem.GstAmount _
                & "," & objInvoiceItem.total_amount & ")"
                
                
    dbConnection.Execute sql

Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsInvoice_s", "InsertInvoice", True)

End Function

Public Function UpdateInvoiceItem(ByVal objInvoiceItem As CMSInvoiceItem.clsInvoiceItem)
On Error GoTo ErrorHandler

    Dim sql As String

    
    sql = "UPDATE InvoiceItem SET" _
                & " Invoice_id =" & objInvoiceItem.Invoice_id _
                & " ,Description =" & "'" & objInvoiceItem.Description & "'" _
                & " ,Amount =" & objInvoiceItem.Amount _
                & " ,Gst  =" & objInvoiceItem.GstAmount _
                & " ,total_amount =" & objInvoiceItem.total_amount
    
    sql = sql & " WHERE Item_Id = " & objInvoiceItem.Item_Id
    
    dbConnection.Execute sql


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsInvoice_s", "UpdateInvoice", True)

End Function

Public Function getByInvoiceId(ByVal lngInvoiceID As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT * FROM InvoiceItem WHERE Invoice_id = " & lngInvoiceID
    
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
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsMember_s", "getByMemberId", True)

End Function

Public Function getByInvoiceItemId(ByVal lngInvoiceItemID As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT * FROM InvoiceItem WHERE Item_id = " & lngInvoiceItemID
    
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly
    
    If rsLocal.EOF = True Then
        Set getByInvoiceItemId = Nothing
    Else
        Set getByInvoiceItemId = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsMember_s", "getByMemberId", True)

End Function

