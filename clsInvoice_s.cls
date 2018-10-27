VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInvoice_s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private dbConnection As ADODB.Connection
'local variable(s) to hold property value(s)
Private mvarNewInvoiceId As Long 'local copy
Private mvarNewInvoiceNo As String
Private mvarNewRefNo As String

Public Property Set DatabaseConnection(vData As ADODB.Connection)
    Set dbConnection = vData
End Property
Public Property Let NewInvoiceId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewInvoiceId = 5
    mvarNewInvoiceId = vData
End Property


Public Property Get NewInvoiceId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewInvoiceId
    NewInvoiceId = mvarNewInvoiceId
End Property
Public Property Let NewInvoiceNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewInvoiceId = 5
    mvarNewInvoiceNo = vData
End Property


Public Property Get NewInvoiceNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewInvoiceId
    NewInvoiceNo = mvarNewInvoiceNo
End Property
Public Property Let NewRefNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewInvoiceId = 5
    mvarNewRefNo = vData
End Property


Public Property Get NewRefNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewInvoiceId
    NewRefNo = mvarNewRefNo
End Property

Public Function InsertInvoice(ByVal objInvoice As CMSInvoice.clsInvoice)
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
   
    
   
        'get next number -- cannot use AutoNumber function
        sql = "SELECT MAX(ID) FROM INVOICE"
        Set rsLocal = New ADODB.Recordset
            rsLocal.Open sql, dbConnection
            
        If IsNull(rsLocal.Fields(0)) Then
            NewInvoiceId = 1
        Else
            NewInvoiceId = rsLocal.Fields(0) + 1
        End If
        NewInvoiceNo = Year(objInvoice.dateofInvoice) & "_" & NewInvoiceId
        
        NewRefNo = Mid(objInvoice.Name1, 1, 3) & "_" & NewInvoiceId
        
        rsLocal.Close
        Set rsLocal = Nothing
    
    'insert new record
    sql = "INSERT INTO INVOICE" _
                & " (ID, Invoice_no, church_id,Ref, Name1, Name2, Address1, Address2" _
                & " ,Address3, Terms,Total_amount,state,Balance" _
                & " ,Phone,Mobile,email,Created_date,over_due_date )" _
                & " VALUES" _
                & " (" & NewInvoiceId _
                & ",'" & NewInvoiceNo & "'" _
                & "," & objInvoice.ChurchId _
                & ",'" & NewRefNo & "','" & objInvoice.Name1 & "'" _
                & ",'" & objInvoice.Name2 & "','" & objInvoice.Address1 & "','" & objInvoice.Address2 & "'" _
                & ",'" & objInvoice.Address3 & "','" & objInvoice.Terms & "'" _
                & "," & objInvoice.Total_amount _
                & ",'" & objInvoice.state & "'" _
                & "," & objInvoice.Balance _
                & ",'" & objInvoice.Phone & "'" _
                & ",'" & objInvoice.Mobile & "'" _
                & ",'" & objInvoice.Email & "'"
                If objInvoice.dateofInvoice = "" Then
                sql = sql & ", NULL"
                Else
                sql = sql & ",'" & Format(objInvoice.dateofInvoice, "yyyy-mm-dd") & "'"
                End If
                
                If objInvoice.over_due_date = "" Then
                sql = sql & ", NULL)"
                Else
                sql = sql & ",'" & Format(objInvoice.over_due_date, "yyyy-mm-dd") & "')"
                End If
    'Debug.Print (sql)
    dbConnection.Execute sql

Exit Function
ErrorHandler:
    Debug.Print (Err.Description)
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsInvoice_s", "InsertInvoice", True)

End Function

Public Function UpdateInvoice(ByVal objInvoice As CMSInvoice.clsInvoice)
On Error GoTo ErrorHandler

    Dim sql As String

    
    sql = "UPDATE INVOICE SET" _
                & " Invoice_no =" & "'" & objInvoice.Invoice_no & "'" _
                & " ,church_id =" & objInvoice.ChurchId _
                & " ,Ref =" & "'" & objInvoice.Ref & "'" _
                & " ,Name1 =" & "'" & objInvoice.Name1 & "'" _
                & " ,Name2 =" & "'" & objInvoice.Name2 & "'" _
                & " ,Address1 =" & "'" & objInvoice.Address1 & "'" _
                & " ,Address2 =" & "'" & objInvoice.Address2 & "'" _
                & " ,Address3 =" & "'" & objInvoice.Address3 & "'" _
                & " ,Terms =" & "'" & objInvoice.Terms & "'" _
                & " ,Total_amount =" & objInvoice.Total_amount _
                & " ,state =" & "'" & objInvoice.state & "'" _
                & " ,Balance =" & objInvoice.Balance _
                & " ,Phone =" & "'" & objInvoice.Phone & "'" _
                & " ,Mobile =" & "'" & objInvoice.Mobile & "'" _
                & " ,Email =" & "'" & objInvoice.Email & "'"
    If objInvoice.dateofInvoice <> "" Then
        sql = sql & " ,Created_date =" & "'" & Format(objInvoice.dateofInvoice, "yyyy-mm-dd") & "'"
    Else
        sql = sql & " ,Created_date = NULL "
    End If
    
    If objInvoice.over_due_date <> "" Then
        sql = sql & " ,over_due_date =" & "'" & Format(objInvoice.over_due_date, "yyyy-mm-dd") & "'"
    Else
        sql = sql & " ,over_due_date = NULL "
    End If
    
    sql = sql & " WHERE ID = " & objInvoice.ID
    'Debug.Print (sql)
    dbConnection.Execute sql


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsInvoice_s", "UpdateInvoice", True)

End Function

Public Function getByInvoiceId(ByVal lngInvoiceID As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT * FROM Invoice WHERE id = " & lngInvoiceID
    
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

Public Function getByInvoiceNo(ByVal lngInvoiceNo As String) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT * FROM Invoice WHERE Invoice_no = '" & lngInvoiceNo & "'"
    
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly
    
    If rsLocal.EOF = True Then
        Set getByInvoiceNo = Nothing
    Else
        Set getByInvoiceNo = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsMember_s", "getByMemberId", True)

End Function

Public Function UpdateBalance(InvoiceId As Long, Balance As Double)
On Error GoTo ErrorHandler

    Dim sql As String

    
    sql = "UPDATE Invoice SET"
    sql = sql & " Balance =" & Balance
    sql = sql & " WHERE ID = " & InvoiceId
    dbConnection.Execute sql


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsMember_s", "UpdateExparydate", True)

End Function




