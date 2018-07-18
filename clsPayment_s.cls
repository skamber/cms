VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPayment_s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarNewPaymentId As Long 'local copy
Private dbConnection As ADODB.Connection
Public Property Set DatabaseConnection(vData As ADODB.Connection)
    Set dbConnection = vData
End Property
Public Property Let NewPaymentId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewPaymentId = 5
    mvarNewPaymentId = vData
End Property


Public Property Get NewPaymentId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewPaymentId
    NewPaymentId = mvarNewPaymentId
End Property

Public Function InsertPayment(ByVal objpayment As CMSPayment.clsPayment)
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
   
    
   
        'get next number -- cannot use AutoNumber function
        sql = "SELECT MAX(Receipt_No) FROM PAYMENT"
        Set rsLocal = New ADODB.Recordset
            rsLocal.Open sql, dbConnection
            
        If IsNull(rsLocal.Fields(0)) Then
            NewPaymentId = 1
        Else
            NewPaymentId = rsLocal.Fields(0) + 1
        End If
        rsLocal.Close
        Set rsLocal = Nothing

    
    
    'insert new record
    sql = "INSERT INTO PAYMENT" _
                & " (MNo, type,Payment, donation_type,Amount_in_Words, church_id,Amount" _
                & ", Receipt_No, user_name, comments,Member_Effective, Member_Expiary" _
                & ", Date_Of_Payment)" _
                & " VALUES" _
                & " (" & objpayment.MNo _
                & ",'" & objpayment.Payment & "'" _
                & ",'" & objpayment.PaymentType & "'" _
                & ",'" & objpayment.DonationType & "'" _
                & ",'" & objpayment.AmountinWords & "'" _
                & "," & objpayment.ChurchId _
                & "," & objpayment.Amount _
                & "," & NewPaymentId & ",'" & objpayment.Accountname & "'"
                
                If objpayment.comments = "" Then
                sql = sql & ", NULL"
                Else
                sql = sql & ",'" & objpayment.comments & "'"
                End If
                
                If objpayment.MemberEffective = "" Then
                sql = sql & ", NULL"
                Else
                sql = sql & ",'" & Format(objpayment.MemberEffective, "dd-mmm-yyyy") & "'"
                End If
                
                If objpayment.MemberExpiary = "" Then
                sql = sql & ", NULL"
                Else
                sql = sql & ",'" & Format(objpayment.MemberExpiary, "dd-mmm-yyyy") & "'"
                End If
                
                If objpayment.DateofPayment = "" Then
                sql = sql & ", NULL)"
                Else
                sql = sql & ",'" & Format(objpayment.DateofPayment, "dd-mmm-yyyy") & "')"
                End If
                
    dbConnection.Execute sql

Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsmember_s", "Insertmember", True)

End Function

Public Function UpdatePayment(ByVal objpayment As CMSPayment.clsPayment)
On Error GoTo ErrorHandler

    Dim sql As String

    
    sql = "UPDATE PAYMENT SET" _
                & " MNo =" & objpayment.MNo _
                & " ,church_id =" & objpayment.ChurchId _
                & " ,Payment =" & "'" & objpayment.PaymentType & "'" _
                & " ,Donation_type =" & "'" & objpayment.DonationType & "'" _
                & " ,Type =" & "'" & objpayment.Payment & "'" _
                & " ,Amount_in_Words =" & "'" & objpayment.AmountinWords & "'" _
                & " ,Amount =" & objpayment.Amount _
                & " ,user_name =" & "'" & objpayment.Accountname & "'"
                
    If objpayment.comments <> "" Then
       sql = sql & ", comments =" & "'" & objpayment.comments & "'"
     Else
       sql = sql & ",comments = NULL "
    End If
                
    If objpayment.MemberEffective <> "" Then
        sql = sql & " ,Member_Effective =" & "'" & Format(objpayment.MemberEffective, "dd-mmm-yyyy") & "'"
    Else
        sql = sql & " ,Member_Effective = NULL "
    End If
    
    If objpayment.MemberExpiary <> "" Then
        sql = sql & " ,Member_Expiary =" & "'" & Format(objpayment.MemberExpiary, "dd-mmm-yyyy") & "'"
    Else
        sql = sql & " ,Member_Expiary = NULL "
    End If
    
    If objpayment.DateofPayment <> "" Then
        sql = sql & " ,Date_Of_Payment =" & "'" & Format(objpayment.DateofPayment, "dd-mmm-yyyy") & "'"
    Else
        sql = sql & " ,Date_Of_Payment = NULL "
    End If
    
    sql = sql & " WHERE Receipt_No = " & objpayment.ReceiptNo
    
    dbConnection.Execute sql


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsPayment_s", "UpdatePayment", True)

End Function


Public Function getByPaymentId(ByVal lngPaymentID As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT * FROM PAYMENT WHERE Receipt_No = " & lngPaymentID
    
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly
    
    If rsLocal.EOF = True Then
        MsgBox "No record found for This Payment.", vbExclamation
        Set getByPaymentId = Nothing
    Else
        Set getByPaymentId = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsPayment_s", "getByPaymentId", True)

End Function



