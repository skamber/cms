VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrganisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private dbConnection As ADODB.Connection

Public Property Set DatabaseConnection(vData As ADODB.Connection)
    Set dbConnection = vData
End Property

Public Function getActionId(strActionName As String) As Long
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT Action_Id, Action_name FROM Actions " _
            & "WHERE Action_name = " & "'" & strActionName & "'"

    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly

    If rsLocal.EOF = True Then
        getActionId = 0
    Else
        getActionId = rsLocal!Action_Id
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getPracticeName", True)

End Function

Public Function getPayment() As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT Id, Payment_type FROM Type"
   
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly

    If rsLocal.EOF = True Then
        'MsgBox "No records found for Office.", vbExclamation
        Set getPayment = Nothing
    Else
        Set getPayment = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getPayment", True)

End Function

Public Function getPaymentType() As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT ID, Payment FROM Payment_type"
   
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly

    If rsLocal.EOF = True Then
        'MsgBox "No records found for Office.", vbExclamation
        Set getPaymentType = Nothing
    Else
        Set getPaymentType = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getPaymentType", True)

End Function

Public Function getAmountType() As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT ID, AmountInWord FROM WordAmount"
   
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly

    If rsLocal.EOF = True Then
        'MsgBox "No records found for Office.", vbExclamation
        Set getAmountType = Nothing
    Else
        Set getAmountType = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getPaymentType", True)

End Function

Public Function getCollectionType() As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT Id, Collection_type FROM CollectionType"
   
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly

    If rsLocal.EOF = True Then
        'MsgBox "No records found for Office.", vbExclamation
        Set getCollectionType = Nothing
    Else
        Set getCollectionType = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getPayment", True)

End Function


Public Function getCashflowItem(CashFlowType As String) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    Select Case CashFlowType
    
    Case "CASHIN": sql = "SELECT Id, itemName FROM CashflowItem where type ='CASHIN'"
    Case "CASHOUT": sql = "SELECT Id, itemName FROM CashflowItem where type ='CASHOUT'"
    End Select
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly

    If rsLocal.EOF = True Then
        'MsgBox "No records found for Office.", vbExclamation
        Set getCashflowItem = Nothing
    Else
        Set getCashflowItem = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getPayment", True)

End Function

Public Function getCashflowItemcodeandGst(ByVal lngItem As String) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT ItemCode, GST FROM CashflowItem WHERE ItemName ='" & lngItem & "'"
   
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly

    If rsLocal.EOF = True Then
        'MsgBox "No records found for Office.", vbExclamation
        Set getCashflowItemcodeandGst = Nothing
    Else
        Set getCashflowItemcodeandGst = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getCashInItemcodeandGst", True)

End Function



Public Function getTerms() As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT ID, TermName FROM Term"
   
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly

    If rsLocal.EOF = True Then
        'MsgBox "No records found for Office.", vbExclamation
        Set getTerms = Nothing
    Else
        Set getTerms = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getPayment", True)

End Function

Public Function getActions() As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT Action_id, Action_Name FROM Actions"

    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly

    If rsLocal.EOF = True Then
       Set getActions = Nothing
    Else
       Set getActions = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getPracticeName", True)

End Function

Public Function getInvoiceItemType() As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT ID, Descriprtion ,GST FROM InvoiceItemType"
   
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly

    If rsLocal.EOF = True Then
        'MsgBox "No records found for Office.", vbExclamation
        Set getInvoiceItemType = Nothing
    Else
        Set getInvoiceItemType = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getPayment", True)

End Function

Public Function getDonationType() As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT ID, Donation FROM Donation_type"
   
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly

    If rsLocal.EOF = True Then
        'MsgBox "No records found for Office.", vbExclamation
        Set getDonationType = Nothing
    Else
        Set getDonationType = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getPaymentType", True)

End Function
Public Function getChurchName() As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT ID, Name FROM Church"
   
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly

    If rsLocal.EOF = True Then
        'MsgBox "No records found for Office.", vbExclamation
        Set getChurchName = Nothing
    Else
        Set getChurchName = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getChurchName", True)

End Function

Public Function getChurchNameById(ByVal lngId As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT Name FROM Church WHERE id =" & lngId
   
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly

    If rsLocal.EOF = True Then
        'MsgBox "No records found for Office.", vbExclamation
        Set getChurchNameById = Nothing
    Else
        Set getChurchNameById = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getChurchName", True)

End Function
