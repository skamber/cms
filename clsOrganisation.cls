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
    
    
    sql = "SELECT Action_Id, Action_name FROM actions " _
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
   ' Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getPracticeName", True)

End Function

Public Function getPayment() As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT Id, Payment_type FROM type"
   
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
    
    
    sql = "SELECT ID, Payment FROM payment_type"
   
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
    
    
    sql = "SELECT ID, AmountInWord FROM wordamount"
   
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
    
    
    sql = "SELECT Id, Collection_type FROM collectiontype"
   
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
    
    Case "CASHIN": sql = "SELECT Id, itemName FROM cashflowitem where type ='CASHIN'"
    Case "CASHOUT": sql = "SELECT Id, itemName FROM cashflowitem where type ='CASHOUT'"
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
    
    
    sql = "SELECT ItemCode, GST FROM cashflowitem WHERE ItemName ='" & lngItem & "'"
   
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
    
    
    sql = "SELECT ID, TermName FROM term"
   
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
    
    
    sql = "SELECT Action_id, Action_Name FROM actions"

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
    
    
    sql = "SELECT ID, Descriprtion ,GST FROM invoiceitemtype"
   
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
    
    
    sql = "SELECT ID, Donation FROM donation_type"
   
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
Public Function getChurchName(ByVal lngCityID As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT ID, Name FROM church WHERE cityId = " & lngCityID
   
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
       
   Debug.Print (Err.Description)
   ' Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getChurchName", True)

End Function

Public Function getCities() As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT ID, CityName, CountryId FROM city"
   
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly

    If rsLocal.EOF = True Then
        'MsgBox "No records found for Office.", vbExclamation
        Set getCities = Nothing
    Else
        Set getCities = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
       
   Debug.Print (Err.Description)
   ' Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getCities", True)

End Function

Public Function getCityNameById(ByVal cityId As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT CityName FROM city WHERE ID =" & cityId
   
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly

    If rsLocal.EOF = True Then
        'MsgBox "No records found for Office.", vbExclamation
        Set getCityNameById = Nothing
    Else
        Set getCityNameById = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
       
   Debug.Print (Err.Description)
   ' Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsOrganisation_s", "getCities", True)

End Function

Public Function getChurchNameById(ByVal lngId As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    
    
    sql = "SELECT Name FROM church WHERE id =" & lngId
   
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
