VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarpayment As String 'local copy
Private mvartypes As String 'local copy
Private mvarDateofcollection As String 'local copy
Private mvarAmountinWords As String 'local copy
Private mvarAmount As Currency 'local copy
Private mvarcomments As String 'local copy
Private mvarUserName As String 'local copy
Private mvarcol_ID As Long 'local copy
Private mvarChurchId As Long 'local copy
Private Const DATE_FORMAT = "dd/mm/yyyy"



Public Property Let col_ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.col_ID = 5
    mvarcol_ID = vData
End Property


Public Property Get col_ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.col_ID
    col_ID = mvarcol_ID
End Property

Public Property Let ChurchId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReceiptNo = 5
    mvarChurchId = vData
End Property

Public Property Get ChurchId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReceiptNo
    ChurchId = mvarChurchId
End Property

Public Property Let UserName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.comments = 5
    mvarUserName = vData
End Property


Public Property Get UserName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.comments
    UserName = mvarUserName
End Property

Public Property Let comments(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.comments = 5
    mvarcomments = vData
End Property


Public Property Get comments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.comments
    comments = mvarcomments
End Property



Public Property Let Amount(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Amount = 5
    mvarAmount = vData
End Property


Public Property Get Amount() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Amount
    Amount = mvarAmount
End Property



Public Property Let AmountinWords(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AmountinWords = 5
    mvarAmountinWords = vData
End Property


Public Property Get AmountinWords() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AmountinWords
    AmountinWords = mvarAmountinWords
End Property



Public Property Let Dateofcollection(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Dateofcollection = 5
    mvarDateofcollection = Format(vData, DATE_FORMAT)
End Property


Public Property Get Dateofcollection() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Dateofcollection
    Dateofcollection = Format(mvarDateofcollection, DATE_FORMAT)
End Property



Public Property Let payment(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.payment = 5
    mvarpayment = vData
End Property


Public Property Get payment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.payment
    payment = mvarpayment
End Property




Public Property Let types(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.payment = 5
    mvartypes = vData
End Property


Public Property Get types() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.payment
    types = mvartypes
End Property


