VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarID As Long 'local copy
Private mvarINV_NO As String 'local copy
Private mvarINV_ID As Long 'local copy
Private mvaramount As Currency 'local copy
Private mvarDateOfReceipt As String 'local copy
Private mvarAccountname As String 'local copy
Private mvarChequeNo As String 'local copy
Private Const DATE_FORMAT = "dd/mm/yyyy"
Private mvarChurchId As Long 'local copy
Private mvarcomments As String 'local copy

Public Property Let ChequeNumber(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Accountname = 5
    mvarChequeNo = vData
End Property


Public Property Get ChequeNumber() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Accountname
    ChequeNumber = mvarChequeNo
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
Public Property Let Accountname(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Accountname = 5
    mvarAccountname = vData
End Property


Public Property Get Accountname() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Accountname
    Accountname = mvarAccountname
End Property


Public Property Let DateOfReceipt(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateOfReceipt = 5
    mvarDateOfReceipt = Format(vData, DATE_FORMAT)
End Property


Public Property Get DateOfReceipt() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateOfReceipt
    DateOfReceipt = Format(mvarDateOfReceipt, DATE_FORMAT)
End Property



Public Property Let amount(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.amount = 5
    mvaramount = vData
End Property


Public Property Get amount() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.amount
    amount = mvaramount
End Property



Public Property Let INV_NO(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.INV_NO = 5
    mvarINV_NO = vData
End Property


Public Property Get INV_NO() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.INV_NO
    INV_NO = mvarINV_NO
End Property


Public Property Let INV_ID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.INV_NO = 5
    mvarINV_ID = vData
End Property


Public Property Get INV_ID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.INV_NO
    INV_ID = mvarINV_ID
End Property

Public Property Let ID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ID = 5
    mvarID = vData
End Property


Public Property Get ID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ID
    ID = mvarID
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
