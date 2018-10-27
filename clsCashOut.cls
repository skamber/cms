VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCashOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarID As Long 'local copy
Private mvarITEM As String 'local copy
Private mvarDateofCashIn As String 'local copy
Private mvarAmount As Currency 'local copy
Private mvarGST As Currency 'local copy
Private mvarTotal_Amount As Currency 'local copy
Private mvarComment As String 'local copy
Private Const DATE_FORMAT = "dd/mm/yyyy"
'local variable(s) to hold property value(s)
Private mvarDateOfCashOut As Date 'local copy
'local variable(s) to hold property value(s)
Private mvarItem_code As String 'local copy
Public Property Let Item_code(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Item_code = 5
    mvarItem_code = vData
End Property


Public Property Get Item_code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Item_code
    Item_code = mvarItem_code
End Property



Public Property Let DateOfCashOut(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateOfCashOut = 5
    mvarDateOfCashOut = Format(vData, DATE_FORMAT)
End Property


Public Property Get DateOfCashOut() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateOfCashOut
    DateOfCashOut = Format(mvarDateOfCashOut, DATE_FORMAT)
End Property




Public Property Let Comment(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Comment = 5
    mvarComment = Trim(vData)
End Property


Public Property Get Comment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Comment
    mvarComment = Replace(mvarComment, "'", " ")
    Comment = mvarComment
End Property



Public Property Let Total_Amount(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Total_Amount = 5
    mvarTotal_Amount = vData
End Property


Public Property Get Total_Amount() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Total_Amount
    Total_Amount = mvarTotal_Amount
End Property



Public Property Let GST(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GST = 5
    mvarGST = vData
End Property


Public Property Get GST() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GST
    GST = mvarGST
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







Public Property Let ITEM(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ITEM = 5
    mvarITEM = vData
End Property


Public Property Get ITEM() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ITEM
    ITEM = mvarITEM
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



