VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInvoiceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'

'local variable(s) to hold property value(s)
Private mvarInvoice_id As Long 'local copy
Private mvarDescription As String 'local copy
Private mvarAmount As Currency 'local copy
Private mvarGstAmount As Currency 'local copy
Private mvartotal_amount As Currency 'local copy
'local variable(s) to hold property value(s)
Private mvarItem_Id As Long 'local copy
Public Property Let Item_Id(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Item_Id = 5
    mvarItem_Id = vData
End Property


Public Property Get Item_Id() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Item_Id
    Item_Id = mvarItem_Id
End Property



Public Property Let total_amount(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.total_amount = 5
    mvartotal_amount = vData
End Property


Public Property Get total_amount() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.total_amount
    total_amount = mvartotal_amount
End Property



Public Property Let GstAmount(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GstAmount = 5
    mvarGstAmount = vData
End Property


Public Property Get GstAmount() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GstAmount
    GstAmount = mvarGstAmount
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



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let Invoice_id(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Invoice_id = 5
    mvarInvoice_id = vData
End Property


Public Property Get Invoice_id() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Invoice_id
    Invoice_id = mvarInvoice_id
End Property



