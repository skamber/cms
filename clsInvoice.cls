VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarID As Long 'local copy
Private mvarInvoice_no As String 'local copy
Private mvarRef As String 'local copy
Private mvardateofInvoice As String 'local copy
Private mvarover_due_date As String 'local copy
Private mvarName1 As String 'local copy
Private mvarName2 As String 'local copy
Private mvarAddress1 As String 'local copy
Private mvarAddress2 As String 'local copy
Private mvarAddress3 As String 'local copy
Private mvarTerms As String 'local copy
Private mvarTotal_amount As Currency 'local copy
Private mvarstate As String 'local copy
Private mvarBalance As Currency 'local copy
Private mvarPhone As String 'local copy
Private mvarMobile As String 'local copy
Private mvarChurchId As Long 'local copy
Private mvarEmail As String 'local copy
Private Const DATE_FORMAT = "dd/mm/yyyy"

Public Property Let Email(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Mobile = 5
    mvarEmail = vData
End Property

Public Property Get Email() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Mobile
    Email = mvarEmail
End Property

Public Property Let Mobile(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Mobile = 5
    mvarMobile = vData
End Property


Public Property Get Mobile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Mobile
    Mobile = mvarMobile
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


Public Property Let Phone(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Phone = 5
    mvarPhone = vData
End Property


Public Property Get Phone() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Phone
    Phone = mvarPhone
End Property



Public Property Let Balance(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Balance = 5
    mvarBalance = vData
End Property


Public Property Get Balance() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Balance
    Balance = mvarBalance
End Property



Public Property Let state(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.state = 5
    mvarstate = vData
End Property


Public Property Get state() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.state
    state = mvarstate
End Property



Public Property Let Total_amount(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Total_amount = 5
    mvarTotal_amount = vData
End Property


Public Property Get Total_amount() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Total_amount
    Total_amount = mvarTotal_amount
End Property



Public Property Let Terms(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Terms = 5
    mvarTerms = vData
End Property


Public Property Get Terms() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Terms
    Terms = mvarTerms
End Property



Public Property Let Address3(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Address3 = 5
    mvarAddress3 = vData
End Property


Public Property Get Address3() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Address3
    Address3 = mvarAddress3
End Property



Public Property Let Address2(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Address2 = 5
    mvarAddress2 = vData
End Property


Public Property Get Address2() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Address2
    Address2 = mvarAddress2
End Property



Public Property Let Address1(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Address1 = 5
    mvarAddress1 = vData
End Property


Public Property Get Address1() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Address1
    Address1 = mvarAddress1
End Property



Public Property Let Name2(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Name2 = 5
    mvarName2 = vData
End Property


Public Property Get Name2() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Name2
    Name2 = mvarName2
End Property



Public Property Let Name1(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Name1 = 5
    mvarName1 = vData
End Property


Public Property Get Name1() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Name1
    Name1 = mvarName1
End Property



Public Property Let over_due_date(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.over_due_date = 5
    mvarover_due_date = Format(vData, DATE_FORMAT)
End Property


Public Property Get over_due_date() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.over_due_date
    over_due_date = Format(mvarover_due_date, DATE_FORMAT)
End Property



Public Property Let dateofInvoice(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dateofInvoice = 5
    mvardateofInvoice = Format(vData, DATE_FORMAT)
End Property


Public Property Get dateofInvoice() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dateofInvoice
    dateofInvoice = Format(mvardateofInvoice, DATE_FORMAT)
End Property



Public Property Let Ref(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Ref = 5
    mvarRef = vData
End Property


Public Property Get Ref() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ref
    Ref = mvarRef
End Property



Public Property Let Invoice_no(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Invoice_no = 5
    mvarInvoice_no = vData
End Property


Public Property Get Invoice_no() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Invoice_no
    Invoice_no = mvarInvoice_no
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



