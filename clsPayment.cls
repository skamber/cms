VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarMNo As Long 'local copy
Private mvarPayment As String 'local copy
Private mvarPaymentType As String 'local copy
Private mvarDonationType As String 'local copy
Private mvarMemberEffective As String 'local copy
Private mvarMemberExpiary As String 'local copy
Private mvarDateofPayment As String 'local copy
Private mvarAmountinWords As String 'local copy
Private mvarAmount As Currency 'local copy
Private mvarcomments As String 'local copy
Private mvarReceiptNo As Long 'local copy
Private mvarAccountname As String 'local copy
Private Const DATE_FORMAT = "dd/mm/yyyy"
Private mvarChurchId As Long 'local copy



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


Public Property Let ReceiptNo(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReceiptNo = 5
    mvarReceiptNo = vData
End Property


Public Property Get ReceiptNo() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReceiptNo
    ReceiptNo = mvarReceiptNo
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



Public Property Let DateofPayment(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DateofPayment = 5
    mvarDateofPayment = Format(vData, DATE_FORMAT)
End Property


Public Property Get DateofPayment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DateofPayment
    DateofPayment = Format(mvarDateofPayment, DATE_FORMAT)
End Property



Public Property Let MemberExpiary(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MemberExpiary = 5
    mvarMemberExpiary = Format(vData, DATE_FORMAT)
End Property


Public Property Get MemberExpiary() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MemberExpiary
    MemberExpiary = Format(mvarMemberExpiary, Dateformat)
End Property



Public Property Let MemberEffective(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MemberEffective = 5
    mvarMemberEffective = Format(vData, DATE_FORMAT)
End Property


Public Property Get MemberEffective() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MemberEffective
    MemberEffective = Format(mvarMemberEffective, DATE_FORMAT)
End Property



Public Property Let PaymentType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PaymentType = 5
    mvarPaymentType = vData
End Property


Public Property Get PaymentType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PaymentType
    PaymentType = mvarPaymentType
End Property


Public Property Let DonationType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PaymentType = 5
    mvarDonationType = vData
End Property


Public Property Get DonationType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PaymentType
    DonationType = mvarDonationType
End Property



Public Property Let Payment(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Payment = 5
    mvarPayment = vData
End Property


Public Property Get Payment() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Payment
    Payment = mvarPayment
End Property



Public Property Let MNo(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MNo = 5
    mvarMNo = vData
End Property


Public Property Get MNo() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MNo
    MNo = mvarMNo
End Property

