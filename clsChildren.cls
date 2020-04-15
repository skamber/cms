VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChildren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarMNo As Long 'local copy
Private mvarChildNo As Long 'local copy
Private mvarFirstName As String 'local copy
Private mvarSurname As String 'local copy
Private mvarGenda As String 'local copy
Private mvarMember As String 'local copy
Private mvarBirthDate As String 'local copy
Private mvarEmail As String 'local copy
Private mvarMemo As String 'local copy
Private mvarCityId As Long
Private mvarMobile As String 'local copy
Private Const DATE_FORMAT = "dd/mm/yyyy"
Public Property Let BirthDate(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BirthDate = 5
    mvarBirthDate = vData
End Property


Public Property Get BirthDate() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BirthDate
    BirthDate = Format(mvarBirthDate, DATE_FORMAT)
End Property



Public Property Let Member(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Member = 5
    mvarMember = Format(vData, DATE_FORMAT)
End Property


Public Property Get Member() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Member
    Member = mvarMember
End Property



Public Property Let Genda(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Genda = 5
    mvarGenda = vData
End Property


Public Property Get Genda() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Genda
    Genda = mvarGenda
End Property



Public Property Let Surname(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Surname = 5
    mvarSurname = vData
End Property


Public Property Get Surname() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Surname
    Surname = mvarSurname
End Property



Public Property Let FirstName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FirstName = 5
    mvarFirstName = vData
End Property


Public Property Get FirstName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FirstName
    FirstName = mvarFirstName
End Property

Public Property Let Email(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FirstName = 5
    mvarEmail = vData
End Property


Public Property Get Email() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FirstName
    Email = mvarEmail
End Property
Public Property Let Memo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FirstName = 5
    mvarMemo = vData
End Property


Public Property Get Memo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FirstName
    Memo = mvarMemo
End Property


Public Property Let ChildNo(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ChildNo = 5
    mvarChildNo = vData
End Property


Public Property Get ChildNo() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ChildNo
    ChildNo = mvarChildNo
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

Public Property Let cityId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mno = 5
    mvarCityId = vData
End Property


Public Property Get cityId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mno
    cityId = mvarCityId
End Property

Public Property Let Mobile(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Phone = 5
    mvarMobile = vData
End Property

Public Property Get Mobile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Phone
    Mobile = mvarMobile
End Property
