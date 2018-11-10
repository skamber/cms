VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'local variable(s) to hold property value(s)
Private mvarmno As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarSurname As String 'local copy
Private mvarGivenName As String 'local copy
Private mvarFullName As String 'local copy
Private mvarSpouseName As String 'local copy
Private mvarAddress1 As String 'local copy
Private mvarAddress2 As String 'local copy
Private mvarState As String 'local copy
Private mvarPostCode As String 'local copy
Private mvarPhone As String 'local copy
Private mvarMembershipExpiary As String 'local copy
Private mvarcreated_date As String
Private mvarDateOfBirth As String
Private mvarcomments As String 'local copy
Private mvarSTATUS As String 'local copy
Private mvarEMAIL As String 'local copy
Private Const DATE_FORMAT = "dd/mm/yyyy"
'local variable(s) to hold property value(s)
Private mvarInitial As String 'local copy
Private mvarCityId As Long




Public Property Let DateOfBirth(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MembershipExpiary = 5
    mvarDateOfBirth = Format(vData, DATE_FORMAT)
End Property


Public Property Get DateOfBirth() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MembershipExpiary
    DateOfBirth = Format(mvarDateOfBirth, DATE_FORMAT)
End Property

Public Property Let Created_Date(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MembershipExpiary = 5
    mvarcreated_date = Format(vData, DATE_FORMAT)
End Property


Public Property Get Created_Date() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MembershipExpiary
    Created_Date = Format(mvarcreated_date, DATE_FORMAT)
End Property


Public Property Let Initial(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Initial = 5
    mvarInitial = vData
End Property


Public Property Get Initial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Initial
    Initial = mvarInitial
End Property




Public Property Let STATUS(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.STATUS = 5
    mvarSTATUS = vData
End Property


Public Property Get STATUS() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.STATUS
    STATUS = mvarSTATUS
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



Public Property Let MembershipExpiary(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MembershipExpiary = 5
    mvarMembershipExpiary = Format(vData, DATE_FORMAT)
End Property


Public Property Get MembershipExpiary() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MembershipExpiary
    MembershipExpiary = Format(mvarMembershipExpiary, DATE_FORMAT)
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



Public Property Let PostCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PostCode = 5
    mvarPostCode = vData
End Property


Public Property Get PostCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PostCode
    PostCode = mvarPostCode
End Property



Public Property Let State(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.State = 5
    mvarState = vData
End Property


Public Property Get State() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.State
    State = mvarState
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



Public Property Let SpouseName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SpouseName = 5
    mvarSpouseName = vData
End Property


Public Property Get SpouseName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SpouseName
    SpouseName = mvarSpouseName
End Property



Public Property Let FullName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FullName = 5
    mvarFullName = vData
End Property


Public Property Get FullName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FullName
    FullName = mvarFullName
End Property



Public Property Let GivenName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GivenName = 5
    mvarGivenName = vData
End Property


Public Property Get GivenName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GivenName
    GivenName = mvarGivenName
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



Public Property Let mno(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.mno = 5
    mvarmno = vData
End Property


Public Property Get mno() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.mno
    mno = mvarmno
End Property


Public Property Let Email(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Surname = 5
    mvarEMAIL = vData
End Property


Public Property Get Email() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Surname
    Email = mvarEMAIL
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
