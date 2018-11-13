VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarLogonPassword As String 'local copy
Private mvarUserName As String 'local copy
Private mvarUserId As Long 'local copy
Private mvarLogonId As String 'local copy
Private mvarPasswordLastChange As Date 'local copy
'local variable(s) to hold property value(s)
'local variable(s) to hold property value(s)
Private mvarPracticeAdministrator As String 'local copy
Private Const DATE_FORMAT = "dd/mm/yyyy"
'local variable(s) to hold property value(s)
Private mvarSystemManager As String 'local copy
Private mvarReportView As String 'local copy
Private mvarCityId As Long

Public Property Let ReportView(ByVal vData As String)

    mvarReportView = vData
End Property


Public Property Get ReportView() As String
    ReportView = mvarReportView
End Property
Public Property Let SystemManager(ByVal vData As String)

    mvarSystemManager = vData
End Property


Public Property Get SystemManager() As String

    SystemManager = mvarSystemManager
End Property





Public Property Let UserId(ByVal vData As Long)
    mvarUserId = vData
End Property

Public Property Get UserId() As Long
    UserId = mvarUserId
End Property

Public Property Let PasswordLastChange(ByVal vData As Date)
    mvarPasswordLastChange = vData
End Property

Public Property Get PasswordLastChange() As Date
    PasswordLastChange = mvarPasswordLastChange
End Property

Public Property Let LogonId(ByVal vData As String)
    mvarLogonId = vData
End Property

Public Property Get LogonId() As String
    LogonId = mvarLogonId
End Property

Public Property Let Username(ByVal vData As String)
    mvarUserName = vData
End Property

Public Property Get Username() As String
    Username = mvarUserName
End Property

Public Property Let LogonPassword(ByVal vData As String)
    mvarLogonPassword = vData
End Property

Public Property Get LogonPassword() As String
    LogonPassword = mvarLogonPassword
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
