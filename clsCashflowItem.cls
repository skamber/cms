VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCashflowItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarID As Long 'local copy
Private mvarItemName As String 'local copy
Private mvarItemCode As String 'local copy
Private mvarGST As Double 'local copy
Private mvarItemType As String 'local copy
Public Property Let ItemType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ItemType = 5
    mvarItemType = vData
End Property


Public Property Get ItemType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ItemType
    ItemType = mvarItemType
End Property



Public Property Let GST(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GST = 5
    mvarGST = vData
End Property


Public Property Get GST() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GST
    GST = mvarGST
End Property



Public Property Let ItemCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ItemCode = 5
    mvarItemCode = vData
End Property


Public Property Get ItemCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ItemCode
    ItemCode = mvarItemCode
End Property



Public Property Let ItemName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ItemName = 5
    mvarItemName = vData
End Property


Public Property Get ItemName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ItemName
    ItemName = mvarItemName
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



