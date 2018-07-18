VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUserPermissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' (c) 2002-2002 Marsh Ltd.                                           }
' All Rights Reserved.                                                   }
 
 
 
' Update History
 
 
'Date        By     Comment

'13-May-02   SAM    Create the unit
'<!<CHECKOUT>!>

'Backed up to 4532 on 16-May-02 by SAM
'<!<PREVIOUS_VERSIONS>!>
'}
Option Explicit
'local variable(s) to hold property value(s)
'local variable(s) to hold property value(s)
Private mvarUserId As Long 'local copy
Private mvarActionId As Long 'local copy
Private mvarReadPermissions As String 'local copy
Private mvarWritePermissions As String 'local copy
Private mvarUpdatePermissions As String 'local copy
Private mvarDeletePermissions As String 'local copy
'local variable(s) to hold property value(s)




Public Property Let DeletePermissions(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DeletePermissions = 5
    mvarDeletePermissions = vData
End Property


Public Property Get DeletePermissions() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeletePermissions
    DeletePermissions = mvarDeletePermissions
End Property



Public Property Let UpdatePermissions(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UpdatePermissions = 5
    mvarUpdatePermissions = vData
End Property


Public Property Get UpdatePermissions() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UpdatePermissions
    UpdatePermissions = mvarUpdatePermissions
End Property



Public Property Let WritePermissions(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WritePermissions = 5
    mvarWritePermissions = vData
End Property


Public Property Get WritePermissions() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WritePermissions
    WritePermissions = mvarWritePermissions
End Property



Public Property Let ReadPermissions(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ReadPermissions = 5
    mvarReadPermissions = vData
End Property


Public Property Get ReadPermissions() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReadPermissions
    ReadPermissions = mvarReadPermissions
End Property



Public Property Let ActionId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ActionId = 5
    mvarActionId = vData
End Property


Public Property Get ActionId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ActionId
    ActionId = mvarActionId
End Property



Public Property Let UserId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserId = 5
    mvarUserId = vData
End Property


Public Property Get UserId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserId
    UserId = mvarUserId
End Property



