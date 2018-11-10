VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMember_s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarNewMemberId As Long 'local copy
Private dbConnection As ADODB.Connection
Public Property Set DatabaseConnection(vData As ADODB.Connection)
    Set dbConnection = vData
End Property
Public Property Let NewMemberId(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NewMemberId = 5
    mvarNewMemberId = vData
End Property


Public Property Get NewMemberId() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NewMemberId
    NewMemberId = mvarNewMemberId
End Property

Public Function InsertMember(ByVal objMember As CMSMember.clsMember)
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
        
   
        'get next number -- cannot use AutoNumber function
        sql = "SELECT MAX(MNO) FROM member"
        Set rsLocal = New ADODB.Recordset
            rsLocal.Open sql, dbConnection
            
        If IsNull(rsLocal.Fields(0)) Then
            NewMemberId = 1
        Else
            NewMemberId = rsLocal.Fields(0) + 1
        End If
        rsLocal.Close
        Set rsLocal = Nothing

    
    
    'insert new record
    sql = "INSERT INTO member" _
                & " (MNo, Mr, Surname, Given_Name, Full_Name, Spouse_Name, Address1, CITY_ID" _
                & " ,Address2, STATE,PostCode, Phone,  comments" _
                & " ,STATUS,Email ,joining_date,Membership_Expiary,DATE_OF_BIRTH)" _
                & " VALUES" _
                & " (" & NewMemberId _
                & ",'" & objMember.Initial & "'" _
                & ",'" & objMember.Surname & "','" & objMember.GivenName & "'" _
                & ",'" & objMember.FullName & "','" & objMember.SpouseName & "','" _
                & objMember.Address1 & "'," & objMember.cityId _
                & ",'" & objMember.Address2 & "','" & objMember.State & "'" _
                & ",'" & objMember.PostCode & "'" _
                & ",'" & objMember.Phone & "'" _
                & ",'" & objMember.comments & "'" _
                & ",'" & objMember.STATUS & "'" _
                & ",'" & objMember.Email & "'"
                
                If objMember.Created_Date = "" Then
                  sql = sql & ", NULL"
                Else
                  sql = sql & ",'" & Format(objMember.Created_Date, "yyyy-mm-dd") & "'"
                End If
                
                If objMember.MembershipExpiary = "" Then
                  sql = sql & ", NULL"
                Else
                  sql = sql & ",'" & Format(objMember.MembershipExpiary, "yyyy-mm-dd") & "'"
                End If
                
                If objMember.DateOfBirth = "" Then
                  sql = sql & ", NULL)"
                Else
                  sql = sql & ",'" & Format(objMember.DateOfBirth, "yyyy-mm-dd") & "')"
                End If
    Debug.Print (sql)
    dbConnection.Execute sql

Exit Function
ErrorHandler:
   
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsmember_s", "Insertmember", True)

End Function

Public Function UpdateMember(ByVal objMember As CMSMember.clsMember)
On Error GoTo ErrorHandler

    Dim sql As String

    
    sql = "UPDATE member SET" _
                & " Mr =" & "'" & objMember.Initial & "'" _
                & " ,Surname =" & "'" & objMember.Surname & "'" _
                & " ,Given_Name =" & "'" & objMember.GivenName & "'" _
                & " ,Full_Name =" & "'" & objMember.FullName & "'" _
                & " ,Spouse_Name =" & "'" & objMember.SpouseName & "'" _
                & " ,Address1 =" & "'" & objMember.Address1 & "'" _
                & " ,Address2 =" & "'" & objMember.Address2 & "'" _
                & " ,STATE =" & "'" & objMember.State & "'" _
                & " ,PostCode =" & "'" & objMember.PostCode & "'" _
                & " ,Phone =" & "'" & objMember.Phone & "'" _
                & " ,comments =" & "'" & objMember.comments & "'" _
                & " ,STATUS =" & "'" & objMember.STATUS & "'" _
                & " ,EMAIL =" & "'" & objMember.Email & "'"
                
    If objMember.Created_Date <> "" Then
        sql = sql & " ,joining_date =" & "'" & Format(objMember.Created_Date, "yyyy-mm-dd") & "'"
    Else
        sql = sql & " ,joining_date = NULL "
    End If
    
                
    If objMember.MembershipExpiary <> "" Then
        sql = sql & " ,Membership_Expiary =" & "'" & Format(objMember.MembershipExpiary, "yyyy-mm-dd") & "'"
    Else
        sql = sql & " ,Membership_Expiary = NULL "
    End If
    
    If objMember.DateOfBirth <> "" Then
        sql = sql & " ,DATE_OF_BIRTH =" & "'" & Format(objMember.DateOfBirth, "yyyy-mm-dd") & "'"
    Else
        sql = sql & " ,DATE_OF_BIRTH = NULL "
    End If
    
    sql = sql & " WHERE Mno = " & objMember.mno
    'Debug.Print (sql)
    dbConnection.Execute sql


Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsMember_s", "UpdateMember", True)

End Function

Public Function UpdateExparydate(MemberNumber As String, ExparyDate As String)
On Error GoTo ErrorHandler

    Dim sql As String

    
    sql = "UPDATE member SET" _
                
                
    If ExparyDate <> "" Then
        sql = sql & " Membership_Expiary =" & "'" & Format(ExparyDate, "yyyy-mm-dd") & "'"
    Else
        sql = sql & " ,Membership_Expiary = NULL "
    End If
    
    sql = sql & " WHERE Mno = " & MemberNumber
    
    dbConnection.Execute sql


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsMember_s", "UpdateExparydate", True)

End Function
Public Function getByMemberId(ByVal lngMemberID As Long) As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT * FROM member WHERE Mno = " & lngMemberID
    
    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly
    
    If rsLocal.EOF = True Then
        MsgBox "No records found for This Member.", vbExclamation
        Set getByMemberId = Nothing
    Else
        Set getByMemberId = rsLocal
    End If
    
    Set rsLocal = Nothing


Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsMember_s", "getByMemberId", True)

End Function

Public Function GetPostCodes() As ADODB.Recordset
On Error GoTo ErrorHandler

    Dim sql As String
    Dim rsLocal As ADODB.Recordset
    

    sql = "SELECT distinct postCode FROM member"

    Set rsLocal = New ADODB.Recordset
        rsLocal.Open sql, dbConnection, adOpenForwardOnly, adLockReadOnly
    
    If rsLocal.EOF = True Then
        Set GetPostCodes = Nothing
    Else
        Set GetPostCodes = rsLocal
    End If

    Set rsLocal = Nothing
    

Exit Function
ErrorHandler:
'    Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsMember_s", "GetPostCode", True)

End Function

Public Function UpdateStatus(MemberNumber As String, STATUS As String)
On Error GoTo ErrorHandler

    Dim sql As String

    
    sql = "UPDATE member SET" _
                
                
    If STATUS <> "" Then
        sql = sql & " STATUS =" & "'" & STATUS & "'"
    Else
        sql = sql & " ,Status = NULL "
    End If
    
    sql = sql & " WHERE Mno = " & MemberNumber
    
    dbConnection.Execute sql


Exit Function
ErrorHandler:
    'Call objError.ErrorRoutine(Err.Number, Err.Description, dbConnection, "clsMember_s", "UpdateExparydate", True)

End Function

