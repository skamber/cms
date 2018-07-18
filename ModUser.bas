Attribute VB_Name = "ModUser"
Option Explicit

Public Function getUserPriveleges(ByVal UserId As Long) As Boolean
On Error GoTo ErrorHandler

    Dim sql As String
    Dim Count As Byte
    Count = 1
    sql = "SELECT * FROM PRIVILEGE WHERE User_Id = " & UserId
    sql = sql & " Order by Action_Id"
    Set Userprivilege = New ADODB.Recordset
        Userprivilege.Open sql, objConnection, adOpenForwardOnly, adLockReadOnly
    
    If Userprivilege.EOF = True Then
        MsgBox "No Privilege found for this user.", vbExclamation
        getUserPriveleges = False
    Else
        getUserPriveleges = True
  
        Do While Not Userprivilege.EOF
            Permissions(Count, 1) = (Userprivilege!read)
            Permissions(Count, 2) = (Userprivilege!Create)
            Permissions(Count, 3) = (Userprivilege!Edit)
            Permissions(Count, 4) = (Userprivilege!Delete)
            Count = Count + 1
            Userprivilege.MoveNext
        Loop
    End If
   
Exit Function
ErrorHandler:
    Call objError.ErrorRoutine(Err.Number, Err.Description, objConnection, "moduser", "getUserPriveleges", True)

End Function
