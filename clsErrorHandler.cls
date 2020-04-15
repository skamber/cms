VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==============================================================================
'Synopsis:  Procedure obtains error details, writes to log, and displays to user.
'Parameters:
'   AppPath - required to be passed otherwise application will write error.log to
'                   path where DLL is started from (most likely client) therefore making
'                   it harder to obtain error details.
'   sVBNumber - Visual Basic generated err.number
'   sVBDescription - Visual Basic generated err.description
'   oConnection - enables rollback for any open transactions
'   sProcedureName - displays current procedure error generated from
'   sModuleName - display current module error generate from
'   bVisibleToUser - indicate if message displayed to user
'==============================================================================
Public Sub ErrorRoutine(sVBNumber As Long, _
                        Optional sVBDescription As String, _
                        Optional ByRef objConnection As ADODB.Connection, _
                        Optional sModuleName As String, _
                        Optional sProcedureName As String, _
                        Optional bVisibleToUser As Boolean)

On Error GoTo InternalError


    Dim iFileNumber As Integer
    Dim sFileName As String

    Dim oError As ADODB.Error
    Dim sDisplayMessage As String
    Dim sErrorMessage As String
    Dim bWriteError As Boolean

    Dim sNativeNumber As String
    Dim sNativeDescription As String


   'Create error.log text file (for developer error tracking)

    iFileNumber = FreeFile
    sFileName = App.Path & "\" & "error.log"

    '[Two types of errors - Visual Basic & Database Provider]
    'Only 1 error will write to VB, however multiple errors for provider will
    'write to connection - write all errors to error.log

    bWriteError = True

    If objConnection.Errors.Count = 0 Then

        'Write error detail to error.log text file
        Open sFileName For Append As #iFileNumber

        Print #iFileNumber, "Date: "; Tab(20); Format(Now, "dd-mmm-yyyy hh:mm:ss")
        Print #iFileNumber, "Procedure: "; Tab(20); sProcedureName
        Print #iFileNumber, "Module: "; Tab(20); sModuleName
        Print #iFileNumber, "Error Number: "; Tab(20); sVBNumber
        Print #iFileNumber, "Error Description: "; Tab(20); sVBDescription
        Print #iFileNumber, "Provider Number: "; Tab(20); sNativeNumber
        Print #iFileNumber, "Provider Description: "; Tab(20); sNativeDescription
        'Print #iFileNumber, "Prospect Id: "; Tab(20); gProspectId
        Print #iFileNumber, vbCrLf

        Close #iFileNumber

    Else

        For Each oError In objConnection.Errors

            '[Note: ORACLE returns number in Description - not NativeError]
            sNativeNumber = oError.NativeError
            sNativeDescription = oError.Description

            'Write error detail to error.log text file
            Open sFileName For Append As #iFileNumber

            Print #iFileNumber, "Date: "; Tab(20); Format(Now, "dd-mmm-yyyy hh:mm:ss")
            Print #iFileNumber, "Procedure: "; Tab(20); sProcedureName
            Print #iFileNumber, "Module: "; Tab(20); sModuleName
            Print #iFileNumber, "Error Number: "; Tab(20); sVBNumber
            Print #iFileNumber, "Error Description: "; Tab(20); sVBDescription
            Print #iFileNumber, "Provider Number: "; Tab(20); sNativeNumber
            Print #iFileNumber, "Provider Description: "; Tab(20); sNativeDescription
            'Print #iFileNumber, "Prospect Id: "; Tab(20); gProspectId
            Print #iFileNumber, vbCrLf

            Close #iFileNumber

        Next

    End If

    'Build error string for display to user
    sDisplayMessage = "An error has occurred during processing.  Please refer below for further details...  " _
                            & vbCrLf _
                            & sErrorMessage & vbCrLf _
                            & "Module: " & vbTab & sModuleName & vbCrLf _
                            & "Procedure: " & sProcedureName & vbCrLf


    If objConnection.Errors.Count = 0 Then
            sDisplayMessage = sDisplayMessage & vbCrLf _
                                    & "Error Number: " & sVBNumber & vbCrLf _
                                    & "Error Description: " & sVBDescription
        Else
            sDisplayMessage = sDisplayMessage & vbCrLf _
                                    & "Error Number: " & sNativeNumber & vbCrLf _
                                    & "Error Description: " & sNativeDescription
    End If


    'Display Form to inform users of type of error occurred.
    If bVisibleToUser = True Then MsgBox sDisplayMessage, vbExclamation, "CMS Error"

    'transaction rollback, clear error collection - if no transaction continue.
    If Not objConnection Is Nothing Then
        On Error Resume Next
        objConnection.RollbackTrans
        objConnection.Errors.Clear
    End If

    Exit Sub


InternalError:
    'if error occurred in errorhandler then close application
    If bVisibleToUser = True Then
        MsgBox "An error has occurred within the Error Handling routine.  The application will be shutdown.", vbExclamation
    End If
    Stop

End Sub


