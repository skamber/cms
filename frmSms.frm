VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSms 
   ClientHeight    =   10575
   ClientLeft      =   660
   ClientTop       =   2040
   ClientWidth     =   15300
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10575
   ScaleWidth      =   15300
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   18
      Top             =   3240
      Width           =   1215
   End
   Begin MSMask.MaskEdBox dteToDateOfBirth 
      Height          =   315
      Left            =   6840
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox dteFromDateOfBirth 
      Height          =   315
      Left            =   4920
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtMnoNumber 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton CmdSendSms 
      Caption         =   "Send SMS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      TabIndex        =   7
      Top             =   8760
      Width           =   1575
   End
   Begin VB.TextBox txtSmsMessage 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   8760
      Width           =   10335
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00800000&
      FillStyle       =   6  'Cross
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   15240
      TabIndex        =   10
      Top             =   0
      Width           =   15300
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "SMS MEMBER"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.ComboBox cmbTypeSearch 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmSms.frx":0000
      Left            =   840
      List            =   "frmSms.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   4335
      Left            =   360
      TabIndex        =   8
      Top             =   3840
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   7646
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSMask.MaskEdBox dteExpairyDate 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      ClipMode        =   1
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label8 
      Caption         =   "To Date of Birth"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   17
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "From Date of Birth"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Mno Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Member Expiry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Message To Send"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   8520
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "List To Send"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Search Criteria"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   9
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "frmSms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InternetCanonicalizeUrl Lib "Wininet.dll" Alias "InternetCanonicalizeUrlW" ( _
    ByVal lpszUrl As Long, _
    ByVal lpszBuffer As Long, _
    ByRef lpdwBufferLength As Long, _
    ByVal dwFlags As Long _
) As Long

Private Sub cmdClear_Click()
cmbTypeSearch.ListIndex = -1
    dteExpairyDate.Text = ""
End Sub



Private Sub cmbTypeSearch_Click()
dteExpairyDate.Text = ""
 Select Case cmbTypeSearch.Text
        Case "Member"
            dteExpairyDate.Enabled = True
            txtMnoNumber.Enabled = True
        Case "Youth"
            txtMnoNumber.Enabled = False
            dteExpairyDate.Enabled = False
    End Select
End Sub

Private Sub cmbTypeSearch_LostFocus()
Select Case cmbTypeSearch.Text
        Case "Member"
            ListView.ListItems.Clear
            ListView.ColumnHeaders.Clear
            ListView.Refresh
            ListView.ColumnHeaders.Add , , "NUMBER", ListView.Width / 13
            ListView.ColumnHeaders.Add , , "NAME", ListView.Width / 13
            ListView.ColumnHeaders.Add , , "SURNAME", ListView.Width / 11
            ListView.ColumnHeaders.Add , , "ADDRESS1", ListView.Width / 6
            ListView.ColumnHeaders.Add , , "ADDRESS2", ListView.Width / 6
            ListView.ColumnHeaders.Add , , "MEMBER EXPIARY DATE", ListView.Width / 7
            ListView.ColumnHeaders.Add , , "STATUS", ListView.Width / 7
            ListView.ColumnHeaders.Add , , "PHONE", ListView.Width / 7
            ListView.ColumnHeaders.Add , , "MOBILE", ListView.Width / 7
            ListView.ColumnHeaders.Add , , "EMAIL", ListView.Width / 5
            ListView.View = lvwReport
        Case "Youth"
            ListView.ListItems.Clear
            ListView.ColumnHeaders.Clear
            ListView.Refresh
            ListView.ColumnHeaders.Add , , "Child Number", ListView.Width / 7
            ListView.ColumnHeaders.Add , , "Member Number", ListView.Width / 7
            ListView.ColumnHeaders.Add , , "First Name", ListView.Width / 5
            ListView.ColumnHeaders.Add , , "Last Name", ListView.Width / 5
            ListView.ColumnHeaders.Add , , "Genda", ListView.Width / 14
            ListView.ColumnHeaders.Add , , "Birth Data", ListView.Width / 6
            ListView.ColumnHeaders.Add , , "Status", ListView.Width / 9
            ListView.ColumnHeaders.Add , , "Mobile", ListView.Width / 7
            ListView.ColumnHeaders.Add , , "Email", ListView.Width / 5
            ListView.View = lvwReport
 End Select
End Sub

Private Sub cmdSearch_Click()
ShowDetais
End Sub


Private Sub CmdSendSms_Click()

Dim intCount, Count As Integer
Dim itmx As ListItem
Dim mobileNumber As String
Dim message As String
Dim Aa As String
Dim mobileColumnId As Integer
Dim lengthOfMessage As Integer
Dim response As String
Dim passSMS As Integer


passSMS = 0
Count = ListView.ListItems.Count
lengthOfMessage = Len(txtSmsMessage.Text)
If txtSmsMessage.Text <> "" And lengthOfMessage > 0 And Count > 0 And lengthOfMessage < 160 Then
    Aa = MsgBox("Are you sure you want to send sms messages?", vbYesNo)
    If Aa = vbNo Then
        Exit Sub
    Else
        Select Case cmbTypeSearch.Text
          Case "Member"
            mobileColumnId = 8
          Case "Youth"
            mobileColumnId = 7
        End Select
        message = URLencshort(txtSmsMessage.Text)
        For intCount = 1 To ListView.ListItems.Count
            Set itmx = ListView.ListItems(intCount)
            mobileNumber = itmx.SubItems(mobileColumnId)
            mobileNumber = Replace(mobileNumber, ")", "")
            mobileNumber = Replace(mobileNumber, "(", "")
            mobileNumber = Replace(mobileNumber, "-", "")
            response = sendSmsViaWebApi(mobileNumber, message)
            If response = "OK" Then
              passSMS = passSMS + 1
            End If
        Next intCount
        MsgBox "Total messages successfuly send is " & passSMS & " from total " & Count, vbExclamation + vbOKOnly
    End If
Else
MsgBox "Please make sure you have mobile list selected and a massage text been entered with max characters of 160", vbExclamation + vbOKOnly
End If

End Sub

Private Function sendSmsViaWebApi(ByRef mobileNumber As String, ByRef message As String)

Dim url As String, data As String
Dim sPostData As String
Dim dctParameters As Scripting.Dictionary

Set dctParameters = New Scripting.Dictionary

dctParameters.Add "username", gSmsUserName
dctParameters.Add "password", gSmsPassword
dctParameters.Add "to", mobileNumber
dctParameters.Add "from", gSmsMessageFrom
dctParameters.Add "message", message
dctParameters.Add "ref", "112233"
dctParameters.Add "maxsplit", "160"
dctParameters.Add "delay", "0"

sPostData = GetPostDataString(dctParameters)
url = "http://api.smsbroadcast.com.au/api-adv.php?"
sendSmsViaWebApi = postdata(url, sPostData)
End Function

Function postdata(url As String, data As String) As String
Dim WinHttpReq As Object, status As String, response As String
On Error GoTo errorfound
Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
url = url + data
WinHttpReq.Open "GET", url
Debug.Print url

WinHttpReq.SetRequestHeader "user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)"

WinHttpReq.Send
'postdata = WinHttpReq.status & " ? " & WinHttpReq.statusText & " ? " & WinHttpReq.responseText
postdata = WinHttpReq.statusText
Exit Function
errorfound:
postdata = Err.Description

End Function


Private Sub Command3_Click()

End Sub

Private Sub Command2_Click()
  cmbTypeSearch.ListIndex = -1
  dteExpairyDate.Text = ""
  txtMnoNumber.Text = ""
  dteFromDateOfBirth.Text = ""
  dteToDateOfBirth.Text = ""
End Sub

Private Sub Form_Activate()
gRecordType = Send_Sms
End Sub

Private Sub Form_Load()
ListView.ListItems.Clear
ListView.Width = Screen.Width - 5000
'ListView.Height = Screen.Height - 5000


End Sub

Private Sub ListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Call SortListView(ListView, ColumnHeader)

End Sub

Private Sub ShowDetais()
    Dim sql As String
    Dim retuns As String
    Dim s As Long
    Dim frm As Object
    Set frm = Me
    Dim strMemberExpiartDate As String
    Dim strFromDateOfBirth As String
    Dim strToDateOfBirth As String
    
    Select Case cmbTypeSearch.Text
        Case "Member"
            sql = "SELECT * FROM member WHERE CITY_ID = " & gCityId
            sql = sql & " AND  Status ='ACTIVE' AND Mobile is not null"
            
            If dteExpairyDate.Text <> "" Then
              strMemberExpiartDate = "Date('" & Format(dteExpairyDate.FormattedText, "yyyy,mm,dd") & "')"
              sql = sql & " AND MEMBERSHIP_EXPIARY <= " & strMemberExpiartDate
            End If
            
            
            If txtMnoNumber.Text <> "" Then
              sql = sql & " AND MNO =  " & txtMnoNumber.Text
            End If
            
            If dteFromDateOfBirth.Text <> "" Then
            strFromDateOfBirth = "Date('" & Format(dteFromDateOfBirth.FormattedText, "yyyy,mm,dd") & "')"
              sql = sql & " AND DATE_OF_BIRTH >= " & strFromDateOfBirth
            End If
            
            If dteToDateOfBirth.Text <> "" Then
            strToDateOfBirth = "Date('" & Format(dteToDateOfBirth.FormattedText, "yyyy,mm,dd") & "')"
              sql = sql & " AND DATE_OF_BIRTH <= " & strToDateOfBirth
            End If
            'Debug.Print sql
            
            Call GenerateMemberList(sql, frm)
            MemberSelected = False
       
        Case "Youth"
            sql = "SELECT * FROM children WHERE CITY_ID = " & gCityId
            sql = sql & " AND MEMBER = 'Y' AND Mobile is not null"
            If dteFromDateOfBirth.Text <> "" Then
              strFromDateOfBirth = "Date('" & Format(dteFromDateOfBirth.FormattedText, "yyyy,mm,dd") & "')"
              sql = sql & " AND BIRTH_DATE >= " & strFromDateOfBirth
            End If
            
            If dteToDateOfBirth.Text <> "" Then
              strToDateOfBirth = "Date('" & Format(dteToDateOfBirth.FormattedText, "yyyy,mm,dd") & "')"
              sql = sql & " AND BIRTH_DATE <= " & strToDateOfBirth
            End If
            
        Call GenerateChildrenList(sql, frm)
        ChildSelected = False
    End Select

End Sub


' Returns post data string based on dictionary.
Private Function GetPostDataString(ByRef the_dctParameters As Scripting.Dictionary) As String

    Dim vName                                   As Variant
    Dim sPostDataString                         As String

    For Each vName In the_dctParameters
        sPostDataString = sPostDataString & UrlEncode(CStr(vName)) & "=" & UrlEncode(CStr(the_dctParameters.Item(vName))) & "&"
    Next vName

    GetPostDataString = Left$(sPostDataString, Len(sPostDataString) - 1)

End Function

Public Function URLencshort(ByRef Text As String) As String
    Dim lngA As Long, strChar As String
    For lngA = 1 To Len(Text)
        strChar = Mid$(Text, lngA, 1)
        If strChar Like "[A-Za-z0-9]" Then
        ElseIf strChar = " " Then
            strChar = "+"
        Else
            strChar = "%" & Right$("0" & Hex$(Asc(strChar)), 2)
        End If
        URLencshort = URLencshort & strChar
    Next lngA
End Function

' Encode the URL data.
Private Function UrlEncode(ByVal the_sURLData As String) As String

    Dim nBufferLen                      As Long
    Dim sBuffer                         As String

    ' Only exception - encode spaces as "+".
    the_sURLData = Replace$(the_sURLData, " ", "+")

    ' Try to #-encode the string.
    ' Reserve a buffer. Maximum size is 3 chars for every 1 char in the input string.
    nBufferLen = Len(the_sURLData) * 3
    sBuffer = Space$(nBufferLen)
    If InternetCanonicalizeUrl(StrPtr(the_sURLData), StrPtr(sBuffer), nBufferLen, 0&) Then
        UrlEncode = Left$(sBuffer, nBufferLen)
    Else
        UrlEncode = the_sURLData
    End If

End Function

Private Sub Command1_Click()
    Dim command As String
    Dim mobile_number As String
    Dim access_token As String
    Dim user_iden As String
    Dim device_iden As String
    Dim message As String
    
    
    mobile_number = "0434282330"
    access_token = "o.V2oatfjWlYWyUh6SsgrmNUbPsFFNxO8F"
    user_iden = "ujvRpovDkd2"
    device_iden = "ujvRpovDkd2sjxi6fjTH3I"
    message = "TestSMS"

    command = "curl --header ""Access-Token:" & access_token & """ --header ""Content-Type:application/json"" --data-binary ""{ """"""push"""""": { """"""type"""""": """"""messaging_extension_reply"""""", """"""package_name"""""": """"""com.pushbullet.android"""""", """"""source_user_iden"""""": """"""" & user_iden & """"""",  """"""target_device_iden"""""": """"""" & device_iden & """"""", """"""conversation_iden"""""": """"""" & mobile_number & """"""", """"""message"""""": """"""" & message & """""""  }, """"""type"""""": """"""push""""""}""  --request POST https://api.pushbullet.com/v2/ephemerals"
    Debug.Print command
    
    Shell "cmd.exe /c " & command
End Sub
