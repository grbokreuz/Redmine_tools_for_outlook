Attribute VB_Name = "RMTC_Module"
' Microsoft Scripting Runtime
Public Setting_Redmine_URL As String
Public Setting_Redmine_APIKEY As String
Public Setting_Redmine_Tracker_For_Granpa As Object
Public Setting_Redmine_Tracker_For_Parents As Object
Public Setting_Redmine_Tracker_For_children As Object
Public Dic_Projects As Object
Public Dic_Projects_ID As Object
Public Dic_Trackers As Object
Public Dic_Story As Object
Public Dic_Activity As Object
Public Dic_Backlog As Object
Public Dic_Statuses As Object
Public Dic_Priority As Object
Public Dic_Asiigned As Object
Public Dic_EstimatedHours As Object
Public Dic_Assigned_To_Me As Object
Public Dic_Users As Object
Public Dic_TimeEntryActivity As Object
Public tmpdelluser As Object
Public LocalSavedSettings As Object
Public TransactionData As Object
Public TransactionTimeEntryData As Object
Public TransactionSearch As Object
Public Mail_Subject As String
Public Mail_Body As String
Public Cal_Title As String
Public Search_Ticket_ID As String
Public Search_Ticket_Subject As String
Public abc As Object
Public Initialized As Integer
Public webincreasemyAPIKey As Integer
Public keywordsearchonAllTrackers As Integer
Public searchContents As Integer
Public debug_ As Boolean
Public RMTS_Search_SingleMode As Boolean
Public for_Japanese As Boolean
Public Function first_initializer()
    debug_ = False
    for_Japanese = True

    If debug_ Then Debug.Print "First Initializer Called"
    Setting_Redmine_URL = ""
    Setting_Redmine_APIKEY = ""
    Set Setting_Redmine_Tracker_For_Granpa = New Dictionary
    Set Setting_Redmine_Tracker_For_Parents = New Dictionary
    Set Setting_Redmine_Tracker_For_children = New Dictionary
    Set Dic_Projects = New Dictionary
    Set Dic_Projects_ID = New Dictionary
    Set Dic_Trackers = New Dictionary
    Set Dic_Story = New Dictionary
    Set Dic_Activity = New Dictionary
    Set Dic_Statuses = New Dictionary
    Set Dic_Priority = New Dictionary
    Set Dic_Asiigned = New Dictionary
    Set Dic_EstimatedHours = New Dictionary
    Set Dic_Users = New Dictionary
    Set Dic_TimeEntryActivity = New Dictionary
    Set Dic_Assigned_To_Me = New Dictionary
    Set LocalSavedSettings = New Dictionary
    Set TransactionData = New Dictionary
    Set tmpdelluser = New Dictionary
    LocalSavedSettings("Nothing") = 1
    TransactionData("Nothing") = 1
    Mail_Subject = ""
    Mail_Body = ""
    Initialized = 1
    RMTS_Search_SingleMode = False

End Function
Public Function CreateHttpObject() As Object
    Dim objweb As Object
    Err.Clear
    Set objweb = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If Err.Number = 0 Then
        Set CreateHttpObject = objweb
        Exit Function
    End If
    Err.Clear
    Set objweb = CreateObject("MSXML2.ServerXMLHTTP")
    If Err.Number = 0 Then
        Set CreateHttpObject = objweb
        Exit Function
    End If
    Err.Clear
    Set objweb = CreateObject("MSXML2.XMLHTTP")
    If Err.Number = 0 Then
        Set CreateHttpObject = objweb
        Exit Function
    End If
    Set CreateHttpObject = Nothing
End Function
Public Function GetData(ByVal URL As String) As String
    Dim data As String
    Dim objweb As Object
    If debug_ Then Debug.Print "REST URL : " & URL
    Set objweb = CreateHttpObject()
    If objweb Is Nothing Then
        GetData = ""
        Exit Function
    End If
    objweb.Open "GET", URL, False
    objweb.Send

    If objweb.responseText = "" Then
        Exit Function
    End If
    GetData = objweb.responseText
End Function
Sub Redmint_CreateTicket()
 Dim myOlExp As Outlook.Explorer
 Dim myOlSel As Outlook.Selection
 Dim X As Integer
 Set myOlExp = Application.ActiveExplorer
 Set myOlSel = myOlExp.Selection
 Call first_initializer
 For X = 1 To myOlSel.Count
    If TypeName(myOlSel.Item(X)) = "AppointmentItem" Then
        PostchkCal myOlSel.Item(X)
        Call RMTC_Creater.rmtc_initializer
        RMTC_Creater.Show
    ElseIf TypeName(myOlSel.Item(X)) = "MailItem" Then
        PostchkMail myOlSel.Item(X)
        Call RMTC_Creater.rmtc_initializer
        RMTC_Creater.Show
    End If
 Next X
 If myOlSel.Count = 0 Then
        Call RMTC_Creater.rmtc_initializer
        RMTC_Creater.Show
 End If
End Sub
Sub Redmint_CreateTimeEntry()
 Dim myOlExp As Outlook.Explorer
 Dim myOlSel As Outlook.Selection
 Dim X As Integer
 Set myOlExp = Application.ActiveExplorer
 Set myOlSel = myOlExp.Selection
 Call first_initializer
 For X = 1 To myOlSel.Count
    If TypeName(myOlSel.Item(X)) = "AppointmentItem" Then
        PostchkCal myOlSel.Item(X)
        Call first_initializer
        Call RMTM_Creater.rmtm_initializer
        RMTM_Creater.Show
    ElseIf TypeName(myOlSel.Item(X)) = "MailItem" Then
        PostchkMail myOlSel.Item(X)
        Call first_initializer
        Call RMTM_Creater.rmtm_initializer
        RMTM_Creater.Show
    End If
 Next X
 If myOlSel.Count = 0 Then
    Call RMTM_Creater.rmtm_initializer
    RMTM_Creater.Show
 End If
End Sub
Sub Redmint_Search()
 Dim myOlExp As Outlook.Explorer
 Dim myOlSel As Outlook.Selection
 Set myOlExp = Application.ActiveExplorer
 Set myOlSel = myOlExp.Selection
 Call first_initializer
 Call RMTS_Search.rmts_initialize
 
 For X = 1 To myOlSel.Count
    If TypeName(myOlSel.Item(X)) = "AppointmentItem" Then
        RMTS_Search_SingleMode = True
        RMTS_Search.TextBox_SearchKey = "==EntryID=" & myOlSel.Item(X).EntryID & "=="
        RMTS_Search.CommandButton_SearchTicket_Click
        RMTS_Search.Show vbModeless
    ElseIf TypeName(myOlSel.Item(X)) = "MailItem" Then

        RMTS_Search_SingleMode = True
        RMTS_Search.TextBox_SearchKey = "==EntryID=" & myOlSel.Item(X).EntryID & "=="
        RMTS_Search.CommandButton_SearchTicket_Click
        RMTS_Search.Show vbModeless
    End If
 Next X

If myOlSel.Count = 0 Then
        RMTS_Search_SingleMode = True
    '   RMTS_Search.CommandButton_SearchTicket_Click
        RMTS_Search.Show vbModeless
End If
End Sub
Sub Dump(Text As String)
Dim k As Long
For k = 1 To Len(Text)
Debug.Print k, Mid(Text, k, 1), Hex(Asc(Mid(Text, k, 1)))
Next
End Sub
Function PostchkMail(obj As MailItem)
    Mail_Subject = ConvertString(obj.subject)
    Mail_Body = ConvertString(obj.Body)
    RMTC_Creater.TextBox_Contetns = vbNewLine & "{{collapse(eMail)" & vbNewLine & Mail_Subject & vbNewLine & Mail_Body & vbNewLine & "}}" & vbNewLine & "{{collapse(EntryID)" & vbNewLine & "==EntryID=" & obj.EntryID & "==" & vbNewLine & "}}" & vbNewLine
    RMTC_Creater.TextBox_Subject = Mail_Subject
End Function
Function PostchkCal(obj As AppointmentItem)
    Mail_Subject = ConvertString(obj.subject)
    Mail_Body = ConvertString(obj.Body)
    RMTC_Creater.TextBox_Contetns = vbNewLine & "{{collapse(Cal)" & vbNewLine & Mail_Subject & vbNewLine & Mail_Body & vbNewLine & "}}" & vbNewLine & "{{collapse(EntryID)" & vbNewLine & "==EntryID=" & obj.EntryID & "==" & vbNewLine & "}}" & vbNewLine
    RMTC_Creater.TextBox_Subject = Mail_Subject
    RMTM_Creater.ScrollBar_timeentry.value = ConvertString(obj.Duration) / 60 / 0.25
    RMTM_Creater.TextBox_Comment.Text = ConvertString(obj.ConversationTopic) & Mail_Subject
    Cal_Title = ConvertString(obj.ConversationTopic)
End Function
Public Function ConvertString(ByVal val As String)
    Dim tmpstr As String
    tmpstr = Mid(val, 1, 6000)
    tmpstr = Replace(tmpstr, "<", "&lt;")
    tmpstr = Replace(tmpstr, ">", "&gt;")
    tmpstr = Replace(tmpstr, """", "&quot;")
    tmpstr = Replace(tmpstr, "'", "&apos;")
    tmpstr = Replace(tmpstr, "&", "&amp;")
    tmpstr = Replace(tmpstr, vbCrLf + vbCrLf + vbCrLf + vbCrLf, vbCrLf + vbCrLf)
    tmpstr = Replace(tmpstr, Chr(13) + Chr(10) + Chr(13) + Chr(10), Chr(13) + Chr(10))
    tmpstr = Replace(tmpstr, vbCr + vbCr, vbCr)
    tmpstr = Replace(tmpstr, Chr(13) + Chr(13), Chr(13))
    tmpstr = Replace(tmpstr, vbLf + vbLf, vbLf)
    tmpstr = Replace(tmpstr, Chr(10) + Chr(10), Chr(10))
    ConvertString = tmpstr
End Function
Sub Delete_Reg()
 DeleteSetting ("OutlookRMTC")
End Sub
Public Function testabs()
    Set abc = New Dictionary
End Function

Public Sub openweb(ByVal urlpath As String)
    Dim WSH As Object
    Set WSH = CreateObject("Wscript.Shell")
    WSH.Run urlpath, 3
    Set WSH = Nothing
End Sub
Public Function get_ticket_subject_for_caption(ByVal ticketnumber As Integer, ByVal URL As String, ByVal apikey As String, ByRef col As Long)
    Dim JSONLib As New JSONLib
    Dim json, tmpdic As Object
    Dim Var As Variant
    Dim total, offset, limit, nextoffset As Integer
    If debug_ Then DebugPrintFile "ÅöstartÅöCalle :: get_ticket_subject_for_caption"
    Dim JsonString As String

    
    JsonString = GetData(URL & "/issues/" & ticketnumber & ".json?key=" & apikey)
    Set json = New Dictionary
    Set json = JSONLib.parse(JsonString)


    If json Is Nothing Then
        MsgBox "not found ticket."
        Exit Function
    End If
    Set Var = json("issue")
    Dim descri As String
    descri = Var("description")

    descri = Replace(descri, vbCrLf + vbCrLf + vbCrLf + vbCrLf, vbCrLf + vbCrLf)
    descri = Replace(descri, Chr(13) + Chr(10) + Chr(13) + Chr(10), Chr(13) + Chr(10))


    descri = Replace(descri, vbCr + vbCr, vbCr)
    descri = Replace(descri, Chr(13) + Chr(13), Chr(13))

    descri = Replace(descri, vbLf + vbLf, vbLf)
    descri = Replace(descri, Chr(10) + Chr(10), Chr(10))
  
  '  descri = Replace(descri, vbCrLf, "")
  '  descri = Replace(descri, Chr(13) + Chr(10), "")

  '  descri = Replace(descri, vbCr, "")
  '  descri = Replace(descri, Chr(13), "")

  '  descri = Replace(descri, vbLf, "")
  '  descri = Replace(descri, Chr(10), "")

    If col >= 0 And RMTM_Creater.ListBox_mytimeentry.List(col, 1) <> "" Then
        MsgBox "#" & Var("id") & ":" & Var("subject") & Chr(13) & "--------------------------------" & Chr(13) & descri & Chr(13) & "--------------------------------", vbOKOnly, "ticket #" & Var("id")
    End If
    Set json = Nothing
    Set JSONLib = Nothing
    
    get_ticket_subject_for_caption = Var("subject")
If debug_ Then DebugPrintFile "ÅöendÅö get_ticket_subject_for_caption"
End Function
Public Sub DebugPrintFile(varData As Variant)

  Dim lngFileNum As Long
  Dim strLogFile As String
  
  strLogFile = "C:\temp\OutlookVBA_DebugPrint.txt"
  lngFileNum = FreeFile()
  Open strLogFile For Append As #lngFileNum
  Print #lngFileNum, varData
  Close #lngFileNum
  
  Debug.Print varData

End Sub
