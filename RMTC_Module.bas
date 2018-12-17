Attribute VB_Name = "RMTC_Module"
' Microsoft Scripting Runtime �ւ̎Q�Ɛݒ肪�K�v�ł�
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
Public Dic_Users As Object
Public Dic_TimeEntryActivity As Object
Public tmpdelluser As Object

Public LocalSavedSettings As Object
Public TransactionData As Object
Public TransactionTimeEntryData As Object

Public Mail_Subject As String
Public Mail_Body As String
Public Cal_Title As String

Public Search_Ticket_ID As String
Public Search_Ticket_Subject As String

Public abc As Object

Public Initialized As Integer
Public webincreasemyAPIKey As Integer
Public debug_ As Boolean

Public Function first_initializer()
debug_ = False
If debug_ Then Debug.Print "First Initializer Called"

Setting_Redmine_URL = ""
Setting_Redmine_APIKEY = ""
' �}�N�� - �c�[�� - �Q�Ɛݒ� - Microsoft Scripting Runtime ��ǉ����邱��
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
Set LocalSavedSettings = New Dictionary
Set TransactionData = New Dictionary
Set tmpdelluser = New Dictionary

LocalSavedSettings("Nothing") = 1
TransactionData("Nothing") = 1
Mail_Subject = ""
Mail_Body = ""

Initialized = 1
End Function

Public Function CreateHttpObject() As Object
    Dim objweb As Object
    '�e�햼�̂�HTTP�I�u�W�F�N�g�̐��������݂�
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

Public Function GetData(ByVal url As String) As String
    Dim data As String
    Dim objweb As Object

    If debug_ Then Debug.Print "REST URL : " & url

    'XMLHTTP�I�u�W�F�N�g�𐶐�
    Set objweb = CreateHttpObject()
    '�I�u�W�F�N�g�̐����Ɏ��s���Ă���Ώ����I��
    If objweb Is Nothing Then
        GetData = ""
        Exit Function
    End If
    objweb.Open "GET", url, False
    objweb.Send

    If objweb.responseText = "" Then
        Exit Function
    End If
    GetData = objweb.responseText
End Function

Sub Redmint_CreateTicket()
 Dim myOlExp As Outlook.Explorer
 Dim myOlSel As Outlook.Selection
 Dim x As Integer
 Set myOlExp = Application.ActiveExplorer
 Set myOlSel = myOlExp.Selection
 
 Call first_initializer
    
 For x = 1 To myOlSel.Count
    If TypeName(myOlSel.Item(x)) = "AppointmentItem" Then
        PostchkCal myOlSel.Item(x)
        Call RMTC_Creater.rmtc_initializer
        RMTC_Creater.Show
    ElseIf TypeName(myOlSel.Item(x)) = "MailItem" Then
        PostchkMail myOlSel.Item(x)
        Call RMTC_Creater.rmtc_initializer
        RMTC_Creater.Show
    End If
 Next x

 If myOlSel.Count = 0 Then
        Call RMTC_Creater.rmtc_initializer
        RMTC_Creater.Show
 End If
 

End Sub
Sub Redmint_CreateTimeEntry()
 Dim myOlExp As Outlook.Explorer
 Dim myOlSel As Outlook.Selection
 Dim x As Integer
 Set myOlExp = Application.ActiveExplorer
 Set myOlSel = myOlExp.Selection
 
 Call first_initializer
    
 For x = 1 To myOlSel.Count
    If TypeName(myOlSel.Item(x)) = "AppointmentItem" Then
        PostchkCal myOlSel.Item(x)
        Call first_initializer
        Call RMTM_Creater.rmtm_initializer
        RMTM_Creater.Show
    ElseIf TypeName(myOlSel.Item(x)) = "MailItem" Then
        PostchkMail myOlSel.Item(x)
        Call first_initializer
        Call RMTM_Creater.rmtm_initializer
        RMTM_Creater.Show
    End If
 Next x

 If myOlSel.Count = 0 Then
    Call first_initializer
    Call RMTM_Creater.rmtm_initializer
    RMTM_Creater.Show
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
    RMTC_Creater.TextBox_Contetns = Mail_Body
    RMTC_Creater.TextBox_Subject = Mail_Subject
    
End Function
Function PostchkCal(obj As AppointmentItem)
    Mail_Subject = ConvertString(obj.subject)
    Mail_Body = ConvertString(obj.Body)
    RMTC_Creater.TextBox_Contetns = Mail_Body
    RMTC_Creater.TextBox_Subject = Mail_Subject

    RMTM_Creater.TextBox_timeentryhours.value = ConvertString(obj.Duration) / 60
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
    ConvertString = tmpstr
End Function
Sub ���W�X�g���폜()
 DeleteSetting ("OutlookRMTC")
End Sub


Public Function testabs()
    Set abc = New Dictionary
End Function



