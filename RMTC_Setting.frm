VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RMTC_Setting 
   Caption         =   "Redmine Ticket Creater Settings"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9555.001
   OleObjectBlob   =   "RMTC_Setting.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "RMTC_Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






' Microsoft Scripting Runtime への参照設定が必要です
Option Explicit

Private Sub ComboBox_Project_for_usermember_Change()
    If debug_ Then Debug.Print "ComboBox_Project_for_usermember_Change"
    Dim selectedItemKey As String
    Dim selectedItemVal As String
    Dim word As Variant
    
    selectedItemKey = ComboBox_Project_for_usermember.value

    For Each word In Dic_Projects.keys
        If word = selectedItemKey Then
            selectedItemVal = Dic_Projects(word)
            Exit For
        End If
    Next word
    
    If selectedItemKey <> "" Then
        Call get_issue_membership_for_project_For_Setting(Setting_Redmine_URL, selectedItemVal, Setting_Redmine_APIKEY)
    End If
    
End Sub

Private Sub CommandButton_LoadSettingForm_Click()
    Dim tmpdic As Object
    Setting_Redmine_URL = TextBox_RedmineURL.value
    Setting_Redmine_APIKEY = TextBox_Redmine_APIKey.value
    Dim tmpstr As String

    If Not Setting_Redmine_URL Like "*/" Then
        Setting_Redmine_URL = Setting_Redmine_URL & "/"
        TextBox_RedmineURL.value = Setting_Redmine_URL
    End If
    
    If Not (Setting_Redmine_URL Like "http://*" Or Setting_Redmine_URL Like "httsp://*") Then
        Setting_Redmine_URL = "http://" & Setting_Redmine_URL
        TextBox_RedmineURL.value = Setting_Redmine_URL
    End If

    If LocalSavedSettings.exists("Setting_Redmine_URL") Then
        LocalSavedSettings("Setting_Redmine_URL") = Setting_Redmine_URL
    Else
        LocalSavedSettings.Add "Setting_Redmine_URL", Setting_Redmine_URL
    End If
    
    If LocalSavedSettings.exists("Setting_Redmine_APIKEY") Then
        LocalSavedSettings("Setting_Redmine_APIKEY") = Setting_Redmine_APIKEY
    Else
        LocalSavedSettings.Add "Setting_Redmine_APIKEY", Setting_Redmine_APIKEY
    End If
 

    ' API通信の成功を確認したら画面を広げる
    
    If Setting_Redmine_URL = "" Or Setting_Redmine_APIKEY = "" Then
        'MsgBox "Enter Redmine URL and Key"
        Exit Sub
    End If
    
' RM_APIから読み込み
    Call get_issue_statuses_For_Setting(Setting_Redmine_URL, Setting_Redmine_APIKEY)
    Call get_issue_priority_For_Setting(Setting_Redmine_URL, Setting_Redmine_APIKEY)
    Call get_issue_tracker_For_Setting(Setting_Redmine_URL, Setting_Redmine_APIKEY)
    Call get_issue_project_For_Setting(Setting_Redmine_URL, Setting_Redmine_APIKEY)
    Call get_timeentry_activity_For_Setting(Setting_Redmine_URL, Setting_Redmine_APIKEY)
' レジストリの値を選択 + 上記プラスuser も
    If LocalSavedSettings.exists("ListBox_setting_priority") Then
        If debug_ Then Debug.Print "Locasetting have ListBox_setting_priority"
        Set tmpdic = LocalSavedSettings("ListBox_setting_priority")
        Call select_listbox(tmpdic, Me.ListBox_setting_priority, Dic_Priority)
    End If
    If LocalSavedSettings.exists("ListBox_setting_Status") Then
        If debug_ Then Debug.Print "Locasetting have ListBox_setting_Status"
        Set tmpdic = LocalSavedSettings("ListBox_setting_Status")
        Call select_listbox(tmpdic, Me.ListBox_setting_Status, Dic_Statuses)
    End If
    If LocalSavedSettings.exists("ListBox_Setting_Status_parents") Then
        If debug_ Then Debug.Print "Locasetting have ListBox_Setting_Status_parents"
        Set tmpdic = LocalSavedSettings("ListBox_Setting_Status_parents")
        Call select_listbox(tmpdic, Me.ListBox_Setting_Status_parents, Dic_Statuses)
    End If
    If LocalSavedSettings.exists("ListBox_Setting_Status_granpa") Then
        If debug_ Then Debug.Print "Locasetting have ListBox_Setting_Status_granpa"
        Set tmpdic = LocalSavedSettings("ListBox_Setting_Status_granpa")
        Call select_listbox(tmpdic, Me.ListBox_Setting_Status_granpa, Dic_Statuses)
    End If
    If LocalSavedSettings.exists("ListBox_Setting_Status_child") Then
        If debug_ Then Debug.Print "Locasetting have ListBox_Setting_Status_child"
        Set tmpdic = LocalSavedSettings("ListBox_Setting_Status_child")
        Call select_listbox(tmpdic, Me.ListBox_Setting_Status_child, Dic_Statuses)
    End If
    If LocalSavedSettings.exists("ListBox_Setting_Tracker_parents") Then
        If debug_ Then Debug.Print "Locasetting have ListBox_Setting_Tracker_parents"
        Set tmpdic = LocalSavedSettings("ListBox_Setting_Tracker_parents")
        Call select_listbox(tmpdic, Me.ListBox_Setting_Tracker_parents, Dic_Trackers)
    End If
    If LocalSavedSettings.exists("ListBox_Setting_Tracker_grandparent") Then
        If debug_ Then Debug.Print "Locasetting have ListBox_Setting_Tracker_grandparent"
        Set tmpdic = LocalSavedSettings("ListBox_Setting_Tracker_grandparent")
        Call select_listbox(tmpdic, Me.ListBox_Setting_Tracker_grandparent, Dic_Trackers)
    End If
    If LocalSavedSettings.exists("ListBox_Setting_Tracker_child") Then
        If debug_ Then Debug.Print "Locasetting have ListBox_Setting_Tracker_child"
        Set tmpdic = LocalSavedSettings("ListBox_Setting_Tracker_child")
        Call select_listbox(tmpdic, Me.ListBox_Setting_Tracker_child, Dic_Trackers)
    End If
    If LocalSavedSettings.exists("ListBox_setting_TimeEntryActivity") Then
        If debug_ Then Debug.Print "Locasetting have ListBox_setting_TimeEntryActivity"
        Set tmpdic = LocalSavedSettings("ListBox_setting_TimeEntryActivity")
        Call select_listbox(tmpdic, Me.ListBox_setting_TimeEntryActivity, Dic_TimeEntryActivity)
    End If
    If LocalSavedSettings.exists("ListBox_New_Ticket_Story") Then
        If debug_ Then Debug.Print "Locasetting have ListBox_New_Ticket_Story"
        tmpstr = LocalSavedSettings("ListBox_New_Ticket_Story")
        Call select_listbox_val(tmpstr, Me.ListBox_New_Ticket_Story, Dic_Trackers)
    End If
    If LocalSavedSettings.exists("ListBox_New_Ticket_Activity") Then
        If debug_ Then Debug.Print "Locasetting have ListBox_New_Ticket_Activity"
        tmpstr = LocalSavedSettings("ListBox_New_Ticket_Activity")
        Call select_listbox_val(tmpstr, Me.ListBox_New_Ticket_Activity, Dic_Trackers)
    End If
    If LocalSavedSettings.exists("ListBox_New_Ticket_Backlog") Then
        If debug_ Then Debug.Print "Locasetting have ListBox_New_Ticket_Backlog"
        tmpstr = LocalSavedSettings("ListBox_New_Ticket_Backlog")
        Call select_listbox_val(tmpstr, Me.ListBox_New_Ticket_Backlog, Dic_Trackers)
    End If

    If LocalSavedSettings.exists("ListBox_User_Settings") Then
        If debug_ Then Debug.Print "Locasetting have ListBox_User_Settings"
        Set tmpdic = LocalSavedSettings("ListBox_User_Settings")
        Call select_listbox_users(tmpdic, Me.ListBox_User_Settings, Dic_Users)
    End If
    
    CommandButton_SaveSetting.Enabled = True

    If webincreasemyAPIKey = 1 Then
         RMTC_Setting.CheckBox_webAccess_KeyIncrease.value = True
    Else
         RMTC_Setting.CheckBox_webAccess_KeyIncrease.value = False
    End If

    If ComboBox_Project_for_usermember.ListRows > 0 Then
        ComboBox_Project_for_usermember = ComboBox_Project_for_usermember.List(0)
    End If
End Sub
Public Sub get_issue_membership_for_project_For_Setting(ByVal url As String, ByVal project As String, ByVal apikey As String)
    If debug_ Then Debug.Print "get_issue_membership_for_project_For_Setting"
    If project = "" Then
        If debug_ Then Debug.Print "Not found; project "
        Exit Sub
        
    End If

    Set Dic_Users = Nothing
    Set Dic_Users = New Dictionary
    ListBox_Users_from_rm.Clear
    Dim subjson As Integer
    Dim jsonstring As String
    
    
    jsonstring = GetData(url & "/projects/" & project & "/memberships.json?key=" & apikey)
    Dim JSONLib As New JSONLib
    Dim json As Object
    Dim Var As Variant
    Dim total, offset, limit
    Set json = New Dictionary
    Set json = JSONLib.parse(jsonstring)
    If json Is Nothing Then
        MsgBox "Cant load rm."
        Exit Sub
    End If
 
    total = json("total_count")
    offset = json("offset")
    limit = json("limit")
    subjson = 0
    
    Do While total > limit + offset
        subjson = 1
        If debug_ Then Debug.Print "limit " & json("limit")
        If debug_ Then Debug.Print "offset " & json("offset")
        If debug_ Then Debug.Print "total " & json("total_count")
        Dim subjsonstr As String
        subjsonstr = GetData(url & "/projects/" & project & "/memberships.json?key=" & apikey & "&offset=" & json("limit") + json("offset"))
        Dim jsonsub As Object
        Set jsonsub = New Dictionary
        Set jsonsub = JSONLib.parse(subjsonstr)
        total = jsonsub("total_count")
        offset = jsonsub("offset")
        limit = jsonsub("limit")
    Loop

    If subjson = 1 Then
        For Each Var In jsonsub("memberships")
            json("memberships").Add Var
        Next Var
    End If

    For Each Var In json("memberships")
        If Var.exists("user") Then
            If Dic_Users.exists(Var("user")("name")) Then
            Else
                Dic_Users.Add Var("user")("name"), Var("user")("id")
            End If
            ListBox_Users_from_rm.AddItem (Var("user")("name"))
        End If
    Next Var
    Set json = Nothing
    Set JSONLib = Nothing

End Sub
Public Sub get_issue_project_For_Setting(ByVal url As String, ByVal apikey As String)
    Set Dic_Projects = Nothing
    Set Dic_Projects = New Dictionary
    ComboBox_Project_for_usermember.Clear
    ListBox_Users_from_rm.Clear
    Set Dic_Projects_ID = New Dictionary
    
    Dim subjson As Integer
    Dim jsonstring As String
    jsonstring = GetData(url & "/projects.json?key=" & apikey)
    Dim JSONLib As New JSONLib
    Dim json As Object
    Dim Var As Variant
    Dim total, offset, limit
    Set json = New Dictionary
    Set json = JSONLib.parse(jsonstring)
    If json Is Nothing Then
        MsgBox "Cant load rm."
        Exit Sub
    End If

    total = json("total_count")
    offset = json("offset")
    limit = json("limit")
    subjson = 0
    
    Do While total > limit + offset
        subjson = 1
        If debug_ Then Debug.Print "limit " & json("limit")
        If debug_ Then Debug.Print "offset " & json("offset")
        If debug_ Then Debug.Print "total " & json("total_count")
        Dim subjsonstr As String
        subjsonstr = GetData(url & "/projects.json?key=" & apikey & "&offset=" & json("limit") + json("offset"))
        Dim jsonsub As Object
        Set jsonsub = New Dictionary
        Set jsonsub = JSONLib.parse(subjsonstr)
        total = jsonsub("total_count")
        offset = jsonsub("offset")
        limit = jsonsub("limit")
    Loop

    If subjson = 1 Then
        For Each Var In jsonsub("projects")
            json("projects").Add Var
        Next Var
    End If

    For Each Var In json("projects")
        Dic_Projects.Add Var("name"), Var("identifier")
        Dic_Projects_ID.Add Var("name"), Var("id")
        ComboBox_Project_for_usermember.AddItem (Var("name"))
    Next Var
    Set json = Nothing
    Set JSONLib = Nothing
End Sub
Public Sub get_timeentry_activity_For_Setting(ByVal url As String, ByVal apikey As String)
    Set Dic_TimeEntryActivity = Nothing
    Set Dic_TimeEntryActivity = New Dictionary
    ListBox_setting_TimeEntryActivity.Clear
    
    Dim jsonstring As String
    jsonstring = GetData(url & "/enumerations/time_entry_activities.json?key=" & apikey)
    Dim JSONLib As New JSONLib
    Dim json As Object
    Dim Var As Variant

    Set json = New Dictionary
    Set json = JSONLib.parse(jsonstring)
    If json Is Nothing Then
        MsgBox "Cant load rm."
        Exit Sub
    End If

    For Each Var In json("time_entry_activities")
        Dic_TimeEntryActivity.Add Var("name"), Var("id")
        ListBox_setting_TimeEntryActivity.AddItem (Var("name"))
    Next Var
    Set json = Nothing
    Set JSONLib = Nothing
End Sub
Public Sub get_issue_statuses_For_Setting(ByVal url As String, ByVal apikey As String)
    Set Dic_Statuses = Nothing
    Set Dic_Statuses = New Dictionary
    ListBox_Setting_Status_parents.Clear
    ListBox_Setting_Status_granpa.Clear
    ListBox_setting_Status.Clear
    ListBox_Setting_Status_child.Clear

    Dim subjson As Integer
    Dim jsonstring As String
    jsonstring = GetData(url & "/issue_statuses.json?key=" & apikey)
      
    Dim JSONLib As New JSONLib
    Dim json As Object
    Dim Var As Variant
    Dim total, offset, limit
    Set json = New Dictionary
    Set json = JSONLib.parse(jsonstring)
    If json Is Nothing Then
        MsgBox "Cant load rm."
        Exit Sub
    End If
    
    
    total = json("total_count")
    offset = json("offset")
    limit = json("limit")
    subjson = 0
    
    Do While total > limit + offset
        subjson = 1
        If debug_ Then Debug.Print "limit " & json("limit")
        If debug_ Then Debug.Print "offset " & json("offset")
        If debug_ Then Debug.Print "total " & json("total_count")
        Dim subjsonstr As String
        subjsonstr = GetData(url & "/issue_statuses.json?key=" & apikey & "&offset=" & json("limit") + json("offset"))
        Dim jsonsub As Object
        Set jsonsub = New Dictionary
        Set jsonsub = JSONLib.parse(subjsonstr)
        total = jsonsub("total_count")
        offset = jsonsub("offset")
        limit = jsonsub("limit")
    Loop

    If subjson = 1 Then
        For Each Var In jsonsub("issue_statuses")
            json("issue_statuses").Add Var
        Next Var
    End If

    For Each Var In json("issue_statuses")
        Dic_Statuses.Add Var("name"), Var("id")
        ListBox_Setting_Status_parents.AddItem (Var("name"))
        ListBox_Setting_Status_granpa.AddItem (Var("name"))
        ListBox_setting_Status.AddItem (Var("name"))
        ListBox_Setting_Status_child.AddItem (Var("name"))
    Next Var
    Set json = Nothing
    Set JSONLib = Nothing
End Sub
Public Sub get_issue_tracker_For_Setting(ByVal url As String, ByVal apikey As String)
    Set Dic_Trackers = Nothing
    Set Dic_Trackers = New Dictionary
    ListBox_Setting_Tracker_grandparent.Clear
    ListBox_Setting_Tracker_parents.Clear
    ListBox_New_Ticket_Story.Clear
    ListBox_New_Ticket_Activity.Clear
    ListBox_New_Ticket_Backlog.Clear
    ListBox_Setting_Tracker_child.Clear

    Dim subjson As Integer
    Dim jsonstring As String
    jsonstring = GetData(url & "/trackers.json?key=" & apikey)
    Dim JSONLib As New JSONLib
    Dim json As Object
    Dim Var As Variant
    Dim total, offset, limit
    Set json = New Dictionary
    Set json = JSONLib.parse(jsonstring)
    If json Is Nothing Then
        MsgBox "Cant load rm."
        Exit Sub
    End If
    
    total = json("total_count")
    offset = json("offset")
    limit = json("limit")
    subjson = 0
    
    Do While total > limit + offset
        subjson = 1
        If debug_ Then Debug.Print "limit " & json("limit")
        If debug_ Then Debug.Print "offset " & json("offset")
        If debug_ Then Debug.Print "total " & json("total_count")
        Dim subjsonstr As String
        subjsonstr = GetData(url & "/trackers.json?key=" & apikey & "&offset=" & json("limit") + json("offset"))
        Dim jsonsub As Object
        Set jsonsub = New Dictionary
        Set jsonsub = JSONLib.parse(subjsonstr)
        total = jsonsub("total_count")
        offset = jsonsub("offset")
        limit = jsonsub("limit")
    Loop

    If subjson = 1 Then
        For Each Var In jsonsub("trackers")
            json("trackers").Add Var
        Next Var
    End If

    For Each Var In json("trackers")
        Dic_Trackers.Add Var("name"), Var("id")
        ListBox_Setting_Tracker_parents.AddItem (Var("name"))
        ListBox_Setting_Tracker_grandparent.AddItem (Var("name"))
        ListBox_Setting_Tracker_child.AddItem (Var("name"))
        ListBox_New_Ticket_Story.AddItem (Var("name"))
        ListBox_New_Ticket_Activity.AddItem (Var("name"))
        ListBox_New_Ticket_Backlog.AddItem (Var("name"))
    Next Var
    Set json = Nothing
    Set JSONLib = Nothing
    
End Sub
Public Sub get_issue_priority_For_Setting(ByVal url As String, ByVal apikey As String)
    Set Dic_Priority = Nothing
    Set Dic_Priority = New Dictionary
    ListBox_setting_priority.Clear
    
    Dim subjson As Integer
    Dim jsonstring As String
    jsonstring = GetData(url & "/enumerations/issue_priorities.json?key=" & apikey)
    Dim JSONLib As New JSONLib
    Dim json As Object
    Dim Var As Variant
    Dim total, offset, limit
    Set json = New Dictionary
    Set json = JSONLib.parse(jsonstring)
    If json Is Nothing Then
        MsgBox "Cant load rm."
        Exit Sub
    End If
 
    total = json("total_count")
    offset = json("offset")
    limit = json("limit")
    subjson = 0
    
    Do While total > limit + offset
        subjson = 1
        If debug_ Then Debug.Print "limit " & json("limit")
        If debug_ Then Debug.Print "offset " & json("offset")
        If debug_ Then Debug.Print "total " & json("total_count")
        Dim subjsonstr As String
        subjsonstr = GetData(url & "/enumerations/issue_priorities.json?key=" & apikey & "&offset=" & json("limit") + json("offset"))
        Dim jsonsub As Object
        Set jsonsub = New Dictionary
        Set jsonsub = JSONLib.parse(subjsonstr)
        total = jsonsub("total_count")
        offset = jsonsub("offset")
        limit = jsonsub("limit")
    Loop

    If subjson = 1 Then
        For Each Var In jsonsub("issue_priorities")
            json("issue_priorities").Add Var
        Next Var
    End If

    For Each Var In json("issue_priorities")
        Dic_Priority.Add Var("name"), Var("id")
        ListBox_setting_priority.AddItem (Var("name"))
    Next Var
    Set json = Nothing
    Set JSONLib = Nothing
    
End Sub


Private Sub CommandButton_Useradd_Click()
    If ListBox_Users_from_rm.value <> "" Then
        Dim alreadyexists, i As Integer
        alreadyexists = 0
        
        tmpdelluser(ListBox_Users_from_rm.value) = 0
        For i = 0 To ListBox_User_Settings.ListCount - 1
            If ListBox_User_Settings.List(i) = ListBox_Users_from_rm.value Then
                ListBox_User_Settings.ListIndex = i
                alreadyexists = 1
                Exit Sub
            End If
        Next i
        If alreadyexists = 0 Then
            ListBox_User_Settings.AddItem ListBox_Users_from_rm.value
        End If
    End If
End Sub
Private Sub CommandButton_Userdel_Click()
    Dim index As Integer
    Dim tmpdic As Object
    Dim tmpstr As String
    index = ListBox_User_Settings.ListIndex
    If index = -1 Then
    Else
        tmpstr = ListBox_User_Settings.List(index)
        ListBox_User_Settings.RemoveItem index
        tmpdelluser(tmpstr) = 1
    End If
End Sub
Private Sub CommandButton_SaveSetting_Click()
    Dim JSONLib As New JSONLib
    Dim json As Object
    Dim RegStr As String
    RegStr = JSONLib.toString(LocalSavedSettings)
    If debug_ Then Debug.Print "★CommandButton_SaveSetting_Click start Now reg is :: "; RegStr
    
    LocalSavedSettings("Setting_Redmine_URL") = TextBox_RedmineURL.Text
    LocalSavedSettings("Setting_Redmine_APIKEY") = TextBox_Redmine_APIKey.Text
    Call ListBox_Packager(LocalSavedSettings, "ListBox_User_Settings", Me.ListBox_User_Settings)
    Call ListBox_Packager(LocalSavedSettings, "ListBox_setting_priority", Me.ListBox_setting_priority)
    
    Call ListBox_Packager(LocalSavedSettings, "ListBox_setting_Status", Me.ListBox_setting_Status)
    Call ListBox_Packager(LocalSavedSettings, "ListBox_Setting_Status_parents", Me.ListBox_Setting_Status_parents)
    Call ListBox_Packager(LocalSavedSettings, "ListBox_Setting_Status_granpa", Me.ListBox_Setting_Status_granpa)
    Call ListBox_Packager(LocalSavedSettings, "ListBox_Setting_Status_child", Me.ListBox_Setting_Status_child)
    
    Call ListBox_Packager(LocalSavedSettings, "ListBox_Setting_Tracker_parents", Me.ListBox_Setting_Tracker_parents)
    Call ListBox_Packager(LocalSavedSettings, "ListBox_Setting_Tracker_grandparent", Me.ListBox_Setting_Tracker_grandparent)
    Call ListBox_Packager(LocalSavedSettings, "ListBox_Setting_Tracker_child", Me.ListBox_Setting_Tracker_child)
    
    Call ListBox_Packager_Val(LocalSavedSettings, "ListBox_New_Ticket_Story", Me.ListBox_New_Ticket_Story)
    Call ListBox_Packager_Val(LocalSavedSettings, "ListBox_New_Ticket_Activity", Me.ListBox_New_Ticket_Activity)
    Call ListBox_Packager_Val(LocalSavedSettings, "ListBox_New_Ticket_Backlog", Me.ListBox_New_Ticket_Backlog)
    Call ListBox_Packager(LocalSavedSettings, "ListBox_setting_TimeEntryActivity", Me.ListBox_setting_TimeEntryActivity)
    Call Dic_Packager(LocalSavedSettings, "Dic_Projects", Dic_Projects)
    Call Dic_Packager(LocalSavedSettings, "Dic_Projects_ID", Dic_Projects_ID)
    
    Debug.Assert Err.Number = 0
    

    If RMTC_Setting.CheckBox_webAccess_KeyIncrease.value = True Then
        webincreasemyAPIKey = 1
        LocalSavedSettings("webincreasemyAPIKey") = 1
    Else
        webincreasemyAPIKey = 0
        LocalSavedSettings("webincreasemyAPIKey") = 0
    End If

    RegStr = JSONLib.toString(LocalSavedSettings)
    
    If debug_ Then Debug.Print "save to reg :: "; RegStr
    SaveSetting "OutlookRMTC", "Settings", "AllSettings", RegStr

    Unload Me
    
    Exit Sub

End Sub

Private Sub ListBox_Packager(ByRef dic_setthing As Object, ByVal keyname As String, ByVal myctrl As MSForms.ListBox)
    Dim i As Integer
    Dim childnode As Object
    Set childnode = New Dictionary

    Dim JSONLib As New JSONLib

    If debug_ Then Debug.Print "Create " & keyname & " Dump: start"
    
  
    ' 基本的には設定は全てクリアして、フォームで選ばれたものを詰め込む
    If ( _
           (keyname = "ListBox_setting_priority") Or _
           (keyname = "ListBox_setting_Status") Or _
           (keyname = "ListBox_Setting_Status_parents") Or _
           (keyname = "ListBox_Setting_Status_granpa") Or _
           (keyname = "ListBox_Setting_Status_child") Or _
           (keyname = "ListBox_Setting_Tracker_parents") Or _
           (keyname = "ListBox_Setting_Tracker_grandparent") Or _
           (keyname = "ListBox_Setting_Tracker_child") Or _
           (keyname = "ListBox_New_Ticket_Story") Or _
           (keyname = "ListBox_New_Ticket_Activity") Or _
           (keyname = "ListBox_New_Ticket_Backlog") Or _
           (keyname = "ListBox_setting_TimeEntryActivity") _
        ) _
    Then
        If dic_setthing.exists(keyname) Then
            dic_setthing.Remove keyname
        End If

        For i = 0 To myctrl.ListCount - 1
            ' If debug_ Then Debug.Print "Member " & myctrl.List(i)
            ' If debug_ Then Debug.Print "Selected " & myctrl.Selected(i)
            ' RMからの辞書にある場合


            If keyname = "ListBox_setting_priority" And Dic_Priority.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                If debug_ Then Debug.Print keyname & " added " & myctrl.List(i) & "=" & Dic_Priority(myctrl.List(i))
                childnode.Add myctrl.List(i), Dic_Priority(myctrl.List(i))
            ElseIf keyname = "ListBox_setting_Status" And Dic_Statuses.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                If debug_ Then Debug.Print keyname & " added " & myctrl.List(i) & "=" & Dic_Statuses(myctrl.List(i))
                childnode.Add myctrl.List(i), Dic_Statuses(myctrl.List(i))
            ElseIf keyname = "ListBox_Setting_Status_parents" And Dic_Statuses.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                If debug_ Then Debug.Print keyname & " added " & myctrl.List(i) & "=" & Dic_Statuses(myctrl.List(i))
                childnode.Add myctrl.List(i), Dic_Statuses(myctrl.List(i))
            ElseIf keyname = "ListBox_Setting_Status_granpa" And Dic_Statuses.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                If debug_ Then Debug.Print keyname & " added " & myctrl.List(i) & "=" & Dic_Statuses(myctrl.List(i))
                childnode.Add myctrl.List(i), Dic_Statuses(myctrl.List(i))
            ElseIf keyname = "ListBox_Setting_Status_child" And Dic_Statuses.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                If debug_ Then Debug.Print keyname & " added " & myctrl.List(i) & "=" & Dic_Statuses(myctrl.List(i))
                childnode.Add myctrl.List(i), Dic_Statuses(myctrl.List(i))
            ElseIf keyname = "ListBox_Setting_Tracker_parents" And Dic_Trackers.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                If debug_ Then Debug.Print keyname & " added " & myctrl.List(i) & "=" & Dic_Trackers(myctrl.List(i))
                childnode.Add myctrl.List(i), Dic_Trackers(myctrl.List(i))
            ElseIf keyname = "ListBox_Setting_Tracker_grandparent" And Dic_Trackers.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                If debug_ Then Debug.Print keyname & " added " & myctrl.List(i) & "=" & Dic_Trackers(myctrl.List(i))
                childnode.Add myctrl.List(i), Dic_Trackers(myctrl.List(i))
            ElseIf keyname = "ListBox_Setting_Tracker_child" And Dic_Trackers.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                If debug_ Then Debug.Print keyname & " added " & myctrl.List(i) & "=" & Dic_Trackers(myctrl.List(i))
                childnode.Add myctrl.List(i), Dic_Trackers(myctrl.List(i))
            ElseIf keyname = "ListBox_New_Ticket_Story" And Dic_Trackers.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                If debug_ Then Debug.Print keyname & " added " & myctrl.List(i) & "=" & Dic_Trackers(myctrl.List(i))
                childnode.Add myctrl.List(i), Dic_Trackers(myctrl.List(i))
            ElseIf keyname = "ListBox_New_Ticket_Activity" And Dic_Trackers.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                If debug_ Then Debug.Print keyname & " added " & myctrl.List(i) & "=" & Dic_Trackers(myctrl.List(i))
                childnode.Add myctrl.List(i), Dic_Trackers(myctrl.List(i))
            ElseIf keyname = "ListBox_New_Ticket_Backlog" And Dic_Trackers.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                If debug_ Then Debug.Print keyname & " added " & myctrl.List(i) & "=" & Dic_Trackers(myctrl.List(i))
                childnode.Add myctrl.List(i), Dic_Trackers(myctrl.List(i))
            ElseIf keyname = "ListBox_setting_TimeEntryActivity" And Dic_TimeEntryActivity.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                If debug_ Then Debug.Print keyname & " added " & myctrl.List(i) & "=" & Dic_TimeEntryActivity(myctrl.List(i))
                childnode.Add myctrl.List(i), Dic_TimeEntryActivity(myctrl.List(i))
            ElseIf keyname = "ListBox_Setting_Tracker_child" And Dic_Trackers.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                If debug_ Then Debug.Print keyname & " added " & myctrl.List(i) & "=" & Dic_Trackers(myctrl.List(i))
                childnode.Add myctrl.List(i), Dic_TimeEntryActivity(myctrl.List(i))
            ElseIf keyname = "ListBox_Setting_Status_child" And Dic_Statuses.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                If debug_ Then Debug.Print keyname & " added " & myctrl.List(i) & "=" & Dic_Statuses(myctrl.List(i))
                childnode.Add myctrl.List(i), Dic_TimeEntryActivity(myctrl.List(i))
            End If
        Next i
        dic_setthing.Add keyname, childnode
        
        If dic_setthing.exists("") Then
          dic_setthing.Remove ""
        End If
    'ListBox_User_Settings は既存のものとマージする。RM辞書にないものが含まれている可能性があるため
    ElseIf (keyname = "ListBox_User_Settings") Then

       Dim testdic As Object
       Dim tmpstr As Variant

       If dic_setthing.exists(keyname) Then
            If debug_ Then Debug.Print "ListBox_User_Settings packager :: keyname " & keyname & " is already exists "
            Set testdic = dic_setthing(keyname)
       Else
            If debug_ Then Debug.Print "ListBox_User_Settings packager :: keyname " & keyname & " is not exists "
            Set testdic = New Dictionary
            dic_setthing.Add keyname, testdic
       End If
       For i = 0 To myctrl.ListCount - 1
            If debug_ Then Debug.Print "checking " & myctrl.List(i)
            'RM loaded 辞書にいて、ローカル辞書に存しない場合
            If Dic_Users.exists(myctrl.List(i)) And (Not testdic.exists(myctrl.List(i))) Then
                    If debug_ Then Debug.Print "this is on dic and not exist localsetting"
                    testdic.Add myctrl.List(i), Dic_Users(myctrl.List(i))
            'RM loaded 辞書にいて、ローカル辞書存在する場合
            ElseIf Dic_Users.exists(myctrl.List(i)) And (testdic.exists(myctrl.List(i))) Then
                    If debug_ Then Debug.Print "this is on dic and not exist localsetting"
                    testdic(myctrl.List(i)) = Dic_Users(myctrl.List(i))
            'RM loaded 辞書に無い、ローカルにはある場合
            ElseIf Dic_Users.exists(myctrl.List(i)) = False And testdic.exists(myctrl.List(i)) = True Then
                    If debug_ Then Debug.Print "this is on localsetting。do nothing"
            Else
                    If debug_ Then Debug.Print "dic is " & Dic_Users.exists(myctrl.List(i))
                    If debug_ Then Debug.Print "local is " & testdic.exists(myctrl.List(i))
            End If
      Next i

      '消されたもの( =辞書にもない場合があり判断できない )を取り除く
      If debug_ Then Debug.Print "delete user :: " & JSONLib.toString(tmpdelluser)
      For Each tmpstr In tmpdelluser
        If tmpdelluser(tmpstr) = 1 And testdic.exists(tmpstr) Then
            If debug_ Then Debug.Print "user deleted " & tmpstr
            testdic.Remove tmpstr
        End If
      Next tmpstr
      
      If testdic.exists("") Then
        testdic.Remove ""
      End If
      
      tmpdelluser.RemoveAll
    Else
    End If

    If debug_ Then Debug.Print keyname & " : LocalData dump packed: " & JSONLib.toString(dic_setthing(keyname))


End Sub
Private Sub ListBox_Packager_Val(ByRef dic_setthing As Object, ByVal keyname As String, ByVal myctrl As MSForms.ListBox)
    Dim i As Integer
    Dim childstr As String
    If debug_ Then Debug.Print "Create " & keyname & " Dump:"
    
    If dic_setthing.exists(keyname) Then
        dic_setthing.Remove keyname
    End If
    
    ' 基本的には設定は全てクリアして、フォームで選ばれたものを詰め込む
    If ( _
           (keyname = "ListBox_New_Ticket_Story") Or _
           (keyname = "ListBox_New_Ticket_Activity") Or _
           (keyname = "ListBox_New_Ticket_Backlog") _
        ) _
    Then
        For i = 0 To myctrl.ListCount - 1
            ' If debug_ Then Debug.Print "Member " & myctrl.List(i)
            ' If debug_ Then Debug.Print "Selected " & myctrl.Selected(i)
            ' RMからの辞書にある場合
    
            If keyname = "ListBox_New_Ticket_Story" And Dic_Trackers.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                childstr = Dic_Trackers(myctrl.List(i))
            ElseIf keyname = "ListBox_New_Ticket_Activity" And Dic_Trackers.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                childstr = Dic_Trackers(myctrl.List(i))
            ElseIf keyname = "ListBox_New_Ticket_Backlog" And Dic_Trackers.exists(myctrl.List(i)) And myctrl.Selected(i) = True Then
                childstr = Dic_Trackers(myctrl.List(i))

            End If
        Next i
        dic_setthing.Add keyname, childstr
    End If
End Sub
Private Sub select_listbox(ByRef setting As Object, ByRef myctrl As MSForms.ListBox, ByVal dic As Object)
'RM辞書登録済みデータ
    Dim JSONLib As New JSONLib
    If debug_ Then Debug.Print "LocalData dump : " & JSONLib.toString(setting)
    Dim i As Integer
        For i = 0 To myctrl.ListCount - 1
            If debug_ Then Debug.Print "now check " & myctrl.List(i)
            If debug_ Then Debug.Print "this dic val " & setting(myctrl.List(i))
            If Not setting(myctrl.List(i)) = "" Then
                myctrl.Selected(i) = True
            End If
        Next i
End Sub

Private Sub select_listbox_users(ByRef setting As Object, ByRef myctrl As MSForms.ListBox, ByVal dic As Object)
'マスタから直に取らないユーザ辞書
    Dim JSONLib As New JSONLib
    Dim Var As Variant
    If debug_ Then Debug.Print "LocalData dump : " & JSONLib.toString(setting)
        For Each Var In setting
            If Var <> "" And selectlist_checkbyval(myctrl, Var) < 0 Then
                myctrl.AddItem Var
            End If
        Next Var
End Sub
Private Sub select_listbox_val(ByRef val As String, ByRef myctrl As MSForms.ListBox, ByVal dic As Object)
'数値データのみを持つデフォルトTrucker項目
    Dim JSONLib As New JSONLib
    If debug_ Then Debug.Print "LocalData dump : " & val
        Dim i As Integer
        For i = 0 To myctrl.ListCount - 1
            If debug_ Then Debug.Print "now check : " & myctrl.List(i)
            If Dic_Trackers(myctrl.List(i)) = val Then
                myctrl.Selected(i) = True
                Exit Sub
            End If
        Next i
End Sub
Private Sub add_listbox(ByVal setting As Object, ByRef myctrl As MSForms.ListBox)
' ユーザリスト用
    Dim Var As Variant
    Dim JSONLib As New JSONLib
    If debug_ Then Debug.Print "LocalData dump : " & JSONLib.toString(setting)
    Dim i As Integer
    For Each Var In setting
        If selectlist_checkbyval(myctrl, Var) < 0 Then
            myctrl.AddItem Var
        End If
    Next Var
End Sub

Private Sub rmtc_setting_initializer()
    If RMTC_Creater.Visible Then
    Else
        Unload Me
        Exit Sub
    End If
    
    If LocalSavedSettings.exists("Setting_Redmine_APIKEY") Then
        TextBox_Redmine_APIKey.value = LocalSavedSettings("Setting_Redmine_APIKEY")
    End If
    If LocalSavedSettings.exists("Setting_Redmine_URL") Then
        TextBox_RedmineURL.value = LocalSavedSettings("Setting_Redmine_URL")
    End If
    
End Sub

Private Sub Dic_Packager(ByRef dic_setthing As Object, ByVal keyname As String, ByVal dic As Object)
    If debug_ Then Debug.Print "Create " & keyname & " Dump:"
    
    If dic_setthing.exists(keyname) Then
        dic_setthing.Remove keyname
    End If
    dic_setthing.Add keyname, dic
End Sub


Private Sub ListBox_Setting_Tracker_parents_Click()

End Sub

Private Sub UserForm_Initialize()
    If Initialized = 1 Then
        TextBox_RedmineURL.Text = Setting_Redmine_URL
        TextBox_Redmine_APIKey.Text = Setting_Redmine_APIKEY
        CommandButton_LoadSettingForm_Click
    Else
        MsgBox "this form is not initialized"
        Unload Me
        Exit Sub
    End If

End Sub
Private Function selectlist_checkbyval(ByRef myctrl As MSForms.ListBox, ByVal val As String)
    Dim i As Long
    For i = 0 To myctrl.ListCount - 1             ''(1)
        If myctrl.List(i) = val Then    ''(2)
            selectlist_checkbyval = i
            Exit Function
        End If
    Next i
    selectlist_checkbyval = -1
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CheckBox_delete_regdata.value = True Then
        DeleteSetting ("OutlookRMTC")
        Initialized = 0
        Unload Me
    End If
End Sub
