VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RMTM_Creater 
   Caption         =   "Redmine Create TimeEntry"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5340
   OleObjectBlob   =   "RMTM_Creater.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "RMTM_Creater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private selected_ticket_id As String

Public Sub set_select_ticket_id(ByRef ticketid As String, ByRef subject As String)
    selected_ticket_id = ticketid
    Label_selected_ticket.Caption = "#" & ticketid & ":" & subject
    Label_selected_ticket.ControlTipText = Label_selected_ticket.Caption
End Sub

Private Sub Button_Settings_Click()
    Call save_transaction_Data_to_reg
    RMTC_Setting.Show vbModeless
    If Initialized = 1 Then
        Call rmtm_initializer
        Call RMTS_Search.rmts_initialize
    Else
        Unload Me
    End If
    
End Sub

Private Sub ComboBox_assigned_me_Change()
    If Dic_Assigned_To_Me.exists(ComboBox_assigned_me.Text) Then
         selected_ticket_id = Dic_Assigned_To_Me(ComboBox_assigned_me.Text)
         Label_selected_ticket.Caption = ComboBox_assigned_me.Text
         Label_selected_ticket.ControlTipText = Label_selected_ticket.Caption
    End If
    Call favorite_initialize(selected_ticket_id)
End Sub

Private Sub ComboBox_fevoritelist_Change()
    Dim Var As Variant

    If TransactionTimeEntryData.exists("favoritelist") Then
        Set tmpdic = TransactionTimeEntryData("favoritelist")
        For Each Var In tmpdic
            If tmpdic(Var) = ComboBox_fevoritelist.value Then
                selected_ticket_id = Var
                Label_selected_ticket.Caption = ComboBox_fevoritelist.value
                Label_selected_ticket.ControlTipText = Label_selected_ticket.Caption
            End If
        Next Var
    End If
    Call favorite_initialize(selected_ticket_id)
 '   TextBox_Comment.Text = ""
End Sub

Private Sub ComboBox_parentActivity_Change()
    ComboBox_parentBacklog.Clear
    If debug_ Then Debug.Print "ComboBox_parentActivity_Change :: select change : Activity = " & ComboBox_parentActivity.value
    If Dic_Activity.exists(ComboBox_parentActivity.value) Then
        If debug_ Then Debug.Print "ComboBox_parentActivity_Change :: select change : Activity id = " & Dic_Activity(ComboBox_parentActivity.value)
        selected_ticket_id = Dic_Activity(ComboBox_parentActivity.value)
        Label_selected_ticket.Caption = ComboBox_parentActivity.value
        Label_selected_ticket.ControlTipText = Label_selected_ticket.Caption
        Call favorite_initialize(selected_ticket_id)

        Call set_backlog_ticket_for_selected_activity(LocalSavedSettings, Setting_Redmine_URL, Setting_Redmine_APIKEY)

        ComboBox_parentBacklog.Enabled = True
    
    End If
End Sub

Private Sub ComboBox_parentBacklog_Change()
   selected_ticket_id = Dic_Backlog(ComboBox_parentBacklog.value)
   Label_selected_ticket.Caption = ComboBox_parentBacklog.value
   Label_selected_ticket.ControlTipText = Label_selected_ticket.Caption
   Call favorite_initialize(selected_ticket_id)
End Sub

Private Sub CommandButton_setfavorite_DBClick()
    Call CommandButton_setfavorite_Click
End Sub
Private Sub CommandButton_setfavorite_Click()
    Dim tmpdic As Dictionary
    Set tmpdic = New Dictionary
    If selected_ticket_id = "" Then
    Else
        If TransactionTimeEntryData.exists("favoritelist") Then
            Set tmpdic = TransactionTimeEntryData("favoritelist")
            If tmpdic.exists(selected_ticket_id) Then
                If tmpdic(selected_ticket_id) <> "" Then
                    CommandButton_setfavorite.Caption = "☆"
                    tmpdic.Remove (selected_ticket_id)
                Else
                    CommandButton_setfavorite.Caption = "★"
                    tmpdic(selected_ticket_id) = Label_selected_ticket.Caption
                End If
            Else
                CommandButton_setfavorite.Caption = "★"
                tmpdic(selected_ticket_id) = Label_selected_ticket.Caption
            End If
        Else
            Set TransactionTimeEntryData("favoritelist") = tmpdic
            CommandButton_setfavorite.Caption = "★"
            tmpdic(selected_ticket_id) = Label_selected_ticket.Caption
        End If
    End If
    Call draw_favorite_box
End Sub

Private Sub CommandButton_StartCaleder_Click()
   Call CalenderForm.setDate(GetToday())
   Call CalenderForm.setCallBackControl(Label_ActivityDate)
   CalenderForm.Show vbModeless
End Sub
Private Function GetToday()
    GetToday = Year(Now) & "/" & Month(Now) & "/" & Day(Now)
End Function
Private Sub CommandButton_SubmitTimeEntry_Click()
    If debug_ Then Debug.Print "CommandButton_SubmitTimeEntry_Click Called"
    Dim JSONLib As New JSONLib
    Dim tmpstoryid, tmpactivityid, tmpbacklogid As String
    
    If selected_ticket_id = "" Then
        MsgBox "find ticket is failed"
        Exit Sub
    End If

    RequestURL = Setting_Redmine_URL & "/time_entries.json?format=xml&key=" & Setting_Redmine_APIKEY

    If debug_ Then Debug.Print "RequestURL is " & RequestURL
    Set xhr = CreateObject("Microsoft.XMLHTTP")
    xhr.Open "POST", RequestURL, False
    xhr.SetRequestHeader "Content-Type", "text/xml"
    
    RequestBody = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>"
    RequestBody = RequestBody & "<time_entry>"


    If selected_ticket_id <> "" Then
        If debug_ Then Debug.Print "  :: issue id is " & selected_ticket_id
        RequestBody = RequestBody & "<issue_id>" & selected_ticket_id & "</issue_id>"
    End If

    If Label_ActivityDate.Caption <> "" Then
        If debug_ Then Debug.Print "  :: spent_on is " & Format(Label_ActivityDate.Caption, "yyyy-mm-dd")
        RequestBody = RequestBody & "<spent_on>" & Format(Label_ActivityDate.Caption, "yyyy-mm-dd") & "</spent_on>"
    End If

    If ComboBox_TimeEntryActivity.Text <> "" Then
        If debug_ Then Debug.Print "  :: activity_id is " & Dic_TimeEntryActivity(ComboBox_TimeEntryActivity.Text) & " its on " & JSONLib.toString(Dic_TimeEntryActivity)
        RequestBody = RequestBody & "<activity_id>" & Dic_TimeEntryActivity(ComboBox_TimeEntryActivity.Text) & "</activity_id>"
    End If


    If TextBox_timeentryhours.value <> "" Then
        If debug_ Then Debug.Print "  :: hours is " & TextBox_timeentryhours.Text
        RequestBody = RequestBody & "<hours>" & TextBox_timeentryhours.value & "</hours>"
    End If
    

    If TextBox_Comment.Text <> "" Then
        If debug_ Then Debug.Print "  :: comments is " & TextBox_Comment.Text
        RequestBody = RequestBody & "<comments>" & TextBox_Comment.Text & "</comments>"
    End If
    
    RequestBody = RequestBody & "</time_entry>"

    If debug_ Then Debug.Print "send xml : " & RequestBody
            
    xhr.Send (RequestBody)

    If xhr.Status = 201 Then
        Call check_my_timeentry_on_today(Setting_Redmine_URL, Setting_Redmine_APIKEY)
        msgreturn = MsgBox("#" & selected_ticket_id & " create time entry . Open Web?", vbYesNo)

            If debug_ Then Debug.Print "user choice is : " & msgreturn

            If msgreturn = vbNo Then
            ElseIf msgreturn = vbYes Then
                If debug_ Then Debug.Print "created backlog open web start : " & Setting_Redmine_URL & "/issue/" & selected_ticket_id & "/time_entries?key=" & Setting_Redmine_APIKEY
                
                If webincreasemyAPIKey = 1 Then
                    openweb (Setting_Redmine_URL & "/issues/" & selected_ticket_id & "/time_entries?key=" & Setting_Redmine_APIKEY)
                Else
                    openweb (Setting_Redmine_URL & "/issues/" & selected_ticket_id & "/time_entries")
                End If
            End If
    Else
        Set json = JSONLib.parse(xhr.responseText)
        MsgBox json("errors").Item(1)
        If debug_ Then Debug.Print "error occured " & json("errors").Item(1)
    End If

End Sub
Private Sub CommandButton2_Click()
    Dim ans As String
    Dim ticketid  As Integer
    
    ans = InputBox("ticket id or kwy word", "get ticket", "")
    If ans = "" Then
        Exit Sub
    End If
    
    If IsNumeric(ans) Then
        ticketid = StrConv(ans, vbNarrow)
        Call get_ticket_subject(ticketid, Setting_Redmine_URL, Setting_Redmine_APIKEY)
        Call favorite_initialize("" & ticketid)
    Else
        Dim myProject, myStory As String
        If LocalSavedSettings.exists("Dic_Projects_ID") And LocalSavedSettings("Dic_Projects_ID").exists(ComboBox_Project.value) Then
            myProject = LocalSavedSettings("Dic_Projects_ID")(ComboBox_Project.value)
        End If
    
        Call RMTS_Search.set_param(myProject, ComboBox_Project.value)
        RMTS_Search.TextBox_SearchKey = ans
        RMTS_Search.CommandButton_SearchTicket_Click
        RMTS_Search.Show vbModeless
        Call favorite_initialize(selected_ticket_id)
    End If
End Sub

Private Sub CommandButton3_Click()
    Call check_my_timeentry_on_today(Setting_Redmine_URL, Setting_Redmine_APIKEY)
End Sub

Private Sub CommandButton4_Click()
    selected_ticket_id = ""
    TextBox_Comment.Text = ""
    Label_selected_ticket.Caption = ""
    Label_selected_ticket.ControlTipText = ""
    ComboBox_parentBacklog.Clear
    ComboBox_parentActivity.Clear
    ComboBox_ParentStory.value = ""
    ComboBox_parentActivity.Enabled = False
    ComboBox_parentBacklog.Enabled = False
End Sub

Private Sub CommandButton5_Click()
    Unload Me
    Redmint_CreateTicket
End Sub

Private Sub Label_assigned_to_me_Click()
    Call set_activity_ticket_for_assigned_id_to_me(LocalSavedSettings, Setting_Redmine_URL, Setting_Redmine_APIKEY)
End Sub

Private Sub Label_selected_ticket_Click()
    If debug_ Then Debug.Print "story open web start : " & Setting_Redmine_URL & "/issue/" & selected_ticket_id & "?key=" & Setting_Redmine_APIKEY
    If selected_ticket_id = "" Then
        Exit Sub
    End If
    If webincreasemyAPIKey = 1 Then
        openweb (Setting_Redmine_URL & "/issues/" & selected_ticket_id & "?key=" & Setting_Redmine_APIKEY)
    Else
        openweb (Setting_Redmine_URL & "/issues/" & selected_ticket_id)
    End If
End Sub
Private Sub Label_selected_ticket_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim buf As Long
    If Button = 2 Then
        If selected_ticket_id = "" Then
            Exit Sub
        End If
        Call get_ticket_subject_for_caption("" + selected_ticket_id, Setting_Redmine_URL, Setting_Redmine_APIKEY)
    End If
End Sub

Private Sub Label_todaytimeentry_reload_Click()
    Call check_my_timeentry_on_today(Setting_Redmine_URL, Setting_Redmine_APIKEY)
End Sub



Private Sub ListBox_mytimeentry_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If ListBox_mytimeentry.value = "" Then
        Exit Sub
    End If
    If webincreasemyAPIKey = 1 Then
        openweb (Setting_Redmine_URL & "/issues/" & ListBox_mytimeentry.value & "/time_entries?key = " & Setting_Redmine_APIKEY)
    Else
        openweb (Setting_Redmine_URL & "/issues/" & ListBox_mytimeentry.value & "/time_entries")

    End If
End Sub
Private Sub ListBox_mytimeentry_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim buf As Long
    If Button = 2 Then
        buf = Int((Y + 1) / ListBox_mytimeentry.Font.Size)
        If buf > ListBox_mytimeentry.ListCount - 1 Then
            buf = ListBox_mytimeentry.ListCount - 1
        End If
        ListBox_mytimeentry.Selected(buf) = True

    Call get_ticket_subject_for_caption(ListBox_mytimeentry.List(buf, 0), Setting_Redmine_URL, Setting_Redmine_APIKEY)
    End If
End Sub
Private Sub ScrollBar_timeentry_Change()
    TextBox_timeentryhours.Text = 0 - ScrollBar_timeentry.value * 0.25
End Sub

Private Sub TextBox_timeentryhours_Change()

    If TextBox_timeentryhours.value = "" Then
        TextBox_timeentryhours.value = "0"
    End If
    ScrollBar_timeentry.value = 0 - CSng(TextBox_timeentryhours.value) / 0.25

End Sub

Private Sub UserForm_Initialize()
    If Initialized = 1 Then
    Else
        MsgBox "failed to load"
        Me.Width = 0
        Me.Height = 0
        Exit Sub
    End If

End Sub
Private Sub UserForm_Activate()
        If time_entries <> 1 Then
        Exit Sub
        End If
        
        If Setting_Redmine_APIKEY = "" Or Setting_Redmine_URL = "" Then
            Button_Settings_Click
        End If
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Initialized = 1 Then
        Dim JSONLib As New JSONLib
        Dim json As Object
        Dim RegStr As String

        RegStr = JSONLib.toString(TransactionTimeEntryData)
        If debug_ Then Debug.Print "UserForm_QueryClose :: save to reg : "; RegStr
        Call save_transaction_Data_to_reg
        If debug_ Then Debug.Print "UserForm_QueryClose ended"
    Else
        If debug_ Then Debug.Print "UserForm_QueryClose called but not initialized or going reload"
        Exit Sub
    End If
    

End Sub


Public Sub rmtm_initializer()
If debug_ Then Debug.Print "rmtm_initializer Called"
    
    Dim Var As Variant
    Dim RegStr As String
    Dim JSONLib As New JSONLib
    Dim tmpdic As Object
    
    RegStr = GetSetting("OutlookRMTC", "Settings", "AllSettings")
    
    If RegStr = "" Then
        Exit Sub
    End If
    If debug_ Then Debug.Print "UserForm_Initialize :: get regset : AllSetting" & RegStr
    Set LocalSavedSettings = JSONLib.parse(RegStr)
    Debug.Assert Err.Number = 0

    If LocalSavedSettings.exists("Setting_Redmine_APIKEY") Then
        Setting_Redmine_APIKEY = LocalSavedSettings("Setting_Redmine_APIKEY")
    Else
        If debug_ Then Debug.Print "can not find LocalSavedSettings(""Setting_Redmine_APIKEY"")"
    End If
    
    If LocalSavedSettings.exists("Setting_Redmine_URL") Then
        Setting_Redmine_URL = LocalSavedSettings("Setting_Redmine_URL")
    Else
        If debug_ Then Debug.Print "can not find LocalSavedSettings(""Setting_Redmine_URL"")"
    End If
    
    If LocalSavedSettings.exists("webincreasemyAPIKey") Then
        webincreasemyAPIKey = LocalSavedSettings("webincreasemyAPIKey")
    Else
        If debug_ Then Debug.Print "can not find LocalSavedSettings(""webincreasemyAPIKey"")"
    End If
    
    If LocalSavedSettings.exists("keywordsearchonAllTrackers") Then
        keywordsearchonAllTrackers = LocalSavedSettings("keywordsearchonAllTrackers")
    Else
        If debug_ Then Debug.Print "can not find LocalSavedSettings(""keywordsearchonAllTrackers"")"
    End If
 
    If LocalSavedSettings.exists("searchContents") Then
        searchContents = LocalSavedSettings("searchContents")
    Else
        If debug_ Then Debug.Print "can not find LocalSavedSettings(""searchContents"")"
    End If

    RMTM_Creater.ComboBox_Project.Clear
    RMTM_Creater.ComboBox_ParentStory.Clear
    RMTM_Creater.ComboBox_parentActivity.Clear
    RMTM_Creater.ComboBox_parentBacklog.Clear
    RMTM_Creater.ComboBox_TimeEntryActivity.Clear

    If LocalSavedSettings.exists("Dic_Projects") Then
        Set tmpdic = LocalSavedSettings("Dic_Projects")
        For Each Var In tmpdic
            RMTM_Creater.ComboBox_Project.AddItem Var
        Next Var
    End If

    If LocalSavedSettings.exists("ListBox_setting_TimeEntryActivity") Then
        Set tmpdic = LocalSavedSettings("ListBox_setting_TimeEntryActivity")
        For Each Var In tmpdic
            RMTM_Creater.ComboBox_TimeEntryActivity.AddItem Var
        Next Var
        Set Dic_TimeEntryActivity = tmpdic
    End If

    RegStr = GetSetting("OutlookRMTC", "Transaction", "TimeEntryTrans")
    If debug_ Then Debug.Print "UserForm_Initialize :: get regset : TransactionTimeEntryData = " & RegStr
    Set TransactionTimeEntryData = JSONLib.parse(RegStr)
    Debug.Assert Err.Number = 0

    If Not TransactionTimeEntryData Is Nothing Then
        If debug_ Then Debug.Print "UserForm_Initialize :: TransactionTimeEntryData data " & JSONLib.toString(TransactionTimeEntryData)
        If debug_ Then Debug.Print "apply TransactionTimeEntryData data to form"
        If TransactionTimeEntryData.exists("Default_TimeEntryActivity") Then
            RMTM_Creater.ComboBox_TimeEntryActivity.value = TransactionTimeEntryData("Default_TimeEntryActivity")
        End If
        If TransactionTimeEntryData.exists("Dic_Projects") Then
            RMTM_Creater.ComboBox_Project.value = TransactionTimeEntryData("Dic_Projects")
        End If
    Else
        If debug_ Then Debug.Print "UserForm_Initialize :: TransactionTimeEntryData data is nothing "
        Set TransactionTimeEntryData = New Dictionary
    End If

    Call draw_favorite_box

    Call check_my_timeentry_on_today(Setting_Redmine_URL, Setting_Redmine_APIKEY)
    
    Call set_activity_ticket_for_assigned_id_to_me(LocalSavedSettings, Setting_Redmine_URL, Setting_Redmine_APIKEY)

    ListBox_mytimeentry.ColumnWidths = "30;65;25;50"
    TextBox_Comment.SetFocus
End Sub
Private Function draw_favorite_box()
    ComboBox_fevoritelist.Clear
    If TransactionTimeEntryData.exists("favoritelist") Then
        For Each Var In TransactionTimeEntryData("favoritelist")
            ComboBox_fevoritelist.AddItem TransactionTimeEntryData("favoritelist")(Var)
        Next Var
    End If
End Function
Private Sub save_transaction_Data_to_reg()
    If debug_ Then Debug.Print "Save TransactionTimeEntryData data to reg."
    
    If TransactionTimeEntryData Is Nothing Then
        If debug_ Then Debug.Print "TransactionTimeEntryData Data is nohing"
        Exit Sub
    End If
    
    If debug_ Then Debug.Print "Save TransactionTimeEntryData data to reg ComboBox_TimeEntryActivity.value = " & ComboBox_TimeEntryActivity.value
    If TransactionTimeEntryData.exists("Default_TimeEntryActivity") Then
        TransactionTimeEntryData("Default_TimeEntryActivity") = RMTM_Creater.ComboBox_TimeEntryActivity.value
    Else
        TransactionTimeEntryData.Add "Default_TimeEntryActivity", RMTM_Creater.ComboBox_TimeEntryActivity.value
    End If

    If debug_ Then Debug.Print "Save TransactionTimeEntryData data to reg ComboBox_Project.value = " & ComboBox_Project.value
    If TransactionTimeEntryData.exists("Dic_Projects") Then
        TransactionTimeEntryData("Dic_Projects") = RMTM_Creater.ComboBox_Project.value
    Else
        TransactionTimeEntryData.Add "Dic_Projects", RMTM_Creater.ComboBox_Project.value
    End If


    Dim JSONLib As New JSONLib
    If debug_ Then Debug.Print "save_transaction_Data_to_reg :: TransactionTimeEntryData data " & JSONLib.toString(TransactionTimeEntryData)
    
    SaveSetting "OutlookRMTC", "Transaction", "TimeEntryTrans", JSONLib.toString(TransactionTimeEntryData)
End Sub
Private Sub ComboBox_Project_Change()
    If debug_ Then Debug.Print "ComboBox_Project_Change :: " & ComboBox_Project.value
    Dic_Story.RemoveAll
    Dic_Activity.RemoveAll
    ComboBox_ParentStory.Clear
    ComboBox_parentActivity.Clear
    ComboBox_parentBacklog.Clear
    
    ComboBox_parentActivity.Enabled = False
    ComboBox_parentBacklog.Enabled = False
 '   TextBox_Comment.Text = ""
    Call set_story_ticket_for_selected_project(LocalSavedSettings, Setting_Redmine_URL, Setting_Redmine_APIKEY)
    If debug_ Then Debug.Print "ComboBox_Project_Change :: ended"
End Sub

Public Sub set_story_ticket_for_selected_project(ByRef LocalSavedSettings As Object, ByVal url As String, ByVal apikey As String)
    If debug_ Then Debug.Print "★start★Calle :: set_story_ticket_for_selected_project"
    Dim JSONLib As New JSONLib
    Dim json As Object
    Dim tmpdic As Object
    Dim Var As Variant
    Dim total, offset, limit, nextoffset As Integer
    If debug_ Then Debug.Print "Calle :: set_story_ticket_for_selected_project"
    Set Dic_Story = New Dictionary
    
    Dim myProject, myStory As String
    If LocalSavedSettings.exists("Dic_Projects_ID") And LocalSavedSettings("Dic_Projects_ID").exists(ComboBox_Project.value) Then
        myProject = LocalSavedSettings("Dic_Projects_ID")(ComboBox_Project.value)
    End If
    If myProject = "" Then
        If debug_ Then Debug.Print "Not found; project "
        Exit Sub
        
    End If
    Dim filterstr, filter_status, filter_tracker  As String
    filterstr = ""
    filter_status = ""
    filter_tracker = ""
    Set Var = New Dictionary
    Set tmpstatusdic = New Dictionary
    Set tmptracksdic = New Dictionary
    
    If LocalSavedSettings.exists("ListBox_Setting_Status_granpa") Then
        Set tmpstatusdic = LocalSavedSettings("ListBox_Setting_Status_granpa")
        For Each Var In tmpstatusdic
            If filter_status = "" Then
                filter_status = "status_id=" & tmpstatusdic(Var)
            Else
                filter_status = filter_status & "|" & tmpstatusdic(Var)
            End If
        Next Var
        
        If filterstr = "" Then
            filterstr = filter_status
        Else
            filterstr = filterstr & "&" & filter_status
        End If
    End If
    If LocalSavedSettings.exists("ListBox_Setting_Tracker_grandparent") Then
        Set tmptracksdic = LocalSavedSettings("ListBox_Setting_Tracker_grandparent")
        For Each Var In tmptracksdic
            If filter_tracker = "" Then
                filter_tracker = "tracker_id=" & tmptracksdic(Var)
            Else
                filter_tracker = filter_tracker & "|" & tmptracksdic(Var)
            End If
        Next Var
        If filterstr = "" Then
            filterstr = filter_tracker
        Else
            filterstr = filterstr & "&" & filter_tracker
        End If
    End If
    Set Dic_Users = Nothing
    Set Dic_Users = New Dictionary

    Dim jsonstring As String
    jsonstring = GetData(url & "/issues.json?key=" & apikey & "&project_id=" & myProject & "&" & filterstr)
    Set json = New Dictionary
    Set json = JSONLib.parse(jsonstring)
    If json Is Nothing Then
        MsgBox "Cant load rm."
        Exit Sub
    End If
    total = json("total_count")
    offset = json("offset")
    limit = json("limit")
    nextoffset = val(limit) + val(offset)

    If debug_ Then Debug.Print "limit " & limit & " / offset " & offset & " / total " & total
    Do While total > nextoffset
        Dim subjsonstr As String
        subjsonstr = GetData(url & "/issues.json?key=" & apikey & "&project_id=" & myProject & "&offset=" & nextoffset & "&" & filterstr)
        Dim jsonsub As Object
        Set jsonsub = New Dictionary
        Set jsonsub = JSONLib.parse(subjsonstr)
        total = jsonsub("total_count")
        offset = jsonsub("offset")
        limit = jsonsub("limit")
        nextoffset = val(limit) + val(offset)
        If debug_ Then Debug.Print "limit " & limit & " / offset " & offset & " / total " & total

        For Each Var In jsonsub("issues")
            json("issues").Add Var
        Next Var
    Loop
    For Each Var In json("issues")
            If Var.exists("parent") = False Then
                If Var("project")("id") = myProject Then
                    If tmpstatusdic(Var("status")("name")) <> "" And tmptracksdic(Var("tracker")("name")) <> "" Then
                        ComboBox_ParentStory.AddItem "#" & Var("id") & ":" & Var("subject")
                        Dic_Story.Add "#" & Var("id") & ":" & Var("subject"), Var("id")
                    End If
                Else
                    If debug_ Then Debug.Print Var("id") & " : project_id  " & Var("project")("id") & " <>  " & myProject & ", story_id " & Var("parent")("id") & " <> " & myStory
                End If
            End If
    Next Var
    Set json = Nothing
    Set JSONLib = Nothing
If debug_ Then Debug.Print "★end★ set_story_ticket_for_selected_project"
End Sub
Public Sub set_activity_ticket_for_selected_story(ByRef LocalSavedSettings As Object, ByVal url As String, ByVal apikey As String)
    Dim JSONLib As New JSONLib
    Dim json, tmpdic As Object
    Dim Var As Variant
    Dim total, offset, limit, nextoffset As Integer
    If debug_ Then Debug.Print "★start★Calle :: set_activity_ticket_for_selected_story"
    Set Dic_Activity = New Dictionary
    
    Dim myProject, myStory As String
    If LocalSavedSettings.exists("Dic_Projects_ID") And LocalSavedSettings("Dic_Projects_ID").exists(ComboBox_Project.value) Then
        myProject = LocalSavedSettings("Dic_Projects_ID")(ComboBox_Project.value)
    End If
    If Dic_Story.exists(ComboBox_ParentStory.value) Then
        myStory = Dic_Story(ComboBox_ParentStory.value)
    End If
    
    If myProject = "" Or myStory = "" Then
        If debug_ Then Debug.Print "Not found; project or story "
        Exit Sub
        
    End If
    Dim filterstr, filter_status, filter_tracker  As String
    filterstr = ""
    filter_status = ""
    filter_tracker = ""
    Set Var = New Dictionary
    Set tmpstatusdic = New Dictionary
    Set tmptracksdic = New Dictionary

    If LocalSavedSettings.exists("ListBox_Setting_Status_parents") Then
        Set tmpstatusdic = LocalSavedSettings("ListBox_Setting_Status_parents")
        For Each Var In tmpstatusdic
            If filter_status = "" Then
                filter_status = "status_id=" & tmpstatusdic(Var)
            Else
                filter_status = filter_status & "|" & tmpstatusdic(Var)
            End If
        Next Var
        
        If filterstr = "" Then
            filterstr = filter_status
        Else
            filterstr = filterstr & "&" & filter_status
        End If
    End If
    If LocalSavedSettings.exists("ListBox_Setting_Tracker_parents") Then
        Set tmptracksdic = LocalSavedSettings("ListBox_Setting_Tracker_parents")
        For Each Var In tmptracksdic
            If filter_tracker = "" Then
                filter_tracker = "tracker_id=" & tmptracksdic(Var)
            Else
                filter_tracker = filter_tracker & "|" & tmptracksdic(Var)
            End If
        Next Var
        If filterstr = "" Then
            filterstr = filter_tracker
        Else
            filterstr = filterstr & "&" & filter_tracker
        End If
    End If
    Set Dic_Users = Nothing
    Set Dic_Users = New Dictionary


    Dim jsonstring As String
    
    jsonstring = GetData(url & "/issues.json?key=" & apikey & "&project_id=" & myProject & "&parent_id=" & myStory & "&" & filterstr)
    
    Set json = New Dictionary
    Set json = JSONLib.parse(jsonstring)
    If json Is Nothing Then
        MsgBox "Cant load rm."
        Exit Sub
    End If
 
    total = json("total_count")
    offset = json("offset")
    limit = json("limit")
    nextoffset = val(limit) + val(offset)

    If debug_ Then Debug.Print "limit " & limit & " / offset " & offset & " / total " & total
    Do While total > nextoffset
        Dim subjsonstr As String
        subjsonstr = GetData(url & "/issues.json?key=" & apikey & "&project_id=" & myProject & "&parent_id=" & myStory & "&offset=" & nextoffset & "&" & filterstr)
        Dim jsonsub As Object
        Set jsonsub = New Dictionary
        Set jsonsub = JSONLib.parse(subjsonstr)
        total = jsonsub("total_count")
        offset = jsonsub("offset")
        limit = jsonsub("limit")
        nextoffset = val(limit) + val(offset)
        If debug_ Then Debug.Print "limit " & limit & " / offset " & offset & " / total " & total
        For Each Var In jsonsub("issues")
            json("issues").Add Var
        Next Var

    Loop
    For Each Var In json("issues")
            If debug_ Then Debug.Print Var("id") & " : localfilter is enable"
            If Var.exists("parent") = True Then
                If debug_ Then Debug.Print Var("id") & " : fined parents"
                If Var("project")("id") = myProject And Var("parent")("id") = myStory Then
                    If tmpstatusdic(Var("status")("name")) <> "" And tmptracksdic(Var("tracker")("name")) <> "" Then
                        ComboBox_parentActivity.AddItem "#" & Var("id") & ":" & Var("subject")
                        Dic_Activity.Add "#" & Var("id") & ":" & Var("subject"), Var("id")
                    End If
                Else
                    If debug_ Then Debug.Print Var("id") & " : project_id  " & Var("project")("id") & " <>  " & myProject & ", story_id " & Var("parent")("id") & " <> " & myStory
                End If
            End If
    Next Var
    Set json = Nothing
    Set JSONLib = Nothing
If debug_ Then Debug.Print "★end★ set_activity_ticket_for_selected_story"
End Sub
Private Sub ComboBox_ParentStory_Change()
    ComboBox_parentActivity.Clear
    If debug_ Then Debug.Print "ComboBox_ParentStory_Change :: select change : Story = " & ComboBox_ParentStory.value
    If Dic_Story.exists(ComboBox_ParentStory.value) Then
        If debug_ Then Debug.Print "ComboBox_ParentStory_Change :: select change : Story id = " & Dic_Story(ComboBox_ParentStory.value)
        selected_ticket_id = Dic_Story(ComboBox_ParentStory.value)
        Label_selected_ticket.Caption = ComboBox_ParentStory.value
        Label_selected_ticket.ControlTipText = ComboBox_ParentStory.value
        Call favorite_initialize(selected_ticket_id)

        Call set_activity_ticket_for_selected_story(LocalSavedSettings, Setting_Redmine_URL, Setting_Redmine_APIKEY)

        ComboBox_parentActivity.Enabled = True
        ComboBox_parentBacklog.Enabled = False
    
    End If
 '   TextBox_Comment.Text = ""
End Sub
Public Sub set_backlog_ticket_for_selected_activity(ByRef LocalSavedSettings As Object, ByVal url As String, ByVal apikey As String)
If debug_ Then Debug.Print "★start★ set_backlog_ticket_for_selected_activity"
    Dim JSONLib As New JSONLib
    Dim json, tmpdic As Object
    Dim Var As Variant
    Dim total, offset, limit, nextoffset As Integer
    If debug_ Then Debug.Print "Calle :: set_backlog_ticket_for_selected_activity"
    Set Dic_Backlog = New Dictionary
    
    Dim myProject, myStory As String
    If LocalSavedSettings.exists("Dic_Projects_ID") And LocalSavedSettings("Dic_Projects_ID").exists(ComboBox_Project.value) Then
        myProject = LocalSavedSettings("Dic_Projects_ID")(ComboBox_Project.value)
    End If
    If Dic_Activity.exists(ComboBox_parentActivity.value) Then
        myStory = Dic_Activity(ComboBox_parentActivity.value)
    End If
    
    If myProject = "" Or myStory = "" Then
        If debug_ Then Debug.Print "Not found; project or story "
        Exit Sub
        
    End If
    Dim filterstr, filter_status, filter_tracker  As String
    filterstr = ""
    filter_status = ""
    filter_tracker = ""
    Set Var = New Dictionary
    Set tmpstatusdic = New Dictionary
    Set tmptracksdic = New Dictionary
    If debug_ Then Debug.Print JSONLib.toString(LocalSavedSettings("ListBox_Setting_Status_child"))
    
    If LocalSavedSettings.exists("ListBox_Setting_Status_child") And (Not IsEmpty(LocalSavedSettings("ListBox_Setting_Status_child"))) Then
        Set tmpstatusdic = LocalSavedSettings("ListBox_Setting_Status_child")
        For Each Var In tmpstatusdic
            If filter_status = "" Then
                filter_status = "status_id=" & tmpstatusdic(Var)
            Else
                filter_status = filter_status & "|" & tmpstatusdic(Var)
            End If
        Next Var
        
        If filterstr = "" Then
            filterstr = filter_status
        Else
            filterstr = filterstr & "&" & filter_status
        End If
    End If
    If LocalSavedSettings.exists("ListBox_Setting_Tracker_child") Then
        Set tmptracksdic = LocalSavedSettings("ListBox_Setting_Tracker_child")
        For Each Var In tmptracksdic
            If filter_tracker = "" Then
                filter_tracker = "tracker_id=" & tmptracksdic(Var)
            Else
                filter_tracker = filter_tracker & "|" & tmptracksdic(Var)
            End If
        Next Var
        If filterstr = "" Then
            filterstr = filter_tracker
        Else
            filterstr = filterstr & "&" & filter_tracker
        End If
    End If
    Set Dic_Users = Nothing
    Set Dic_Users = New Dictionary

    Dim jsonstring As String
    jsonstring = GetData(url & "/issues.json?key=" & apikey & "&project_id=" & myProject & "&parent_id=" & myStory & "&" & filterstr)
    Set json = New Dictionary
    Set json = JSONLib.parse(jsonstring)
    If json Is Nothing Then
        MsgBox "Cant load rm."
        Exit Sub
    End If
    total = json("total_count")
    offset = json("offset")
    limit = json("limit")
    nextoffset = val(limit) + val(offset)

    If debug_ Then Debug.Print "limit " & limit & " / offset " & offset & " / total " & total
    Do While total > nextoffset

        Dim subjsonstr As String
        subjsonstr = GetData(url & "/issues.json?key=" & apikey & "&project_id=" & myProject & "&parent_id=" & myStory & "&offset=" & nextoffset & "&" & filterstr)

        Dim jsonsub As Object
        Set jsonsub = New Dictionary
        Set jsonsub = JSONLib.parse(subjsonstr)
        total = jsonsub("total_count")
        offset = jsonsub("offset")
        limit = jsonsub("limit")
        nextoffset = val(limit) + val(offset)
        If debug_ Then Debug.Print "limit " & limit & " / offset " & offset & " / total " & total
        For Each Var In jsonsub("issues")
            json("issues").Add Var
        Next Var
    Loop
    For Each Var In json("issues")
            If debug_ Then Debug.Print Var("id") & " : localfilter is enable"
            If Var.exists("parent") = True Then
                If debug_ Then Debug.Print "this ticket have parents"
                If Var("project")("id") = myProject And Var("parent")("id") = myStory Then
                    If tmpstatusdic(Var("status")("name")) <> "" And tmptracksdic(Var("tracker")("name")) <> "" Then
                        ComboBox_parentBacklog.AddItem "#" & Var("id") & ":" & Var("subject")
                        Dic_Backlog.Add "#" & Var("id") & ":" & Var("subject"), Var("id")
                    Else
                        If debug_ Then Debug.Print Var("id") & " : its not my children "
                    End If
                Else
                    If debug_ Then Debug.Print Var("id") & " : project_id  " & Var("project")("id") & " <>  " & myProject & ", story_id " & Var("parent")("id") & " <> " & myStory
                End If
            End If
    Next Var
    Set json = Nothing
    Set JSONLib = Nothing
If debug_ Then Debug.Print "★end★ set_backlog_ticket_for_selected_activity"
End Sub

Public Sub get_ticket_subject(ByRef ticketnumber As Integer, ByVal url As String, ByVal apikey As String)
    Dim JSONLib As New JSONLib
    Dim json, tmpdic As Object
    Dim Var As Variant
    Dim total, offset, limit, nextoffset As Integer
    If debug_ Then Debug.Print "★start★Calle :: get_ticket_subject"
    Dim jsonstring As String
    jsonstring = GetData(url & "/issues/" & ticketnumber & ".json?key=" & apikey)
    Set json = New Dictionary
    Set json = JSONLib.parse(jsonstring)
    If json Is Nothing Then
        MsgBox "not found ticket."
        Call CommandButton2_Click
        Exit Sub
    End If
    Set Var = json("issue")
    selected_ticket_id = Var("id")
    Label_selected_ticket.Caption = "#" & Var("id") & ":" & Var("subject")
    Label_selected_ticket.ControlTipText = "#" & Var("id") & ":" & Var("subject")
    TextBox_Comment.Text = Var("subject")
    Set json = Nothing
    Set JSONLib = Nothing
If debug_ Then Debug.Print "★end★ get_ticket_subject"
End Sub

Private Sub check_my_timeentry_on_today(ByVal url As String, ByVal apikey As String)
    Dim JSONLib As New JSONLib
    Dim json, tmpdic As Object
    Dim Var As Variant
    Dim total, offset, limit, nextoffset As Integer
    If debug_ Then Debug.Print "★start★Calle :: check_my_timeentry_on_today"

    Dim jsonstring As String
    ListBox_mytimeentry.Clear
    jsonstring = GetData(url & "time_entries.json?user_id=me&spent_on=" & Format(GetToday(), "yyyy-mm-dd") & "&key=" & apikey)
    Set json = New Dictionary
    Set json = JSONLib.parse(jsonstring)
    If json Is Nothing Then
        Exit Sub
    End If
    total = json("total_count")
    offset = json("offset")
    limit = json("limit")
    nextoffset = val(limit) + val(offset)

    If debug_ Then Debug.Print "limit " & limit & " / offset " & offset & " / total " & total
    Do While total > nextoffset

        Dim subjsonstr As String
        subjsonstr = GetData(url & "/time_entries.json?user_id=me&spent_on=" & Format(GetToday(), "yyyy-mm-dd") & "?key=" & apikey & "&offset=" & nextoffset)
        Dim jsonsub As Object
        Set jsonsub = New Dictionary
        Set jsonsub = JSONLib.parse(subjsonstr)
        total = jsonsub("total_count")
        offset = jsonsub("offset")
        limit = jsonsub("limit")
        nextoffset = val(limit) + val(offset)
        If debug_ Then Debug.Print "limit " & limit & " / offset " & offset & " / total " & total
        For Each Var In jsonsub("issues")
            json("issues").Add Var
        Next Var
    Loop
    Dim totaltimeentry As Single
    Dim myname As String
    totaltimeentry = 0
    Dim listline As Integer
    listline = 0
    For Each Var In json("time_entries")
        myname = Var("user")("name")
        totaltimeentry = totaltimeentry + CSng(Var("hours"))
        If debug_ Then Debug.Print Var("issue")("id") & "/" & Var("activity")("name") & "/" & Var("hours") & "/" & Var("comments") & "++=" & totaltimeentry
        ListBox_mytimeentry.AddItem ""
        ListBox_mytimeentry.List(listline, 0) = Var("issue")("id")
        ListBox_mytimeentry.List(listline, 1) = Var("activity")("name")
        ListBox_mytimeentry.List(listline, 2) = Var("hours")
        ListBox_mytimeentry.List(listline, 3) = Var("comments")
        listline = listline + 1
    Next Var
    Set json = Nothing
    Set JSONLib = Nothing
    Label_spentontoday_hours.Caption = "[" & totaltimeentry & "] hours spent on today"
If debug_ Then Debug.Print "★end★ check_my_timeentry_on_today"

End Sub
Private Function favorite_initialize(ByRef ticketid As String)
    If ticketid = "" Then
        Exit Function
    End If
    CommandButton_setfavorite.Caption = "☆"
    If TransactionTimeEntryData.exists("favoritelist") Then
      If TransactionTimeEntryData("favoritelist").exists(ticketid) Then
        If TransactionTimeEntryData("favoritelist")(ticketid) <> "" Then
            CommandButton_setfavorite.Caption = "★"
        Else
            CommandButton_setfavorite.Caption = "☆"
        End If
      End If
    End If
End Function
Public Sub set_activity_ticket_for_assigned_id_to_me(ByRef LocalSavedSettings As Object, ByVal url As String, ByVal apikey As String)
If debug_ Then Debug.Print "★start★ set_activity_ticket_for_assigned_id_to_me"
    Dim JSONLib As New JSONLib
    Dim json, tmpdic As Object
    Dim Var As Variant
    Dim total, offset, limit, nextoffset As Integer
    If debug_ Then Debug.Print "Calle :: set_backlog_ticket_for_selected_activity"
    Set Dic_Assigned_To_Me = New Dictionary
    ComboBox_assigned_me.Clear
    
    Dim filterstr, filter_status, filter_tracker  As String
    filterstr = ""
    filter_tracker = ""
    Set Var = New Dictionary
    Set tmpstatusdic = New Dictionary
    Set tmptracksdic = New Dictionary
    If debug_ Then Debug.Print JSONLib.toString(LocalSavedSettings("ListBox_Setting_Status_child"))
    
    If LocalSavedSettings.exists("ListBox_Setting_Tracker_child") Then
        Set tmptracksdic = LocalSavedSettings("ListBox_Setting_Tracker_child")
        For Each Var In tmptracksdic
            If filter_tracker = "" Then
                filter_tracker = "tracker_id=" & tmptracksdic(Var)
            Else
                filter_tracker = filter_tracker & "|" & tmptracksdic(Var)
            End If
        Next Var
        If filterstr = "" Then
            filterstr = filter_tracker
        Else
            filterstr = filterstr & "&" & filter_tracker
        End If
    End If
    Set Dic_Users = Nothing
    Set Dic_Users = New Dictionary

    Dim jsonstring As String
    jsonstring = GetData(url & "/issues.json?key=" & apikey & "&assigned_to_id=me&status=open&" & filterstr)
    Set json = New Dictionary
    Set json = JSONLib.parse(jsonstring)
    If json Is Nothing Then
        MsgBox "Cant load rm."
        Exit Sub
    End If
    total = json("total_count")
    offset = json("offset")
    limit = json("limit")
    nextoffset = val(limit) + val(offset)

    If debug_ Then Debug.Print "limit " & limit & " / offset " & offset & " / total " & total
    Do While total > nextoffset

        Dim subjsonstr As String
        subjsonstr = GetData(url & "/issues.json?key=" & apikey & "&assigned_to_id=me&status=open&offset=" & nextoffset & "&" & filterstr)

        Dim jsonsub As Object
        Set jsonsub = New Dictionary
        Set jsonsub = JSONLib.parse(subjsonstr)
        total = jsonsub("total_count")
        offset = jsonsub("offset")
        limit = jsonsub("limit")
        nextoffset = val(limit) + val(offset)
        If debug_ Then Debug.Print "limit " & limit & " / offset " & offset & " / total " & total

        For Each Var In jsonsub("issues")
            json("issues").Add Var
        Next Var

    Loop
    For Each Var In json("issues")
       ComboBox_assigned_me.AddItem "#" & Var("id") & ":" & Var("subject")
       Dic_Assigned_To_Me.Add "#" & Var("id") & ":" & Var("subject"), Var("id")
    Next Var
    Set json = Nothing
    Set JSONLib = Nothing
If debug_ Then Debug.Print "★end★ set_activity_ticket_for_assigned_id_to_me"
End Sub
