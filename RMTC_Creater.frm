VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RMTC_Creater 
   Caption         =   "Redmine Ticket Creater"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   OleObjectBlob   =   "RMTC_Creater.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "RMTC_Creater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Microsoft Scripting Runtime
Option Explicit
Private RMTM_Con_parentid As String

Private Sub ComboBox_parentActivity_Change()
    If debug_ Then Debug.Print "ComboBox_ParentActivity_Change :: select change : Activity = " & ComboBox_parentActivity.value
    If Dic_Activity.exists(ComboBox_parentActivity.value) Then
        If debug_ Then Debug.Print "ComboBox_ParentStory_Change :: select change : Activity id = " & Dic_Activity(ComboBox_parentActivity.value)
        LabelLabel_GotWeb_parent.Caption = Dic_Activity(ComboBox_parentActivity.value)
        LabelLabel_GotWeb_parent.ForeColor = &H8000000D
    Else
        LabelLabel_GotWeb_parent.Caption = "New"
        LabelLabel_GotWeb_parent.ForeColor = &H80000007
    End If
        
End Sub

Private Sub ComboBox_Project_Change()
    If debug_ Then Debug.Print "ComboBox_Project_Change :: " & ComboBox_Project.value
    Dic_Story.RemoveAll
    Dic_Activity.RemoveAll
    ComboBox_ParentStory.Clear
    ComboBox_parentActivity.Clear
    Call set_story_ticket_for_selected_project(LocalSavedSettings, Setting_Redmine_URL, Setting_Redmine_APIKEY)
    If debug_ Then Debug.Print "ComboBox_Project_Change :: ended"
End Sub

Private Sub ComboBox_Status_Change()
    If debug_ Then Debug.Print "ComboBox_status_Change :: " & ComboBox_Status.value
    If debug_ Then Debug.Print "ComboBox_status_Change :: ended"
End Sub
Private Sub ComboBox_Asignedto_Change()
    If debug_ Then Debug.Print "ComboBox_Asignedto_Change :: " & ComboBox_Asignedto.value
    If debug_ Then Debug.Print "ComboBox_Asignedto_Change :: ended"
End Sub
Private Sub ComboBox_Priority_Change()
    If debug_ Then Debug.Print "ComboBox_Priority_Change :: " & ComboBox_Priority.value
    If debug_ Then Debug.Print "ComboBox_Priority_Change :: ended"
End Sub

Private Sub ComboBox_ParentStory_Change()
    ComboBox_parentActivity.Clear
    If debug_ Then Debug.Print "ComboBox_ParentStory_Change :: select change : Story = " & ComboBox_ParentStory.value
    If Dic_Story.exists(ComboBox_ParentStory.value) Then
        If debug_ Then Debug.Print "ComboBox_ParentStory_Change :: select change : Story id = " & Dic_Story(ComboBox_ParentStory.value)
        Label_GotoWeb_grapa.Caption = Dic_Story(ComboBox_ParentStory.value)
        Label_GotoWeb_grapa.ForeColor = &H8000000D
        Call set_activity_ticket_for_selected_project(LocalSavedSettings, Setting_Redmine_URL, Setting_Redmine_APIKEY)
    Else
        Label_GotoWeb_grapa.Caption = "New"
        Label_GotoWeb_grapa.ForeColor = &H80000007
    End If
    
End Sub

Private Sub CommandButton_ClearActivity_Click()
    ComboBox_parentActivity = ""
End Sub

Private Sub CommandButton_clearStory_Click()
    ComboBox_ParentStory = ""
End Sub

Private Sub CommandButton_cleartext_Click()
    TextBox_Contetns.Text = ""
    TextBox_Subject.Text = ""
End Sub

Private Sub CommandButton_Edit2_Click()
Dim tmpstr As String
tmpstr = TextBox_Contetns.value

    TextBox_Contetns.value = _
        vbCrLf & _
        "{{collapse(contents)" & _
        vbCrLf & _
        tmpstr & _
        vbCrLf & _
        "}}"
    RMTC_Creater.TextBox_Contetns.SelStart = 0
    RMTC_Creater.TextBox_Contetns.SetFocus
    RMTC_Creater.TextBox_Subject.SelStart = 0
End Sub

Private Sub CommandButton_submit_Click()
    If debug_ Then Debug.Print "CommandButton_submit_Click Called"
    Dim JSONLib As New JSONLib
    Dim tmpstoryid, tmpactivityid, tmpbacklogid As String
    If LocalSavedSettings.exists("ListBox_New_Ticket_Story") Then
        tmpstoryid = LocalSavedSettings("ListBox_New_Ticket_Story")
    Else
        MsgBox "not set default categ tracker"
        Exit Sub
    End If
    If LocalSavedSettings.exists("ListBox_New_Ticket_Activity") Then
        tmpactivityid = LocalSavedSettings("ListBox_New_Ticket_Activity")
    Else
        MsgBox "not set default subcateg tracker"
        Exit Sub
    End If
    If LocalSavedSettings.exists("ListBox_New_Ticket_Backlog") Then
        tmpbacklogid = LocalSavedSettings("ListBox_New_Ticket_Backlog")
    Else
        MsgBox "not set default subsubcateg. tracker"
        Exit Sub
    End If
    If debug_ Then Debug.Print "activity dic is " & JSONLib.toString(Dic_Activity) & " check " & Dic_Activity.exists(ComboBox_parentActivity.Text)
    If debug_ Then Debug.Print "activity dic value " & Dic_Activity(ComboBox_parentActivity.Text) & " check " & Dic_Activity(ComboBox_parentActivity.Text)
    Dim parentid As String
    Dim msgreturn As Integer
        If Dic_Story(ComboBox_ParentStory.Text) = "" Then
            parentid = postredmineJson(tmpstoryid, "")
        Else
            parentid = Dic_Story(ComboBox_ParentStory.Text)
        End If
        If parentid < 0 Then
           Exit Sub
        End If
        If Dic_Activity(ComboBox_parentActivity.Text) = "" Then
            parentid = postredmineJson(tmpactivityid, parentid)
        Else
            parentid = Dic_Activity(ComboBox_parentActivity.Text)
        End If
        If parentid < 0 Then
           Exit Sub
        End If
        parentid = postredmineJson(tmpbacklogid, parentid)
        msgreturn = MsgBox("#" & parentid & " is created. open web?", vbYesNo)
            If debug_ Then Debug.Print "user choice is : " & msgreturn
            If msgreturn = vbNo Then
            ElseIf msgreturn = vbYes Then
                If debug_ Then Debug.Print "created backlog open web start : " & Setting_Redmine_URL & "/issue/" & parentid & "?key=" & Setting_Redmine_APIKEY
                
                If webincreasemyAPIKey = 1 Then
                    openweb (Setting_Redmine_URL & "/issues/" & parentid & "?key=" & Setting_Redmine_APIKEY)
                Else
                    openweb (Setting_Redmine_URL & "/issues/" & parentid & "")
                End If
            End If
        If parentid <> "" Then
            CommandButton_toTimeentry.Enabled = True
            RMTM_Con_parentid = parentid
        End If
        Exit Sub
End Sub

Private Sub CommandButton_toTimeentry_Click()
    Call RMTM_Creater.rmtm_initializer
    Call RMTM_Creater.set_select_ticket_id(RMTM_Con_parentid, TextBox_Subject.Text)
    Unload Me
    RMTM_Creater.Show
    Call rmtc_initializer
End Sub

Private Sub CommandButton2_Click()
    Unload Me
    Redmint_CreateTimeEntry
End Sub

Private Sub Label_GotoWeb_grapa_Click()
    If debug_ Then Debug.Print "story open web start : " & Setting_Redmine_URL & "/issue/" & Dic_Story(ComboBox_ParentStory) & "?key=" & Setting_Redmine_APIKEY
    If Dic_Story(ComboBox_ParentStory.Text) = "" Then
        Exit Sub
    End If
    If webincreasemyAPIKey = 1 Then
        openweb (Setting_Redmine_URL & "/issues/" & Dic_Story(ComboBox_ParentStory.Text) & "?key=" & Setting_Redmine_APIKEY)
    Else
        openweb (Setting_Redmine_URL & "/issues/" & Dic_Story(ComboBox_ParentStory.Text))
    End If
End Sub

Private Sub LabelLabel_GotWeb_parent_Click()
    If debug_ Then Debug.Print "activity open web start : " & Setting_Redmine_URL & "/issue/" & Dic_Activity(ComboBox_parentActivity) & "?key=" & Setting_Redmine_APIKEY
    If Dic_Activity(ComboBox_parentActivity.Text) = "" Then
        Exit Sub
    End If
    If webincreasemyAPIKey = 1 Then
        openweb (Setting_Redmine_URL & "/issues/" & Dic_Activity(ComboBox_parentActivity.Text) & "?key=" & Setting_Redmine_APIKEY)
    Else
        openweb (Setting_Redmine_URL & "/issues/" & Dic_Activity(ComboBox_parentActivity.Text))
    End If

End Sub

Private Sub TextBox_Contetns_Change()
    Label_MaxLength_count.Caption = TextBox_Contetns.TextLength
    If TextBox_Contetns.TextLength > 6000 Then
        CommandButton_submit.Enabled = False
    Else
        CommandButton_submit.Enabled = True
    End If
End Sub
Private Sub Button_Settings_Click()
    Call save_transaction_Data_to_reg
    RMTC_Setting.Show
    If Initialized = 1 Then
        Call rmtc_initializer
        Call RMTS_Search.rmts_initialize
    Else
        Unload Me
    End If
End Sub
Private Sub CommandButton_DueCaleder_Click()
   Call CalenderForm.setDate(GetToday())
   Call CalenderForm.setCallBackControl(Label_DueDate)
   CalenderForm.Show
End Sub
Private Sub CommandButton_StartCaleder_Click()
   Call CalenderForm.setDate(GetToday())
   Call CalenderForm.setCallBackControl(Label_StartDate)
   CalenderForm.Show
End Sub
Private Function GetToday()
    GetToday = Year(Now) & "/" & Month(Now) & "/" & Day(Now)
End Function

Private Sub save_transaction_Data_to_reg()
    Dim JSONLib As New JSONLib
    If debug_ Then Debug.Print "Save transaction data to reg."
    If TransactionData Is Nothing Then
        Set TransactionData = New Dictionary
    End If
    If debug_ Then Debug.Print "Save transaction data to reg ComboBox_Project.value = " & ComboBox_Project.value
    If TransactionData.exists("Default_Project") Then
        TransactionData("Default_Project") = ComboBox_Project.value
    Else
        TransactionData.Add "Default_Project", ComboBox_Project.value
    End If
    If debug_ Then Debug.Print "Save transaction data to reg ComboBox_Status.value = " & ComboBox_Status.value
    If TransactionData.exists("Default_Status") Then
        TransactionData("Default_Status") = ComboBox_Status.value
    Else
        TransactionData.Add "Default_Status", ComboBox_Status.value
    End If
    If debug_ Then Debug.Print "Save transaction data to reg ComboBox_Asignedto.value = " & ComboBox_Asignedto.value
    If TransactionData.exists("Default_Asignedto") Then
        TransactionData("Default_Asignedto") = ComboBox_Asignedto.value
    Else
        TransactionData.Add "Default_Asignedto", ComboBox_Asignedto.value
    End If
    If debug_ Then Debug.Print "Save transaction data to reg ComboBox_Priority.value = " & ComboBox_Priority.value
    If TransactionData.exists("Default_Priority") Then
        TransactionData("Default_Priority") = ComboBox_Priority.value
    Else
        TransactionData.Add "Default_Priority", ComboBox_Priority.value
    End If
    If debug_ Then Debug.Print "save_transaction_Data_to_reg :: transaction data " & JSONLib.toString(TransactionData)
    SaveSetting "OutlookRMTC", "Transaction", "TransactionData", JSONLib.toString(TransactionData)
End Sub

Private Sub UserForm_Activate()
        If Setting_Redmine_APIKEY = "" Or Setting_Redmine_URL = "" Then
            Button_Settings_Click
        End If
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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Initialized = 1 Then
        Dim JSONLib As New JSONLib
        Dim json As Object
        Dim RegStr As String
        RegStr = JSONLib.toString(TransactionData)
        If debug_ Then Debug.Print "UserForm_QueryClose :: save to reg : "; RegStr
        Call save_transaction_Data_to_reg
        If debug_ Then Debug.Print "UserForm_QueryClose ended"
    Else
        If debug_ Then Debug.Print "UserForm_QueryClose called but not initialized or going reload"
        Exit Sub
    End If
End Sub

Public Sub set_story_ticket_for_selected_project(ByRef LocalSavedSettings As Object, ByVal url As String, ByVal apikey As String)
    Dim JSONLib As New JSONLib
    Dim json, tmpdic As Object
    Dim Var As Variant
    Dim total, offset, limit, nextoffset As Integer
    If debug_ Then Debug.Print "Calle :: set_story_ticket_for_selected_project"
    Set Dic_Story = New Dictionary
    Dim myProject As String
    If LocalSavedSettings.exists("Dic_Projects_ID") And LocalSavedSettings("Dic_Projects_ID").exists(RMTC_Creater.ComboBox_Project.value) Then
        myProject = LocalSavedSettings("Dic_Projects_ID")(RMTC_Creater.ComboBox_Project.value)
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
    Set tmpdic = New Dictionary
    If debug_ Then Debug.Print JSONLib.toString(LocalSavedSettings("ListBox_Setting_Status_granpa"))
    If LocalSavedSettings.exists("ListBox_Setting_Status_granpa") Then
        Set tmpdic = LocalSavedSettings("ListBox_Setting_Status_granpa")
        For Each Var In tmpdic
            If filter_status = "" Then
                filter_status = "status_id=" & tmpdic(Var)
            Else
                filter_status = filter_status & "|" & tmpdic(Var)
            End If
        Next Var
        If filterstr = "" Then
            filterstr = filter_status
        Else
            filterstr = filterstr & "&" & filter_status
        End If
    End If
    If LocalSavedSettings.exists("ListBox_Setting_Tracker_grandparent") Then
        Set tmpdic = LocalSavedSettings("ListBox_Setting_Tracker_grandparent")
        For Each Var In tmpdic
            If filter_tracker = "" Then
                filter_tracker = "tracker_id=" & tmpdic(Var)
            Else
                filter_tracker = filter_tracker & "|" & tmpdic(Var)
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
        If Var("project")("id") = myProject Then
            ComboBox_ParentStory.AddItem "#" & Var("id") & ":" & Var("subject")
            Dic_Story.Add "#" & Var("id") & ":" & Var("subject"), Var("id")
        Else
            If debug_ Then Debug.Print "this ticket " & Var("id") & " is not match which project_id is not " & myProject
        End If
    Next Var
    Set json = Nothing
    Set JSONLib = Nothing
End Sub
Public Sub set_activity_ticket_for_selected_project(ByRef LocalSavedSettings As Object, ByVal url As String, ByVal apikey As String)
    Dim JSONLib As New JSONLib
    Dim json, tmpdic As Object
    Dim Var As Variant
    Dim total, offset, limit, nextoffset As Integer
    If debug_ Then Debug.Print "Calle :: set_activity_ticket_for_selected_project"
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
    Set tmpdic = New Dictionary
    If debug_ Then Debug.Print JSONLib.toString(LocalSavedSettings("ListBox_Setting_Status_granpa"))
    If LocalSavedSettings.exists("ListBox_Setting_Status_parents") Then
        Set tmpdic = LocalSavedSettings("ListBox_Setting_Status_parents")
        For Each Var In tmpdic
            If filter_status = "" Then
                filter_status = "status_id=" & tmpdic(Var)
            Else
                filter_status = filter_status & "|" & tmpdic(Var)
            End If
        Next Var
        If filterstr = "" Then
            filterstr = filter_status
        Else
            filterstr = filterstr & "&" & filter_status
        End If
    End If
    If LocalSavedSettings.exists("ListBox_Setting_Tracker_parents") Then
        Set tmpdic = LocalSavedSettings("ListBox_Setting_Tracker_parents")
        For Each Var In tmpdic
            If filter_tracker = "" Then
                filter_tracker = "tracker_id=" & tmpdic(Var)
            Else
                filter_tracker = filter_tracker & "|" & tmpdic(Var)
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
        If Var.exists("parent") Then
            If Var("project")("id") = myProject And Var("parent")("id") = myStory Then
                ComboBox_parentActivity.AddItem "#" & Var("id") & ":" & Var("subject")
                Dic_Activity.Add "#" & Var("id") & ":" & Var("subject"), Var("id")
            Else
                If debug_ Then Debug.Print "this ticket " & Var("id") & " is not match which project_id <>  " & myProject & " story_id <>  " & myStory
            End If
        End If
    Next Var
    Set json = Nothing
    Set JSONLib = Nothing
End Sub
Public Function postredmineJson(ByVal tmyracker As String, ByVal parentid As String)
    Dim Bodystr
    Dim SubjStr
    Dim xhr
    Dim RequestURL As String
    Dim RequestBody As String
    Dim bPmary() As Byte
    Dim JSONLib As New JSONLib
    Dim json As Object
    RequestURL = Setting_Redmine_URL & "/issues.json?format=xml&key=" & Setting_Redmine_APIKEY
    If debug_ Then Debug.Print "RequestURL is " & RequestURL
    Set xhr = CreateObject("Microsoft.XMLHTTP")
    xhr.Open "POST", RequestURL, False
    xhr.SetRequestHeader "Content-Type", "text/xml"
    RequestBody = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>"
    RequestBody = RequestBody & "<issue>"
    If debug_ Then Debug.Print "post data :: project_id is " & LocalSavedSettings("Dic_Projects")(ComboBox_Project.Text)
    RequestBody = RequestBody & "<project_id>" & LocalSavedSettings("Dic_Projects")(ComboBox_Project.Text) & "</project_id>"
    If debug_ Then Debug.Print "post data :: subject is " & TextBox_Subject.Text
    RequestBody = RequestBody & "<subject>" & ConvertString(TextBox_Subject.Text) & "</subject>"
    If LocalSavedSettings("ListBox_User_Settings")(ComboBox_Asignedto.Text) <> "" Then
        If debug_ Then Debug.Print "post data :: ComboBox_Asignedto is " & LocalSavedSettings("ListBox_User_Settings")(ComboBox_Asignedto.Text)
        RequestBody = RequestBody & "<assigned_to_id>" & LocalSavedSettings("ListBox_User_Settings")(ComboBox_Asignedto.Text) & "</assigned_to_id>"
    End If
    If ComboBox_Status.value <> "" Then
        If debug_ Then Debug.Print "post data :: ComboBox_Status is " & ComboBox_Status.Text
        RequestBody = RequestBody & "<status_id>" & LocalSavedSettings("ListBox_setting_Status")(ComboBox_Status.Text) & "</status_id>"
    End If
    If ComboBox_Estimated.value <> "" Then
        If debug_ Then Debug.Print "post data :: ComboBox_Estimated is " & ComboBox_Estimated.Text
        RequestBody = RequestBody & "<estimated_hours>" & ComboBox_Estimated.Text & "</estimated_hours>"
    End If
    If Label_StartDate.Caption <> "" Then
        If debug_ Then Debug.Print "post data :: Label_StartDate is " & Format(Label_StartDate.Caption, "yyyy-mm-dd")
        RequestBody = RequestBody & "<start_date>" & Format(Label_StartDate.Caption, "yyyy-mm-dd") & "</start_date>"
    End If
    If Label_DueDate.Caption <> "" Then
        If debug_ Then Debug.Print "post data :: Label_DueDate is " & Format(Label_DueDate.Caption, "yyyy-mm-dd")
        RequestBody = RequestBody & "<start_date>" & Format(Label_DueDate.Caption, "yyyy-mm-dd") & "</start_date>"
    End If
    If parentid = "" Then
    Else
        If debug_ Then Debug.Print "post data :: parentid is " & parentid
        RequestBody = RequestBody & "<parent_issue_id>" & parentid & "</parent_issue_id>"
    End If
    If debug_ Then Debug.Print "post data :: tmyracker is " & tmyracker
    RequestBody = RequestBody & "<tracker_id>" & tmyracker & "</tracker_id>"
    RequestBody = RequestBody & "<description>" & ConvertString(TextBox_Contetns.Text) & "</description>"
    RequestBody = RequestBody & "</issue>"
    xhr.Send (RequestBody)
    If xhr.Status = 201 Then
        Set json = JSONLib.parse(xhr.responseText)
        postredmineJson = json("issue")("id")
    Else
        Set json = JSONLib.parse(xhr.responseText)
        MsgBox json("errors").Item(1)
        If debug_ Then Debug.Print "error occured " & json("errors").Item(1)
        postredmineJson = -1
    End If

End Function

Public Sub rmtc_initializer()
If debug_ Then Debug.Print "rmtc_initializer Called"
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
    RMTC_Creater.ComboBox_Project.Clear
    RMTC_Creater.ComboBox_Asignedto.Clear
    RMTC_Creater.ComboBox_Priority.Clear
    RMTC_Creater.ComboBox_Status.Clear
    Set Dic_Story = New Dictionary
    If LocalSavedSettings.exists("Dic_Projects") Then
        Set tmpdic = LocalSavedSettings("Dic_Projects")
        For Each Var In tmpdic
            RMTC_Creater.ComboBox_Project.AddItem Var
        Next Var
    End If
    RMTC_Creater.ComboBox_Asignedto.AddItem ""
    If LocalSavedSettings.exists("ListBox_User_Settings") Then
        Set tmpdic = LocalSavedSettings("ListBox_User_Settings")
        For Each Var In tmpdic
                RMTC_Creater.ComboBox_Asignedto.AddItem Var
        Next Var
    End If
    If LocalSavedSettings.exists("ListBox_setting_priority") Then
        Set tmpdic = LocalSavedSettings("ListBox_setting_priority")
        For Each Var In tmpdic
                RMTC_Creater.ComboBox_Priority.AddItem Var
        Next Var
    End If
    If LocalSavedSettings.exists("ListBox_Setting_Status_parents") Then
        Set tmpdic = LocalSavedSettings("ListBox_Setting_Status_parents")
        For Each Var In tmpdic
                RMTC_Creater.ComboBox_Status.AddItem Var
        Next Var
    End If
    RMTC_Creater.ComboBox_Estimated.AddItem 1
    RMTC_Creater.ComboBox_Estimated.AddItem 2
    RMTC_Creater.ComboBox_Estimated.AddItem 3
    RMTC_Creater.ComboBox_Estimated.AddItem 5
    RMTC_Creater.ComboBox_Estimated.AddItem 8
    RMTC_Creater.ComboBox_Estimated.AddItem 13
    RMTC_Creater.ComboBox_Estimated.AddItem 21
    RMTC_Creater.ComboBox_Estimated.AddItem 34
    RMTC_Creater.ComboBox_Estimated.AddItem 55
    RMTC_Creater.ComboBox_Estimated.AddItem 89
    RMTC_Creater.ComboBox_Estimated.value = RMTC_Creater.ComboBox_Estimated.List(0)
    RegStr = GetSetting("OutlookRMTC", "Transaction", "TransactionData")
    If debug_ Then Debug.Print "UserForm_Initialize :: get regset : TransactionData = " & RegStr
    Set TransactionData = JSONLib.parse(RegStr)
    Debug.Assert Err.Number = 0
    If Not TransactionData Is Nothing Then
        If debug_ Then Debug.Print "UserForm_Initialize :: transaction data " & JSONLib.toString(TransactionData)
        If debug_ Then Debug.Print "apply transaction data to form"
        If TransactionData.exists("Default_Project") Then
            RMTC_Creater.ComboBox_Project.value = TransactionData("Default_Project")
        End If
        If TransactionData.exists("Default_Asignedto") Then
            RMTC_Creater.ComboBox_Asignedto.value = TransactionData("Default_Asignedto")
        End If
        If TransactionData.exists("Default_Priority") Then
            RMTC_Creater.ComboBox_Priority.value = TransactionData("Default_Priority")
        End If
        If TransactionData.exists("Default_Status") Then
            RMTC_Creater.ComboBox_Status.value = TransactionData("Default_Status")
        End If
    Else
        If debug_ Then Debug.Print "UserForm_Initialize :: transaction data is nothing "
    End If
    CommandButton_toTimeentry.Enabled = False
    RMTC_Creater.TextBox_Contetns.SelStart = 0
    RMTC_Creater.TextBox_Contetns.SetFocus
    RMTC_Creater.TextBox_Subject.SelStart = 0
End Sub

