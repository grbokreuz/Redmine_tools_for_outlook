VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RMTS_Search 
   Caption         =   "RMTS_Search"
   ClientHeight    =   5325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8475.001
   OleObjectBlob   =   "RMTS_Search.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "RMTS_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' for redmine api under 3.3 , upper virsion is not support search api
Public project As String
Public Function set_param(ByVal inp_id As String, ByVal inp_name As String)
    project = inp_id
    ComboBox_Project.value = inp_name
End Function
Public Sub get_ticket_for_keyword_categ(ByRef project As String, ByRef keyword As String, ByVal url As String, ByVal apikey As String)
    Dim JSONLib As New JSONLib
    Dim json, tmpdic As Object
    Dim Var As Variant
    Dim total, offset, limit, nextoffset As Integer
    If debug_ Then Debug.Print "Calle :: get_ticket_for_keyword_categ"
    Set Dic_Story = New Dictionary

    Dim myProject As String
    myProject = project
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
    Dim subjson As Integer
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
    subjson = 0
    If debug_ Then Debug.Print "limit " & limit & " / offset " & offset & " / total " & total
    Do While total > nextoffset
        subjson = 1
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
    Loop
    If subjson = 1 Then
        For Each Var In jsonsub("issues")
            json("issues").Add Var
        Next Var
    End If
    listline = ListBox_TicketList.ListCount
    For Each Var In json("issues")
        Dim tmpsubject As String
        Dim tmpdescrip As String
        tmpsubject = Var("subject")
        tmpdescrip = Var("description")
        
        If searchContents = 1 Then
            If (Var("project")("id") = myProject And ((reg_hit(tmpsubject, TextBox_SearchKey.Text) > 0) Or (reg_hit(tmpdescrip, TextBox_SearchKey.Text) > 0))) Then
                If debug_ Then Debug.Print "this ticket " & Var("id") & " is hit"
                ListBox_TicketList.AddItem ""
                ListBox_TicketList.List(listline, 0) = Var("id")
                ListBox_TicketList.List(listline, 1) = Var("tracker")("name")
                ListBox_TicketList.List(listline, 2) = Var("subject")
                ListBox_TicketList.List(listline, 3) = Mid(Var("description"), 1, 50)
                listline = listline + 1
            Else
                If debug_ Then Debug.Print "this ticket " & Var("id") & " is not match which project_id is " & Var("project")("id") & " not " & myProject & " and subject is " & tmpsubject & " not " & TextBox_SearchKey.Text
            End If
        ElseIf (Var("project")("id") = myProject And (reg_hit(tmpsubject, TextBox_SearchKey.Text) > 0)) Then
        
            If debug_ Then Debug.Print "this ticket " & Var("id") & " is hit"
            ListBox_TicketList.AddItem ""
            ListBox_TicketList.List(listline, 0) = Var("id")
            ListBox_TicketList.List(listline, 1) = Var("tracker")("name")
            ListBox_TicketList.List(listline, 2) = Var("subject")
            ListBox_TicketList.List(listline, 3) = Mid(Var("description"), 1, 50)
            listline = listline + 1
        Else
            If debug_ Then Debug.Print "this ticket " & Var("id") & " is not match which project_id is " & Var("project")("id") & " not " & myProject & " and subject is " & tmpsubject & " not " & TextBox_SearchKey.Text
        End If
        tmpsubject = ""
        tmpdescrip = ""
    Next Var
    Set json = Nothing
    Set JSONLib = Nothing
End Sub
Public Sub get_ticket_for_keyword_subcat(ByRef project As String, ByRef keyword As String, ByVal url As String, ByVal apikey As String)
    Dim JSONLib As New JSONLib
    Dim json, tmpdic As Object
    Dim Var As Variant
    Dim total, offset, limit, nextoffset As Integer
    If debug_ Then Debug.Print "Calle :: get_ticket_for_keyword_subcat"
    Set Dic_Story = New Dictionary

    Dim myProject As String
    myProject = project
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
    If debug_ Then Debug.Print JSONLib.toString(LocalSavedSettings("ListBox_Setting_Status_parents"))
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
    Dim subjson As Integer
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
    subjson = 0
    If debug_ Then Debug.Print "limit " & limit & " / offset " & offset & " / total " & total
    Do While total > nextoffset
        subjson = 1
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
    Loop
    If subjson = 1 Then
        For Each Var In jsonsub("issues")
            json("issues").Add Var
        Next Var
    End If
    listline = ListBox_TicketList.ListCount
    For Each Var In json("issues")
        Dim tmpsubject As String
        Dim tmpdescrip As String
        tmpsubject = Var("subject")
        tmpdescrip = Var("description")
        If searchContents = 1 Then
            If (Var("project")("id") = myProject And ((reg_hit(tmpsubject, TextBox_SearchKey.Text) > 0) Or (reg_hit(tmpdescrip, TextBox_SearchKey.Text) > 0))) Then
                If debug_ Then Debug.Print "this ticket " & Var("id") & " is hit"
                ListBox_TicketList.AddItem ""
                ListBox_TicketList.List(listline, 0) = Var("id")
                ListBox_TicketList.List(listline, 1) = Var("tracker")("name")
                ListBox_TicketList.List(listline, 2) = Var("subject")
                ListBox_TicketList.List(listline, 3) = Mid(Var("description"), 1, 50)
                listline = listline + 1
            Else
                If debug_ Then Debug.Print "this ticket " & Var("id") & " is not match which project_id is " & Var("project")("id") & " not " & myProject & " and subject is " & tmpsubject & " not " & TextBox_SearchKey.Text
            End If
        ElseIf (Var("project")("id") = myProject And (reg_hit(tmpsubject, TextBox_SearchKey.Text) > 0)) Then
        
            If debug_ Then Debug.Print "this ticket " & Var("id") & " is hit"
            ListBox_TicketList.AddItem ""
            ListBox_TicketList.List(listline, 0) = Var("id")
            ListBox_TicketList.List(listline, 1) = Var("tracker")("name")
            ListBox_TicketList.List(listline, 2) = Var("subject")
            ListBox_TicketList.List(listline, 3) = Mid(Var("description"), 1, 50)
            listline = listline + 1
        Else
            If debug_ Then Debug.Print "this ticket " & Var("id") & " is not match which project_id is " & Var("project")("id") & " not " & myProject & " and subject is " & tmpsubject & " not " & TextBox_SearchKey.Text
        End If
        tmpsubject = ""
        tmpdescrip = ""
    Next Var
    Set json = Nothing
    Set JSONLib = Nothing
End Sub
Public Sub get_ticket_for_keyword_subsub(ByRef project As String, ByRef keyword As String, ByVal url As String, ByVal apikey As String)
    Dim JSONLib As New JSONLib
    Dim json, tmpdic As Object
    Dim Var As Variant
    Dim total, offset, limit, nextoffset As Integer
    If debug_ Then Debug.Print "Calle :: get_ticket_for_keyword_subsub"
    Set Dic_Story = New Dictionary

    Dim myProject As String
    myProject = project
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
    If debug_ Then Debug.Print JSONLib.toString(LocalSavedSettings("ListBox_Setting_Status_child"))
    If LocalSavedSettings.exists("ListBox_Setting_Status_child") Then
        Set tmpdic = LocalSavedSettings("ListBox_Setting_Status_child")
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
    If LocalSavedSettings.exists("ListBox_Setting_Tracker_child") Then
        Set tmpdic = LocalSavedSettings("ListBox_Setting_Tracker_child")
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
    Dim subjson As Integer
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
    subjson = 0
    If debug_ Then Debug.Print "limit " & limit & " / offset " & offset & " / total " & total
    Do While total > nextoffset
        subjson = 1
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
    Loop
    If subjson = 1 Then
        For Each Var In jsonsub("issues")
            json("issues").Add Var
        Next Var
    End If
    listline = ListBox_TicketList.ListCount
    For Each Var In json("issues")
        Dim tmpsubject As String
        Dim tmpdescrip As String
        tmpsubject = Var("subject")
        tmpdescrip = Var("description")
        If searchContents = 1 Then
            If (Var("project")("id") = myProject And ((reg_hit(tmpsubject, TextBox_SearchKey.Text) > 0) Or (reg_hit(tmpdescrip, TextBox_SearchKey.Text) > 0))) Then
                If debug_ Then Debug.Print "this ticket " & Var("id") & " is hit"
                ListBox_TicketList.AddItem ""
                ListBox_TicketList.List(listline, 0) = Var("id")
                ListBox_TicketList.List(listline, 1) = Var("tracker")("name")
                ListBox_TicketList.List(listline, 2) = Var("subject")
                ListBox_TicketList.List(listline, 3) = Mid(Var("description"), 1, 50)
                listline = listline + 1
            Else
                If debug_ Then Debug.Print "this ticket " & Var("id") & " is not match which project_id is " & Var("project")("id") & " not " & myProject & " and subject is " & tmpsubject & " not " & TextBox_SearchKey.Text
            End If
        ElseIf (Var("project")("id") = myProject And (reg_hit(tmpsubject, TextBox_SearchKey.Text) > 0)) Then
        
            If debug_ Then Debug.Print "this ticket " & Var("id") & " is hit"
            ListBox_TicketList.AddItem ""
            ListBox_TicketList.List(listline, 0) = Var("id")
            ListBox_TicketList.List(listline, 1) = Var("tracker")("name")
            ListBox_TicketList.List(listline, 2) = Var("subject")
            ListBox_TicketList.List(listline, 3) = Mid(Var("description"), 1, 50)
            listline = listline + 1
        Else
            If debug_ Then Debug.Print "this ticket " & Var("id") & " is not match which project_id is " & Var("project")("id") & " not " & myProject & " and subject is " & tmpsubject & " not " & TextBox_SearchKey.Text
        End If
        tmpsubject = ""
        tmpdescrip = ""
    Next Var
    Set json = Nothing
    Set JSONLib = Nothing
End Sub
Private Function reg_hit(ByRef str As String, ByRef ptrn As String)
    Dim regex
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = ptrn
    Dim result
    Set result = regex.Execute(str)
    reg_hit = result.Count
    Set result = Nothing
    Set regex = Nothing
End Function

Public Sub CommandButton_SearchTicket_Click()
    If RMTS_Search.CommandButton_SearchTicket.Enabled = False Then
        Exit Sub
    End If

    CommandButton_SearchTicket.Enabled = False

    If LocalSavedSettings.exists("Dic_Projects_ID") Then
        If LocalSavedSettings("Dic_Projects_ID").exists(ComboBox_Project.value) Then
            project = LocalSavedSettings("Dic_Projects_ID")(ComboBox_Project.value)
        End If
    End If

    ListBox_TicketList.Clear

    If project = "" Then
        CommandButton_SearchTicket.Enabled = True
        Exit Sub
    End If

    If keywordsearchonAllTrackers = 1 Then
        Call get_ticket_for_keyword_categ(project, TextBox_SearchKey.Text, Setting_Redmine_URL, Setting_Redmine_APIKEY)
        Call get_ticket_for_keyword_subcat(project, TextBox_SearchKey.Text, Setting_Redmine_URL, Setting_Redmine_APIKEY)
        Call get_ticket_for_keyword_subsub(project, TextBox_SearchKey.Text, Setting_Redmine_URL, Setting_Redmine_APIKEY)
    Else
        Call get_ticket_for_keyword_subsub(project, TextBox_SearchKey.Text, Setting_Redmine_URL, Setting_Redmine_APIKEY)
    End If
    CommandButton_SearchTicket.Enabled = True
End Sub

Private Sub ListBox_TicketList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If ListBox_TicketList.value = "" Then
        Exit Sub
    End If
    Dim myindex As Integer
    myindex = ListBox_TicketList.ListIndex
    If debug_ Then Debug.Print "listbox index " & myindex
    Call RMTM_Creater.set_select_ticket_id(ListBox_TicketList.List(myindex, 0), ListBox_TicketList.List(myindex, 2))
    Unload Me
End Sub

Private Sub TextBox_SearchKey_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode <> 13 Then Exit Sub
    Call CommandButton_SearchTicket_Click
End Sub
Public Sub rmts_initialize()
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
    If keywordsearchonAllTrackers = 1 Then
        Label_serachTarget = "For All Category"
    ElseIf keywordsearchonAllTrackers = 0 Then
        Label_serachTarget = "For SubSubCateg."
    End If
    If LocalSavedSettings.exists("searchContents") Then
        searchContents = LocalSavedSettings("searchContents")
    Else
       If debug_ Then Debug.Print "can not find LocalSavedSettings(""searchContents"")"
    End If
    If LocalSavedSettings.exists("Dic_Projects") Then
        Set tmpdic = LocalSavedSettings("Dic_Projects")
        For Each Var In tmpdic
            RMTS_Search.ComboBox_Project.AddItem Var
        Next Var
    End If
End Sub

Private Sub UserForm_Activate()
        If Setting_Redmine_APIKEY = "" Or Setting_Redmine_URL = "" Then
            Button_Settings_Click
        End If
End Sub
Private Sub Button_Settings_Click()
    RMTC_Setting.Show
    If Initialized = 1 Then
        Call rmts_initialize
    Else
        Unload Me
    End If
End Sub
Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    If Initialized = 1 Then
        Call rmts_initialize
        ListBox_TicketList.ColumnWidths = "30;65;80;200"
    Else
        MsgBox "Failed to Load"
        Me.Width = 0
        Me.Height = 0
        Exit Sub
    End If
End Sub

