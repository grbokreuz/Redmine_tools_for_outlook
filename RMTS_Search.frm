VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RMTS_Search 
   Caption         =   "RMTS_Search"
   ClientHeight    =   5325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
   OleObjectBlob   =   "RMTS_Search.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "RMTS_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' for redmine api under 3.3 , upper virsion is not support search api
' /redmine/search.xml?q=querystring&all_words=1&titles_only=0&attachments=1&options=1&open_issues=1&scope=my_projects
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
    filter_tracker = ""
    Set Var = New Dictionary
    Set tmpdic = New Dictionary

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
    jsonstring = GetData(url & "/issues.json?key=" & apikey & "&status_id=open&project_id=" & myProject & "&" & filterstr)
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
        subjsonstr = GetData(url & "/issues.json?key=" & apikey & "&status_id=open&project_id=" & myProject & "&offset=" & nextoffset & "&" & filterstr)
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
                ListBox_TicketList.List(listline, 3) = convert_no_return("" & Var("description"))
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
            ListBox_TicketList.List(listline, 3) = convert_no_return("" & Var("description"))
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

    filter_tracker = ""
    Set Var = New Dictionary
    Set tmpdic = New Dictionary

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
    jsonstring = GetData(url & "/issues.json?key=" & apikey & "&status_id=open&project_id=" & myProject & "&" & filterstr)
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
        subjsonstr = GetData(url & "/issues.json?key=" & apikey & "&status_id=open&project_id=" & myProject & "&offset=" & nextoffset & "&" & filterstr)
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
                ListBox_TicketList.List(listline, 3) = convert_no_return("" & Var("description"))
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
            ListBox_TicketList.List(listline, 3) = convert_no_return("" & Var("description"))
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
    filter_tracker = ""
    Set Var = New Dictionary
    Set tmpdic = New Dictionary

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

    Dim jsonstring As String
    jsonstring = GetData(url & "/issues.json?key=" & apikey & "&status_id=open&project_id=" & myProject & "&" & filterstr)
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
        subjsonstr = GetData(url & "/issues.json?key=" & apikey & "&status_id=open&project_id=" & myProject & "&offset=" & nextoffset & "&" & filterstr)
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
                ListBox_TicketList.List(listline, 3) = convert_no_return("" & Var("description"))
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
            ListBox_TicketList.List(listline, 3) = convert_no_return("" & Var("description"))
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

Private Sub ComboBox_Project_Change()
    TransactionSearch("Default_Projects") = ComboBox_Project.value
End Sub

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

    DoEvents

    If keywordsearchonAllTrackers = 1 Then
        Call get_ticket_for_keyword_categ(project, TextBox_SearchKey.Text, Setting_Redmine_URL, Setting_Redmine_APIKEY)
        Call get_ticket_for_keyword_subcat(project, TextBox_SearchKey.Text, Setting_Redmine_URL, Setting_Redmine_APIKEY)
        Call get_ticket_for_keyword_subsub(project, TextBox_SearchKey.Text, Setting_Redmine_URL, Setting_Redmine_APIKEY)
    Else
        Call get_ticket_for_keyword_subsub(project, TextBox_SearchKey.Text, Setting_Redmine_URL, Setting_Redmine_APIKEY)
    End If

    DoEvents
    CommandButton_SearchTicket.Enabled = True
End Sub


Private Sub ListBox_TicketList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If ListBox_TicketList.value = "" Then
        Exit Sub
    End If
    Dim myindex As Integer
    myindex = ListBox_TicketList.ListIndex
    If debug_ Then Debug.Print "listbox index " & myindex
    
    If RMTS_Search_SingleMode = True Then
        If debug_ Then Debug.Print "ListBox_TicketList_DblClick open web start : " & Setting_Redmine_URL & "/issue/" & ListBox_TicketList.List(myindex, 0) & "?key=" & Setting_Redmine_APIKEY
        If ListBox_TicketList.List(myindex, 0) = "" Then
            Exit Sub
        End If
        If webincreasemyAPIKey = 1 Then
            openweb (Setting_Redmine_URL & "/issues/" & ListBox_TicketList.List(myindex, 0) & "?key=" & Setting_Redmine_APIKEY)
        Else
            openweb (Setting_Redmine_URL & "/issues/" & ListBox_TicketList.List(myindex, 0))
        End If
    
    Else
        Call RMTM_Creater.set_select_ticket_id(ListBox_TicketList.List(myindex, 0), ListBox_TicketList.List(myindex, 2))
        Unload Me
    End If
End Sub
Private Sub ListBox_TicketList_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim buf As Long
    If Button = 2 Then
        buf = Int((Y + 1) / ListBox_TicketList.Font.Size)
        If buf > ListBox_TicketList.ListCount - 1 Then
            buf = ListBox_TicketList.ListCount - 1
        End If
        ListBox_TicketList.Selected(buf) = True
        MsgBox ListBox_TicketList.List(buf, 3)
    End If
End Sub
Private Sub TextBox_SearchKey_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If ((Shift And olShiftStateShiftMask) > 0) Or ((Shift And olShiftStateAltMask) > 0) Or ((Shift And olShiftStateCtrlMask) > 0) Or _
        KeyCode Is Nothing Or KeyCode = vbKeyUp Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or _
        KeyCode = vbKeyNumlock Or KeyCode = vbKeyPrint Or KeyCode = vbKeyShift Or KeyCode = vbKeyEscape Or KeyCode = vbKeyCapital Or KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        Exit Sub
    End If

    If KeyCode = 13 Then
        Call CommandButton_SearchTicket_Click
        Exit Sub
    End If
    If for_Japanese = True And 33 <= KeyCode And KeyCode <= 126 Then
        KeyCode = 0
        TextBox_SearchKey.IMEMode = vbIMEModeHiragana
        Exit Sub
    End If
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

    If debug_ Then Debug.Print "rmts_initialize :: get regset : AllSetting" & RegStr
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

    RegStr = GetSetting("OutlookRMTC", "Transaction", "TransactionSearch")
    If debug_ Then Debug.Print "UserForm_Initialize :: get regset : TransactionSearch = " & RegStr
    Set TransactionSearch = JSONLib.parse(RegStr)
    Debug.Assert Err.Number = 0

    If Not TransactionSearch Is Nothing Then
        If debug_ Then Debug.Print "UserForm_Initialize :: TransactionSearch data " & JSONLib.toString(TransactionSearch)
        If debug_ Then Debug.Print "apply TransactionSearch data to form"
        If TransactionSearch.exists("Default_Projects") Then
            RMTS_Search.ComboBox_Project.value = TransactionSearch("Default_Projects")
        End If
    Else
        If debug_ Then Debug.Print "UserForm_Initialize :: TransactionSearch data is nothing "
        Set TransactionSearch = New Dictionary
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
Private Sub UserForm_Initialize()
    If Initialized = 1 Then
        Call rmts_initialize
        ListBox_TicketList.ColumnWidths = "30;45;100;0"
    Else
        MsgBox "Failed to Load"
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

        RegStr = JSONLib.toString(TransactionSearch)
        If debug_ Then Debug.Print "UserForm_QueryClose :: save to reg : "; RegStr
        Call save_transaction_Data_to_reg
        If debug_ Then Debug.Print "UserForm_QueryClose ended"
    Else
        If debug_ Then Debug.Print "UserForm_QueryClose called but not initialized or going reload"
        Exit Sub
    End If
End Sub
Private Sub save_transaction_Data_to_reg()
    If debug_ Then Debug.Print "Save TransactionSearch data to reg."
    
    If TransactionSearch Is Nothing Then
        If debug_ Then Debug.Print "TransactionSearch Data is nohing"
        Exit Sub
    End If
    
    If TransactionSearch.exists("Default_Projects") Then
        TransactionSearch("Default_Projects") = RMTS_Search.ComboBox_Project.value
    Else
        TransactionSearch.Add "Default_Projects", RMTS_Search.ComboBox_Project.value
    End If

    Dim JSONLib As New JSONLib
    If debug_ Then Debug.Print "save_transaction_Data_to_reg :: Else data " & JSONLib.toString(TransactionSearch)
    
    SaveSetting "OutlookRMTC", "Transaction", "TransactionSearch", JSONLib.toString(TransactionSearch)
End Sub
Private Function convert_no_return(ByRef str As String)
    Dim outstr As String
    Dim siriesstr As String
    siriesstr = "AorD"
    str = Mid(str, 1, 500)
    outstr = ""
    For k = 1 To Len(str)
        If Hex(Asc(Mid(str, k, 1))) = "D" Or Hex(Asc(Mid(str, k, 1))) = "A" Then
            If siriesstr = "AorD" Then
            Else
                outstr = outstr & Mid(str, k, 1)
                siriesstr = "AorD"
            End If
        Else
            outstr = outstr & Mid(str, k, 1)
            siriesstr = Hex(Asc(Mid(str, k, 1)))
        End If

    Next
    convert_no_return = outstr
End Function
