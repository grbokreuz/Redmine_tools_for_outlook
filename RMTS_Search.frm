VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RMTS_Search 
   Caption         =   "RMTS_Search"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "RMTS_Search.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "RMTS_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Search_Ticket_Subject
'Search_Ticket_ID

Public project As String
Public Function set_param(ByVal inp_id As String, ByVal inp_name As String)
    project = inp_id
    Label_Projectname.Caption = inp_name
End Function
Public Sub get_ticket_for_keyword(ByRef project As String, ByRef keyword As String, ByVal url As String, ByVal apikey As String)
    Dim JSONLib As New JSONLib
    Dim json, tmpdic As Object
    Dim Var As Variant
    Dim total, offset, limit, nextoffset As Integer
    If debug_ Then Debug.Print "Calle :: set_story_ticket_for_selected_project"
    Set Dic_Story = New Dictionary

    ListBox_TicketList.Clear

    Dim myProject As String
    myProject = project
    
    If myProject = "" Then
        If debug_ Then Debug.Print "Not found; project "
        Exit Sub
        
    End If

    ' create filter
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

    listline = 0
    For Each Var In json("issues")
        Dim tmpsubject As String
        tmpsubject = Var("subject")
        If Var("project")("id") = myProject And (reg_hit(tmpsubject, TextBox_SearchKey.Text) > 0) Then
            If debug_ Then Debug.Print "this ticket " & Var("id") & " is hit"
            ListBox_TicketList.AddItem ""
            ListBox_TicketList.List(listline, 0) = Var("id")
            ListBox_TicketList.List(listline, 1) = Var("tracker")("name")
            ListBox_TicketList.List(listline, 2) = Var("subject")
            listline = listline + 1
        Else
            If debug_ Then Debug.Print "this ticket " & Var("id") & " is not match which project_id is " & Var("project")("id") & " not " & myProject & " and subject is " & tmpsubject & " not " & TextBox_SearchKey.Text
        End If
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
End Function

Public Sub CommandButton_SearchTicket_Click()
    Call get_ticket_for_keyword(project, Label_Projectname.Caption, Setting_Redmine_URL, Setting_Redmine_APIKEY)
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

Private Sub UserForm_Initialize()
   ListBox_TicketList.ColumnWidths = "30;65;80"
End Sub


Private Sub TextBox_SearchKey_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode <> 13 Then Exit Sub
    Call CommandButton_SearchTicket_Click
End Sub

