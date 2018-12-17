VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalenderForm 
   Caption         =   "Calender"
   ClientHeight    =   2700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3375
   OleObjectBlob   =   "CalenderForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CalenderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private current_date As Date
Private accept_date As Date
Private cntrol_type As String
Private callback_form As Control
Private callback_cell As String


Sub setCallBackControl(cb As Control)
   cntrol_type = "UserForm"
   Set callback_form = cb
End Sub

Sub setCallBackCell(cb As String)
   cntrol_type = "Cell"
   callback_cell = cb
End Sub


Private Sub CommandButton1_Click()
    current_date = DateAdd("m", 1, current_date)
    DateLabel.Caption = Format(current_date, "yyyy/mm")
    Call createDays
End Sub

Private Sub CommandButton2_Click()
    current_date = DateAdd("m", -1, current_date)
    DateLabel.Caption = Format(current_date, "yyyy/mm")
    Call createDays
End Sub

Sub createDays()
    Dim cal_date As Date
    
    current_month = Month(current_date)
    DateLabel.Caption = Format(current_date, "yyyy/mm")
    'w = Weekday(current_date)

    Call clearDateLabel
    cal_date = Format(current_date, "yyyy/mm") & "/1"
    w = Weekday(cal_date)
    cnt = 1
    For Each ctl In Controls
        If TypeName(ctl) = "Label" Then
            If cnt >= w Then
                ctl.Caption = Day(cal_date)
                If Year(accept_date) = Year(cal_date) And Month(accept_date) = Month(cal_date) And Day(accept_date) = Day(cal_date) Then
                  ctl.BackColor = RGB(190, 200, 255)
                End If
                cal_date = cal_date + 1
                If Month(cal_date) <> current_month Then Exit For
            End If
            cnt = cnt + 1
        End If
    Next

End Sub


Sub clearDateLabel()
    For Each ctl In Controls
        ctl.BackColor = RGB(235, 235, 255)
        If ctl <> DateLabel Then
            If TypeName(ctl) = "Label" Then ctl.Caption = ""
        End If
    Next
    DateLabel.BackColor = RGB(255, 255, 255)
End Sub

Sub setDate(d0)

   On Error GoTo ErrHandler
   d = CDate(d0)
   If d = 0 Then d = Now()
   accept_date = d
   current_date = d
   Call createDays
   Exit Sub

ErrHandler:
   d = Now()
   accept_date = d
   current_date = d
   Call createDays

End Sub

Sub sendDate(d)
   If d = "" Then Exit Sub

   If cntrol_type = "UserForm" Then
      If TypeName(callback_form) = "TextBox" Then callback_form.Text = Format(current_date, "yyyy/m/") & d
      If TypeName(callback_form) = "Label" Then callback_form.Caption = Format(current_date, "yyyy/m/") & d
   End If
   
'   If cntrol_type = "Cell" Then
'      Range(callback_cell) = Format(current_date, "yyyy/m/") & d
'   End If

   CalenderForm.Hide
End Sub



Private Sub Label1_Click()
   Label1.BackColor = RGB(255, 190, 225)
   Call sendDate(Label1.Caption)
End Sub


Private Sub Label2_Click()
   Label2.BackColor = RGB(255, 225, 225)
   Call sendDate(Label2.Caption)
End Sub

Private Sub Label3_Click()
   Label3.BackColor = RGB(255, 225, 225)
   Call sendDate(Label3.Caption)
End Sub

Private Sub Label4_Click()
   Label4.BackColor = RGB(255, 225, 225)
   Call sendDate(Label4.Caption)
End Sub

Private Sub Label5_Click()
   Label5.BackColor = RGB(255, 225, 225)
   Call sendDate(Label5.Caption)
End Sub

Private Sub Label6_Click()
   Label6.BackColor = RGB(255, 225, 225)
   Call sendDate(Label6.Caption)
End Sub

Private Sub Label7_Click()
   Label7.BackColor = RGB(255, 225, 225)
   Call sendDate(Label7.Caption)
End Sub

Private Sub Label8_Click()
   Label8.BackColor = RGB(255, 225, 225)
   Call sendDate(Label8.Caption)
End Sub

Private Sub Label9_Click()
   Label9.BackColor = RGB(255, 225, 225)
   Call sendDate(Label9.Caption)
End Sub

Private Sub Label10_Click()
   Label10.BackColor = RGB(255, 225, 225)
   Call sendDate(Label10.Caption)
End Sub

Private Sub Label11_Click()
   Label11.BackColor = RGB(255, 225, 225)
   Call sendDate(Label11.Caption)
End Sub

Private Sub Label12_Click()
   Label12.BackColor = RGB(255, 225, 225)
   Call sendDate(Label12.Caption)
End Sub

Private Sub Label13_Click()
   Label13.BackColor = RGB(255, 225, 225)
   Call sendDate(Label13.Caption)
End Sub

Private Sub Label14_Click()
   Label14.BackColor = RGB(255, 225, 225)
   Call sendDate(Label14.Caption)
End Sub

Private Sub Label15_Click()
   Label15.BackColor = RGB(255, 225, 225)
   Call sendDate(Label15.Caption)
End Sub

Private Sub Label16_Click()
   Label16.BackColor = RGB(255, 225, 225)
   Call sendDate(Label16.Caption)
End Sub

Private Sub Label17_Click()
   Label17.BackColor = RGB(255, 225, 225)
   Call sendDate(Label17.Caption)
End Sub

Private Sub Label18_Click()
   Label18.BackColor = RGB(255, 225, 225)
   Call sendDate(Label18.Caption)
End Sub

Private Sub Label19_Click()
   Label19.BackColor = RGB(255, 225, 225)
   Call sendDate(Label19.Caption)
End Sub

Private Sub Label20_Click()
   Label20.BackColor = RGB(255, 225, 225)
   Call sendDate(Label20.Caption)
End Sub

Private Sub Label21_Click()
   Label21.BackColor = RGB(255, 225, 225)
   Call sendDate(Label21.Caption)
End Sub

Private Sub Label22_Click()
   Label22.BackColor = RGB(255, 225, 225)
   Call sendDate(Label22.Caption)
End Sub

Private Sub Label23_Click()
   Label23.BackColor = RGB(255, 225, 225)
   Call sendDate(Label23.Caption)
End Sub

Private Sub Label24_Click()
   Label24.BackColor = RGB(255, 225, 225)
   Call sendDate(Label24.Caption)
End Sub

Private Sub Label25_Click()
   Label25.BackColor = RGB(255, 225, 225)
   Call sendDate(Label25.Caption)
End Sub

Private Sub Label26_Click()
   Label26.BackColor = RGB(255, 225, 225)
   Call sendDate(Label26.Caption)
End Sub

Private Sub Label27_Click()
   Label27.BackColor = RGB(255, 225, 225)
   Call sendDate(Label27.Caption)
End Sub

Private Sub Label28_Click()
   Label28.BackColor = RGB(255, 225, 225)
   Call sendDate(Label28.Caption)
End Sub

Private Sub Label29_Click()
   Label29.BackColor = RGB(255, 225, 225)
   Call sendDate(Label29.Caption)
End Sub

Private Sub Label30_Click()
   Label30.BackColor = RGB(255, 225, 225)
   Call sendDate(Label30.Caption)
End Sub

Private Sub Label31_Click()
   Label31.BackColor = RGB(255, 225, 225)
   Call sendDate(Label31.Caption)
End Sub

Private Sub Label32_Click()
   Label32.BackColor = RGB(255, 225, 225)
   Call sendDate(Label32.Caption)
End Sub

Private Sub Label33_Click()
   Label33.BackColor = RGB(255, 225, 225)
   Call sendDate(Label33.Caption)
End Sub

Private Sub Label34_Click()
   Label34.BackColor = RGB(255, 225, 225)
   Call sendDate(Label34.Caption)
End Sub

Private Sub Label35_Click()
   Label35.BackColor = RGB(255, 225, 225)
   Call sendDate(Label35.Caption)
End Sub

Private Sub Label36_Click()
   Label36.BackColor = RGB(255, 225, 225)
   Call sendDate(Label36.Caption)
End Sub

Private Sub Label37_Click()
   Label37.BackColor = RGB(255, 225, 225)
   Call sendDate(Label37.Caption)
End Sub

Private Sub Label38_Click()
   Label38.BackColor = RGB(255, 225, 225)
   Call sendDate(Label38.Caption)
End Sub

Private Sub Label39_Click()
   Label39.BackColor = RGB(255, 225, 225)
   Call sendDate(Label39.Caption)
End Sub

Private Sub Label40_Click()
   Label40.BackColor = RGB(255, 225, 225)
   Call sendDate(Label40.Caption)
End Sub

Private Sub Label41_Click()
   Label41.BackColor = RGB(255, 225, 225)
   Call sendDate(Label41.Caption)
End Sub

Private Sub Label42_Click()
   Label42.BackColor = RGB(255, 225, 225)
   Call sendDate(Label42.Caption)
End Sub

