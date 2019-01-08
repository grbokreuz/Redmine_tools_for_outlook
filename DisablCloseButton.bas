Attribute VB_Name = "DisablCloseButton"
'Outlookの [閉じる] ボタンを無効に設定
 
'Win32API
Public Declare PtrSafe Function FindWindow Lib "user32" _
Alias "FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long
Public Declare PtrSafe Function GetSystemMenu Lib "user32" _
(ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare PtrSafe Function DeleteMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, _
ByVal wFlags As Long) As Long
Public Declare PtrSafe Function DrawMenuBar Lib "user32" _
(ByVal hwnd As Long) As Long
 
Public Const SC_CLOSE = &HF060&
Public Const MF_BYCOMMAND = &H0&
 
Public Sub DisableCloseButton()
Dim hwnd As Long
Dim hMenu As Long
Dim rc As Long
hwnd = FindWindow("rctrl_renwnd32", vbNullString) 'rctrl_renwnd32=Outoook
hMenu = GetSystemMenu(hwnd, 0&)
rc = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
rc = DrawMenuBar(hwnd)

' Application.ActiveExplorer.WindowState = olMinimized
End Sub
'参考
'//park11.wakwak.com/~miko/Excel_Note/11-01_userform.htm#11-01-11
 
'注意
'PtrSafe はVBA7で搭載されたキーワード。VBA6.5以前の場合はこのキーワードを削除。


