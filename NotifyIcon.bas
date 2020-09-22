Attribute VB_Name = "MNotifyIcon"
      'user defined type required by Shell_NotifyIcon API call
      Public Type NOTIFYICONDATA
       cbSize As Long
       hwnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hIcon As Long
       szTip As String * 64
      End Type

      'constants required by Shell_NotifyIcon API call:
      Public Const NIM_ADD = &H0
      Public Const NIM_MODIFY = &H1
      Public Const NIM_DELETE = &H2
      Public Const NIF_MESSAGE = &H1
      Public Const NIF_ICON = &H2
      Public Const NIF_TIP = &H4
      Public Const Mouse_Move = 512
      Public Const Mouse_Left_Down = 513
      Public Const Mouse_Left_Click = 514
      Public Const Mouse_Left_DbClick = 515
      Public Const Mouse_Right_Down = 516
      Public Const Mouse_Right_Click = 517
      Public Const Mouse_Right_DbClick = 518
      Public Const Mouse_Button_Down = 519
      Public Const Mouse_Button_Click = 520
      Public Const Mouse_Button_DbClick = 521
      
      Public Declare Function Shell_NotifyIcon Lib "SHELL32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      Public abd As NOTIFYICONDATA
