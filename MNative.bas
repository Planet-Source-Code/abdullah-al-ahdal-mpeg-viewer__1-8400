Attribute VB_Name = "MNative"
Option Base 1
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long


Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type SizeRECT
IWidth As Long
IHight As Long
End Type

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

Global PathFull As String
Type ClassV
Class(6) As String
'TitleClass(6) As String
End Type
Sub GetPlaceMe()
On Error Resume Next
'////////////////// //////////////////////////////////////
'//////////GetPlace me.exe /////////////////////////////
If Len(App.Path) > 3 Then
PathFull = App.Path & "\" & App.EXEName & ".exe"
Else
PathFull = App.Path & App.EXEName & ".exe"
End If
'///////////////////////////////////////////////////////
End Sub
Public Function GetSize(hwnd) As SizeRECT
Dim VRECT As RECT
Call GetWindowRect(hwnd, VRECT)
Dim IWidth As Long
Dim IHight As Long
GetSize.IWidth = VRECT.Right - VRECT.Left
GetSize.IHight = VRECT.Bottom - VRECT.Top
End Function
Public Function GetHwndByClass(Class1 As String, Class2 As String, Class3 As String, Class4 As String, Class5 As String, Class6 As String) As Long
Dim FindClass(6) As Long

FindClass(1) = FindWindow(Class1, vbNullString)

If Class2 = "0" Then GetHwndByClass = FindClass(1): Exit Function


FindClass(2) = FindWindowEx(FindClass(1), 0, Class2, vbNullString)
If Class3 = "0" Then GetHwndByClass = FindClass(2): Exit Function

FindClass(3) = FindWindowEx(FindClass(2), 0, Class3, vbNullString)
If Class4 = "0" Then GetHwndByClass = FindClass(3): Exit Function

FindClass(4) = FindWindowEx(FindClass(3), 0, Class4, vbNullString)
If Class5 = "0" Then GetHwndByClass = FindClass(4): Exit Function

FindClass(5) = FindWindowEx(FindClass(4), 0, Class5, vbNullString)
If Class6 = "0" Then GetHwndByClass = FindClass(5): Exit Function

FindClass(6) = FindWindowEx(FindClass(5), 0, Class6, vbNullString)
GetHwndByClass = FindClass(6): Exit Function
End Function
Sub OpenVideo(PathVideo As String, HandleVideo As String, TypeVideoAviOrMpeg As String, WherePlayVideo_HWnd As Long, VWidth As Long, VHight As Long, VTop As Long, VLeft As Long)
Dim ToDo As String
    Last$ = WherePlayVideo_HWnd & " Style " & &H40000000
    
    ToDo$ = "open " & PathVideo & " Type " & TypeVideoAviOrMpeg & "video Alias video parent " & Last$
    'MsgBox ToDo$
    X% = mciSendString(ToDo$, 0&, 0, 0)
    
    If VWidth = 0 And VHight = 0 Then
    ToDo$ = "put video window at " & VLeft & " " & VTop & " "
    'MsgBox ToDo$
        X% = mciSendString(ToDo$, 0&, 0, 0)
    Else
    ToDo$ = "put video window at " & VLeft & " " & VTop & " " & VWidth & " " & VHight
    'MsgBox ToDo$
        X% = mciSendString(ToDo$, 0&, 0, 0)
    End If
End Sub
Sub CloseVideo()
X% = mciSendString("Close video", 0&, 0, 0&)
End Sub
Sub ResumeVideo()
X% = mciSendString("Resume video", 0&, 0, 0&)
End Sub
Sub StopVideo()
X% = mciSendString("Stop video", 0&, 0, 0&)
End Sub
Sub PlayVideo(FromWhereStartPlayVideo As String, ToWherePlayVideo As String)
If FromWhereStartPlayVideo = "None" Then
    ToDo$ = "play video"
    'MsgBox ToDo$
        X% = mciSendString(ToDo$, 0&, 0, 0)
    ElseIf Not FromWhereStartPlayVideo = "None" And Not ToWherePlayVideo = "None" Then
    ToDo$ = "play video from " & FromWhereStartPlayVideo & " to " & ToWherePlayVideo
    'MsgBox ToDo$
    X% = mciSendString(ToDo$, 0&, 0, 0)
    ElseIf ToWherePlayVideo = "None" Then
    ToDo$ = "play video from " & FromWhereStartPlayVideo
    'MsgBox ToDo$
    X% = mciSendString(ToDo$, 0&, 0, 0)
End If

End Sub
Sub ReSizeVideo(VWidth As Long, VHight As Long, VTop As Long, VLeft As Long)
ToDo$ = "put video window at " & VLeft & " " & VTop & " " & VWidth & " " & VHight
        X% = mciSendString(ToDo$, 0&, 0, 0)
End Sub
Sub PauseVideo()
    X% = mciSendString("Pause video", 0&, 0, 0&)
End Sub
Sub SeekTo(Where As Long)
    X% = mciSendString("seek video to " & Where, 0&, 0, 0)
End Sub
Public Function GetTotalFrames() As String
Dim mssg As String * 255
  X% = mciSendString("set video time format frames", mssg, 255, 0)
  X% = mciSendString("status video length", mssg, 255, 0)
  GetTotalFrames = Str(mssg)
End Function
Public Function GetTotalTimeBymilliseconds() As String
  Dim mssg As String * 255
  X% = mciSendString("set video time format ms", mssg, 255, 0)
  X% = mciSendString("status video length", mssg, 255, 0)
GetTotalTimeBymilliseconds = Str(mssg)
End Function
Public Function GetVideoStats() As String
  Dim mssg As String * 255
  X% = mciSendString("status video mode", mssg, 255, 0)
  GetVideoStats = mssg
End Function

'//regedit
Public Sub SaveString(hKey As Long, StrPath As String, StrValue As String, StrData As String)
   Dim KeyH&
    r = RegCreateKey(hKey, StrPath, KeyH&)
    r = RegSetValueEx(KeyH&, StrValue, 0, 1, ByVal StrData, Len(StrData))
    r = RegCloseKey(KeyH&)
End Sub


