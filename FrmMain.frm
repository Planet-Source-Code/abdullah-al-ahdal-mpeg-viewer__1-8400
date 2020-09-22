VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Professional Viewer Video By Abdullah Al-ahdal. E-Mail:a_ahdal@yahoo.com"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerCheckhWnd 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3570
      Top             =   2670
   End
   Begin VB.Timer TimerVideo 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2190
      Top             =   2670
   End
   Begin VB.Frame FrameControlVideo 
      Height          =   5595
      Left            =   180
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   7245
      Begin VB.CommandButton CmdHFCV 
         Caption         =   "Hide"
         Height          =   375
         Left            =   810
         TabIndex        =   42
         Top             =   5070
         Width           =   5715
      End
      Begin VB.Frame Frame8 
         Caption         =   "Command Video"
         Height          =   1425
         Left            =   780
         TabIndex        =   37
         Top             =   3570
         Width           =   5685
         Begin VB.CommandButton CmdControlVideo 
            Caption         =   "&Back Video"
            Height          =   525
            Index           =   5
            Left            =   1560
            TabIndex        =   44
            Top             =   840
            Width           =   1245
         End
         Begin VB.CommandButton CmdControlVideo 
            Caption         =   "&Next Video"
            Height          =   525
            Index           =   4
            Left            =   2880
            TabIndex        =   43
            Top             =   840
            Width           =   1245
         End
         Begin VB.CommandButton CmdControlVideo 
            Caption         =   "P&ause Video"
            Height          =   525
            Index           =   3
            Left            =   2880
            TabIndex        =   41
            Top             =   300
            Width           =   1245
         End
         Begin VB.CommandButton CmdControlVideo 
            Caption         =   "&Close Video"
            Height          =   525
            Index           =   2
            Left            =   4200
            TabIndex        =   40
            Top             =   300
            Width           =   1245
         End
         Begin VB.CommandButton CmdControlVideo 
            Caption         =   "&Play Video"
            Height          =   525
            Index           =   1
            Left            =   1560
            TabIndex        =   39
            Top             =   300
            Width           =   1245
         End
         Begin VB.CommandButton CmdControlVideo 
            Caption         =   "&Open Video"
            Height          =   525
            Index           =   0
            Left            =   210
            TabIndex        =   38
            Top             =   300
            Width           =   1245
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Seek on video"
         Height          =   765
         Left            =   780
         TabIndex        =   35
         Top             =   2730
         Width           =   5655
         Begin VB.HScrollBar HScrollSeek 
            Height          =   345
            Left            =   240
            TabIndex        =   36
            Top             =   270
            Width           =   5145
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Size Video"
         Height          =   1485
         Left            =   1560
         TabIndex        =   25
         Top             =   1140
         Width           =   4035
         Begin VB.CheckBox ChUseOrginalWindowSize 
            Caption         =   "Use Orginal Window Size."
            Height          =   255
            Left            =   210
            TabIndex        =   34
            Top             =   1170
            Width           =   3495
         End
         Begin VB.TextBox TxtLeft 
            Height          =   315
            Left            =   2880
            TabIndex        =   32
            Text            =   "0"
            Top             =   810
            Width           =   825
         End
         Begin VB.TextBox TxtTop 
            Height          =   315
            Left            =   750
            TabIndex        =   30
            Text            =   "0"
            Top             =   810
            Width           =   795
         End
         Begin VB.TextBox TxtWidth 
            Height          =   315
            Left            =   2880
            TabIndex        =   28
            Top             =   300
            Width           =   825
         End
         Begin VB.TextBox TxtHight 
            Height          =   315
            Left            =   750
            TabIndex        =   26
            Top             =   300
            Width           =   795
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Left :"
            Height          =   195
            Index           =   1
            Left            =   2250
            TabIndex        =   33
            Top             =   870
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Top :"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   31
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Width :"
            Height          =   195
            Index           =   0
            Left            =   2250
            TabIndex        =   29
            Top             =   360
            Width           =   510
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hight :"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   27
            Top             =   330
            Width           =   465
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Stats Video"
         Height          =   705
         Left            =   1560
         TabIndex        =   24
         Top             =   240
         Width           =   4065
         Begin VB.Label Label4 
            Caption         =   "Stats Video:"
            Height          =   255
            Left            =   510
            TabIndex        =   65
            Top             =   300
            Width           =   885
         End
         Begin VB.Label LbStatsVideo 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   1440
            TabIndex        =   45
            Top             =   300
            Width           =   45
         End
      End
   End
   Begin VB.Frame FrameSelectPath 
      Height          =   5595
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   7245
      Begin VB.CommandButton CmdRemoveFileFromList 
         Caption         =   "Remove from list"
         Height          =   525
         Left            =   1710
         TabIndex        =   7
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CommandButton CmdAddFile 
         Caption         =   "Add to list"
         Height          =   525
         Left            =   150
         TabIndex        =   6
         Top             =   4920
         Width           =   1545
      End
      Begin VB.CommandButton CmdHideFrameSelectFile 
         Caption         =   "Hide"
         Height          =   525
         Left            =   5550
         TabIndex        =   5
         Top             =   4920
         Width           =   1545
      End
      Begin VB.FileListBox File1 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   3810
         MultiSelect     =   2  'Extended
         Pattern         =   "*.mpg;*.dat;*.mpeg;*.mpe"
         System          =   -1  'True
         TabIndex        =   4
         Top             =   660
         Width           =   3285
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2130
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   3555
      End
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   6945
      End
      Begin VB.ListBox ListFiles 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   1
         Top             =   2880
         Width           =   6945
      End
      Begin VB.Label LbCountVideos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Count Videos"
         Height          =   195
         Left            =   3900
         TabIndex        =   21
         Top             =   5100
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5595
      Left            =   180
      TabIndex        =   55
      Top             =   120
      Visible         =   0   'False
      Width           =   7245
      Begin VB.ListBox ListClasses 
         Height          =   2010
         Index           =   0
         Left            =   5580
         TabIndex        =   62
         Top             =   390
         Width           =   1545
      End
      Begin VB.ListBox ListClasses 
         Height          =   2010
         Index           =   6
         Left            =   2310
         TabIndex        =   61
         Top             =   2520
         Width           =   1485
      End
      Begin VB.ListBox ListClasses 
         Height          =   2010
         Index           =   5
         Left            =   3930
         TabIndex        =   60
         Top             =   2520
         Width           =   1545
      End
      Begin VB.ListBox ListClasses 
         Height          =   2010
         Index           =   4
         Left            =   5580
         TabIndex        =   59
         Top             =   2580
         Width           =   1545
      End
      Begin VB.ListBox ListClasses 
         Height          =   2010
         Index           =   3
         Left            =   690
         TabIndex        =   58
         Top             =   390
         Width           =   1515
      End
      Begin VB.ListBox ListClasses 
         Height          =   2010
         Index           =   2
         Left            =   2310
         TabIndex        =   57
         Top             =   390
         Width           =   1485
      End
      Begin VB.ListBox ListClasses 
         Height          =   2010
         Index           =   1
         Left            =   3930
         TabIndex        =   56
         Top             =   390
         Width           =   1515
      End
   End
   Begin VB.Frame FrameMain 
      Height          =   5595
      Left            =   180
      TabIndex        =   8
      Top             =   120
      Width           =   7245
      Begin VB.CommandButton CmdShowFCV 
         Caption         =   "Video Control && Infromation"
         Height          =   525
         Left            =   1860
         TabIndex        =   22
         Top             =   4980
         Width           =   3615
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         Height          =   525
         Left            =   5520
         TabIndex        =   20
         Top             =   4980
         Width           =   1515
      End
      Begin VB.CommandButton CmdMinimizeWindow 
         Caption         =   "Minimize Window"
         Height          =   525
         Left            =   300
         TabIndex        =   19
         Top             =   4980
         Width           =   1515
      End
      Begin VB.CommandButton CmdAddingVideotoList 
         Caption         =   "Adding Video to List"
         Height          =   555
         Left            =   300
         TabIndex        =   18
         Top             =   4380
         Width           =   6735
      End
      Begin VB.Frame Frame4 
         Caption         =   "Options"
         Height          =   975
         Left            =   240
         TabIndex        =   13
         Top             =   150
         Width           =   6825
         Begin VB.CheckBox ChPlayAutomaticOnStart 
            Caption         =   "Play Video Automatic on Start"
            Height          =   345
            Left            =   3960
            TabIndex        =   67
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox ChPlayNextVideoAutomatic 
            Caption         =   "Play Next Video Automatic"
            Height          =   285
            Left            =   1740
            TabIndex        =   17
            Top             =   390
            Width           =   2265
         End
         Begin VB.CheckBox ChRunWithSystem 
            Caption         =   "Run With System"
            Height          =   255
            Left            =   150
            TabIndex        =   16
            Top             =   390
            Width           =   1665
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Adding Places"
         Height          =   1455
         Left            =   270
         TabIndex        =   10
         Top             =   2760
         Width           =   6765
         Begin VB.CommandButton CmdCancel 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   345
            Left            =   5220
            TabIndex        =   66
            Top             =   1050
            Width           =   1125
         End
         Begin VB.CommandButton CmdUpdate 
            Caption         =   "&Update"
            Enabled         =   0   'False
            Height          =   345
            Left            =   5220
            TabIndex        =   64
            Top             =   690
            Width           =   1125
         End
         Begin VB.TextBox txtClass 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   1470
            TabIndex        =   53
            Top             =   1080
            Width           =   3645
         End
         Begin VB.TextBox txtClass 
            Enabled         =   0   'False
            Height          =   315
            Index           =   5
            Left            =   3990
            TabIndex        =   52
            Top             =   690
            Width           =   1125
         End
         Begin VB.TextBox txtClass 
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   2760
            TabIndex        =   51
            Top             =   690
            Width           =   1125
         End
         Begin VB.TextBox txtClass 
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   5220
            TabIndex        =   50
            Top             =   300
            Width           =   1125
         End
         Begin VB.TextBox txtClass 
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   3990
            TabIndex        =   49
            Top             =   300
            Width           =   1125
         End
         Begin VB.TextBox txtClass 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   2760
            TabIndex        =   48
            Top             =   300
            Width           =   1125
         End
         Begin VB.TextBox txtClass 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   47
            Top             =   300
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Title of this place :"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   54
            Top             =   1110
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Place Play video:"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   11
            Top             =   330
            Width           =   1230
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Menu of places"
         Height          =   1575
         Left            =   240
         TabIndex        =   9
         Top             =   1170
         Width           =   6825
         Begin VB.CommandButton CmdAddPlace 
            Caption         =   "&Add Place"
            Height          =   525
            Left            =   210
            TabIndex        =   63
            Top             =   600
            Width           =   825
         End
         Begin VB.CommandButton CmdSave 
            Caption         =   "&Save"
            Height          =   525
            Left            =   3000
            TabIndex        =   46
            Top             =   600
            Width           =   3645
         End
         Begin VB.CommandButton CmdDeletePlace 
            Caption         =   "&Delete Place"
            Height          =   525
            Left            =   2040
            TabIndex        =   15
            Top             =   600
            Width           =   825
         End
         Begin VB.ComboBox ComboPlaces 
            Height          =   315
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   210
            Width           =   6495
         End
         Begin VB.CommandButton CmdEditPlace 
            Caption         =   "&Edit Place"
            Height          =   525
            Left            =   1140
            TabIndex        =   12
            Top             =   600
            Width           =   825
         End
      End
   End
   Begin VB.Menu MnuMain 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu MnuMainShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu MnuMainHide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu MnuMainS1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMainNext 
         Caption         =   "&Next"
      End
      Begin VB.Menu MnuMainBack 
         Caption         =   "&Back"
      End
      Begin VB.Menu MnuMainS2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMainClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NFile As Long
Dim ClassWindow As ClassV
Dim Classes(5) As String
Dim PathFileClasses As String
Dim AddingNew As Boolean
Dim Editing As Boolean
Dim hWndVideo As Long
Dim SizeVideo As SizeRECT

Private Sub ChPlayAutomaticOnStart_Click()
SaveSetting App.Title, App.Title, "PlayAutomaticOnStart", ChPlayAutomaticOnStart.Value
End Sub

Private Sub ChPlayNextVideoAutomatic_Click()
SaveSetting App.Title, App.Title, "PlayNextVideoAutomatic", ChPlayNextVideoAutomatic.Value
End Sub

Private Sub ChRunWithSystem_Click()
If ChRunWithSystem.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Profissonal Viewer", PathFull
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Profissonal Viewer", 0
End If
End Sub

Private Sub ChUseOrginalWindowSize_Click()
SaveSetting App.Title, App.Title, "UseOrginalWindowSize", ChUseOrginalWindowSize.Value
End Sub

Private Sub CmdAddFile_Click()
File1_DblClick
End Sub

Private Sub CmdAddingVideotoList_Click()
FrameMain.Visible = False
FrameSelectPath.Visible = True
End Sub

Private Sub CmdAddPlace_Click()
AddingNewClassOrEditing "Add"
End Sub

Private Sub CmdCancel_Click()
EnableCMD
SaveOrGetClassesinReg False
End Sub

Private Sub CmdControlVideo_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
hWndVideo = GetHwndByClass(txtClass(0), txtClass(1), txtClass(2), txtClass(3), txtClass(4), txtClass(5))
If hWndVideo = 0 Then Exit Sub
'MsgBox hWndVideo
SizeVideo = GetSize(hWndVideo)
TxtHight = SizeVideo.IHight
TxtWidth = SizeVideo.IWidth

HScrollSeek.Value = 0
OpenVideo ListFiles.List(NFile), "video", "Mpeg", hWndVideo, TxtWidth, TxtHight, 0, 0
PlayVideo "None", "None"
TimerVideo.Enabled = True
HScrollSeek.Max = GetTotalFrames
SaveOrGetClassesinReg True   'ÍÝÙ ØÑíÞ ãÞÈÖ ÇáäÇÝÐÉ
Case 1
PlayVideo "None", "None"

Case 2
CloseVideo
Case 3
PauseVideo
Case 4
NextMpeg
Case 5
BackMpeg
End Select
End Sub

Private Sub CmdDeletePlace_Click()
On Error Resume Next
If ComboPlaces.Text = "" Then MsgBox "Please First Select a number from the list", vbInformation: Exit Sub
For i = 0 To 6
ListClasses(i).RemoveItem Val(ComboPlaces.Text)
txtClass(i) = ""
DoEvents
Next i
ComboPlaces.RemoveItem Val(ComboPlaces.Text)
ComboPlaces.Refresh

End Sub

Private Sub CmdEditPlace_Click()
On Error Resume Next
If ComboPlaces.Text = "" Then MsgBox "Please First Select a number from the list", vbInformation: Exit Sub

AddingNewClassOrEditing "Edit"

End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdHFCV_Click()
FrameControlVideo.Visible = False
FrameMain.Visible = True

End Sub

Private Sub CmdHideFrameSelectFile_Click()
FrameMain.Visible = True
FrameSelectPath.Visible = False
SaveFileNameMpeg
End Sub


Private Sub CmdMinimizeWindow_Click()
Me.Hide
End Sub

Private Sub CmdRemoveFileFromList_Click()
ListFiles_DblClick
End Sub

Private Sub CmdSave_Click()
On Error Resume Next

Kill PathFileClasses
Dim ClassName As ClassV

'////////////////// //////////////////////////////////////


Open PathFileClasses For Random As #1

For n = 0 To ListClasses(0).ListCount - 1 'ÚÏÏ ÇáãÑÇÊ ÇáãÓÌáÉ

For i = 0 To 6
ClassName.Class(i) = ListClasses(i).List(n)
'MsgBox ClassName.Class(i)
DoEvents
Next i

Put #1, n + 1, ClassName

DoEvents

Next n




'///////////////////////////////////////////////////////
End Sub





Private Sub CmdShowFCV_Click()
FrameControlVideo.Visible = True
FrameMain.Visible = False
End Sub



Private Sub CmdUpdate_Click()
If AddingNew = True Then
For i = 0 To 6
txtClass_LostFocus (i)
DoEvents
Next i

ComboPlaces.AddItem ListClasses(0).ListCount

For n = 0 To 6
ListClasses(n).AddItem txtClass(n)
DoEvents
Next n
AddingNew = False
End If

If Editing = True Then
For n = 0 To 6
ListClasses(n).AddItem txtClass(n), ComboPlaces.Text
DoEvents
Next n
Editing = False
End If

EnableCMD
End Sub

Private Sub ComboPlaces_Click()
For i = 0 To 6
txtClass(i).Text = ListClasses(i).List(Val(ComboPlaces.Text))
DoEvents
Next i
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dim StoreDirive As Variant
StoreDirive = Dir1.Path
On Error Resume Next
Dir1.Path = Drive1.Drive
If Err = 68 Then Drive1.Drive = StoreDirive

End Sub

Private Sub Form_Load()

   With abd
      .cbSize = Len(abd)
      .hwnd = Me.hwnd
      .uId = vbNull
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      .uCallBackMessage = Mouse_Move
      .hIcon = Me.Icon
      .szTip = App.Title & vbNullChar
   End With
   Shell_NotifyIcon NIM_ADD, abd

CommandOnBeginProgram

ReadFilesMpeg



CreateShortcut

'/////////////////////
If ListFiles.List(0) = "" Then ListFiles.AddItem "Spears.mpg"
If txtClass(0).Text = "" Then txtClass(0) = "Progman": txtClass(1) = "SHELLDLL_DefView": txtClass(2) = "SysListView32": txtClass(3) = 0: txtClass(4) = 0: txtClass(5) = 0: TimerCheckhWnd.Enabled = True: ChPlayAutomaticOnStart.Value = 1


'////////////////////

If Not txtClass(0).Text = "" Or Not txtClass(0).Text = "0" Then Hide
hWndVideo = GetHwndByClass(txtClass(0), txtClass(1), txtClass(2), txtClass(3), txtClass(4), txtClass(5))
If hWndVideo = 0 Then TimerCheckhWnd.Enabled = True
If Not hWndVideo = 0 Then If ChPlayAutomaticOnStart.Value = 1 Then CmdControlVideo_Click (0)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result As Long
Dim msg As Long
If Me.ScaleMode = vbPixels Then
     msg = X
Else
     msg = X / Screen.TwipsPerPixelX
End If

Select Case msg
             Case Mouse_Right_Down
            'Case mouse right down
            
            Case Mouse_Right_Click
            'Case mouse right click
            Me.PopupMenu MnuMain
            'Me.PopupMenu MnuMain
            Case Mouse_Right_DbClick
            'Case mouse right DbClick
 '           NextMpeg
            Case Mouse_Left_Down
            'Case mouse left Dwon
           
            Case Mouse_Left_Click
            'Case mouse left click
            Case Mouse_Left_DbClick
            'Case mouse left dbclick
            Me.WindowState = vbNormal
            Me.Show
            Case Mouse_Button_Down
            'Case mouse Button Down
            
            Case Mouse_Button_Click
            'Case mouse Button click
            Case Mouse_Button_DbClick
            'Case mouse Button Dbclick
            NextMpeg
    End Select
        
            
   
End Sub

Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then Me.Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)
CmdSave_Click
Shell_NotifyIcon NIM_DELETE, abd
CloseVideo
End: End
End Sub





'================================================================================================
'================================================================================================
Sub RunFirstFileMpeg()
LbCountVideos.Caption = "Count Videos: " & ListFiles.ListCount
Moves.FullScreenMode = True
Moves.filename = ListFiles.List(0)
End Sub
Sub NextMpeg()
Dim CountM As Long
CountM = ListFiles.ListCount - 1
NFile = NFile + 1
If NFile > CountM Then NFile = 0
CloseVideo
CmdControlVideo_Click (0)
End Sub
Sub BackMpeg()
Dim CountM As Long
CountM = ListFiles.ListCount - 1
NFile = NFile - 1
If NFile < 0 Then NFile = CountM
CloseVideo
CmdControlVideo_Click (0)
End Sub
Sub ReadFilesMpeg()
Dim NameF As String
On Error Resume Next
Open "c:\Video.ini" For Input As #1
If Err = 53 Then GoTo ends
For i = 0 To 2
If i = 1 Then i = 0
Input #1, NameF
If Err = 62 And NameF = "" Then GoTo ends
If Err = 62 Then Exit For
ListFiles.AddItem NameF
Next i
Close #1

Exit Sub
ends:
Close #1
FrameControl.Visible = True
Moves.Visible = False
End Sub
Sub SaveFileNameMpeg()
Open "c:\Video.ini" For Output As #1
For i = 0 To ListFiles.ListCount - 1
Print #1, ListFiles.List(i)
DoEvents
Next i
Close #1
End Sub
Sub ClearFileFromList()
'On Error Resume Next
For a = ListFiles.ListCount - 1 To 0 Step -1
If ListFiles.Selected(a) = True Then
ListFiles.RemoveItem a
End If
DoEvents
Next a
LbCountVideos.Caption = "Count Videos: " & ListFiles.ListCount
End Sub

Sub AddFileToListBox()
For a = 0 To File1.ListCount - 1
If File1.Selected(a) = True Then
b = File1.List(a)
ListFiles.AddItem File1.Path & "\" & b
End If
DoEvents
Next a
LbCountVideos.Caption = "Count Videos: " & ListFiles.ListCount
End Sub
Private Sub File1_DblClick()
AddFileToListBox
End Sub


Private Sub Frame1_DblClick()
FrameMain.Visible = True
Frame1.Visible = False

End Sub

Private Sub FrameMain_DblClick()
'FrameMain.Visible = False
'Frame1.Visible = True

End Sub

Private Sub HScrollSeek_Change()
HScrollSeek_Scroll
End Sub

Private Sub HScrollSeek_Scroll()
SeekTo HScrollSeek.Value
PlayVideo "None", "None"
End Sub

Private Sub ListFiles_DblClick()
ClearFileFromList
End Sub

Sub CommandOnBeginProgram()
On Error Resume Next
GetPlaceMe
ChRunWithSystem.Value = GetSetting(App.Title, App.Title, "RunWithSystem", 1)
If ChRunWithSystem.Value = 1 Then ChRunWithSystem_Click 'ÅÐÇ Úáì ÚáÇãÉ ÊÔÛíá ãÚ ÇáäÙÇã ÕÍ äßÊÈ ãÑÉ ÃÎÑì ÇáãÓÇÑ Ýí ãÍÑÑ ÇáÊÓÌíá
ChPlayNextVideoAutomatic = GetSetting(App.Title, App.Title, "PlayNextVideoAutomatic", 1)
ChUseOrginalWindowSize.Value = GetSetting(App.Title, App.Title, "UseOrginalWindowSize", 1)
ChPlayAutomaticOnStart.Value = GetSetting(App.Title, App.Title, "PlayAutomaticOnStart", 0)

SaveOrGetClassesinReg False


If Len(App.Path) > 3 Then
PathFileClasses = App.Path & "\" & "Classes" & ".dat"
Else
PathFileClasses = App.Path & "Classes" & ".dat"
End If
ReadFileClasses
End Sub


Private Sub MnuMainBack_Click()
BackMpeg
End Sub

Private Sub MnuMainClose_Click()
CmdExit_Click
End Sub

Private Sub MnuMainHide_Click()
Me.Hide
End Sub

Private Sub MnuMainNext_Click()
NextMpeg
End Sub

Private Sub MnuMainShow_Click()
Me.WindowState = vbNormal
Me.Show
End Sub

Private Sub TimerCheckhWnd_Timer()
hWndVideo = GetHwndByClass(txtClass(0), txtClass(1), txtClass(2), txtClass(3), txtClass(4), txtClass(5))
If Not hWndVideo = 0 And Not txtClass(0).Enabled = True Then TimerCheckhWnd.Enabled = False: CmdControlVideo_Click (0)

End Sub

Private Sub TimerVideo_Timer()

'If hWndVideo = 0 Then Exit Sub
'MsgBox hWndVideo

Stats = GetVideoStats

LbStatsVideo.Caption = Stats
If LbStatsVideo.Caption = "" Then TimerCheckhWnd.Enabled = True: CmdControlVideo_Click (2)
If ChPlayNextVideoAutomatic.Value = 1 Then If LbStatsVideo.Caption = "stopped" Then NextMpeg

If ChUseOrginalWindowSize.Value = 1 Then
SizeVideo = GetSize(hWndVideo)
TxtHight = SizeVideo.IHight
TxtWidth = SizeVideo.IWidth
TxtLeft = 0: TxtTop = 0

End If

ReSizeVideo Val(TxtWidth), Val(TxtHight), Val(TxtTop), Val(TxtLeft)
'Dim Stats As String
End Sub

Private Sub txtClass_LostFocus(Index As Integer)
If Index = 6 Then If txtClass(6).Text = "" Then txtClass(6).Text = "abd"


If txtClass(Index).Text = "" Then txtClass(Index).Text = 0
 
End Sub

Sub ReadFileClasses()
On Error Resume Next
Dim ClassName As ClassV

'////////////////// //////////////////////////////////////


Open PathFileClasses For Random As #1
For i = 1 To 50


Get #1, i, ClassName
If ClassName.Class(0) = "" Then Close #1: Exit For
DoEvents

For n = 0 To 6
ListClasses(n).AddItem ClassName.Class(n)
DoEvents
Next n

ComboPlaces.AddItem i - 1




Next i

Close #1

'///////////////////////////////////////////////////////
End Sub

Sub AddingNewClassOrEditing(AddOrEdit As String)
If AddOrEdit = "Add" Then
AddingNew = True
For ic = 0 To 6
txtClass(ic).Enabled = True
txtClass(ic).Text = ""
DoEvents
Next ic
Else
For ic = 0 To 6
txtClass(ic).Enabled = True
DoEvents
Next ic

Editing = True
End If

CmdCancel.Enabled = True
CmdUpdate.Enabled = True
ComboPlaces.Enabled = False
CmdAddPlace.Enabled = False
CmdEditPlace.Enabled = False
CmdSave.Enabled = False
CmdDeletePlace.Enabled = False
txtClass(0).SetFocus
CloseVideo
TimerVideo.Enabled = False
TimerCheckhWnd.Enabled = False
End Sub


Sub EnableCMD()

For ic = 0 To 6
txtClass(ic).Enabled = False
DoEvents
Next ic
CmdUpdate.Enabled = False
CmdCancel.Enabled = False
ComboPlaces.Enabled = True
CmdAddPlace.Enabled = True
CmdEditPlace.Enabled = True
CmdSave.Enabled = True
CmdDeletePlace.Enabled = True

TimerCheckhWnd.Enabled = True 'áÊÔÛíá ÇáÝíÏíæ ÈÚÏ ÇáÅäÊåÇÁ ãä ÇáÅÖÇÝÉ

End Sub


Sub SaveOrGetClassesinReg(Save As Boolean)
If Save = True Then
For ic = 0 To 6
SaveSetting App.Title, App.Title, "txtClass" & ic, txtClass(ic)
DoEvents
Next ic

Else

For ic = 0 To 6
txtClass(ic) = GetSetting(App.Title, App.Title, "txtClass" & ic, "")
DoEvents
Next ic

End If

End Sub

Sub CreateShortcut()
On Error Resume Next
Dim ResultCS As Long
ResultCS = GetSetting(App.Title, App.Title, "CreatedShortcut", 0)
If Not ResultCS = 0 Then Exit Sub

SaveSetting App.Title, App.Title, "CreatedShortcut", 1
Dim Grp$, grppath$, a$, prog$, Title$
Grp$ = "Viewer Video"  'áÅÖÇÝÉ ãÌãæÚÉ Åáì ãÏíÑ ÇáÈÑÇãÌ
FrmMain.Label1(0).LinkTopic = "progman|progman"
FrmMain.Label1(0).LinkMode = 2
FrmMain.Label1(0).LinkTimeout = 100
FrmMain.Label1(0).LinkExecute "[creategroup(" + Grp$ + Chr$(44) + grppath$ + ")]"
FrmMain.Label1(0).LinkTimeout = 50
FrmMain.Label1(0).LinkMode = 0

a$ = App.Path
prog$ = PathFull
Title$ = App.Title
FrmMain.Label1(0).LinkTopic = "progman|progman"
FrmMain.Label1(0).LinkMode = 2
FrmMain.Label1(0).LinkTimeout = 100
FrmMain.Label1(0).LinkExecute "[additem(" + prog$ + Chr$(44) + Title$ + Chr$(44) + ",,)]"
FrmMain.Label1(0).LinkTimeout = 50
FrmMain.Label1(0).LinkMode = 0
End Sub

