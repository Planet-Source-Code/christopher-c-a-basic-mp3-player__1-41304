VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "WinSound"
   ClientHeight    =   6990
   ClientLeft      =   3000
   ClientTop       =   2580
   ClientWidth     =   7710
   Height          =   7680
   Left            =   2940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   Top             =   1950
   Width           =   7830
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Volume"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   2280
      TabIndex        =   26
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton cmdVolDecrease 
         Caption         =   "Decrease"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdVolIncrease 
         Caption         =   "Increase"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   41.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1560
         TabIndex        =   28
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblVol 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "75"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   41.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Find Music"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6000
      Top             =   240
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete File"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "File To Playlist"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play File"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   5400
      Width           =   1455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   11
      Top             =   4920
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   4440
      TabIndex        =   17
      Top             =   4920
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1830
      Left            =   1800
      TabIndex        =   16
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   1680
      Width           =   6135
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1980
      ItemData        =   "Form1.frx":0000
      Left            =   240
      List            =   "Form1.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   2040
      Width           =   7215
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "File List"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      TabIndex        =   18
      Top             =   4680
      Width           =   7455
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Play List"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Width           =   7455
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Single Play"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton opCont 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Continuous Play"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   10
         Top             =   2640
         Width           =   1695
      End
      Begin VB.OptionButton opRepeat 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Repeat Play"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Selected:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Now Playing"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   2055
      Begin VB.Label lblMin 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   41.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   41.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   24
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblSec2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   41.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   840
         TabIndex        =   23
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblSec 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   41.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1320
         TabIndex        =   22
         Top             =   120
         Width           =   615
      End
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   3375
      Left            =   7440
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   4290
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   
      DisplayForeColor=   
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   1
      VideoBorderColor=   
      VideoBorder3D   =   0   'False
      Volume          =   -1500
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuScan 
         Caption         =   "Scan For Music"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuControls 
      Caption         =   "Controls"
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
      End
   End
   Begin VB.Menu mnuPlayOp 
      Caption         =   "Play Options"
      Begin VB.Menu mnuSinglePlay 
         Caption         =   "Single Play"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRepeatPlay 
         Caption         =   "Repeat Play"
      End
      Begin VB.Menu mnuContPlay 
         Caption         =   "Continuous Play"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      Begin VB.Menu mnuAbout2 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim allowplay As String
Dim paused As Boolean
Dim allowpause As Boolean





Private Sub cmdPause_Click()
If allowpause = True Then
On Error Resume Next
If paused = False Then
MediaPlayer1.Pause
paused = True
allowplay = "no"
Exit Sub
End If

If paused = True Then
MediaPlayer1.Play
paused = False
allowplay = "yes"
End If

End If
End Sub


Private Sub cmdPlay_Click()
If paused = True Then
MediaPlayer1.Play
paused = False
allowplay = "yes"
Exit Sub
End If
If paused = False Then
MediaPlayer1.Open Text1.Text
lblSec = "0"
lblMin = "0"
lblSec2 = "0"
Exit Sub
End If
End Sub


Private Sub cmdScan_Click()
frmScan.Visible = True
frmScan.Show
End Sub

Private Sub cmdStop_Click()
MediaPlayer1.Stop
allowplay = "no"
lblSec = "0"
lblMin = "0"
lblSec2 = "0"
allowpause = False
End Sub

Private Sub cmdVolDecrease_Click()
If lblVol > 0 Then
lblVol = lblVol - 5
End If
Select Case lblVol
Case "100"
MediaPlayer1.Volume = 0
Exit Sub
Case "95"
MediaPlayer1.Volume = -300
Exit Sub
Case "90"
MediaPlayer1.Volume = -600
Exit Sub
Case "85"
MediaPlayer1.Volume = -900
Exit Sub
Case "80"
MediaPlayer1.Volume = -1200
Exit Sub
Case "75"
MediaPlayer1.Volume = -1500
Exit Sub
Case "70"
MediaPlayer1.Volume = -1800
Exit Sub
Case "65"
MediaPlayer1.Volume = -2100
Exit Sub
Case "60"
MediaPlayer1.Volume = -2400
Exit Sub
Case "55"
MediaPlayer1.Volume = -2700
Exit Sub
Case "50"
MediaPlayer1.Volume = -3000
Exit Sub
Case "45"
MediaPlayer1.Volume = -3300
Exit Sub
Case "40"
MediaPlayer1.Volume = -3600
Exit Sub
Case "35"
MediaPlayer1.Volume = -3900
Exit Sub
Case "30"
MediaPlayer1.Volume = -4200
Exit Sub
Case "25"
MediaPlayer1.Volume = -4500
Exit Sub
Case "20"
MediaPlayer1.Volume = -4800
Exit Sub
Case "15"
MediaPlayer1.Volume = -5100
Exit Sub
Case "10"
MediaPlayer1.Volume = -5400
Exit Sub
Case "5"
MediaPlayer1.Volume = -5700
Exit Sub
Case "0"
MediaPlayer1.Volume = -6000
Exit Sub
End Select
End Sub

Private Sub cmdVolIncrease_Click()
If lblVol < 100 Then
lblVol = lblVol + 5
End If
Select Case lblVol
Case "100"
MediaPlayer1.Volume = 0
Exit Sub
Case "95"
MediaPlayer1.Volume = -300
Exit Sub
Case "90"
MediaPlayer1.Volume = -600
Exit Sub
Case "85"
MediaPlayer1.Volume = -900
Exit Sub
Case "80"
MediaPlayer1.Volume = -1200
Exit Sub
Case "75"
MediaPlayer1.Volume = -1500
Exit Sub
Case "70"
MediaPlayer1.Volume = -1800
Exit Sub
Case "65"
MediaPlayer1.Volume = -2100
Exit Sub
Case "60"
MediaPlayer1.Volume = -2400
Exit Sub
Case "55"
MediaPlayer1.Volume = -2700
Exit Sub
Case "50"
MediaPlayer1.Volume = -3000
Exit Sub
Case "45"
MediaPlayer1.Volume = -3300
Exit Sub
Case "40"
MediaPlayer1.Volume = -3600
Exit Sub
Case "35"
MediaPlayer1.Volume = -3900
Exit Sub
Case "30"
MediaPlayer1.Volume = -4200
Exit Sub
Case "25"
MediaPlayer1.Volume = -4500
Exit Sub
Case "20"
MediaPlayer1.Volume = -4800
Exit Sub
Case "15"
MediaPlayer1.Volume = -5100
Exit Sub
Case "10"
MediaPlayer1.Volume = -5400
Exit Sub
Case "5"
MediaPlayer1.Volume = -5700
Exit Sub
Case "0"
MediaPlayer1.Volume = -6000
Exit Sub
End Select
End Sub

Private Sub Command1_Click()
MediaPlayer1.Open File1.path & "\" & File1.FileName
End Sub



Private Sub Command2_Click()
oldsongs = ""
List1.AddItem File1.path & "\" & File1.FileName
newsong = File1.path & "\" & File1.FileName
On Error Resume Next
Open "c:/PlayList.txt" For Append As #1
Print #1, "" & newsong & ""
Close #1
End Sub





Private Sub Command3_Click()
Kill File1.path & "\" & File1.FileName
File1.Refresh
End Sub



Private Sub Dir1_Change()
File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
End Sub

Private Sub Form_Load()
allowpause = False
Dir1.path = Drive1.Drive
paused = False
On Error Resume Next
If FileLen("c:/PlayList.txt") = 0 Then
Open "c:/PlayList.txt" For Output As #1
Close #1
End If

Open "c:/PlayList.txt" For Input As #1
Do Until EOF(1)
Input #1, playlistitem
List1.AddItem playlistitem
Loop
Close #1
allowplay = "no"
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmScan.Visible = False
End Sub


Private Sub List1_Click()
Text1.Text = List1.Text
End Sub


Private Sub List1_DblClick()

Text1.Text = List1.Text
MediaPlayer1.Stop
lblSec = "0"
lblMin = "0"
lblSec2 = "0"
MediaPlayer1.Open Text1.Text
 

End Sub


Private Sub List1_KeyPress(KeyAscii As Integer)
Text1.Text = List1.Text
End Sub





Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
allowplay = "no"
lblSec = "0"
lblSec2 = "0"
lblMin = "0"
allowpause = False
If opCont.Value = True Then
On Error GoTo error1
allowpause = True
List1.ListIndex = List1.ListIndex + 1
MediaPlayer1.Open Text1.Text
Exit Sub
error1:
List1.ListIndex = 0
MediaPlayer1.Open Text1.Text
End If

If opRepeat.Value = True Then
allowpause = True
MediaPlayer1.Open Text1.Text
End If

End Sub


Private Sub MediaPlayer1_NewStream()
lblSec = "0"
lblSec2 = "0"
lblMin = "0"
allowplay = "yes"
allowpause = True
End Sub


Private Sub mnuAbout2_Click()
If allowplay = "yes" Then
MediaPlayer1.Pause
MsgBox "WinSound Developed By Curran Tech.", vbInformation + vbOKOnly, "About"
MediaPlayer1.Play
Else
MsgBox "WinSound Developed By Curran Tech.", vbInformation + vbOKOnly, "About"
End If
End Sub

Private Sub mnuContPlay_Click()
opCont.Value = True
mnuSinglePlay.Checked = False
mnuRepeatPlay.Checked = False
mnuContPlay.Checked = True
End Sub

Private Sub mnuFileExit_Click()
End
End Sub

Private Sub mnuPause_Click()
If allowpause = True Then
On Error Resume Next
If paused = False Then
MediaPlayer1.Pause
paused = True
allowplay = "no"
Exit Sub
End If

If paused = True Then
MediaPlayer1.Play
paused = False
allowplay = "yes"
End If

End If
End Sub

Private Sub mnuPlay_Click()
If paused = True Then
MediaPlayer1.Play
paused = False
allowplay = "yes"
Exit Sub
End If
If paused = False Then
MediaPlayer1.Open Text1.Text
lblSec = "0"
lblMin = "0"
lblSec2 = "0"
Exit Sub
End If
End Sub

Private Sub mnuRepeatPlay_Click()
opRepeat.Value = True
mnuSinglePlay.Checked = False
mnuRepeatPlay.Checked = True
mnuContPlay.Checked = False
End Sub

Private Sub mnuScan_Click()
frmScan.Visible = True
frmScan.Show
End Sub

Private Sub mnuSinglePlay_Click()
Option1.Value = True
mnuSinglePlay.Checked = True
mnuRepeatPlay.Checked = False
mnuContPlay.Checked = False
End Sub

Private Sub mnuStop_Click()
MediaPlayer1.Stop
allowplay = "no"
lblSec = "0"
lblMin = "0"
lblSec2 = "0"
allowpause = False
End Sub

Private Sub opCont_Click()
mnuSinglePlay.Checked = False
mnuRepeatPlay.Checked = False
mnuContPlay.Checked = True
End Sub

Private Sub opRepeat_Click()
mnuSinglePlay.Checked = False
mnuRepeatPlay.Checked = True
mnuContPlay.Checked = False
End Sub

Private Sub Option1_Click()
mnuSinglePlay.Checked = True
mnuRepeatPlay.Checked = False
mnuContPlay.Checked = False
End Sub

Private Sub Timer1_Timer()
If allowplay = "yes" Then
If lblSec = "9" Then
lblSec = "0"
lblSec2 = lblSec2 + 1
If lblSec2 = "6" Then
lblSec2 = "0"
lblMin = lblMin + 1
End If
Else
lblSec = lblSec + 1
End If
End If
End Sub


