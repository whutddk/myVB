VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form game1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   12240
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "重来"
      Height          =   855
      Left            =   5880
      TabIndex        =   5
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   855
      Left            =   3480
      TabIndex        =   4
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   735
      Left            =   1080
      TabIndex        =   3
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   2520
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   2520
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "每次可报1个或两个数，按自然数顺序从1开始报，先到300者为胜"
      ClipControls    =   0   'False
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2640
      TabIndex        =   0
      Top             =   720
      Width           =   6495
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   615
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1085
      _cy             =   450
   End
   Begin VB.Image Image2 
      Height          =   2055
      Left            =   7680
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   960
      Top             =   1200
      Width           =   2295
   End
End
Attribute VB_Name = "game1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim filename3 As String
Dim number As Integer
Private Sub Form_Initialize()
filename3 = App.Path & "\record.txt"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
number = 297
Call compuuter_rndstep
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture()
Image2.Picture = LoadPicture()
number = 0
Command1.Enabled = False
Command2.Enabled = False
End Sub
Private Sub start()
Randomize
Dim coin As Single
coin = Rnd * 10
Dim ret As Integer
If coin > 5 Then
ret = MsgBox("看你太菜，你先行！", 0)
Command1.Caption = number + 1
Command2.Caption = number + 2
Else
ret = MsgBox("我先上", 0)
Call computer_rndstep
End If
Command1.Enabled = True
Command2.Enabled = True
WindowsMediaPlayer1.URL = App.Path & "\data\music\蓝色天空.mp3"
WindowsMediaPlayer1.Controls.play
Command1.Enabled = True
Command2.Enabled = True
End Sub

Private Sub computer_rndstep()
Dim step As Single
Randomize
step = Rnd * 10
If step > 5 Then
number = number + 1

Else
number = number + 2
End If
Command1.Caption = number + 1
Command2.Caption = number + 2
Frame3.Caption = number
Call recording
End Sub

Private Sub computer_winstep()
Dim gostep As Integer
gostep = number Mod 3
number = number + 3 - gostep
Frame3.Caption = number
If number = 300 Then
Command1.Enabled = False
Command2.Enabled = False
Exit Sub
End If
Command1.Caption = number + 1
Command2.Caption = number + 2
End Sub

Private Sub recording()
Open filename3 For Append As #3 Len = 50
Print #3, number
Close #3
End Sub

Private Sub Command1_Click()
number = number + 1
Frame2.Caption = number
Call recording
If number = 300 Then
Frame1.Caption = "你赢了"
grecord2(usernum) = True

Load mainmeun
mainmeun.Show
Unload Me
Else
    If number Mod 3 = 0 Then
    Call computer_rndstep
    Else
    Call computer_winstep
    End If
End If


End Sub

Private Sub Command2_Click()
number = number + 2
If number = 300 Then
Frame1.Caption = "你赢了"
Frame2.Caption = number
Call recording
grecord2(usernum) = True
Load mainmeun
mainmeun.Show
Unload Me
Else
    If number Mod 3 = 0 Then
    Call computer_rndstep
    Else
    Call computer_winstep
    End If
End If

End Sub

Private Sub Command3_Click()
Call start
End Sub
