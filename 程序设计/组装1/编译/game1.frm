VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form game1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12240
   Icon            =   "game1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   12240
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "投降"
      Height          =   495
      Left            =   9840
      TabIndex        =   8
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "开始/重来"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   5
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      TabIndex        =   4
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   3
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   2
      Top             =   2520
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.Label Label1 
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      TabIndex        =   7
      Top             =   2280
      Width           =   1935
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
      Left            =   9120
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   240
      Top             =   1080
      Width           =   2295
   End
End
Attribute VB_Name = "game1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'这个游戏程序的设计思路来源于我小学时奥赛书上的一道题，原题如下：
    '两人按自然数顺序轮流报数，每人每次只能报1个或2个数。比如第1个人可以报1，第2个人可以报2或2、3；
    '第1个人也可以报1、2，第二人可以报3、4.这样继续下去，谁报到30，谁就胜。请问谁有必胜的策略？
    '解：后报者有必胜策略。如果第1个人报1，第2个人就报2、3；如果第1个人报1、2，第2个人就报3.
    '接着，当第1个人报一个数时，第2个人就报两个数，使第2个人始终报3的倍数。这样后报数者必胜。
    '如果交换报数顺序，但对方未掌握必胜策略，那么，第1个报数的人一旦抓住机会报出3的倍数，先报数者也能稳操胜券
    '(来自小学奥赛书）
'小学初中一直拿这个来捉弄同学，上了高中，知道30很容易被看出破绽，于是把30改成300

'本游戏仅供同学间娱乐

Option Explicit
Dim filename3 As String
Dim number As Integer

Private Sub Command4_Click()    '退出按钮
Load mainmeun
Unload Me
mainmeun.Show
End Sub

Private Sub Form_Initialize()   '游戏记录，不是核心
filename3 = App.Path & "\data\user\record.txt"
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture()  '加载玩家相片，不是核心未加入地址
Image2.Picture = LoadPicture()
number = 0                      '初始化
Command1.Enabled = False        '两个选择钮不可用
Command2.Enabled = False
End Sub

Private Sub computer_rndstep()  '计算机随机选数过程，当玩家选对时跳转
Dim step As Single
Randomize                       '取随机数
step = Rnd * 10
If step > 5 Then                '随机选数
number = number + 1
Else
number = number + 2
End If
Command1.Caption = number + 1   '调整按钮选项
Command2.Caption = number + 2
Frame3.Caption = number
Call recording                  '记录
End Sub

Private Sub computer_winstep()   '计算机胜利选数过程，当玩家选错时跳转
Dim gostep As Integer
gostep = number Mod 3           '计算应选的数
number = number + 3 - gostep
Frame3.Caption = number
If number = 300 Then            '判断是否胜利
Command1.Enabled = False        '如果是，按钮不可用
Command2.Enabled = False
Exit Sub
End If
Command1.Caption = number + 1   '调整按钮选项
Command2.Caption = number + 2
End Sub

Private Sub recording()         '记录，不是核心
Open filename3 For Append As #3 Len = 2
Print #3, number
Close #3
End Sub

Private Sub Command1_Click()    '玩家的选择
number = number + 1
Frame2.Caption = number
Call recording
If number = 300 Then            '判断玩家是否胜利
Frame1.Caption = "你赢了"
'grecord2(usernum) = True

Load game2                      '加载下一个游戏，不是核心
game2.Show
Unload Me
Else
    If number Mod 3 = 0 Then    '判断玩家是否选错
    Call computer_rndstep
    Else
    Call computer_winstep
    End If
End If

End Sub

Private Sub Command2_Click()    '同上
number = number + 2
Frame2.Caption = number
Call recording
If number = 300 Then
Frame1.Caption = "你赢了"
'grecord2(usernum) = True
Load game2
game2.Show
Unload Me
Else
    If number Mod 3 = 0 Then
    Call computer_rndstep
    Else
    Call computer_winstep
    End If
End If
End Sub

Private Sub Command3_Click()    '初始化，复位
number = 0
Randomize
Dim coin As Single              '取随机数
coin = Rnd * 10
Dim ret As Integer
If coin > 5 Then                '取一半概率判断先行，要坑同学可以修改
ret = MsgBox("看你太菜，你先行！", 0)
Command1.Caption = number + 1
Command2.Caption = number + 2
Else
ret = MsgBox("我先上", 0)
Call computer_rndstep
End If
Command1.Enabled = True         '按钮可用
Command2.Enabled = True
WindowsMediaPlayer1.URL = App.Path & "\data\music\蓝色天空.mp3"     '播放背景音乐（来自网络)，不是核心
WindowsMediaPlayer1.Controls.play
End Sub

'事实证明，30改300是个馊主意，与其让我从1到300算3的倍数，不如让我从1写到300。这个程序我仅测试了3次。不过我同学好像没几个受得了。
