VERSION 5.00
Begin VB.Form lunch 
   Caption         =   "登录"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "lunch.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "lunch.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   4755
   ScaleWidth      =   13140
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "登录"
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "新建用户"
      Height          =   420
      Left            =   1080
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   540
      Left            =   4560
      TabIndex        =   2
      Top             =   1440
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   540
      ItemData        =   "lunch.frx":1194
      Left            =   240
      List            =   "lunch.frx":1196
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "密码"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "用户"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "lunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recmax As Integer
Dim i As Integer
Dim check() As String


Private Sub Form_Load()
filename1 = App.Path & "\data\user\fileuser.txt"
filenum1 = FreeFile()
Open filename1 For Random As filenum1 Len = 50
recmax = LOF(filenum1) / 50
    For i = 1 To recmax
    ReDim Preserve check(recmax)
    ReDim Preserve grecord1(recmax)
    ReDim Preserve grecord2(recmax)
    ReDim Preserve grecord3(recmax)
    Get #filenum1, i, user
    Combo1.AddItem user.name
    check(i) = user.password
    grecord1(i) = user.gamerecord1
    grecord2(i) = user.gamerecord2
    grecord3(i) = user.gamerecord3
    Next i
    Close filenum1
    Combo1.Refresh
End Sub


Private Sub Command1_Click()
Load Userbuild
Unload lunch
Userbuild.Show

End Sub

Private Sub Command2_Click()
ReDim Preserve check(recmax)
filenum2 = FreeFile()
filename2 = App.Path & "\data\user\time.txt"
Dim launchtime As Date
usernum = Combo1.ListIndex
usernum = usernum + 1
Dim iput As String * 16
iput = Text1.Text
    If iput = check(usernum) Then
    Open filename2 For Append As filenum2 Len = 10
    launchtime = Now
    Print #filenum2, launchtime
    Close #filenum2
    Load mainmeun
    Unload lunch
   mainmeun.Show
    Else
    Text1.Text = ""
    End If
End Sub



