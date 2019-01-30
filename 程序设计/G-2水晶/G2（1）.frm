VERSION 5.00
Begin VB.Form Userbuild 
   Caption         =   "新建用户"
   ClientHeight    =   3240
   ClientLeft      =   7230
   ClientTop       =   4605
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   21.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000011&
   Icon            =   "G2（1）.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   7320
   Begin VB.CommandButton Command2 
      Caption         =   "放弃"
      Height          =   615
      Left            =   1080
      TabIndex        =   7
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确认"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000011&
      Caption         =   "确认密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000011&
      Caption         =   "密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000011&
      Caption         =   "用户名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -120
      TabIndex        =   3
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Userbuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
filename1 = App.Path & "\data\user\fileuser.txt"
filenum1 = FreeFile()
    If Text2.Text = Text3.Text Then
        If Text1.Text <> "" Then
            user.name = Text1.Text
            user.password = Text2.Text
            user.buildtime = Now
            user.gamerecord1 = True
            user.gamerecord2 = False
            user.gamerecord3 = False
            Open filename1 For Random As filenum1 Len = 50 'Len(uesr)
            recordnum = LOF(filenum1) / 50
            recordnum = recordnum + 1
            Put #filenum1, recordnum, user
            Close filenum1
            Load lunch
            Unload Userbuild
            lunch.Show
'Put #filenum, recordnum, user.password
'recordnum = recordnum + 1
'Put #filenum, recordnum, user.buildtime
'recordnum = recordnum + 1
        Else
    Text1.Text = "无名氏"
        End If
    Else
        Text2.Text = ""
        Text3.Text = ""
        
    End If
End Sub



Private Sub Command2_Click()
Load lunch
Unload Userbuild
lunch.Show
End Sub

