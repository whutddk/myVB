VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form game3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   11175
   ClientLeft      =   1095
   ClientTop       =   1155
   ClientWidth     =   19635
   Icon            =   "game3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "game3.frx":08CA
   ScaleHeight     =   11175
   ScaleWidth      =   19635
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   375
      Left            =   12480
      TabIndex        =   2
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   495
      Left            =   8280
      TabIndex        =   0
      Top             =   8160
      Width           =   2775
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   14160
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
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
      _cx             =   2143
      _cy             =   873
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   35
      Left            =   2640
      Top             =   5280
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   34
      Left            =   9120
      Top             =   2640
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   33
      Left            =   11760
      Top             =   1200
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   32
      Left            =   11880
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   31
      Left            =   12240
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   30
      Left            =   10080
      Top             =   3360
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   29
      Left            =   8040
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   28
      Left            =   5880
      Top             =   3240
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   27
      Left            =   1080
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   26
      Left            =   6480
      Top             =   1680
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   25
      Left            =   5640
      Top             =   7320
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   24
      Left            =   3720
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   23
      Left            =   8520
      Top             =   6120
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   22
      Left            =   5400
      Top             =   2400
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   21
      Left            =   6480
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   20
      Left            =   7680
      Top             =   5160
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   19
      Left            =   9240
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   18
      Left            =   8880
      Top             =   4320
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   17
      Left            =   10200
      Top             =   1920
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   16
      Left            =   3000
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   15
      Left            =   2280
      Top             =   8040
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   14
      Left            =   13200
      Top             =   4440
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   13
      Left            =   11760
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   12
      Left            =   6960
      Top             =   4440
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   11
      Left            =   9720
      Top             =   6120
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   10
      Left            =   2040
      Top             =   6840
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   9
      Left            =   11760
      Top             =   7320
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   8
      Left            =   7680
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   7
      Left            =   5520
      Top             =   4440
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   6
      Left            =   6960
      Top             =   7680
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   5
      Left            =   4200
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   4
      Left            =   6960
      Top             =   480
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   3
      Left            =   10200
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   2
      Left            =   5160
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   1
      Left            =   960
      Top             =   7680
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   0
      Left            =   4440
      Top             =   8040
      Width           =   735
   End
   Begin VB.Image choose 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   5760
      Top             =   600
      Width           =   735
   End
   Begin VB.Image messioncomplete 
      Height          =   735
      Left            =   240
      Top             =   360
      Width           =   255
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   35
      Left            =   4200
      Top             =   4200
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   34
      Left            =   3480
      Top             =   4200
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   33
      Left            =   2760
      Top             =   4200
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   32
      Left            =   2040
      Top             =   4200
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   29
      Left            =   4200
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   28
      Left            =   3480
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   27
      Left            =   2760
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   26
      Left            =   2040
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   23
      Left            =   4200
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   22
      Left            =   3480
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   21
      Left            =   2760
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   20
      Left            =   2040
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   17
      Left            =   4200
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   16
      Left            =   3480
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   15
      Left            =   2760
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   31
      Left            =   1320
      Top             =   4200
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   25
      Left            =   1320
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   19
      Left            =   1320
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   13
      Left            =   1320
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   11
      Left            =   4200
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   10
      Left            =   3480
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   9
      Left            =   2760
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   8
      Left            =   2040
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   7
      Left            =   1320
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   14
      Left            =   2040
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   30
      Left            =   600
      Top             =   4200
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   24
      Left            =   600
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   18
      Left            =   600
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   12
      Left            =   600
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   6
      Left            =   600
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   5
      Left            =   4200
      Top             =   600
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   4
      Left            =   3480
      Top             =   600
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   3
      Left            =   2760
      Top             =   600
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   2
      Left            =   2040
      Top             =   600
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   1
      Left            =   1320
      Top             =   600
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   0
      Left            =   600
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "game3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'����һ��ƴͼ��Ϸ����Ҫǰ�ڴ�����Ƭ��ͼ��������Ӧ��С��ͼƬ�鲢������������ͼ�������磩

Option Explicit
Private wfull(35) As Boolean    '�����������������ж�ͼƬ���Ƿ��Ѿ�����ͼƬ
Private cfull(35) As Boolean
Private r As Integer

Private Sub Command2_Click()    '�˳�ť�����Ǻ���
Load mainmeun
Unload Me
mainmeun.Show
End Sub

Private Sub Form_Initialize()
Dim i As Integer
For i = 0 To 35
cfull(i) = True                 'ѡ�����Ѿ�����ͼ��������Ϊ��
wfull(i) = False                '������δ����ͼ��������Ϊ��
Next i
r = 37                          '����Ѿ�ѡ���ͼƬ����
End Sub

Private Sub Form_Load()
Command1.Enabled = False
Dim i As Integer
For i = 0 To 35                 '����ѭ��
cpart(i).Picture = LoadPicture(App.Path & "\data\2picture\choose\" & i + 1 & ".jpg")        '��֮ǰ׼���õ�ͼƬ�����ѡ����
Next i
WindowsMediaPlayer1.URL = App.Path & "\data\music\��մ�.mp3"                               '�������֣��������磩�����Ǻ���
WindowsMediaPlayer1.Controls.play
End Sub

Private Sub cpart_Click(Index As Integer)           'ѡ��ͼƬ��ʱ
If r = 37 Then                                      '�ж��Ƿ��Ѿ�ѡ��
r = Index
    If cfull(r) = True Then
    choose.Picture = cpart(r).Picture               '��ѡ������ͼƬ�������ʾ��
    cpart(r).Picture = LoadPicture()                'ȥ��ѡ������ͼƬ��
    cfull(r) = False                                '�޸Ĳ�����
    Else
    cpart(r).Picture = choose.Picture
    choose.Picture = LoadPicture()
    r = 37
    End If
End If
End Sub

Private Sub Image1_Click(Index As Integer)          '���빤����
If r < 37 Then
Dim k As Integer
k = Index
    If wfull(k) = False Then                        '�жϹ������Ƿ��Ѿ�����ͼƬ��
        If k = r Then                               '�ж�ѡ���Ƿ���ȷ
        Image1(k).Picture = choose.Picture          '��ʾ
        choose.Picture = LoadPicture()
        wfull(k) = True                             '�޸Ĳ�����
        r = 37
        Dim c As Integer
        For c = 0 To 35
            If wfull(c) = False Then                '�ж��Ƿ��������ƴͼ����������ť����
            Exit Sub
            End If
        Next c
        Command1.Enabled = True                     '���ť
        Else
        cpart(r).Picture = choose.Picture           'ѡ����󣬰�ͼƬ��Ż�ѡ����
        choose.Picture = LoadPicture()
        cfull(r) = True
        r = 37
        End If
    End If
End If
End Sub

Private Sub Command1_Click()                        '���ť����ɺ���ʾ�޷�ԭͼ
Command1.Visible = False
choose.Visible = False
Dim i As Integer
For i = 0 To 35
cpart(i).Visible = False
Next i
messioncomplete.Picture = LoadPicture(App.Path & "\data\2picture\ԭ��Сͼ.jpg")
Command2.Caption = "���"
End Sub

