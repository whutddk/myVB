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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "Ͷ��"
      Height          =   495
      Left            =   9840
      TabIndex        =   8
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "��ʼ/����"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "������"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "ÿ�οɱ�1����������������Ȼ��˳���1��ʼ�����ȵ�300��Ϊʤ"
      ClipControls    =   0   'False
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Name            =   "������"
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
'�����Ϸ��������˼·��Դ����Сѧʱ�������ϵ�һ���⣬ԭ�����£�
    '���˰���Ȼ��˳������������ÿ��ÿ��ֻ�ܱ�1����2�����������1���˿��Ա�1����2���˿��Ա�2��2��3��
    '��1����Ҳ���Ա�1��2���ڶ��˿��Ա�3��4.����������ȥ��˭����30��˭��ʤ������˭�б�ʤ�Ĳ��ԣ�
    '�⣺�����б�ʤ���ԡ������1���˱�1����2���˾ͱ�2��3�������1���˱�1��2����2���˾ͱ�3.
    '���ţ�����1���˱�һ����ʱ����2���˾ͱ���������ʹ��2����ʼ�ձ�3�ı��������������߱�ʤ��
    '�����������˳�򣬵��Է�δ���ձ�ʤ���ԣ���ô����1����������һ��ץס���ᱨ��3�ı������ȱ�����Ҳ���Ȳ�ʤȯ
    '(����Сѧ�����飩
'Сѧ����һֱ�������׽Ūͬѧ�����˸��У�֪��30�����ױ��������������ǰ�30�ĳ�300

'����Ϸ����ͬѧ������

Option Explicit
Dim filename3 As String
Dim number As Integer

Private Sub Command4_Click()    '�˳���ť
Load mainmeun
Unload Me
mainmeun.Show
End Sub

Private Sub Form_Initialize()   '��Ϸ��¼�����Ǻ���
filename3 = App.Path & "\data\user\record.txt"
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture()  '���������Ƭ�����Ǻ���δ�����ַ
Image2.Picture = LoadPicture()
number = 0                      '��ʼ��
Command1.Enabled = False        '����ѡ��ť������
Command2.Enabled = False
End Sub

Private Sub computer_rndstep()  '��������ѡ�����̣������ѡ��ʱ��ת
Dim step As Single
Randomize                       'ȡ�����
step = Rnd * 10
If step > 5 Then                '���ѡ��
number = number + 1
Else
number = number + 2
End If
Command1.Caption = number + 1   '������ťѡ��
Command2.Caption = number + 2
Frame3.Caption = number
Call recording                  '��¼
End Sub

Private Sub computer_winstep()   '�����ʤ��ѡ�����̣������ѡ��ʱ��ת
Dim gostep As Integer
gostep = number Mod 3           '����Ӧѡ����
number = number + 3 - gostep
Frame3.Caption = number
If number = 300 Then            '�ж��Ƿ�ʤ��
Command1.Enabled = False        '����ǣ���ť������
Command2.Enabled = False
Exit Sub
End If
Command1.Caption = number + 1   '������ťѡ��
Command2.Caption = number + 2
End Sub

Private Sub recording()         '��¼�����Ǻ���
Open filename3 For Append As #3 Len = 2
Print #3, number
Close #3
End Sub

Private Sub Command1_Click()    '��ҵ�ѡ��
number = number + 1
Frame2.Caption = number
Call recording
If number = 300 Then            '�ж�����Ƿ�ʤ��
Frame1.Caption = "��Ӯ��"
'grecord2(usernum) = True

Load game2                      '������һ����Ϸ�����Ǻ���
game2.Show
Unload Me
Else
    If number Mod 3 = 0 Then    '�ж�����Ƿ�ѡ��
    Call computer_rndstep
    Else
    Call computer_winstep
    End If
End If

End Sub

Private Sub Command2_Click()    'ͬ��
number = number + 2
Frame2.Caption = number
Call recording
If number = 300 Then
Frame1.Caption = "��Ӯ��"
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

Private Sub Command3_Click()    '��ʼ������λ
number = 0
Randomize
Dim coin As Single              'ȡ�����
coin = Rnd * 10
Dim ret As Integer
If coin > 5 Then                'ȡһ������ж����У�Ҫ��ͬѧ�����޸�
ret = MsgBox("����̫�ˣ������У�", 0)
Command1.Caption = number + 1
Command2.Caption = number + 2
Else
ret = MsgBox("������", 0)
Call computer_rndstep
End If
Command1.Enabled = True         '��ť����
Command2.Enabled = True
WindowsMediaPlayer1.URL = App.Path & "\data\music\��ɫ���.mp3"     '���ű������֣���������)�����Ǻ���
WindowsMediaPlayer1.Controls.play
End Sub

'��ʵ֤����30��300�Ǹ������⣬�������Ҵ�1��300��3�ı������������Ҵ�1д��300����������ҽ�������3�Ρ�������ͬѧ����û�����ܵ��ˡ�
