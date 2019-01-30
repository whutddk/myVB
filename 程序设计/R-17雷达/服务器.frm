VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   10800
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "指令"
      Height          =   1575
      Left            =   2280
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
      Begin VB.CommandButton Command2 
         Caption         =   "指令2"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "指令1"
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "连接"
      Height          =   1575
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "计算机名"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "端口号"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "发送"
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   3000
      Width           =   3375
   End
   Begin VB.ListBox List1 
      Height          =   1320
      Left            =   7320
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   635
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "连接"
      TabPicture(0)   =   "服务器.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "指令"
      TabPicture(1)   =   "服务器.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "文件"
      TabPicture(2)   =   "服务器.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "对话"
      TabPicture(3)   =   "服务器.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
   End
   Begin MSWinsockLib.Winsock radar 
      Index           =   0
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
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
      _cx             =   1296
      _cy             =   661
   End
   Begin VB.Label Label1 
      Caption         =   "状态"
      Height          =   255
      Left            =   7440
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private filename1 As String '发
Private filename2 As String '收
Private sent As Long
Private max_sent As Long
Private arrival As Boolean
Private filedata() As Byte
Private wsindex As Integer
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const thescreen = 0
Const theform = 0


Private Sub Form_Load()
ReDim filedata(1 To 9500)
radar(0).LocalPort = 1111
radar(0).Listen
arrival = False
sent = 0
wsindex = 0
End Sub


Private Sub radar_Close(Index As Integer)
On Error Resume Next
Dim i  As Integer
For i = 1 To wsindex
radar(i).Close
Unload radar(i)
Next i

End Sub

Private Sub radar_ConnectionRequest(Index As Integer, ByVal requestID As Long)
wsindex = wsindex + 1
Load radar(wsindex)
radar(wsindex).Accept requestID
List1.AddItem radar(wsindex).RemoteHost + radar(wsindex).RemoteHostIP
List1.Refresh
End Sub

Private Sub radar_DataArrival(Index As Integer, ByVal bytesTotal As Long)
ReDim filedata(1 To 9500)
Select Case Index
Case 2
On Error Resume Next
Close #2
Call keybd_event(vbKeySnapshot, thescreen, 1, 1)
filename2 = Str(Year(Date) & Month(Date) & Day(Date) & Hour(Time) & Minute(Time) & Second(Time)) '"abc" 'Str(CVar(Now))
radar(2).SendData arrival
WindowsMediaPlayer1.URL = App.Path & "\data\sound\msg.wav"
WindowsMediaPlayer1.Controls.play
Case 4
If arrival = False Then
radar(4).GetData filename2, vbString
arrival = True

Open App.Path & "\data\receive\" & filename2 & ".bmp" For Binary As #2  '收
radar(4).SendData arrival
Else
ReDim filedata(1 To 9500)
radar(4).GetData filedata(), vbByte
Put #2, , filedata()
radar(4).SendData arrival
End If

Case 3
filename1 = Str(Year(Date) & Month(Date) & Day(Date) & Hour(Time) & Minute(Time) & Second(Time)) '
SavePicture Clipboard.GetData(vbCFBitmap), App.Path & "\data\post\" & filename1 & ".bmp"  '发
Open App.Path & "\data\post\" & filename1 & ".bmp" For Binary As #1
max_sent = LOF(1)
radar(3).SendData filename1
arrival = False
Case 6
If sent > max_sent Then
Close #1
radar(6).SendData arrival
arrival = False
Exit Sub
End If
ReDim filedata(1 To 9500)
Get #1, sent + 1, filedata()
sent = sent + 9500
radar(3).SendData filedata()


End Select


End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
Select Case SSTab1.Tab
Case 0

Frame1.Visible = True
Frame2.Visible = False
Command3.Visible = True
Text3.Visible = False
Case 1
Frame1.Visible = False
Frame2.Visible = True
Command3.Visible = False
Text3.Visible = False
Case 2
Frame1.Visible = False
Frame2.Visible = False
Command3.Visible = True
Text3.Visible = True
Case 3
Frame1.Visible = False
Frame2.Visible = False
Command3.Visible = True
Text3.Visible = True
End Select
End Sub






