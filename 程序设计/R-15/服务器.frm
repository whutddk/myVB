VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13290
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   13290
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame4 
      Caption         =   "平均速度"
      Height          =   495
      Left            =   9960
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   9960
      TabIndex        =   11
      Text            =   "100000"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "内存"
      Height          =   255
      Left            =   9960
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1500
      Left            =   840
      TabIndex        =   9
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      Caption         =   "发送文件"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "端口"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   5040
      TabIndex        =   4
      Text            =   "改1.jpg"
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "发送"
      Height          =   495
      Left            =   8160
      TabIndex        =   3
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "断开"
      Height          =   735
      Left            =   5040
      TabIndex        =   2
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "connect"
      Height          =   735
      Left            =   2760
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "侦听"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   5760
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock tcpserver 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   1320
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private wsindex As Integer
Private filename As String
Private filedata() As Byte
Private filelen As Long
Private sent() As Long


Private Sub Form_load()
Text1.Enabled = 1
Command2.Enabled = 0
Command4.Enabled = False
Text2.Text = 0
wsindex = 0

End Sub

Private Sub Command1_Click() '侦听
If Val(Text2.Text) > 0 Then
tcpserver(wsindex).LocalPort = Val(Text2.Text)
Else
tcpserver(wsindex).LocalPort = 2000
End If
Text2.Text = tcpserver(wsindex).LocalPort
tcpserver(wsindex).Bind
tcpserver(wsindex).Listen
Command1.Enabled = 0
Label1.Caption = "listennow"
End Sub

Private Sub tcpserver_ConnectionRequest(index As Integer, ByVal requestID As Long)
wsindex = wsindex + 1
Load tcpserver(wsindex)
Label1.Caption = "requestnow"
Command2.Enabled = 1
tcpserver(wsindex).Accept requestID
List1.AddItem tcpserver(wsindex).RemoteHost + tcpserver(wsindex).RemoteHostIP
List1.Refresh
Command4.Enabled = 1

End Sub

Private Sub tcpserver_Connect(index As Integer)
Label1.Caption = "connectnow"
Text1.Enabled = True
Command4.Enabled = True
End Sub

Private Sub tcpserver_DataArrival(index As Integer, ByVal bytesTotal As Long)
DoEvents
ReDim filedata(1 To index, 1 To 8000)
ReDim sent(2 To wsindex)
If index = 1 Then
Dim i As Integer
Dim j As Integer
For i = 2 To wsindex
    For j = 1 To 8000
    Get #i, , filedata(i, j)
    tcpserver(i).SendData filedata(i, j)
    Next j
sent(i) = 0
Next i
Else





















Get #wsindex, sent(index) + 1, filedata()
sent(index) = sent(index) + 1
tcpserver(index).SendData filedata()
End If
'If sent > filelen Then
'Close #1
'Exit Sub
'End If
'sent = sent + 1

End Sub

Private Sub Command4_click()  '发送
If Text1.Text = "" Then
Exit Sub
Else
End If
filename = Text1.Text
Open App.Path & "\" & filename For Binary As #1
filelen = LOF(1)
Dim copydata() As Byte
Static len_file As Integer
len_file = CInt(filelen / (wsindex - 1))
Dim i As Integer
ReDim copydata(1 To len_file)
For i = 2 To wsindex
Open App.Path & "\" & i For Binary As #i
Get #1, , copydata()
Put #i, , copydata()
Next i
tcpserver(1).SendData filename
Label1.Caption = "sending"
End Sub

Private Sub tcpserver_SendComplete(wsindex As Integer)
Label1.Caption = "completenow"
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim i As Integer
For i = 0 To wsindex
If tcpserver(i).State <> 0 Then
tcpserver(i).Close
End If
Next i
Command1.Enabled = 1
End Sub

Private Sub tcpserver_Close(wsindex As Integer)
Label1.Caption = "closeing"
End Sub
