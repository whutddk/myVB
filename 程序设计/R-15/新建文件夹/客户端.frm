VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13350
   LinkTopic       =   "Form2"
   ScaleHeight     =   6825
   ScaleWidth      =   13350
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Text            =   "3"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "线程"
      Height          =   255
      Left            =   8880
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "字节"
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "发送"
      Height          =   615
      Left            =   3120
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "断开"
      Height          =   735
      Left            =   3000
      TabIndex        =   4
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "连接"
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   10680
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock tcpclient 
      Index           =   0
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private arrival As Boolean
Private filename As String
Private filedata() As Byte
Private wsindex As Integer




Private Sub form_load()
arrival = False
wsindex = 1
Text1.Text = 0
Command3.Enabled = 0
End Sub
Private Sub Command1_Click()  '连接
'If tcpclient.State <> 0 Then
'tcpclient.Close
'End If
wsindex = Text3.Text + 1
Dim i As Integer
For i = 1 To wsindex
Load tcpclient(i)

If Val(Text1.Text) > 0 Then
tcpclient(i).RemotePort = Val(Text1.Text)
Else
tcpclient(i).RemotePort = 2000
End If
Text1.Text = tcpclient(i).RemotePort
tcpclient(i).RemoteHost = "sys-PC" 'tcpclient.LocalHostName
tcpclient(i).Connect
Label1.Caption = "connecting" & tcpclient(i).RemoteHost
Text1.Enabled = 0
Command1.Enabled = 0
Next i
ReDim filedata(2 To wsindex, 1 To 8000)
End Sub

Private Sub tcpclient_Connect(index As Integer)
'Dim sign As String
'sign = "good job"
'tcpclient.SendData sign
Label1.Caption = "connection"
Command3.Enabled = 1
End Sub

Private Sub Command2_Click() '断开
Dim i As Integer
For i = 1 To wsindex
If tcpclient(i).State <> 0 Then
tcpclient(i).Close
End If
Next i
Command1.Enabled = 1
Dim copydata(1 To 100000) As Byte
For i = 2 To wsindex
Get #i, , copydata()
Put #1, , copydata()
Close #i
Next i
Close #1
End Sub


'Private Sub Command3_Click() '发送
'If Text2.Text <> "" Then
'tcpclient.SendData Text2.Text
'End If
'End Sub

Private Sub tcpclient_DataArrival(index As Integer, ByVal bytesTotal As Long)
DoEvents
'Dim filedata(1 To 8000) As Byte
'Dim getfilename As String
'getfilename = App.Path & "\getname.jpg"
If arrival = False Then
tcpclient(index).GetData filename, vbString
arrival = True
Open App.Path & "\" & filename For Binary As #1
Dim i As Integer
For i = 2 To wsindex
Open App.Path & "\" & i For Binary As #i
Next i
tcpclient(index).SendData arrival
Else
For i = 1 To 8000
tcpclient(index).GetData filedata(index, i), vbByte
Put #index, , filedata(index, i)
Next i
tcpclient(index).SendData arrival
End If
'Dim tcdata As String
'tcpclient.GetData tcdata, vbString
'Text2.Text = tcdata
End Sub

