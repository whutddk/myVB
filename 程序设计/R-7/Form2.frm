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
      Text            =   "Text2"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   10680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock tcpclient 
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
Dim cut As Boolean

Private Sub form_load()
cut = False
Text1.Text = 0
Command3.Enabled = 0
End Sub
Private Sub Command1_Click()  '连接
If tcpclient.State <> 0 Then
tcpclient.Close
End If
If Val(Text1.Text) > 0 Then
tcpclient.RemotePort = Val(Text1.Text)
Else
tcpclient.RemotePort = 2000
End If
Text1.Text = tcpclient.RemotePort
tcpclient.RemoteHost = tcpclient.LocalHostName
tcpclient.Connect
Label1.Caption = "connecting"
Text1.Enabled = 0

Command1.Enabled = 0
End Sub



Private Sub tcpclient_Connect()
Dim sign As String
sign = "good job"
tcpclient.SendData sign
Label1.Caption = "complete"
Command3.Enabled = 1

End Sub
Private Sub Command2_Click() '断开
If tcpclient.State <> 0 Then
tcpclient.Close
End If
Command1.Enabled = 1
End Sub


Private Sub Command3_Click() '发送
If Text2.Text <> "" Then
tcpclient.SendData Text2.Text
End If
End Sub

Private Sub tcpclient_DataArrival(ByVal bytesTotal As Long)
Dim tcdata As String
tcpclient.GetData tcdata, vbString
Text2.Text = tcdata
End Sub

