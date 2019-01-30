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
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   5040
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "send"
      Height          =   495
      Left            =   8160
      TabIndex        =   3
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "close"
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
      Caption         =   "listen"
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
Dim wsindex As Integer

Private Sub form_load()
Text1.Enabled = 0
Command2.Enabled = 0
Command4.Enabled = False
Text2.Text = 0
wsindex = 0
End Sub
Private Sub Command1_Click() 'ÕìÌý
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
Private Sub tcpserver_ConnectionRequest(wsindex As Integer, ByVal requestID As Long)
wsindex = wsindex + 1
Load tcpserver(wsindex)
Label1.Caption = "requestnow"
Command2.Enabled = 1
tcpserver(wsindex).Accept requestID
End Sub

Private Sub tcpserver_Connect(wsindex As Integer)
Text1.Enabled = 1
Label1.Caption = "connectnow"
Command4.Enabled = True
Text1.Enabled = True
End Sub

Private Sub tcpserver_DataArrival(wsindex As Integer, ByVal bytesTotal As Long)
Dim tcdata As String
tcpserver(wsindex).GetData tcdata, vbString
Text1.Text = tcdata
End Sub

Private Sub Command4_click()
If Text1.Text = "" Then
Exit Sub
Else
tcpserver(wsindex).SendData Text1.Text
Label1.Caption = "sending"
End If
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
