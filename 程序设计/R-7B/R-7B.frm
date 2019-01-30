VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   8280
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "发送"
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock flight 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label3 
      Caption         =   "本地端口"
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "远方端口"
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "远程主机IP"
      Height          =   255
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub Command1_Click()
'If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then
   ' flight.RemoteHostIP = Text1.Text
   ' flight.RemotePort = Text2.Text
    'flight.Bind = Text3.Text
'Else
    '
'End If
'End Sub



Private Sub Command2_Click()
Dim fsenddata As String
fsenddata = Text5.Text
flight.SendData fsenddata
End Sub

Private Sub flight_DataArrival(ByVal bytesTotal As Long)
Dim fgetdata As String
flight.GetData fgetdata, vbString
Text4.Text = fgetdata
End Sub

Private Sub Form_Load()
flight.RemoteHost = flight.LocalHostName
flight.RemotePort = 3000
flight.Bind 2000
End Sub
