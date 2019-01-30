VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17580
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   17580
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   240
   End
   Begin VB.PictureBox Picture1 
      Height          =   7695
      Left            =   1440
      ScaleHeight     =   7635
      ScaleWidth      =   16155
      TabIndex        =   0
      Top             =   1200
      Width           =   16215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Const thescreen = 0
'Const theform = 0



Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False

End Sub

Private Sub Form_Load()
'Call keybd_event(vbKeySnapshot, thescreen, 1, 1)
'Picture1.Picture = Clipboard.GetData(vbCFBitmap)
'SavePicture Picture1.Picture, App.Path & "\picture.jpg"


'Open App.Path & "picture.jpg" For Binary As #1
'Dim filedata(1 To 8000) As Byte
'filedata() = CByte(Picture1.Picture)
'Put #1, , filedata()
'Close #1
Timer1.Interval = 10000

'Call keybd_event(116, 0, 0, 0)
End Sub

Private Sub Timer1_Timer()
'Call keybd_event(116, 0, 0, 0)
'SendKeys "{ENTER}"
SendKeys "{F5}"
LeftClick 1
End Sub

