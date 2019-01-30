VERSION 5.00
Begin VB.Form mainmeun 
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5070
   Icon            =   "mainmeun.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "mainmeun.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   2895
   ScaleWidth      =   5070
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "关卡3"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "关卡2"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "关卡1"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "mainmeun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Load game1
Unload Me
game1.Show
End Sub

Private Sub Command2_Click()
Load game2
Unload Me
game2.Show
End Sub
Private Sub Command3_Click()
Load game3
Unload Me
game3.Show
End Sub

'Private Sub Form_Load()
'If grecord1(usernum) = True Then
'Command1.Enabled = True
'End If
'If grecord2(usernum) = True Then
'Command2.Enabled = True
'End If
'If grecord3(usernum) = True Then
'Command3.Enabled = True
'End If
'End Sub
Private Sub Form_Load()

End Sub
