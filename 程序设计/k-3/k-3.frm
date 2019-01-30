VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   13425
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "记录"
      Height          =   855
      Left            =   2280
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "##"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "登录时间"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "姓名"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim filename As String
Dim filenum As Integer



Private Sub Command1_Click()


    If Text1.Text <> "" Then 'And Text2.Text <> ""
   filename = App.Path & "\center.txt"
   filenum = FreeFile()
  
   Dim recordnum As Integer
   ouruser.name = Text1.Text
   ouruser.comeingtime = Text2.Text
   Open filename For Random As filenum Len = 30

   recordnum = LOF(filenum) / 22
   recordnum = recordnum + 1
  Put #filenum, recordnum, ouruser
Close
    Else
End If
End Sub
