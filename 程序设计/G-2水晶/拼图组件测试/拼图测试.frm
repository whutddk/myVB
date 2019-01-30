VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10515
   ClientLeft      =   1620
   ClientTop       =   2115
   ClientWidth     =   18150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "拼图测试.frx":0000
   ScaleHeight     =   10515
   ScaleWidth      =   18150
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   645
      Left            =   2040
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "DDK行动组"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_click()
Load game2
Unload Me
game2.Show
End Sub
