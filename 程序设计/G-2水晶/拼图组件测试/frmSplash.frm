VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4590
   ClientLeft      =   6900
   ClientTop       =   4725
   ClientWidth     =   7425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmSplash.frx":000C
   Picture         =   "frmSplash.frx":057A
   ScaleHeight     =   4590
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   645
      Left            =   2040
      Picture         =   "frmSplash.frx":7582
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G-2      1.0.1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5040
      TabIndex        =   1
      Top             =   4080
      Width           =   1485
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
      Left            =   4800
      TabIndex        =   0
      Top             =   3720
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
Load lunch
Unload Me
lunch.Show
End Sub
