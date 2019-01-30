VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form game2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   ClientHeight    =   12225
   ClientLeft      =   3810
   ClientTop       =   450
   ClientWidth     =   14700
   Icon            =   "game2.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "game2.frx":08CA
   ScaleHeight     =   12225
   ScaleWidth      =   14700
   Begin VB.CommandButton Command1 
      Caption         =   "投降"
      Height          =   255
      Left            =   13080
      TabIndex        =   15
      Top             =   11760
      Width           =   1095
   End
   Begin VB.CommandButton cc4 
      Caption         =   "叠加"
      Height          =   615
      Left            =   9480
      TabIndex        =   13
      Top             =   9480
      Width           =   255
   End
   Begin VB.CommandButton nextlevel 
      Caption         =   "完成"
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   8760
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton gu 
      Caption         =   "后退"
      Height          =   855
      Left            =   4680
      TabIndex        =   11
      Top             =   9480
      Width           =   615
   End
   Begin VB.CommandButton cc1 
      Caption         =   "叠加"
      Height          =   615
      Left            =   1920
      TabIndex        =   10
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cr1 
      Caption         =   "逆时针"
      Height          =   615
      Left            =   3120
      TabIndex        =   9
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cf2 
      Caption         =   "顺时针"
      Height          =   615
      Left            =   5280
      TabIndex        =   8
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cc2 
      Caption         =   "叠加"
      Height          =   615
      Left            =   6480
      TabIndex        =   7
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cr2 
      Caption         =   "逆时针"
      Height          =   615
      Left            =   7680
      TabIndex        =   6
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cf3 
      Caption         =   "顺时针"
      Height          =   615
      Left            =   10080
      TabIndex        =   5
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command1"
      Height          =   615
      Left            =   17040
      TabIndex        =   4
      Top             =   13800
      Width           =   975
   End
   Begin VB.CommandButton gg 
      Caption         =   "后退"
      Height          =   615
      Left            =   4560
      TabIndex        =   3
      Top             =   13200
      Width           =   975
   End
   Begin VB.CommandButton cr3 
      Caption         =   "逆时针"
      Height          =   615
      Left            =   12480
      TabIndex        =   2
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cc3 
      Caption         =   "叠加"
      Height          =   615
      Left            =   11280
      TabIndex        =   1
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cf1 
      Caption         =   "顺时针"
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   5640
      Width           =   975
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   375
      Left            =   11640
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   3413
      _cy             =   661
   End
   Begin VB.Image wpart4 
      BorderStyle     =   1  'Fixed Single
      Height          =   4815
      Left            =   480
      Top             =   6600
      Width           =   4455
   End
   Begin VB.Image wpart3 
      Height          =   5055
      Left            =   480
      Top             =   6600
      Width           =   4095
   End
   Begin VB.Image wpart2 
      Height          =   5055
      Left            =   480
      Top             =   6600
      Width           =   4095
   End
   Begin VB.Image aim 
      BorderStyle     =   1  'Fixed Single
      Height          =   5055
      Left            =   9840
      Top             =   6600
      Width           =   4095
   End
   Begin VB.Image cpart4 
      BorderStyle     =   1  'Fixed Single
      Height          =   5055
      Left            =   5280
      Top             =   6600
      Width           =   4095
   End
   Begin VB.Image wpart1 
      Height          =   5055
      Left            =   480
      Top             =   6600
      Width           =   4095
   End
   Begin VB.Image cpart3 
      BorderStyle     =   1  'Fixed Single
      Height          =   5055
      Left            =   9840
      Top             =   480
      Width           =   4095
   End
   Begin VB.Image cpart2 
      BorderStyle     =   1  'Fixed Single
      Height          =   5055
      Left            =   5160
      Top             =   480
      Width           =   4095
   End
   Begin VB.Image cpart1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5055
      Left            =   360
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "game2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private round1 As Integer
Private round2 As Integer
Private round3 As Integer
Private full1 As Integer
Private full2 As Integer
Private full3 As Integer
Private full4 As Boolean
Private onshow1 As Integer
Private onshow2 As Integer
Private onshow3 As Integer



Private Sub Command1_Click()
Load mainmeun
Unload Me
mainmeun.Show
End Sub

Private Sub Form_Load()
round1 = 101
round2 = 102
round3 = 102
full1 = 0
full2 = 0
full3 = 0
cpart1.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\a2.gif")
cpart2.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\b3.gif")
cpart3.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\c3.gif")
cpart4.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\d.gif")
aim.Picture = LoadPicture(App.Path & "\data\game1\IMG_0146改.jpg")
gu.Enabled = False
nextlevel.Visible = False
WindowsMediaPlayer1.URL = App.Path & "\data\music\LedZeper - AboveTheTreetops.mp3"
WindowsMediaPlayer1.Controls.play

End Sub



Private Sub cf1_Click()
round1 = round1 + 1
onshow1 = round1 Mod 4
If onshow1 = 0 Then
    cpart1.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\a1.gif")
    ElseIf onshow1 = 1 Then
    cpart1.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\a2.gif")
    ElseIf onshow1 = 2 Then
    cpart1.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\a3.gif")
    ElseIf onshow1 = 3 Then
    cpart1.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\a4.gif")
    End If
    
End Sub
Private Sub cr1_Click()
round1 = round1 - 1
onshow1 = round1 Mod 4
If onshow1 = 0 Then
    cpart1.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\a1.gif")
    ElseIf onshow1 = 1 Then
    cpart1.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\a2.gif")
    ElseIf onshow1 = 2 Then
    cpart1.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\a3.gif")
    ElseIf onshow1 = 3 Then
    cpart1.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\a4.gif")
    End If
    
End Sub
Private Sub cc1_click()
cf1.Enabled = False
cc1.Enabled = False
cr1.Enabled = False
If full1 = 0 Then
    wpart1.Picture = cpart1.Picture
    full1 = 1
    ElseIf full2 = 0 Then
    wpart2.Picture = cpart1.Picture
    full2 = 1
    Else
    wpart3.Picture = cpart1.Picture
    full3 = 1
End If
gu.Enabled = True

End Sub
Private Sub cf2_Click()
round2 = round2 + 1
onshow2 = round2 Mod 4
If onshow2 = 0 Then
    cpart2.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\b1.gif")
    ElseIf onshow2 = 1 Then
    cpart2.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\b2.gif")
    ElseIf onshow2 = 2 Then
    cpart2.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\b3.gif")
    ElseIf onshow2 = 3 Then
    cpart2.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\b4.gif")
    End If
    
End Sub

Private Sub cr2_Click()
round2 = round2 - 1
onshow2 = round2 Mod 4
If onshow2 = 0 Then
    cpart2.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\b1.gif")
    ElseIf onshow2 = 1 Then
    cpart2.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\b2.gif")
    ElseIf onshow2 = 2 Then
    cpart2.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\b3.gif")
    ElseIf onshow2 = 3 Then
    cpart2.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\b4.gif")
    End If
    
End Sub
Private Sub cc2_click()
cf2.Enabled = False
cc2.Enabled = False
cr2.Enabled = False
If full1 = 0 Then
    wpart1.Picture = cpart2.Picture
    full1 = 2
    ElseIf full2 = 0 Then
    wpart2.Picture = cpart2.Picture
    full2 = 2
    Else
    wpart3.Picture = cpart2.Picture
    full3 = 2
End If
gu.Enabled = True

End Sub
Private Sub cf3_Click()
round3 = round3 + 1
onshow3 = round3 Mod 4
If onshow3 = 0 Then
    cpart3.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\c1.gif")
    ElseIf onshow3 = 1 Then
    cpart3.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\c2.gif")
    ElseIf onshow3 = 2 Then
    cpart3.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\c3.gif")
    ElseIf onshow3 = 3 Then
    cpart3.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\c4.gif")
    End If
    
End Sub

Private Sub cr3_Click()
round3 = round3 - 1
onshow3 = round3 Mod 4
If onshow3 = 0 Then
    cpart3.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\c1.gif")
    ElseIf onshow3 = 1 Then
    cpart3.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\c2.gif")
    ElseIf onshow3 = 2 Then
    cpart3.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\c3.gif")
    ElseIf onshow3 = 3 Then
    cpart3.Picture = LoadPicture(App.Path & "\data\game1\1.0.2\c4.gif")
    End If
    
End Sub

Private Sub cc3_click()
cf3.Enabled = False
cc3.Enabled = False
cr3.Enabled = False
If full1 = 0 Then
    wpart1.Picture = cpart3.Picture
    full1 = 3
    ElseIf full2 = 0 Then
    wpart2.Picture = cpart3.Picture
    full2 = 3
    Else
    wpart3.Picture = cpart3.Picture
    full3 = 3
End If
gu.Enabled = True

End Sub



Private Sub cc4_click()
cc4.Enabled = False
gu.Enabled = True
full4 = 1
wpart4.Picture = cpart4.Picture
    If full1 = 1 And full2 = 2 And full3 = 3 And onshow1 = 0 And onshow2 = 0 And onshow3 = 0 Then
    nextlevel.Visible = True
    End If

End Sub

Private Sub gu_Click()
If full4 = True Then
wpart4.Picture = LoadPicture()
cc4.Enabled = True
full4 = False
ElseIf full3 <> 0 Then
    Select Case full3
    Case 1
    cc1.Enabled = True
    cf1.Enabled = True
    cr1.Enabled = True
    Case 2
    cc2.Enabled = True
    cf2.Enabled = True
    cr2.Enabled = True
    Case 3
    cc3.Enabled = True
    cf3.Enabled = True
    cr3.Enabled = True
    End Select
wpart3.Picture = LoadPicture()
full3 = 0
ElseIf full2 <> 0 Then
    Select Case full2
    Case 1
    cc1.Enabled = True
    cf1.Enabled = True
    cr1.Enabled = True
    Case 2
    cc2.Enabled = True
    cf2.Enabled = True
    cr2.Enabled = True
    Case 3
    cc3.Enabled = True
    cf3.Enabled = True
    cr3.Enabled = True
    End Select
wpart2.Picture = LoadPicture()
full2 = 0
Else
    Select Case full1
    Case 1
    cc1.Enabled = True
    cf1.Enabled = True
    cr1.Enabled = True
    Case 2
    cc2.Enabled = True
    cf2.Enabled = True
    cr2.Enabled = True
    Case 3
    cc3.Enabled = True
    cf3.Enabled = True
    cr3.Enabled = True
    End Select
wpart1.Picture = LoadPicture()
full1 = 0
End If
If full1 = 0 And full2 = 0 And full3 = 0 And full4 = False Then
gu.Enabled = False
End If
End Sub






Private Sub nextlevel_Click()
Load mainmeun
Unload Me
mainmeun.Show
End Sub


