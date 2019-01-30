VERSION 5.00
Begin VB.Form game2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   13080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   21510
   Icon            =   "game2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13080
   ScaleWidth      =   21510
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "完成"
      Height          =   495
      Left            =   8280
      TabIndex        =   0
      Top             =   8160
      Width           =   2775
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   35
      Left            =   8160
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   34
      Left            =   9960
      Top             =   1680
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   33
      Left            =   11760
      Top             =   1200
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   32
      Left            =   11880
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   31
      Left            =   11040
      Top             =   3120
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   30
      Left            =   10080
      Top             =   3360
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   29
      Left            =   8760
      Top             =   3360
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   28
      Left            =   7920
      Top             =   3720
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   27
      Left            =   1080
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   26
      Left            =   2040
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   25
      Left            =   2880
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   24
      Left            =   3720
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   23
      Left            =   4560
      Top             =   5400
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   22
      Left            =   5520
      Top             =   4920
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   21
      Left            =   6600
      Top             =   4800
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   20
      Left            =   7680
      Top             =   5160
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   19
      Left            =   8520
      Top             =   5160
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   18
      Left            =   9480
      Top             =   5160
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   17
      Left            =   10800
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   16
      Left            =   11640
      Top             =   5160
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   15
      Left            =   12600
      Top             =   5160
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   14
      Left            =   12600
      Top             =   6240
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   13
      Left            =   11760
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   12
      Left            =   10800
      Top             =   6120
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   11
      Left            =   9720
      Top             =   6120
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   10
      Left            =   8400
      Top             =   6240
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   9
      Left            =   7440
      Top             =   6240
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   8
      Left            =   6480
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   7
      Left            =   5640
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   6
      Left            =   4920
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   5
      Left            =   4200
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   4
      Left            =   3480
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   3
      Left            =   2760
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   2
      Left            =   2040
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   1
      Left            =   1320
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image cpart 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   0
      Left            =   600
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image choose 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   5760
      Top             =   600
      Width           =   735
   End
   Begin VB.Image messioncomplete 
      Height          =   735
      Left            =   240
      Top             =   360
      Width           =   255
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   35
      Left            =   4200
      Top             =   4200
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   34
      Left            =   3480
      Top             =   4200
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   33
      Left            =   2760
      Top             =   4200
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   32
      Left            =   2040
      Top             =   4200
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   29
      Left            =   4200
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   28
      Left            =   3480
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   27
      Left            =   2760
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   26
      Left            =   2040
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   23
      Left            =   4200
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   22
      Left            =   3480
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   21
      Left            =   2760
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   20
      Left            =   2040
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   17
      Left            =   4200
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   16
      Left            =   3480
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   15
      Left            =   2760
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   31
      Left            =   1320
      Top             =   4200
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   25
      Left            =   1320
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   19
      Left            =   1320
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   13
      Left            =   1320
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   11
      Left            =   4200
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   10
      Left            =   3480
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   9
      Left            =   2760
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   8
      Left            =   2040
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   7
      Left            =   1320
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   14
      Left            =   2040
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   30
      Left            =   600
      Top             =   4200
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   24
      Left            =   600
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   18
      Left            =   600
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   12
      Left            =   600
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   6
      Left            =   600
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   5
      Left            =   4200
      Top             =   600
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   4
      Left            =   3480
      Top             =   600
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   3
      Left            =   2760
      Top             =   600
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   2
      Left            =   2040
      Top             =   600
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   1
      Left            =   1320
      Top             =   600
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   0
      Left            =   600
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "game2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private wfull(35) As Boolean
Private cfull(35) As Boolean
Private r As Integer

Private Sub Form_Initialize()
Dim i As Integer
For i = 0 To 35
cfull(i) = True
wfull(i) = False
Next i
r = 37
End Sub

Private Sub Form_Load()
Command1.Enabled = False
Dim i As Integer
For i = 0 To 35
cpart(i).Picture = LoadPicture(App.Path & "\data\2picture\choose\" & i + 1 & ".jpg")
Next i
End Sub


Private Sub cpart_Click(Index As Integer)
If r = 37 Then
r = Index
    If cfull(r) = True Then
    choose.Picture = cpart(r).Picture
    'game2.MousePointer = 99
    cpart(r).Picture = LoadPicture()
    cfull(r) = False
    Else
    cpart(r).Picture = choose.Picture
    choose.Picture = LoadPicture()
    r = 37
    End If
End If
End Sub


Private Sub Image1_Click(Index As Integer)
If r < 37 Then
Dim k As Integer
k = Index
    If wfull(k) = False Then
        If k = r Then
        Image1(k).Picture = choose.Picture
        choose.Picture = LoadPicture()
        wfull(k) = True
        r = 37
        
        Dim c As Integer
        For c = 0 To 35
            If wfull(c) = False Then
            Exit Sub
            End If
        Next c
        Command1.Enabled = True
        Else
        cpart(r).Picture = choose.Picture
        choose.Picture = LoadPicture()
        cfull(r) = True
        r = 37
        End If
    End If
End If
End Sub

Private Sub Command1_Click()
Command1.Visible = False
choose.Visible = False
Dim i As Integer
For i = 0 To 35
cpart(i).Visible = False
Next i

messioncomplete.Picture = LoadPicture(App.Path & "\data\2picture\原大小图.jpg")
End Sub

