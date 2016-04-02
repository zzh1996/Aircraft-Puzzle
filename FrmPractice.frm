VERSION 5.00
Begin VB.Form FrmPractice 
   BackColor       =   &H00FFC060&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Practice Mode"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7785
   Icon            =   "FrmPractice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   549
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   519
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   24.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7500
      Left            =   120
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   0
      Top             =   600
      Width           =   7500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Answer"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Puzzle"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Puzzle NO.="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "FrmPractice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim StageNO As Integer, Step As Integer
Dim Status As Map, CMap As Map
Dim Playing As Boolean

Private Sub Command1_Click()
    NewPuzzle
    Pic.SetFocus
End Sub

Sub NewPuzzle()
    Playing = True
    Step = 0
    StageNO = RndInt(0, 11967)
    Erase Status.Dat
    CMap = Stages(StageNO)
    Reshow
End Sub

Private Sub Command2_Click()
    ShowAnswer
    Pic.SetFocus
End Sub

Sub ShowAnswer()
    Playing = False
    Reshow
End Sub

Private Sub Form_Load()
    NewPuzzle
End Sub

Sub Reshow()
    If Playing Then
        ShowMap Pic, CMap, Status, True
    Else
        Pic.Cls
        GreyMap Pic, Status
        ShowMap Pic, CMap, All1, False
    End If
    Label1.Caption = "Puzzle NO." & StageNO
    Label2.Caption = "Step=" & Step
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmMain.Show
End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 500 Or Y >= 500 Then Exit Sub
    If Not Playing Then Exit Sub
    Dim PX As Integer, PY As Integer
    PX = Int(X / 50)
    PY = Int(Y / 50)
    If Status.Dat(PX, PY) = 0 Then
        Status.Dat(PX, PY) = 1
        Step = Step + 1
        Reshow
        JudgeWin
    End If
End Sub

Sub JudgeWin()
    If IsWin(CMap, Status) Then 'win
        ShowAnswer
        MsgBox "You have won in " & Step & " steps!"
    End If
End Sub
