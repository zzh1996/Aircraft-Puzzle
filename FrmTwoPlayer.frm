VERSION 5.00
Begin VB.Form FrmTwoPlayer 
   BackColor       =   &H00FFC060&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Two Players"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15465
   Icon            =   "FrmTwoPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   584
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1031
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Pic1 
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
   Begin VB.PictureBox Pic2 
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
      Left            =   7800
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   1
      Top             =   600
      Width           =   7500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Game"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10800
      TabIndex        =   4
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label Label3 
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
      Left            =   1560
      TabIndex        =   3
      Top             =   8280
      Width           =   675
   End
End
Attribute VB_Name = "FrmTwoPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim P1Map As Map, P2Map As Map, P1Status As Map, P2Status As Map
Dim Step As Integer
Dim Won As Boolean
Dim Turn As Integer

Function NewGame() As Boolean
    Dim TempMap As Integer
    Design "Player 1 please design a puzzle for Player 2 to guess"
    If ReturnOK Then
        TempMap = ReturnMap
        Design "Player 2 please design a puzzle for Player 1 to guess"
        If ReturnOK Then
            P2Map = Stages(TempMap)
            P1Map = Stages(ReturnMap)
            Erase P1Status.Dat
            Erase P2Status.Dat
            Step = 0
            Turn = 1
            Reshow
        End If
    End If
    NewGame = ReturnOK
End Function

Private Sub Command1_Click()
    NewGame
    Pic1.SetFocus
End Sub

Private Sub Form_Load()
    ExitTwoPlayer = Not NewGame
End Sub

Sub Reshow()
    If Turn = 1 Then
        Label1.ForeColor = vbRed
        Label2.ForeColor = vbBlack
    Else
        Label1.ForeColor = vbBlack
        Label2.ForeColor = vbRed
    End If
    Label3.Caption = "Step=" & Step
    ShowMap Pic1, P1Map, P1Status, True
    ShowMap Pic2, P2Map, P2Status, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmMain.Show
End Sub

Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 500 Or Y >= 500 Then Exit Sub
    If Turn = 1 Then
        Dim PX As Integer, PY As Integer
        PX = Int(X / 50)
        PY = Int(Y / 50)
        If P1Status.Dat(PX, PY) = 0 Then
            P1Status.Dat(PX, PY) = 1
            Step = Step + 1
            Reshow
            If Not JudgeWin Then Turn = 2
        End If
    Else
        MsgBox "It's Player 2 's turn!"
    End If
End Sub

Private Sub Pic2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 500 Or Y >= 500 Then Exit Sub
    If Turn = 2 Then
        Dim PX As Integer, PY As Integer
        PX = Int(X / 50)
        PY = Int(Y / 50)
        If P2Status.Dat(PX, PY) = 0 Then
            P2Status.Dat(PX, PY) = 1
            Reshow
            If Not JudgeWin Then Turn = 1
        End If
    Else
        MsgBox "It's Player 1 's turn!"
    End If
End Sub

Function JudgeWin() As Boolean
    Won = False
    If IsWin(P1Map, P1Status) Then
        ShowAnswer
        ShowStageNO
        MsgBox "Player 1 have won in " & Step & " steps!"
        Won = True
    ElseIf IsWin(P2Map, P2Status) Then
        ShowAnswer
        ShowStageNO
        MsgBox "Player 2 have won in " & Step & " steps!"
        Won = True
    End If
    If Won Then
        If Not NewGame Then Unload Me
    End If
    JudgeWin = Won
End Function

Sub ShowStageNO()
    Label3.Caption = "Player 1 stage NO.=" & SearchMap(P1Map) & " Player 2 stage NO.=" & SearchMap(P2Map)
End Sub

Sub ShowAnswer()
    Dim i As Integer, j As Integer
    Pic1.Cls
    GreyMap Pic1, P1Status
    ShowMap Pic1, P1Map, All1, False
    Pic2.Cls
    GreyMap Pic2, P2Status
    ShowMap Pic2, P2Map, All1, False
End Sub

