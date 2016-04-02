VERSION 5.00
Begin VB.Form FrmAI 
   BackColor       =   &H00FFC060&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Play with AI"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15465
   Icon            =   "FrmAI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   584
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1031
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox Difficulty 
      Height          =   300
      ItemData        =   "FrmAI.frx":08CA
      Left            =   13560
      List            =   "FrmAI.frx":08DA
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Game"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   8280
      Width           =   1215
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulty:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   12120
      TabIndex        =   7
      Top             =   8280
      Width           =   1320
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
      TabIndex        =   4
      Top             =   8280
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Computer"
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
      Left            =   10920
      TabIndex        =   3
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player"
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
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "FrmAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PMap As Map, CMap As Map, PStatus As Map, CStatus As Map, Mark As Map
Dim Step As Integer
Dim Won As Boolean
Dim LastButton As Integer, LastX As Single, LastY As Single

Function NewGame() As Boolean
    Design "Design your puzzle for computer to guess"
    If ReturnOK Then
        CMap = Stages(ReturnMap)
        Erase PStatus.Dat
        Erase CStatus.Dat
        PMap = Stages(RndInt(0, 11967))
        Erase Mark.Dat
        Step = 0
        Difficulty.Enabled = True
        Reshow
    End If
    NewGame = ReturnOK
End Function

Private Sub Command1_Click()
    NewGame
    Pic1.SetFocus
End Sub

Private Sub Difficulty_Click()
    If Me.Visible Then Pic1.SetFocus
End Sub

Private Sub Form_Load()
    Difficulty.ListIndex = 2
    ExitAI = Not NewGame
End Sub

Sub Reshow()
    Label3.Caption = "Step=" & Step
    ShowMap Pic1, PMap, PStatus, True
    ShowMark Pic1, Mark, PStatus
    Dim i As Integer, j As Integer
    Pic2.Cls
    GreyMap Pic2, CStatus
    ShowMap Pic2, CMap, All1, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmMain.Show
End Sub

Private Sub Pic1_DblClick()
    Pic1_MouseDown LastButton, 0, LastX, LastY
End Sub

Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 500 Or Y >= 500 Then Exit Sub
    Dim PX As Integer, PY As Integer
    PX = Int(X / 50)
    PY = Int(Y / 50)
    If Button = 1 Then
        If PStatus.Dat(PX, PY) = 0 Then
            If Difficulty.ListIndex = 3 Then 'Evil
                If PMap.Dat(PX, PY) = 2 Then
                    CalcMatch PMap, PStatus
                    Dim s As Integer
                    Dim NewMaps(11967) As Integer, NewMapCount As Integer
                    NewMapCount = 0
                    For s = 0 To SolutionCount - 1
                        If Stages(Solution(s)).Dat(PX, PY) <> 2 Then '(PX,PY) is not head
                            NewMaps(NewMapCount) = Solution(s)
                            NewMapCount = NewMapCount + 1
                        End If
                    Next
                    If NewMapCount > 0 Then 'Change map
                        PMap = Stages(NewMaps(RndInt(0, NewMapCount - 1)))
                    End If
                End If
            End If
            PStatus.Dat(PX, PY) = 1
            Step = Step + 1
            Difficulty.Enabled = False
            Reshow
            JudgeWin
            If Not Won Then
                Label3.Caption = "AI at work, please wait..."
                AIGo
                Reshow
                JudgeWin
            End If
        End If
    ElseIf Button = 2 Then
        Mark.Dat(PX, PY) = Mark.Dat(PX, PY) + 1
        If Mark.Dat(PX, PY) > 3 Then Mark.Dat(PX, PY) = 0
        Reshow
    End If
    LastButton = Button
    LastX = X
    LastY = Y
End Sub

Sub JudgeWin()
    Won = False
    If IsWin(PMap, PStatus) Then
        ShowAnswer
        ShowStageNO
        MsgBox "Player have won in " & Step & " steps!"
        Won = True
    ElseIf IsWin(CMap, CStatus) Then
        ShowAnswer
        ShowStageNO
        MsgBox "Computer have won in " & Step & " steps!"
        Won = True
    End If
    If Won Then
        If Not NewGame Then Unload Me
    End If
End Sub

Sub ShowStageNO()
    Label3.Caption = "Player stage NO.=" & SearchMap(PMap) & " Computer stage NO.=" & SearchMap(CMap)
End Sub

Sub AIGo()
    Dim MaxX(99) As Integer, MaxY(99) As Integer, Max As Integer
    Dim Pointer As Integer, RndPointer As Integer
    Dim i As Integer, j As Integer
    CalcPossibility CMap, CStatus
    Max = 0
    For i = 0 To 9
        For j = 0 To 9
            If CStatus.Dat(i, j) = 0 Then
                If Possibility.Dat(i, j) > Max Then Max = Possibility.Dat(i, j)
            End If
        Next
    Next
    Select Case Difficulty.ListIndex
    Case 0 'Easy
        Pointer = 0
        For i = 0 To 9
            For j = 0 To 9
                If (Possibility.Dat(i, j) > 0 Or Rnd() < 0.1) And CStatus.Dat(i, j) = 0 Then
                    MaxX(Pointer) = i
                    MaxY(Pointer) = j
                    Pointer = Pointer + 1
                End If
            Next
        Next
    Case 1 'Medium
        Pointer = 0
        For i = 0 To 9
            For j = 0 To 9
                If Possibility.Dat(i, j) > 0 And CStatus.Dat(i, j) = 0 Then '(i.j) is possible and unknown
                    MaxX(Pointer) = i
                    MaxY(Pointer) = j
                    Pointer = Pointer + 1
                End If
            Next
        Next
    Case 2, 3 'Hard & Evil
        Pointer = 0
        For i = 0 To 9
            For j = 0 To 9
                If Possibility.Dat(i, j) = Max And CStatus.Dat(i, j) = 0 Then  '(i.j) is max and unknown
                    MaxX(Pointer) = i
                    MaxY(Pointer) = j
                    Pointer = Pointer + 1
                End If
            Next
        Next
    End Select
    RndPointer = RndInt(0, Pointer - 1)
    CStatus.Dat(MaxX(RndPointer), MaxY(RndPointer)) = 1
End Sub

Sub ShowAnswer()
    Dim i As Integer, j As Integer
    Pic1.Cls
    GreyMap Pic1, PStatus
    ShowMap Pic1, PMap, All1, False
End Sub

