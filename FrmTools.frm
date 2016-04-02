VERSION 5.00
Begin VB.Form FrmTools 
   BackColor       =   &H00FFC060&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Puzzle Tools"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7770
   Icon            =   "FrmTools.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   580
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   518
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "Possibility View"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Solve"
      Default         =   -1  'True
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   8160
      Width           =   975
   End
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
      Top             =   480
      Width           =   7500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click or right click to change the status of each square."
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6840
   End
End
Attribute VB_Name = "FrmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CMap As Map, Status As Map
Dim LastX As Integer, LastY As Integer, LastButton As Integer

Private Sub Command1_Click()
    Clear
    Pic.SetFocus
End Sub

Private Sub Command2_Click()
    CalcMatch CMap, Status
    If SolutionCount > 0 Then
        SendStatus = Status
        FrmSolution.Show 1
    Else
        MsgBox "No solution!"
    End If
    Pic.SetFocus
End Sub

Private Sub Command3_Click()
    CalcPossibility CMap, Status
    If SolutionCount > 0 Then
        SendMap = CMap
        SendStatus = Status
        FrmPossibility.Show 1
    Else
        MsgBox "No solution!"
    End If
    Pic.SetFocus
End Sub

Private Sub Form_Load()
    Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmMain.Show
End Sub

Sub Reshow()
    ShowMap Pic, CMap, Status, True
End Sub

Sub Clear()
    Erase CMap.Dat
    Erase Status.Dat
    Reshow
End Sub

Private Sub Pic_DblClick()
    Pic_MouseDown LastButton, 0, LastX * 50, LastY * 50
End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < 0 Or Y < 0 Or X >= 500 Or Y >= 500 Then Exit Sub
    Dim PX As Integer, PY As Integer
    PX = Int(X / 50)
    PY = Int(Y / 50)
    If Button = 1 Then
        If Status.Dat(PX, PY) = 0 Then
            Status.Dat(PX, PY) = 1
            CMap.Dat(PX, PY) = 0
        Else
            Select Case CMap.Dat(PX, PY)
            Case 0
                CMap.Dat(PX, PY) = 2
            Case 1
                Status.Dat(PX, PY) = 0
            Case 2
                CMap.Dat(PX, PY) = 1
            End Select
        End If
    ElseIf Button = 2 Then
        If Status.Dat(PX, PY) = 0 Then
            Status.Dat(PX, PY) = 1
            CMap.Dat(PX, PY) = 1
        Else
            Select Case CMap.Dat(PX, PY)
            Case 0
                Status.Dat(PX, PY) = 0
            Case 1
                CMap.Dat(PX, PY) = 2
            Case 2
                CMap.Dat(PX, PY) = 0
            End Select
        End If
    End If
    LastX = PX
    LastY = PY
    LastButton = Button
    Reshow
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim PX As Integer, PY As Integer
    PX = Int(X / 50)
    PY = Int(Y / 50)
    If PX <> LastX Or PY <> LastY Then
        Pic_MouseDown Button, Shift, X, Y
    End If
End Sub
