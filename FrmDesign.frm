VERSION 5.00
Begin VB.Form FrmDesign 
   BackColor       =   &H00FFC060&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Design your puzzle"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7755
   Icon            =   "FrmDesign.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   7755
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command4 
      Caption         =   "Random"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Use code"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Complete"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   8520
      Width           =   1095
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
      Top             =   840
      Width           =   7500
   End
   Begin VB.Label Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click to place your aircrafts, right click to rotate."
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
      TabIndex        =   3
      Top             =   480
      Width           =   6360
   End
End
Attribute VB_Name = "FrmDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CMap As Map, Rotate As Integer

Private Sub Command1_Click()
    ReturnMap = SearchMap(CMap)
    If ReturnMap < 0 Then 'illegal
        MsgBox "Illegal map"
    Else
        ReturnOK = True
        Unload Me
    End If
End Sub

Private Sub Command2_Click()
    Erase CMap.Dat
    Reshow 0, 0
    Pic.SetFocus
End Sub

Private Sub Command3_Click()
    On Error GoTo Err
    Dim Code As String, CodeNum As Long
    Code = InputBox("Please input your code:")
    If Code <> "" Then
        If IsNumeric(Code) Then
            CodeNum = Int(Val(Code))
            If CodeNum < 0 Or CodeNum > 11967 Then
                MsgBox "Stage NO. out of range! (0~11967 expected)"
            Else
                CMap = Stages(CodeNum)
                Reshow 0, 0
            End If
        Else
            If Len(Code) = 9 Then
                Code = LCase(Code)
                If IsNum(Mid(Code, 1, 1)) And IsNum(Mid(Code, 2, 1)) And IsNum(Mid(Code, 4, 1)) And IsNum(Mid(Code, 5, 1)) And IsNum(Mid(Code, 7, 1)) And IsNum(Mid(Code, 8, 1)) And IsRotate(Mid(Code, 3, 1)) And IsRotate(Mid(Code, 6, 1)) And IsRotate(Mid(Code, 9, 1)) Then
                    Dim TempMap As Map
                    Erase TempMap.Dat
                    PastePlane TempMap, GetRotate(Mid(Code, 3, 1)), Val(Mid(Code, 1, 1)), Val(Mid(Code, 2, 1))
                    PastePlane TempMap, GetRotate(Mid(Code, 6, 1)), Val(Mid(Code, 4, 1)), Val(Mid(Code, 5, 1))
                    PastePlane TempMap, GetRotate(Mid(Code, 9, 1)), Val(Mid(Code, 7, 1)), Val(Mid(Code, 8, 1))
                    If SearchMap(TempMap) >= 0 Then
                        CMap = TempMap
                        Reshow 0, 0
                    Else
                        MsgBox "Aircrafts cannot cover each other!"
                    End If
                Else
                    MsgBox "Illegal character in code! (only 0~9,U,D,L,R)"
                End If
            Else
                MsgBox "Code length error! (9 expected)"
            End If
        End If
    End If
    Pic.SetFocus
    Exit Sub
Err:
    MsgBox "Illegal input!"
    Pic.SetFocus
    Exit Sub
End Sub

Function IsRotate(X As String) As Boolean
    IsRotate = (X = "u") Or (X = "d") Or (X = "l") Or (X = "r")
End Function

Function GetRotate(X As String) As Integer
    GetRotate = IIf(X = "u", 0, IIf(X = "r", 1, IIf(X = "d", 2, 3)))
End Function

Function IsNum(X As String) As Boolean
    IsNum = (X >= "0" And X <= "9")
End Function

Private Sub Command4_Click()
    CMap = Stages(RndInt(0, 11967))
    Reshow 0, 0
    Pic.SetFocus
End Sub

Private Sub Form_Load()
    ReturnOK = False
    Title.Caption = SendText
    Erase CMap.Dat
    Rotate = 0
    Reshow 0, 0
End Sub

Sub Reshow(PX As Integer, PY As Integer)
    Dim i As Integer, j As Integer
    Dim SquareColor As Long
    Pic.Cls
    For i = 0 To 4
        For j = 0 To 4
            Select Case PlaneData(Rotate, i, j)
            Case 0
                SquareColor = vbWhite
            Case 1
                SquareColor = &HFF7777
            Case 2
                SquareColor = &H7777FF
            End Select
            ColorBox Pic, PX + i, PY + j, SquareColor
        Next
    Next
    ShowMap Pic, CMap, All1, False
End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 500 Or Y >= 500 Then Exit Sub
    If Button = 2 Then
        Rotate = Rotate + 1
        If Rotate > 3 Then Rotate = 0
    ElseIf Button = 1 Then
        Dim TempMap As Map
        TempMap = CMap
        Dim PX As Integer, PY As Integer
        PX = Int(X / 50) - 2
        PY = Int(Y / 50) - 2
        If PX < 0 Then PX = 0
        If PY < 0 Then PY = 0
        If PX > 5 Then PX = 5
        If PY > 5 Then PY = 5
        If PastePlane(TempMap, Rotate, PX, PY) Then
            CMap = TempMap
        Else
            MsgBox "Aircrafts cannot cover each other!"
        End If
    End If
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 500 Or Y >= 500 Then Exit Sub
    Dim PX As Integer, PY As Integer
    PX = Int(X / 50) - 2
    PY = Int(Y / 50) - 2
    If PX < 0 Then PX = 0
    If PY < 0 Then PY = 0
    If PX > 5 Then PX = 5
    If PY > 5 Then PY = 5
    Reshow PX, PY
End Sub
