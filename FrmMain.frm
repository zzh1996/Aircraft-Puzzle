VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00FFC060&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aircraft Puzzle 1.2"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   4680
      Width           =   3855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   4
      Top             =   3600
      Width           =   3855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Two Players"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play with AI"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Practice Mode"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "负一的平方根 April,2014"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   2280
      TabIndex        =   0
      Top             =   5760
      Width           =   2070
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
    FrmPractice.Show
End Sub

Private Sub Command2_Click()
    Unload Me
    FrmAI.Show
    If ExitAI Then Unload FrmAI
End Sub

Private Sub Command3_Click()
    Unload Me
    FrmTwoPlayer.Show
    If ExitTwoPlayer Then Unload FrmTwoPlayer
End Sub

Private Sub Command4_Click()
    Unload Me
    FrmTools.Show
End Sub

Private Sub Command5_Click()
    Unload Me
    FrmHelp.Show
End Sub
