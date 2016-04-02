VERSION 5.00
Begin VB.Form FrmPossibility 
   BackColor       =   &H00FFC060&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Possibility"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7740
   Icon            =   "FrmPossibility.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   515
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   516
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "ºÚÌå"
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
      Top             =   120
      Width           =   7500
   End
End
Attribute VB_Name = "FrmPossibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OldFont As New StdFont
Dim NewFont As New StdFont

Private Sub Form_Load()
    Reshow
End Sub

Sub Reshow()
    Dim i As Integer, j As Integer
    Dim Percent As Single, Clr As Integer
    Set OldFont = Pic.Font
    NewFont.Bold = False
    NewFont.Size = 12
    Set Pic.Font = NewFont
    Pic.Cls
    For i = 0 To 9
        For j = 0 To 9
            If SendStatus.Dat(i, j) = 0 Then
                Percent = Possibility.Dat(i, j) / SolutionCount
                Clr = Int(255 - 255 * Percent)
                ColorBox Pic, i, j, RGB(255, Clr, Clr)
                Pic.ForeColor = vbBlack
                Pic.CurrentX = i * 50 + 1
                Pic.CurrentY = j * 50 + 20
                Pic.Print Format(Percent, "0.0%")
            End If
        Next
    Next
    Set Pic.Font = OldFont
    GreyMap Pic, SendStatus
    ShowMap Pic, SendMap, SendStatus, False
End Sub

