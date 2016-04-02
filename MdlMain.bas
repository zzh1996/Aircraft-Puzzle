Attribute VB_Name = "MdlMain"
Option Explicit

Public Type Map
    Dat(9, 9) As Integer
End Type

Public PlaneData(3, 4, 4) As Integer
Public Stages(11967) As Map

Public ReturnMap As Integer, ReturnOK As Boolean, SendText As String
Public All1 As Map

Public ExitAI As Boolean
Public ExitTwoPlayer As Boolean

Public Possibility As Map, SendMap As Map, SendStatus As Map

Public Solution(11967) As Integer, SolutionCount As Integer

Sub Main()
    If App.PrevInstance Then
        MsgBox "This game is already running!"
    Else
        Init
        FrmMain.Show
    End If
End Sub

Sub Init()
    Randomize
    LoadPlaneData
    LoadStages
    Dim i As Integer, j As Integer
    For i = 0 To 9
        For j = 0 To 9
            All1.Dat(i, j) = 1
        Next
    Next
End Sub

Sub LoadPlaneData()
    Dim Plane(3) As String
    Plane(0) = "0020011111001000010001110"
    Plane(1) = "0001010010111121001000010"
    Plane(2) = "0111000100001001111100200"
    Plane(3) = "0100001001211110100101000"
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    For i = 0 To 3
        For j = 0 To 4
            For k = 0 To 4
                PlaneData(i, j, k) = Val(Mid(Plane(i), j + k * 5 + 1, 1))
            Next
        Next
    Next
End Sub

Sub LoadStages()
    Dim a As Integer, b As Integer, c As Integer
    Dim d As Integer, e As Integer, f As Integer
    Dim g As Integer, h As Integer, i As Integer
    Dim s As Integer
    Dim Map1 As Map, Map2 As Map, Map3 As Map
    s = 0
    For a = 0 To 5
        For b = 0 To 5
            For c = 0 To 3
                Erase Map1.Dat
                PastePlane Map1, c, a, b
                For d = 0 To a
                    For e = 0 To 5
                        If a > d Or (a = d And b > e) Then
                            For f = 0 To 3
                                Map2 = Map1
                                If PastePlane(Map2, f, d, e) Then
                                    For g = 0 To d
                                        For h = 0 To 5
                                            If d > g Or (d = g And e > h) Then
                                                For i = 0 To 3
                                                    Map3 = Map2
                                                    If PastePlane(Map3, i, g, h) Then
                                                        Stages(s) = Map3
                                                        s = s + 1
                                                    End If
                                                Next
                                            End If
                                        Next
                                    Next
                                End If
                            Next
                        End If
                    Next
                Next
            Next
        Next
    Next
End Sub

Function PastePlane(ByRef Map0 As Map, N As Integer, X As Integer, Y As Integer) As Boolean
    Dim i As Integer, j As Integer
    For i = 0 To 4
        For j = 0 To 4
            If PlaneData(N, i, j) > 0 Then
                If Map0.Dat(X + i, Y + j) > 0 Then
                    PastePlane = False
                    Exit Function
                Else
                    Map0.Dat(X + i, Y + j) = PlaneData(N, i, j)
                End If
            End If
        Next
    Next
    PastePlane = True
End Function

Sub P(Map0 As Map)
    Dim i As Integer, j As Integer
    Debug.Print "---------"
    For j = 0 To 9
        For i = 0 To 9
            Debug.Print Map0.Dat(i, j);
        Next
        Debug.Print
    Next
End Sub

Function MapEqual(Map1 As Map, Map2 As Map) As Boolean
    MapEqual = True
    Dim i As Integer, j As Integer
    For i = 0 To 9
        For j = 0 To 9
            If Map1.Dat(i, j) <> Map2.Dat(i, j) Then
                MapEqual = False
                Exit Function
            End If
        Next
    Next
End Function

Function RndInt(a As Integer, b As Integer)
    RndInt = Int(Rnd() * (b - a + 1)) + a
End Function

Sub ShowMap(Pic As PictureBox, Map0 As Map, Status As Map, Cls As Boolean)
    Dim i As Integer
    Dim j As Integer
    If Cls Then Pic.Cls
    Pic.ForeColor = vbBlack
    For i = 0 To 9
        Pic.Line (i * 50, 0)-(i * 50, 500)
        Pic.Line (0, i * 50)-(500, i * 50)
        Pic.Line (i * 50 + 49, 0)-(i * 50 + 49, 500)
        Pic.Line (0, i * 50 + 49)-(500, i * 50 + 49)
    Next
    For i = 0 To 9
        For j = 0 To 9
            If Status.Dat(i, j) = 1 Then
                Pic.CurrentX = i * 50 + 8
                Pic.CurrentY = j * 50 + 10
                Select Case Map0.Dat(i, j)
                Case 0 'empty
                    Pic.ForeColor = vbBlack
                    Pic.Print "¡¤"
                Case 1 'body
                    Pic.ForeColor = vbBlue
                    Pic.Print "¡Á"
                Case 2 'head
                    Pic.ForeColor = vbRed
                    Pic.Print "¡Ì"
                End Select
            End If
        Next
    Next
End Sub

Function Design(Title As String) As Boolean
    SendText = Title
    FrmDesign.Show 1
    Design = ReturnOK
End Function

Function SearchMap(Map0 As Map) As Integer
    Dim i As Integer, j As Integer, s As Integer
    Dim Equal As Boolean
    For s = 0 To 11967
        Equal = True
        For i = 0 To 9
            For j = 0 To 9
                If Stages(s).Dat(i, j) <> Map0.Dat(i, j) Then Equal = False
            Next
        Next
        If Equal Then
            SearchMap = s
            Exit Function
        End If
    Next
    SearchMap = -1
End Function

Function IsWin(Map0 As Map, Status As Map) As Boolean
    Dim i As Integer, j As Integer, Count As Integer
    For i = 0 To 9
        For j = 0 To 9
            If Status.Dat(i, j) = 1 And Map0.Dat(i, j) = 2 Then Count = Count + 1
        Next
    Next
    IsWin = (Count = 3)
End Function

Sub CalcMatch(Map0 As Map, Status As Map)
    Dim s As Integer, i As Integer, j As Integer
    Dim Match As Boolean
    SolutionCount = 0
    For s = 0 To 11967
        Match = True
        For i = 0 To 9
            For j = 0 To 9
                If Status.Dat(i, j) = 1 Then
                    If Map0.Dat(i, j) <> Stages(s).Dat(i, j) Then
                        Match = False
                    End If
                End If
            Next
        Next
        If Match Then
            Solution(SolutionCount) = s
            SolutionCount = SolutionCount + 1
        End If
    Next
End Sub

Sub CalcPossibility(Map0 As Map, Status As Map)
    Dim s As Integer, i As Integer, j As Integer
    CalcMatch Map0, Status
    Erase Possibility.Dat
    For s = 0 To SolutionCount - 1
        For i = 0 To 9
            For j = 0 To 9
                If Stages(Solution(s)).Dat(i, j) = 2 Then Possibility.Dat(i, j) = Possibility.Dat(i, j) + 1
            Next
        Next
    Next
End Sub

Sub GreyMap(Pic As PictureBox, Status As Map)
    Dim i As Integer, j As Integer
    For i = 0 To 9
        For j = 0 To 9
            If Status.Dat(i, j) = 1 Then ColorBox Pic, i, j, &H777777
        Next
    Next
End Sub

Sub ColorBox(Pic As PictureBox, X As Integer, Y As Integer, Clr As Long)
    Pic.Line (X * 50, Y * 50)-(X * 50 + 49, Y * 50 + 49), Clr, BF
End Sub

Sub ShowMark(Pic As PictureBox, Mark As Map, Status As Map)
    Dim i As Integer, j As Integer
    Pic.ForeColor = &HBBBBBB
    For i = 0 To 9
        For j = 0 To 9
            If Status.Dat(i, j) = 0 Then
                Pic.CurrentX = i * 50 + 8
                Pic.CurrentY = j * 50 + 10
                Select Case Mark.Dat(i, j)
                Case 2 'empty
                    Pic.Print "¡¤"
                Case 1 'body
                    Pic.Print "¡Á"
                Case 3 'head
                    Pic.Print "¡Ì"
                End Select
            End If
        Next
    Next
End Sub
