VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private clsBase    As New clsBase
Private colExplode As New Collection
Private colUfo     As New Collection
Private HName      As String, Level    As Long
Private MLoop      As Boolean, Running As Boolean
Private Score      As Currency, Paused As Boolean
Private MX         As Long
Private Function AimWay(PW As String) As String
    Dim A As Long, B As Byte, C As Byte, D As Byte
    
    AimWay = Space$(Len(PW))
    For A = 1 To Len(PW)
        C = A Mod 16
        B = Asc(Mid$(PW, A, 1)) \ 16
        If C = 4 Or C = 9 Or C = 11 Or C = 14 Then B = B + IIf(B Mod 2 < 1, 1, -1)
        If C = 5 Or C = 10 Or C = 15 Or C = 0 Then B = B + IIf(B Mod 4 < 2, 2, -2)
        If C = 1 Or C = 6 Or C = 11 Or C = 12 Then B = B + IIf(B Mod 8 < 4, 4, -4)
        If C = 2 Or C = 7 Or C = 13 Or C = 16 Then B = B + IIf(B < 8, 8, -8)
        If C = 12 Or C = 13 Or C = 14 Or C = 15 Then B = 15 - B
        D = B * 16
        B = Asc(Mid$(PW, A, 1)) Mod 16
        If C = 5 Or C = 7 Or C = 9 Or C = 10 Or C = 0 Then B = B + IIf(B Mod 2 < 1, 1, -1)
        If C = 1 Or C = 6 Or C = 9 Or C = 11 Or C = 12 Then B = B + IIf(B Mod 4 < 2, 2, -2)
        If C = 2 Or C = 7 Or C = 8 Or C = 9 Or C = 13 Then B = B + IIf(B Mod 8 < 4, 4, -4)
        If C = 3 Or C = 12 Or C = 14 Then B = B + IIf(B < 8, 8, -8)
        If C = 8 Or C = 10 Or C = 11 Then B = 15 - B
        Mid$(AimWay, A, 1) = Chr$(B + D)
    Next
End Function
Private Function Decript(Str As String) As String
    Str = XorWay(Str)
    Str = AimWay(Str)
    Decript = Str
End Function
Private Sub DrawStuff()
    Dim A As Long
    
    BitBlt frmGame.picMain.hDC, 0, 0, frmGame.picBG.ScaleWidth, frmGame.picBG.ScaleHeight, frmGame.picBG.hDC, 0, 0, vbSrcCopy
    For A = 1 To colUfo.Count
        colUfo(A).Draw
    Next
    clsBase.Draw
    For A = 1 To colExplode.Count
        colExplode(A).Draw
    Next
    frmGame.picMain.Refresh
End Sub
Private Function Encript(Str As String) As String
    Str = AimWay(Str)
    Str = XorWay(Str)
    Encript = Str
End Function
Public Sub EndGame()
    Running = False
End Sub
Public Sub GetHigh()
    On Error Resume Next
    Dim A As Long, FF As Integer, FileData As String
    
    FF = FreeFile
    Open MyPath & "Data.dat" For Binary As FF
        FileData = Decript(Input(LOF(FF), FF))
    Close FF
    If Len(FileData) <> 400 Then Error 1
    If Err.Number <> 0 Then
        Err.Clear
        MakeNewHigh
        Exit Sub
    Else
        For A = 1 To Len(FileData) Step 40
            ANames((A - 1) / 40) = Mid(FileData, A, 25)
            AScores((A - 1) / 40) = Mid(FileData, A + 25, 15)
        Next
    End If
    If Err.Number <> 0 Then
        Err.Clear
        MakeNewHigh
    End If
End Sub
Public Function inMLoop() As Boolean
    inMLoop = MLoop
End Function
Public Function isPause() As Boolean
    isPause = Paused
End Function
Private Sub MakeNewHigh()
    Dim A As Long, B As Long, FF As Integer, FileData As String
    
    FileData = ""
    For A = 0 To 9
        FileData = FileData & Chr(0)
        For B = 1 To 24
            FileData = FileData & Chr(RandNumLng(0, 255))
        Next
        FileData = FileData & Right(String(15, "0") & RandNumCur(0, 922337203685476#), 15)
    Next
    If FileExist(MyPath & "Data.dat") Then Kill MyPath & "Data.dat"
    FF = FreeFile
    Open MyPath & "Data.dat" For Binary As FF
        Put #FF, , Encript(FileData)
    Close FF
    For A = 0 To 9
        ANames(A) = String(25, 0)
        AScores(A) = 0
    Next
End Sub
Public Sub KillRocket()
    clsBase.KillRocket
End Sub
Private Sub MainLoop()
    Dim A As Long, TempScore As String, Delay As Long
    Dim Wave As Long, Killes As Long, UfoDelay As Long, UfoNum As Long, UfoCount As Long
    
    Randomize
    Delay = timeGetTime
    Do While Running
        If UfoDelay > 0 Then UfoDelay = UfoDelay - 1
        If UfoNum < Level And UfoDelay = 0 Then
            colUfo.Add New clsUfo
            UfoNum = UfoNum + 1
            UfoDelay = 25
        End If
        UfoCount = colUfo.Count
        MoveStuff
        If UfoCount <> colUfo.Count Then Killes = Killes + UfoCount - colUfo.Count
        If Killes = Level Then
            Killes = 0
            Wave = Wave + 1
            UfoNum = 0
        End If
        If Wave = 2 Then
            Wave = 0
            Score = Score + 10 * Val(Level)
            Level = Level + 1
        End If
        TempScore = "Score: " & FormatNumber(Score, 0)
        If frmGame.lblScore <> TempScore Then frmGame.lblScore = TempScore
        If frmGame.lblLevel <> "Level: " & Level Then frmGame.lblLevel = "Level: " & Level
        If Level = 257 Then Running = False
        Do
            DoEvents
            If Paused Then Delay = timeGetTime + 500
            If frmGame.WindowState = 1 Then Delay = timeGetTime + 1000
        Loop While timeGetTime - Delay < 25 And Running
        Delay = timeGetTime
        DrawStuff
    Loop
    If Level = 257 Then
        MsgBox "You Win!", vbSystemModal, "You Win!"
    Else
        MsgBox "Game Over!           ", vbSystemModal, "Game Over!"
    End If
    NewHigh
    For A = 1 To colExplode.Count
        colExplode.Remove 1
    Next
    For A = 1 To colUfo.Count
        colUfo.Remove 1
    Next
    clsBase.Reset
    BitBlt frmGame.picMain.hDC, 0, 0, frmGame.picBG.ScaleWidth, frmGame.picBG.ScaleHeight, frmGame.picBG.hDC, 0, 0, vbSrcCopy
    frmGame.picMain.Refresh
    Paused = False
    frmGame.mnuEnd.Enabled = False
    frmGame.mnuPause.Checked = False
    frmGame.mnuPause.Enabled = False
    frmGame.mnuPopPause.Checked = False
    frmGame.mnuPopPause.Enabled = False
    frmGame.mnuNew.Enabled = True
    Level = 1
    Score = 0
    frmGame.lblLevel = "Level: 1"
    frmGame.lblScore = "Score: 0"
End Sub
Public Sub MouseClick()
    If Not Running Then Exit Sub
    clsBase.Fire
End Sub
Public Sub MouseMove(X As Long)
    If Not Running Then Exit Sub
    MX = X
End Sub
Private Sub MoveStuff()
    Dim A As Long
    
    clsBase.Move MX
    For A = 1 To colExplode.Count
        colExplode(A).Move
    Next
    For A = 1 To colUfo.Count
        colUfo(A).Move
    Next
    clsBase.CheckColition
    For A = colExplode.Count To 1 Step -1
        If colExplode(A).isDone Then colExplode.Remove A
    Next
    For A = colUfo.Count To 1 Step -1
        colUfo(A).CheckColition
        If colUfo(A).isDone Then
            colUfo.Remove A
            Score = Score + Level
        End If
    Next
End Sub
Public Sub NewExplode(X As Long, Y As Long)
    colExplode.Add New clsExplode
    colExplode(colExplode.Count).SetInfo X, Y
End Sub
Public Sub NewGame()
    If Running Then Exit Sub
    Running = True
    Level = 1
    Score = 0
    frmGame.lblLevel = "Level: 1"
    frmGame.lblScore = "Score: 0"
    MLoop = True
    MainLoop
    MLoop = False
    If Not frmGame.mnuSndOnAllTime.Checked Then frmGame.MediaPlayer1.Stop
End Sub
Public Sub NewHigh()
    Dim A As Long, B As Long, FF As Integer
    Dim FileData As String, Str As String
    
    If Score = 0 Then Exit Sub
    GetHigh
    For A = 0 To 9
        If Score >= AScores(A) Or Left(ANames(A), 1) = Chr(0) Then
            Str = InputBox("New High Score: " & FormatNumber(Score, 0) & vbNewLine & vbNewLine & "Enter Your Name:", "New High Score", HName)
            If Str = "" Then Exit Sub
            HName = Left$(Str & String(24, 0), 25)
            For B = 8 To A Step -1
                ANames(B + 1) = ANames(B)
                AScores(B + 1) = AScores(B)
            Next
            ANames(A) = HName
            AScores(A) = Score
            Score = 0
            FileData = ""
            For B = 0 To 9
                FileData = FileData & ANames(B)
                FileData = FileData & Right$(String(15, "0") & AScores(B), 15)
            Next
            If FileExist(MyPath & "Data.dat") Then Kill MyPath & "Data.dat"
            FF = FreeFile
            Open MyPath & "Data.dat" For Binary As FF
                Put #FF, , Encript(FileData)
            Close FF
            frmGame.mnuHigh_Click
            Exit Sub
        End If
    Next
End Sub
Public Sub Pause()
    If Not Running Then Exit Sub
    Paused = Not Paused
    frmGame.mnuPause.Checked = Paused
    frmGame.mnuPopPause.Checked = Paused
End Sub
Private Function XorWay(Str As String) As String
    Dim A As Long
    
    XorWay = Space$(Len(Str))
    For A = 1 To Len(Str)
        Mid$(XorWay, A, 1) = Chr$(Asc(Mid$(Key, A, 1)) Xor Asc(Mid$(Str, A, 1)))
    Next
End Function
Private Sub Class_Initialize()
    Level = 1
End Sub
