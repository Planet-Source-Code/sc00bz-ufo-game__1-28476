VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ufo"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4635
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   309
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picUfo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      Picture         =   "frmGame.frx":1272
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   672
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   10080
   End
   Begin VB.PictureBox picRocket 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   60
      Picture         =   "frmGame.frx":2404
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   114
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.PictureBox picExplode 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   1020
      Picture         =   "frmGame.frx":2BCA
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   420
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.PictureBox picUfoMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      Picture         =   "frmGame.frx":5D84
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picRocketMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   60
      Picture         =   "frmGame.frx":5E02
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   114
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.PictureBox picExplodeMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   1020
      Picture         =   "frmGame.frx":603C
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   420
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.PictureBox picBase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   120
      Picture         =   "frmGame.frx":6D06
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   0
      ScaleHeight     =   457
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   309
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   4635
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   1635
         TabIndex        =   12
         Top             =   3270
         Width           =   1365
      End
   End
   Begin VB.PictureBox picBG 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   0
      ScaleHeight     =   457
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   309
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level: 1"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Score: 0"
      Height          =   195
      Left            =   2340
      TabIndex        =   2
      Top             =   0
      Width           =   2295
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -250
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "En&d Game"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPause 
         Caption         =   "&Pause"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSnd 
         Caption         =   "&Sound"
         Checked         =   -1  'True
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSndOnAllTime 
         Caption         =   "On &All The Time"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuNextSong 
         Caption         =   "Next Song"
         Enabled         =   0   'False
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuVol 
         Caption         =   "Master &Volume"
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHigh 
         Caption         =   "&High Scores"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "&Minimize"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuDash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuMainHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuPopUpExit 
      Caption         =   "PopUpExit"
      Visible         =   0   'False
      Begin VB.Menu mnuPopVol 
         Caption         =   "Master &Volume"
      End
      Begin VB.Menu mnuPopPause 
         Caption         =   "&Pause"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPopMinimize 
         Caption         =   "&Minimize"
      End
      Begin VB.Menu mnuDash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    Dim A As Long
    
    Key = Space$(400)
    For A = 1 To 400
        Mid$(Key, A, 1) = Chr$(RandNumLng(0, 255))
    Next
    Randomize
    MyPath = Replace(App.Path & "\", "\\", "\")
    FadePic picBG, 0, RGB(0, 0, 128)
    picBG.Line (0, picBG.ScaleHeight)-(picBG.ScaleWidth, picBG.ScaleHeight - 7), RGB(0, 255, 0), BF
    BitBlt picMain.hDC, 0, 0, picBG.ScaleWidth, picBG.ScaleHeight, picBG.hDC, 0, 0, vbSrcCopy
    OnTop hWnd, True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mnuExit_Click
End Sub
Private Sub lblMsg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    X = X \ 15 + lblMsg.Left
    Y = Y \ 15 + lblMsg.Top
    If Button = 2 And X >= 0 And Y >= 0 And X < picMain.ScaleWidth And Y < picMain.ScaleHeight Then
        PopupMenu mnuPopUpExit, , X, Y + picMain.Top
    End If
End Sub
Private Sub MediaPlayer1_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
    On Error Resume Next
    
    mnuNextSong.Enabled = MediaPlayer1.PlayState = mpPlaying
    If mnuNextSong.Enabled Or Not mnuSnd.Checked Or Not clsMain.inMLoop And Not mnuSndOnAllTime.Checked Then Exit Sub
    MediaPlayer1.Open RndSnd
    MediaPlayer1.Play
    Do
        DoEvents
    Loop Until MediaPlayer1.PlayState = mpPlaying
    mnuNextSong.Enabled = True
End Sub
Private Sub mnuEnd_Click()
    mnuPause.Enabled = False
    mnuPopPause.Enabled = False
    mnuEnd.Enabled = False
    If clsMain.isPause Then clsMain.Pause
    clsMain.EndGame
End Sub
Private Sub mnuExit_Click()
    mnuFile.Enabled = False
    clsMain.NewHigh
    MediaPlayer1.Stop
    DoEvents
    End
End Sub
Private Sub mnuHelp_Click()
    MsgBox "How To Play:" & vbNewLine & "Move The Mouse To Change The Position Of Your Base and Rocket." & vbNewLine & "Click The Mouse Or Press Space Bar To Fire!", vbSystemModal, "Help!"
End Sub
Public Sub mnuHigh_Click()
    Dim A As Long, B As Long
    
    clsMain.GetHigh
    For A = 0 To 9
        If Left(ANames(A), 1) = Chr(0) Then
            For B = A To 9
                frmHigh.lblName(B) = String(49, "-")
                frmHigh.lblScore(B) = String(34, "-")
            Next
            Exit For
        Else
            frmHigh.lblName(A) = ANames(A)
            frmHigh.lblScore(A) = FormatNumber(AScores(A), 0)
        End If
    Next
    frmHigh.Show 1, frmGame
End Sub
Private Sub mnuMinimize_Click()
    WindowState = 1
End Sub
Private Sub mnuNew_Click()
    On Error Resume Next
    
    mnuNew.Enabled = False
    mnuPause.Enabled = True
    mnuPopPause.Enabled = True
    mnuEnd.Enabled = True
    If mnuSnd.Checked Then
        If MediaPlayer1.PlayState <> mpPlaying Then
            MediaPlayer1.Open RndSnd
            MediaPlayer1.Play
            mnuSndOnAllTime.Enabled = True
            mnuNextSong.Enabled = True
        End If
    End If
    clsMain.NewGame
End Sub
Private Sub mnuNextSong_Click()
    MediaPlayer1.Stop
End Sub
Private Sub mnuPause_Click()
    clsMain.Pause
    If clsMain.isPause Then lblMsg = "Pause"
    lblMsg.Visible = clsMain.isPause
End Sub
Private Sub mnuPopExit_Click()
    mnuExit_Click
End Sub
Private Sub mnuPopMinimize_Click()
    mnuMinimize_Click
End Sub
Private Sub mnuPopPause_Click()
    mnuPause_Click
End Sub
Private Sub mnuPopVol_Click()
    mnuVol_Click
End Sub
Private Sub mnuSndOnAllTime_Click()
    On Error Resume Next
    
    mnuSndOnAllTime.Checked = Not mnuSndOnAllTime.Checked
    If mnuSndOnAllTime.Checked Then
        If MediaPlayer1.PlayState <> mpPlaying Then
            MediaPlayer1.Open RndSnd
            MediaPlayer1.Play
            mnuNextSong.Enabled = True
        End If
    ElseIf Not clsMain.inMLoop Then
        MediaPlayer1.Stop
        mnuNextSong.Enabled = False
    End If
End Sub
Private Sub mnuSnd_Click()
    On Error Resume Next
    
    mnuSnd.Checked = Not mnuSnd.Checked
    If mnuSnd.Checked Then
        If MediaPlayer1.PlayState <> mpPlaying And clsMain.inMLoop Then
            MediaPlayer1.Open RndSnd
            MediaPlayer1.Play
            mnuNextSong.Enabled = True
        End If
        mnuSndOnAllTime.Enabled = True
    Else
        MediaPlayer1.Stop
        mnuNextSong.Enabled = False
        mnuSndOnAllTime.Enabled = False
    End If
End Sub
Private Sub mnuVol_Click()
    If Not clsMain.isPause Then mnuPause_Click
    Shell "SNDVOL32 -t", vbNormalFocus
End Sub
Private Sub picMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then clsMain.MouseClick
End Sub
Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clsMain.MouseMove CLng(X)
End Sub
Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        clsMain.MouseClick
    ElseIf Button = 2 And X >= 0 And Y >= 0 And X < picMain.ScaleWidth And Y < picMain.ScaleHeight Then
        PopupMenu mnuPopUpExit, , X, Y + picMain.Top
    End If
End Sub
