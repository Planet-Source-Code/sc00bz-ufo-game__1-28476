VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private X As Long, Y As Long
Private Frame As Long, Done As Boolean
Public Sub BlowUp()
    clsMain.NewExplode X - 11, Y - 17
    Done = True
End Sub
Public Sub CheckColition()
    Dim Rect As Rect
    
    Rect.Bottom = Y + 13
    Rect.Left = X
    Rect.Right = X + 28
    Rect.Top = Y
    If inRect(Base, Rect) Then BlowUp
    If inRect(Rocket, Rect) Then
        BlowUp
        clsMain.KillRocket
    End If
    If Not Done And Y = frmGame.picMain.ScaleHeight - 20 Then
        clsMain.EndGame
    End If
End Sub
Public Sub Draw()
    CopyPic frmGame.picMain.hDC, X, Y, 28, 13, frmGame.picUfoMask.hDC, frmGame.picUfo.hDC, 0, 0, 28 * Frame, 0
End Sub
Public Function isDone() As Boolean
    isDone = Done
End Function
Public Sub Move()
    Dim tX As Long
    
    tX = Abs(X - frmGame.picMain.ScaleWidth + 28)
    If tX < 2 Then
        X = X + RandNumLng(-2, tX)
    ElseIf X < 2 Then
        X = X + RandNumLng(Int(-X), 2)
    Else
        X = X + RandNumLng(-2, 2)
    End If
    Y = Y + 2
    If Y >= frmGame.picMain.ScaleHeight - 20 Then
        Y = frmGame.picMain.ScaleHeight - 20
    End If
    Frame = (Frame + 1) Mod 24
End Sub
Private Sub Class_Initialize()
    X = RandNumLng(0, frmGame.picMain.ScaleWidth - 28)
    Y = -13
End Sub
