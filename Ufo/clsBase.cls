VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private clsRocket As New clsRocket
Private X As Long, Y As Long
Private Done As Boolean, Frame As Long
Public Sub CheckColition()
    clsRocket.CheckColition
    If clsRocket.isFiring Then
        Base.Left = X
        Base.Right = X + 27
    Else
        Base.Left = X + 4
        Base.Right = X + 24
    End If
    Base.Bottom = Y + 9
    Base.Top = Y
End Sub
Public Sub Draw()
    If clsRocket.isFiring Then
        CopyPic frmGame.picMain.hDC, X, Y + 4, 27, 4, frmGame.picBase.hDC, frmGame.picBase.hDC, 0, 4, 0, 14
    Else
        CopyPic frmGame.picMain.hDC, X, Y, 27, 5, frmGame.picBase.hDC, frmGame.picBase.hDC, 0, 0, 0, 10
    End If
    clsRocket.Draw
End Sub
Public Sub Fire()
    clsRocket.Fire
End Sub
Public Function isDone() As Boolean
    isDone = Done
End Function
Public Sub KillRocket()
    clsRocket.BlowUp
End Sub
Public Sub Move(ByVal MX As Long)
    MX = MX - 13.5
    If Abs(X - MX) < 5 Then
        X = MX
    ElseIf X < MX Then
        X = X + 5
    Else
        X = X - 5
    End If
    clsRocket.Move X
End Sub
Public Sub Reset()
    X = frmGame.picMain.ScaleWidth / 2 - 13
    Y = frmGame.picMain.ScaleHeight - 16
    Done = False
    Frame = 0
    clsRocket.Reset
End Sub
Private Sub Class_Initialize()
    X = frmGame.picMain.ScaleWidth / 2 - 13
    Y = frmGame.picMain.ScaleHeight - 16
End Sub
