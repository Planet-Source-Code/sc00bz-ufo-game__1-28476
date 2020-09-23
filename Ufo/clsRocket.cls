VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private X As Long, Y As Long
Private Firing As Boolean, Frame As Long
Public Sub BlowUp()
    Y = frmGame.picMain.ScaleHeight - 38
    Firing = False
End Sub
Public Sub CheckColition()
    Rocket.Bottom = Y + 31
    Rocket.Left = X + 5
    Rocket.Right = X + 14
    Rocket.Top = Y
End Sub
Public Sub Draw()
    If Firing Then
        CopyPic frmGame.picMain.hDC, X, Y, 19, 31, frmGame.picRocketMask.hDC, frmGame.picRocket.hDC, 19 * Frame, 0, 19 * Frame, 0
    Else
        CopyPic frmGame.picMain.hDC, X, Y, 19, 26, frmGame.picRocketMask.hDC, frmGame.picRocket.hDC, 19 * Frame, 0, 19 * Frame, 0
    End If
End Sub
Public Sub Fire()
    Firing = True
End Sub
Public Function isFiring() As Boolean
    isFiring = Firing
End Function
Public Sub Move(ByVal MX As Long)
    Frame = (Frame + 1) Mod 2
    MX = MX + 4
    If Firing Then
        Y = Y - 5
        If Y < -31 Then
            Y = frmGame.picMain.ScaleHeight - 38
            Frame = Frame + 2
            Firing = False
        Else
            If X > MX Then
                Frame = Frame + 4
            ElseIf X = MX Then
                Frame = Frame + 2
            End If
        End If
    Else
        Frame = Frame + 2
    End If
    X = MX
End Sub
Public Sub Reset()
    X = frmGame.picMain.ScaleWidth \ 2 - 13
    Y = frmGame.picMain.ScaleHeight - 38
    Firing = False
    Frame = 2
End Sub
Private Sub Class_Initialize()
    X = frmGame.picMain.ScaleWidth \ 2 - 13
    Y = frmGame.picMain.ScaleHeight - 38
    Frame = 2
End Sub
