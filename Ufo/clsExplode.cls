VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExplode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private X As Long, Y As Long
Private Frame As Long, Done As Boolean
Public Sub Draw()
    CopyPic frmGame.picMain.hDC, X, Y, 50, 50, frmGame.picExplodeMask.hDC, frmGame.picExplode.hDC, 50 * Frame, 0, 50 * Frame, 0
End Sub
Public Function isDone() As Boolean
    isDone = Done
End Function
Public Sub Move()
    Frame = (Frame + 1) Mod 10
    If Frame = 0 Then Done = True
End Sub
Public Sub SetInfo(NewX As Long, NewY As Long)
    X = NewX
    Y = NewY
End Sub
