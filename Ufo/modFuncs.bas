Attribute VB_Name = "modFuncs"
Option Explicit
Public Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type RGB
    B As Double
    G As Double
    R As Double
End Type

Private SndNum As Long
Public Base As Rect, Rocket As Rect
Public MyPath As String, Key As String, clsMain As New clsMain
Public ANames(9) As String, AScores(9) As Currency

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Function Color2RGB(Color As Long) As RGB
    Color2RGB.B = Color \ 65536
    Color2RGB.G = Color \ 256 Mod 256
    Color2RGB.R = Color Mod 256
End Function
Public Sub CopyPic(DeshDC As Long, DesX As Long, DesY As Long, Width As Long, Height As Long, mhDC As Long, nhDC As Long, mSorX As Long, mSorY As Long, nSorX As Long, nSorY As Long)
    BitBlt DeshDC, DesX, DesY, Width, Height, mhDC, mSorX, mSorY, vbSrcAnd
    BitBlt DeshDC, DesX, DesY, Width, Height, nhDC, nSorX, nSorY, vbSrcPaint
End Sub
Public Sub FadePic(Pic As PictureBox, Color1 As Long, Color2 As Long)
    Dim A As Long, RGB1 As RGB, RGB2 As RGB, CDraw As RGB, CStep As RGB
    
    RGB1 = Color2RGB(Color1)
    RGB2 = Color2RGB(Color2)
    CDraw = Color2RGB(Color1)
    CStep.R = (RGB2.R - RGB1.R) / Pic.ScaleHeight
    CStep.G = (RGB2.G - RGB1.G) / Pic.ScaleHeight
    CStep.B = (RGB2.B - RGB1.B) / Pic.ScaleHeight
    For A = 0 To Pic.ScaleHeight - 1
        Pic.Line (0, A)-(Pic.ScaleWidth, A), RGB(Int(CDraw.R), Int(CDraw.G), Int(CDraw.B))
        CDraw.R = CDraw.R + CStep.R
        CDraw.G = CDraw.G + CStep.G
        CDraw.B = CDraw.B + CStep.B
    Next
End Sub
Public Function FileExist(FilePath As String) As Boolean
    FileExist = Len(Dir(FilePath, 47)) > 0
End Function
Public Function inRect(Rect1 As Rect, Rect2 As Rect) As Boolean
    Dim R1CenterX As Long, R2CenterX As Long, R1CenterY As Long, R2CenterY As Long
    Dim MaxDifX   As Long, MaxDifY   As Long, DifX      As Long, DifY      As Long
    
    R1CenterX = (Rect1.Right - Rect1.Left) / 2
    R1CenterY = (Rect1.Bottom - Rect1.Top) / 2
    R2CenterX = (Rect2.Right - Rect2.Left) / 2
    R2CenterY = (Rect2.Bottom - Rect2.Top) / 2
    MaxDifX = R1CenterX + R2CenterX
    MaxDifY = R1CenterY + R2CenterY
    DifX = Abs(R1CenterX + Rect1.Left - R2CenterX - Rect2.Left)
    DifY = Abs(R1CenterY + Rect1.Top - R2CenterY - Rect2.Top)
    inRect = DifX < MaxDifX And DifY < MaxDifY
End Function
Public Sub OnTop(hWnd As Long, Top As Boolean)
    SetWindowPos hWnd, IIf(Top, -1, -2), 0, 0, 0, 0, 3
End Sub
Public Function RandNumCur(ByVal Num1 As Currency, ByVal Num2 As Currency) As Currency
    Dim A As Currency
    
    If Num1 > Num2 Then
        A = Rnd * (Num1 - Num2 + 1)
        RandNumCur = Int(A) + Num2
    Else
        A = Rnd * (Num2 - Num1 + 1)
        RandNumCur = Int(A) + Num1
    End If
End Function
Public Function RandNumLng(ByVal Num1 As Long, ByVal Num2 As Long) As Long
    Dim A As Double
    
    If Num1 > Num2 Then
        A = Rnd * (Num1 - Num2 + 1)
        RandNumLng = Int(A) + Num2
    Else
        A = Rnd * (Num2 - Num1 + 1)
        RandNumLng = Int(A) + Num1
    End If
End Function
Public Function RndSnd() As String
    If SndNum = 0 Then SndNum = RandNumLng(1, 6)
    RndSnd = MyPath & "snd\" & SndNum & ".mid"
    SndNum = SndNum + 1
    If SndNum > 6 Then SndNum = 1
End Function
