Attribute VB_Name = "modMain"
Option Explicit

Public Const VK_DOWN As Long = &H28
Public Const VK_LEFT As Long = &H25
Public Const VK_RIGHT As Long = &H27
Public Const VK_UP As Long = &H26
Public Const VK_SPACE As Long = &H20
Public Const VK_LBUTTON As Long = &H1

Public Const SHIP_TOP_SPRITE = 101
Public Const SHIP_TOP_MASK = 102
Public Const SHIP_RIGHT_SPRITE = 103
Public Const SHIP_RIGHT_MASK = 104
Public Const SHIP_LEFT_SPRITE = 105
Public Const SHIP_LEFT_MASK = 106
Public Const BALL_SPRITE = 107
Public Const BALL_MASK = 108
Public Const ASTEROID_SPRITE = 109
Public Const ASTEROID_MASK = 110
Public Const BULLET_SPRITE = 111
Public Const BULLET_MASK = 112

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const FLOODFILLBORDER = 0
Public Const FLOODFILLSURFACE = 1

Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Public Declare Function GetActiveWindow Lib "user32.dll" () As Long
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Declare Function IntersectRect Lib "user32.dll" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Function PI() As Single
    PI = Atn(1) * 4
End Function

Public Function CalcRectArea(ByVal X As Integer, ByVal Y As Integer) As Single
    CalcRectArea = X * Y
End Function

Public Sub AJBFloodFill(ByVal lhDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long)
    Dim lhBrush    As Long
    Dim lhOldBrush As Long
    
    lhBrush = CreateSolidBrush(Color)
    lhOldBrush = SelectObject(lhDC, lhBrush)
    ExtFloodFill lhDC, X, Y, vbWhite, FLOODFILLSURFACE
    SelectObject lhDC, lhOldBrush
    DeleteObject lhBrush
End Sub



