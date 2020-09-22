Attribute VB_Name = "modMain"
Option Explicit

Public Type Size
    cx As Long
    cy As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal E As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal Op As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Public Const FLOODFILLBORDER = 0
Public Const FLOODFILLSURFACE = 1

Public Function PI() As Single
    PI = Atn(1) * 4
End Function

Public Function CustomFont(ByVal hgt As Long, _
                           ByVal wid As Long, _
                           ByVal Escapement As Long, _
                           ByVal orientation As Long, _
                           ByVal wgt As Long, _
                           ByVal is_italic As Long, _
                           ByVal is_underscored As Long, _
                           ByVal is_striken_out As Long, _
                           ByVal face As String) As Long
                           
    Const CLIP_LH_ANGLES = 16

    CustomFont = CreateFont( _
        hgt, wid, Escapement, orientation, wgt, _
        is_italic, is_underscored, is_striken_out, _
        0, 0, CLIP_LH_ANGLES, 2, 0, face)
End Function

