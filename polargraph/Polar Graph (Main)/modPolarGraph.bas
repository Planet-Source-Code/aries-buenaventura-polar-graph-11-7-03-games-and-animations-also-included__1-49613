Attribute VB_Name = "modPolarGraph"
Option Explicit

Public Type Size
    cx As Long
    cy As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal E As Long, ByVal O As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal Op As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetTextExtentPoint Lib "gdi32.dll" Alias "GetTextExtentPointA" (ByVal hdc As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As Size) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public MathLib As New MathLibrary

Public Sub Main()
    Dim OldTimer As Single
    
    OldTimer = Timer
    
    frmSplash.Show
    frmSplash.Refresh
    
    Load frmMain
    
    ' make sure that the splash screen will be shown
    ' exactly 2 minutes or more.
    Do While Abs(Timer - OldTimer) < 2
        DoEvents
    Loop
    
    Unload frmSplash
    frmMain.Show
End Sub

Public Function CustomFont(ByVal hgt As Long, _
                           ByVal wid As Long, _
                           ByVal escapement As Long, _
                           ByVal orientation As Long, _
                           ByVal wgt As Long, _
                           ByVal is_italic As Long, _
                           ByVal is_underscored As Long, _
                           ByVal is_striken_out As Long, _
                           ByVal face As String) As Long
                           
    Const CLIP_LH_ANGLES = 16

    CustomFont = CreateFont( _
        hgt, wid, escapement, orientation, wgt, _
        is_italic, is_underscored, is_striken_out, _
        0, 0, CLIP_LH_ANGLES, 1, 0, face)
End Function
