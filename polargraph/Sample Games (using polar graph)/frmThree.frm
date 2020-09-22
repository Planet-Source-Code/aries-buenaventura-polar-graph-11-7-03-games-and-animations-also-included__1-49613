VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmThree 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test #3"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3750
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPause 
      Caption         =   "&Pause"
      Height          =   315
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   735
   End
   Begin VB.CheckBox chkDetails 
      Caption         =   "&Details"
      Height          =   315
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4620
      Width           =   735
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   3480
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   315
      Left            =   3000
      TabIndex        =   4
      Top             =   4320
      Width           =   735
   End
   Begin VB.PictureBox picWindow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   0
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   1
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Default         =   -1  'True
         Height          =   375
         Left            =   1380
         TabIndex        =   0
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left Button - (Fire)"
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   17
         Top             =   2820
         Width           =   1260
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arrow Keys"
         Height          =   195
         Index           =   0
         Left            =   1320
         TabIndex        =   16
         Top             =   2340
         Width           =   795
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Space Bar - (Fire)"
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   15
         Top             =   2580
         Width           =   1230
      End
      Begin VB.Label lblDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Radius            = "
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   14
         Tag             =   "Radius            = "
         Top             =   1080
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lblDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "note: open the POLAR GRAPH to plot the details."
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   7
         Left            =   60
         TabIndex        =   13
         Tag             =   "Unit                 = "
         Top             =   3960
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.Label lblDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit                 = "
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   12
         Tag             =   "Unit                 = "
         Top             =   900
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lblDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coefficient      = "
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Tag             =   "Coefficient      = "
         Top             =   0
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step                = "
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   10
         Tag             =   "Step                = "
         Top             =   720
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lblDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Angle  = "
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   9
         Tag             =   "Ending Angle  = "
         Top             =   540
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lblDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Angle = "
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   8
         Tag             =   "Starting Angle = "
         Top             =   360
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lblDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "f(t)                   = "
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Tag             =   "f(t)                   = "
         Top             =   180
         Visible         =   0   'False
         Width           =   1170
      End
   End
   Begin MSComctlLib.Slider sldGameSpeed 
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   4500
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   344
      _Version        =   393216
      Min             =   1
      Max             =   5
      SelStart        =   1
      Value           =   1
   End
   Begin VB.Label lblGameSpeed 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Game Speed      :"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   4320
      Width           =   1245
   End
End
Attribute VB_Name = "frmThree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ShipInfo
    XPos    As Long      ' x position of the ship
    YPos    As Long      ' y position of the ship
    Width   As Long      ' Ship width
    Height  As Long      ' Ship height
    CenterX As Long      ' x center of the ship
    CenterY As Long      ' y center of the ship
End Type

Dim Ship       As ShipInfo
Dim Amplitude  As Single
Dim IsShipFire As Boolean
Dim WinW       As Integer   ' Window width
Dim WinH       As Integer   ' Window height
Dim OldShipX   As Integer   ' controls the position of the ship
Dim BulletW    As Integer   ' Bullet widht
Dim BulletH    As Integer   ' Bullet height

Private Sub chkDetails_Click()
    Dim i As Integer
    
    For i = lblDetails.LBound To lblDetails.UBound
        lblDetails(i).Visible = CBool(chkDetails.Value)
    Next i
End Sub

Private Sub chkPause_Click()
    tmrTimer.Enabled = Not CBool(chkPause.Value)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Render()
    Dim X          As Single ' x position of the ship
    Dim Y          As Single ' y position of the ship
    
    ' get the position of the ship and make sure that
    ' the ship is in the right position
    X = Ship.XPos + Ship.CenterX
    Y = Ship.YPos + Ship.CenterY
    
    picWindow.Cls
    If IsShipFire Then
        Call ShipFire(Amplitude)
        Amplitude = Amplitude + sldGameSpeed.Value
        If Amplitude > 150 Then Amplitude = 150
    End If
    
    If OldShipX < Ship.XPos Then
        BitBlt picWindow.hDC, X, Y, Ship.Width, Ship.Height, _
               frmMain.picRShipMask.hDC, 0, 0, vbSrcAnd
        BitBlt picWindow.hDC, X, Y, Ship.Width, Ship.Height, _
               frmMain.picRShipSprite.hDC, 0, 0, vbSrcInvert
    ElseIf OldShipX > Ship.XPos Then
        BitBlt picWindow.hDC, X, Y, Ship.Width, Ship.Height, _
               frmMain.picLShipMask.hDC, 0, 0, vbSrcAnd
        BitBlt picWindow.hDC, X, Y, Ship.Width, Ship.Height, _
               frmMain.picLShipSprite.hDC, 0, 0, vbSrcInvert
    Else
        BitBlt picWindow.hDC, X, Y, Ship.Width, Ship.Height, _
               frmMain.picTShipMask.hDC, 0, 0, vbSrcAnd
        BitBlt picWindow.hDC, X, Y, Ship.Width, Ship.Height, _
               frmMain.picTShipSprite.hDC, 0, 0, vbSrcInvert
    End If
    
    OldShipX = Ship.XPos
    
    lblDetails(0).Caption = lblDetails(0).Tag & "A = " & Amplitude
    lblDetails(1).Caption = lblDetails(1).Tag & "A*Sin(16*t)"
    lblDetails(2).Caption = lblDetails(2).Tag & "0°"
    lblDetails(3).Caption = lblDetails(3).Tag & "360°"
    lblDetails(4).Caption = lblDetails(4).Tag & "4"
    lblDetails(5).Caption = lblDetails(5).Tag & "it's up to you!"
    lblDetails(6).Caption = lblDetails(6).Tag & "it's up to you!"
End Sub

Private Sub ShipFire(ByVal Amp As Single)
    Dim px     As Single
    Dim py     As Single
    Dim Angle  As Integer
    Dim Radian As Single
    Dim IsBulletVisible As Boolean
    
    Static sx  As Single
    Static sy  As Single
    
    On Error Resume Next
    
    If Amplitude = 0 Then
        sx = Ship.XPos
        sy = Ship.YPos
    End If
    
    IsBulletVisible = False
    For Angle = 0 To 360 Step 4
        ' 1° = PI/180 (radian)
        Radian = Angle * PI / 180 ' convert this to radian
        
        ' polar graph formula (r = radius)
        ' (Geometery)
        '    x = r CosØ
        '    y = r SinØ
        ' (Computer)
        '    x = r *  CosØ
        '    y = r * -SinØ
        px = Amp * Sin(16 * Radian) * Cos(Radian)
        py = Amp * Sin(16 * Radian) * -Sin(Radian)
       
        px = sx + (WinW - BulletW) / 2 + px
        py = sy + (WinH - BulletH) / 2 + py
        
        If (px >= -BulletW) And (px <= WinW) And _
           (py >= -BulletH) And (py <= WinH) Then
            IsBulletVisible = True
            BitBlt picWindow.hDC, px, py, BulletW, BulletH, _
                   frmMain.picBulletMask.hDC, 0, 0, vbSrcAnd
            BitBlt picWindow.hDC, px, py, BulletW, BulletH, _
                   frmMain.picBulletSprite.hDC, 0, 0, vbSrcInvert
        End If
    Next Angle
    
    If Amplitude = 150 Then IsBulletVisible = False
    IsShipFire = IsBulletVisible
End Sub

Private Sub cmdStart_Click()
    cmdStart.Visible = False
    lblMsg(0).Visible = False
    lblMsg(1).Visible = False
    lblMsg(2).Visible = False
    tmrTimer.Enabled = True
    picWindow.SetFocus
End Sub

Private Sub Form_Load()
    ' get ship width and hegiht
    Ship.Width = ScaleX(frmMain.picTShipSprite.Picture.Width, vbHimetric, vbPixels)
    Ship.Height = ScaleX(frmMain.picTShipSprite.Picture.Height, vbHimetric, vbPixels)
    
    ' get window width and height
    WinW = picWindow.ScaleWidth
    WinH = picWindow.ScaleHeight
    
    Ship.CenterX = (WinW - Ship.Width) / 2   ' get x-center of the ship
    Ship.CenterY = (WinH - Ship.Height) / 2  ' get y-center of the ship
    
    BulletW = ScaleX(frmMain.picBulletSprite.Picture.Width, vbHimetric, vbPixels)
    BulletH = ScaleY(frmMain.picBulletSprite.Picture.Height, vbHimetric, vbPixels)
    
    OldShipX = Ship.XPos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If tmrTimer.Enabled Then tmrTimer.Enabled = False
End Sub

Private Sub sldGameSpeed_Click()
    If Not CBool(chkPause.Value) Then picWindow.SetFocus
End Sub

Private Sub tmrTimer_Timer()
    ' if window lost focus then exit
    If Not CBool(GetActiveWindow) Or _
      (GetActiveWindow <> Me.hWnd) Then Exit Sub
      
    Dim X As Integer
    Dim Y As Integer
    
    X = Ship.CenterX
    Y = Ship.CenterY
    
    ' get keypress
    If GetAsyncKeyState(VK_LEFT) And (Ship.XPos >= -X) Then
        Ship.XPos = Ship.XPos - sldGameSpeed.Value
        If Ship.XPos < -X Then Ship.XPos = -X
    ElseIf GetAsyncKeyState(VK_RIGHT) And (Ship.XPos <= X) Then
        Ship.XPos = Ship.XPos + sldGameSpeed.Value
        If Ship.XPos > X Then Ship.XPos = X
    ElseIf GetAsyncKeyState(VK_UP) And (Ship.YPos >= -Y) Then
        Ship.YPos = Ship.YPos - sldGameSpeed.Value
        If Ship.YPos < -Y Then Ship.YPos = -Y
    ElseIf GetAsyncKeyState(VK_DOWN) And (Ship.YPos <= Y) Then
        Ship.YPos = Ship.YPos + sldGameSpeed.Value
        If Ship.YPos >= Y Then Ship.YPos = Y
    End If
    
    If GetAsyncKeyState(VK_SPACE) Or GetAsyncKeyState(VK_LBUTTON) Then
        If Not IsShipFire Then
            Amplitude = 0: IsShipFire = True
        End If
    End If
    
    Call Render
End Sub
