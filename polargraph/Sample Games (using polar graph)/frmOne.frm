VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOne 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test #1"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPause 
      Caption         =   "&Pause"
      Height          =   315
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5220
      Width           =   735
   End
   Begin VB.CheckBox chkHitArea 
      Caption         =   "&Hit Area"
      Height          =   315
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4920
      Width           =   735
   End
   Begin VB.CheckBox chkDetails 
      Caption         =   "&Details"
      Height          =   315
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4620
      Width           =   735
   End
   Begin MSComctlLib.Slider sldRadius 
      Height          =   255
      Left            =   -60
      TabIndex        =   2
      Top             =   4500
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      Min             =   50
      Max             =   100
      SelStart        =   50
      Value           =   50
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   3420
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   315
      Left            =   3000
      TabIndex        =   7
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
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Default         =   -1  'True
         Height          =   375
         Left            =   1380
         TabIndex        =   19
         Top             =   1980
         Width           =   855
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arrow Keys"
         Height          =   195
         Left            =   1380
         TabIndex        =   20
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label lblDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Radius            = "
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   18
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Tag             =   "f(t)                   = "
         Top             =   180
         Visible         =   0   'False
         Width           =   1170
      End
   End
   Begin MSComctlLib.Slider sldRingSpeed 
      Height          =   195
      Left            =   1260
      TabIndex        =   4
      Top             =   5040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   344
      _Version        =   393216
      Min             =   1
      SelStart        =   5
      Value           =   5
   End
   Begin MSComctlLib.Slider sldGameSpeed 
      Height          =   195
      Left            =   1260
      TabIndex        =   6
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   344
      _Version        =   393216
      Min             =   1
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
      TabIndex        =   5
      Top             =   5280
      Width           =   1245
   End
   Begin VB.Label lblBallSpeed 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ring speed (Ball):"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   5040
      Width           =   1245
   End
   Begin VB.Label lblRadius 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Radius:"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   4320
      Width           =   540
   End
End
Attribute VB_Name = "frmOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type AsteroidInfo
    XPos   As Long      ' x position of the asteroid
    YPos   As Long      ' y position of the asteroid
    Width  As Long      ' Asteroid width
    Height As Long      ' Asteroid height
    Speed  As Long      ' Asteroid speed
End Type

Private Type ShipInfo
    XPos    As Long      ' x position of the ship
    YPos    As Long      ' y position of the ship
    Width   As Long      ' Ship width
    Height  As Long      ' Ship height
    CenterX As Long      ' x center of the ship
    CenterY As Long      ' y center of the ship
End Type

Private Type BallInfo
    XPos   As Single     ' x position of the ball
    YPos   As Single     ' y position of the ball
    Angle  As Single     ' angle of the ball
    Width  As Long       ' Ball width
    Height As Long       ' Ball height
End Type

Dim WinW     As Integer   ' Window width
Dim WinH     As Integer   ' Window height
Dim OldShipX As Integer   ' controls the position of the ship

Dim Ship             As ShipInfo
Dim Ball(0 To 1)     As BallInfo
Dim Asteroid(0 To 1) As AsteroidInfo

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

Private Sub cmdStart_Click()
    cmdStart.Visible = False
    lblMsg.Visible = False
    tmrTimer.Enabled = True
    picWindow.SetFocus
End Sub

Private Sub Form_Load()
    ' get asteroid width and height
    Asteroid(0).Width = ScaleX(frmMain.picAsteroidSprite.Picture.Width, vbHimetric, vbPixels)
    Asteroid(0).Height = ScaleY(frmMain.picAsteroidSprite.Picture.Height, vbHimetric, vbPixels)
    
    Asteroid(1).Width = ScaleX(frmMain.picAsteroidSprite.Picture.Width, vbHimetric, vbPixels)
    Asteroid(1).Height = ScaleY(frmMain.picAsteroidSprite.Picture.Height, vbHimetric, vbPixels)
    
    ' get ball width and height
    Ball(0).Width = ScaleX(frmMain.picBallSprite.Picture.Width, vbHimetric, vbPixels)
    Ball(0).Height = ScaleY(frmMain.picBallSprite.Picture.Height, vbHimetric, vbPixels)
    
    Ball(1).Width = ScaleX(frmMain.picBallSprite.Picture.Width, vbHimetric, vbPixels)
    Ball(1).Height = ScaleY(frmMain.picBallSprite.Picture.Height, vbHimetric, vbPixels)
    
    ' get ship width and hegiht
    Ship.Width = ScaleX(frmMain.picTShipSprite.Picture.Width, vbHimetric, vbPixels)
    Ship.Height = ScaleX(frmMain.picTShipSprite.Picture.Height, vbHimetric, vbPixels)
    
    ' get window width and height
    WinW = picWindow.ScaleWidth
    WinH = picWindow.ScaleHeight
    
    Asteroid(0).XPos = CInt(Rnd * WinW) - Asteroid(0).Width
    Asteroid(1).XPos = CInt(Rnd * WinW) - Asteroid(1).Width
    Asteroid(0).Speed = CInt(Rnd * 3) + 1
    Asteroid(1).Speed = CInt(Rnd * 3) + 1
    
    Ball(0).Angle = 0    ' starting angle of 1st ball
    Ball(1).Angle = 180  ' starting angle of 2nd ball
    
    Ship.CenterX = (WinW - Ship.Width) / 2   ' get x-center of the ship
    Ship.CenterY = (WinH - Ship.Height) / 2  ' get y-center of the ship
    
    OldShipX = Ship.XPos
End Sub

Private Sub Render()
    Dim i          As Integer ' Iteration
    Dim X          As Single  ' x position of the ship
    Dim Y          As Single  ' y position of the ship
    
    ' get the position of the ship and make sure that
    ' the ship is in the right position
    X = Ship.XPos + Ship.CenterX
    Y = Ship.YPos + Ship.CenterY
    
    picWindow.Cls
    Call AsteroidMovement(0)
    Call AsteroidMovement(1)

    If Asteroid(0).YPos > WinH Then
        Asteroid(0).XPos = CInt(Rnd * WinW) - Asteroid(0).Width
        Asteroid(0).YPos = -Asteroid(0).Width
        Asteroid(0).Speed = CInt(Rnd * 3) + 1
    Else
        Asteroid(0).YPos = Asteroid(0).YPos + _
                           Asteroid(0).Speed + sldGameSpeed.Value
    End If
    
    If Asteroid(1).YPos > WinH Then
        Asteroid(1).XPos = CInt(Rnd * WinW) - Asteroid(1).Width
        Asteroid(1).YPos = -Asteroid(1).Width
        Asteroid(1).Speed = CInt(Rnd * 3) + 1
    Else
        Asteroid(1).YPos = Asteroid(1).YPos + _
                           Asteroid(1).Speed + sldGameSpeed.Value
    End If
    
    Call BallMovement(0) ' draw ball 1 in different angle
    Call BallMovement(1) ' draw ball 2 in different angle
    
    If Ball(0).Angle > 360 Then
        ' since that 360° = 0°, 375° = 15°...by getting
        ' the reminder of a given angle divided by 360°
        ' we can also get the same result.
        ' Ex. 450° mod 360° = 90°
        
        Ball(0).Angle = Ball(0).Angle Mod 360
    Else
        Ball(0).Angle = Ball(0).Angle + sldRingSpeed.Value
    End If
    
    If Ball(1).Angle > 360 Then
        Ball(1).Angle = Ball(1).Angle Mod 360
    Else
        Ball(1).Angle = Ball(1).Angle + sldRingSpeed.Value
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
    
    Call TestCollisions

    OldShipX = Ship.XPos
    
    If CBool(chkHitArea.Value) Then Call ShowHitArea
    
    lblDetails(0).Caption = lblDetails(0).Tag & "None"
    lblDetails(1).Caption = lblDetails(1).Tag & sldRadius.Value
    lblDetails(2).Caption = lblDetails(2).Tag & "(B1 = 0°), (B2 = 180°)"
    lblDetails(3).Caption = lblDetails(3).Tag & "360° (Both)"
    lblDetails(4).Caption = lblDetails(4).Tag & sldRingSpeed.Value
    lblDetails(5).Caption = lblDetails(5).Tag & "less than or equal to 50"
    lblDetails(6).Caption = lblDetails(6).Tag & sldRadius.Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If tmrTimer.Enabled Then tmrTimer.Enabled = False
End Sub

Private Sub sldGameSpeed_Change()
    If Not CBool(chkPause.Value) Then picWindow.SetFocus
End Sub

Private Sub sldRadius_Change()
    If Not CBool(chkPause.Value) Then picWindow.SetFocus
End Sub

Private Sub sldRingSpeed_Change()
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
    
    Call Render
End Sub

Private Sub BallMovement(ByVal Index As Integer)
    Dim px     As Single
    Dim py     As Single
    Dim Radian As Single
    
    ' 1° = PI/180 (radian)
    Radian = Ball(Index).Angle * PI / 180 ' convert this to radian
    
    ' polar graph formula (r = radius)
    ' (Geometery)
    '    x = r CosØ
    '    y = r SinØ
    ' (Computer)
    '    x = r *  CosØ
    '    y = r * -SinØ
    px = sldRadius.Value * Cos(Radian)
    py = sldRadius.Value * -Sin(Radian)
    
    ' get the position of the ball and make sure that
    ' the ball is in the right position
    Ball(Index).XPos = Ship.XPos + (WinW - Ball(Index).Width) / 2 + px
    Ball(Index).YPos = Ship.YPos + (WinH - Ball(Index).Height) / 2 + py
        
    ' draw the ball
    BitBlt picWindow.hDC, Ball(Index).XPos, Ball(Index).YPos, _
                          Ball(Index).Width, Ball(Index).Height, _
           frmMain.picBallMask.hDC, 0, 0, vbSrcAnd
    BitBlt picWindow.hDC, Ball(Index).XPos, Ball(Index).YPos, _
                          Ball(Index).Width, Ball(Index).Height, _
           frmMain.picBallSprite.hDC, 0, 0, vbSrcInvert
End Sub

Private Sub AsteroidMovement(ByVal Index As Integer)
    BitBlt picWindow.hDC, Asteroid(Index).XPos, Asteroid(Index).YPos, _
                         Asteroid(Index).Width, Asteroid(Index).Height, _
           frmMain.picAsteroidMask.hDC, 0, 0, vbSrcAnd
    BitBlt picWindow.hDC, Asteroid(Index).XPos, Asteroid(Index).YPos, _
                          Asteroid(Index).Width, Asteroid(Index).Height, _
           frmMain.picAsteroidSprite.hDC, 0, 0, vbSrcInvert
End Sub

Private Sub TestCollisions()
    Dim i             As Integer ' iteration
    Dim j             As Integer
    Dim DRect         As RECT    ' Destination rect
    Dim SRect         As RECT    ' Ship rect
    Dim ARect(0 To 1) As RECT    ' Asteroid rect
    Dim BRect(0 To 1) As RECT    ' Ball rect
        
    SetRect SRect, Ship.XPos + Ship.CenterX, _
                   Ship.YPos + Ship.CenterY, _
                   Ship.XPos + Ship.CenterX + Ship.Width, _
                   Ship.YPos + Ship.CenterY + Ship.Height
    
    For i = LBound(ARect()) To UBound(ARect())
        SetRect ARect(i), Asteroid(i).XPos, _
                          Asteroid(i).YPos - Asteroid(i).Speed, _
                          Asteroid(i).XPos + Asteroid(i).Width, _
                          Asteroid(i).YPos + Asteroid(i).Height - Asteroid(i).Speed
    Next i
    
    For i = LBound(BRect()) To UBound(BRect())
        SetRect BRect(i), Ball(i).XPos, _
                          Ball(i).YPos, _
                          Ball(i).XPos + Ball(i).Width, _
                          Ball(i).YPos + Ball(i).Height
    Next i
    
    For i = LBound(ARect()) To UBound(ARect())
        If IntersectRect(DRect, SRect, ARect(i)) Then
            ' if ship is been hit then do something
            
            Dim msg As String
            
            msg = "You're ship is been hit!"
            picWindow.CurrentX = (WinW - picWindow.TextWidth(msg)) / 2
            picWindow.CurrentY = (WinH - picWindow.TextHeight(msg)) / 2
            picWindow.Print msg
            
            Debug.Print "***********"
            Debug.Print "Asteroid      : " & i
            Debug.Print "Damage (Ship) : " & CalcRectArea(DRect.Right - DRect.Left, _
                                                          DRect.Bottom - DRect.Top)
        Else
            ' nothing is hit
        End If
    Next i
    
    For j = LBound(ARect()) To UBound(ARect())
        For i = LBound(BRect()) To UBound(BRect())
            If IntersectRect(DRect, BRect(i), ARect(j)) Then
                ' if asteroid is been hit by the ring then do something
                Debug.Print "***********"
                Debug.Print "Ring              : " & i + 1
                Debug.Print "Asteroid          : " & j + 1
                Debug.Print "Damage (Asteroid) : " & CalcRectArea(DRect.Right - DRect.Left, _
                                                                  DRect.Bottom - DRect.Top)
            Else
                ' nothing is hit
            End If
        Next i
    Next j
End Sub

Private Sub ShowHitArea()
    Dim i As Integer
    
    For i = LBound(Asteroid()) To UBound(Asteroid())
        picWindow.Line (Asteroid(i).XPos, Asteroid(i).YPos - Asteroid(i).Speed)- _
                       (Asteroid(i).XPos + Asteroid(i).Width, _
                       (Asteroid(i).YPos + Asteroid(i).Height - _
                        Asteroid(i).Speed)), vbRed, B
    Next i
    
    For i = LBound(Ball()) To UBound(Ball())
        picWindow.Line (Ball(i).XPos, Ball(i).YPos)- _
                       (Ball(i).XPos + Ball(i).Width, _
                       (Ball(i).YPos + Ball(i).Height)), vbRed, B
    Next i

    picWindow.Line (Ship.XPos + Ship.CenterX, Ship.YPos + Ship.CenterY)- _
                   (Ship.XPos + Ship.CenterX + Ship.Width, _
                   (Ship.YPos + Ship.CenterY + Ship.Height)), vbRed, B
End Sub
