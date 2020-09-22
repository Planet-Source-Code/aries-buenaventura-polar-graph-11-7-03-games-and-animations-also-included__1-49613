VERSION 5.00
Begin VB.Form frmFour 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test #5"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3765
   ControlBox      =   0   'False
   FillColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDetails 
      Caption         =   "&Details"
      Height          =   315
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4620
      Width           =   735
   End
   Begin VB.CommandButton cmdSpin 
      Caption         =   "&Spin"
      Default         =   -1  'True
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   4320
      Width           =   735
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1620
      Top             =   2040
   End
   Begin VB.PictureBox picWindow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
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
      Begin VB.Label lblDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Radius            = "
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
         Tag             =   "f(t)                   = "
         Top             =   180
         Visible         =   0   'False
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmFour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DURATION = 30
Private Const RADIUS = 100

Private Type WheelInfo
    Angle As Single
    Color As Long
    Kind  As String
End Type

Dim OffsetAngle As Single
Dim Delay       As Single
Dim Speed       As Single
Dim WinW        As Long
Dim WinH        As Long
Dim Wheel(1 To 12) As WheelInfo

Private Sub chkDetails_Click()
    Dim i As Integer
    
    For i = lblDetails.LBound To lblDetails.UBound
        lblDetails(i).Visible = CBool(chkDetails.Value)
    Next i
End Sub

Private Sub cmdSpin_Click()
    If tmrTimer.Enabled Then Exit Sub
    
    Call Randomize
    
    Delay = DURATION
    Speed = CSng(Rnd * 6) + 6 ' set maximum speed of the wheel when spin
    tmrTimer.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    ' starting angle
    OffsetAngle = 360 / UBound(Wheel())
    
    For i = LBound(Wheel()) To UBound(Wheel())
        Wheel(i).Angle = (i * OffsetAngle) Mod 360 + 15
        Wheel(i).Kind = Chr$(i + 64)
        If i Mod 4 = 0 Then
            Wheel(i).Color = vbRed
        ElseIf i Mod 4 = 1 Then
            Wheel(i).Color = vbBlue
        ElseIf i Mod 4 = 2 Then
            Wheel(i).Color = vbCyan
        Else
            Wheel(i).Color = vbYellow
        End If
    Next i
    
    WinW = picWindow.ScaleWidth
    WinH = picWindow.ScaleHeight

    Call Render
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If tmrTimer.Enabled Then tmrTimer.Enabled = False
End Sub

Private Sub tmrTimer_Timer()
    If Delay <= 0 Then
        Speed = Speed - 1
        Delay = DURATION * (Speed * 0.1)
    Else
        Delay = Delay - 1
    End If
    
    If Speed > 0 Then
        Call Render
    Else
        Dim i As Integer
        
        For i = LBound(Wheel()) To UBound(Wheel())
            If (Round(Wheel(i).Angle) + OffsetAngle / 2 >= 75) And _
               (Round(Wheel(i).Angle) + OffsetAngle / 2 <= 105) Then
                    picWindow.Line (0, WinH - picWindow.TextHeight(Wheel(i).Kind))- _
                                   (WinW, WinH), Wheel(i).Color, BF
                    
                    picWindow.ForeColor = Wheel((i + 1) Mod UBound(Wheel) + 1).Color
                    picWindow.CurrentX = (WinW - picWindow.TextWidth(Wheel(i).Kind)) / 2
                    picWindow.CurrentY = WinH - picWindow.TextHeight(Wheel(i).Kind)
                    picWindow.Print Wheel(i).Kind
            End If
        Next i
        
        tmrTimer.Enabled = False
    End If
End Sub

Private Sub Render()
    Dim i  As Integer
    Dim cx As Integer
    Dim cy As Integer
    Dim px As Single
    Dim py As Single
    Dim Radian As Single
    
    cx = WinW / 2
    cy = WinH / 2
    
    picWindow.Cls
    picWindow.ForeColor = vbBlack
    picWindow.FillStyle = vbFSTransparent
    picWindow.Circle (cx, cy), RADIUS
    
    For i = LBound(Wheel()) To UBound(Wheel())
        Radian = Wheel(i).Angle * PI / 180
        
        px = RADIUS * Cos(Radian)
        py = RADIUS * -Sin(Radian)
        
        picWindow.Line (cx, cy)-(cx + px, cy + py)
        picWindow.FillStyle = vbSolid
        picWindow.FillColor = Wheel((i + 1) Mod UBound(Wheel) + 1).Color
        picWindow.Circle (cx + px, cy + py), 4
             
        Radian = (Wheel(i).Angle + OffsetAngle) * PI / 180
        
        px = (RADIUS - 10) * Cos(Radian)
        py = (RADIUS - 10) * -Sin(Radian)
        
        Wheel(i).Angle = Wheel(i).Angle + Speed
        If Wheel(i).Angle > 360 Then
            Wheel(i).Angle = Wheel(i).Angle Mod 360
        End If
    Next i
    
    For i = LBound(Wheel()) To UBound(Wheel())
        Radian = (Wheel(i).Angle + OffsetAngle / 2 - Speed) * PI / 180
        
        px = (RADIUS - 20) * Cos(Radian)
        py = (RADIUS - 20) * -Sin(Radian)
            
        AJBFloodFill picWindow.hDC, cx + px, cy + py, Wheel(i).Color
        
        picWindow.ForeColor = Wheel((i + 1) Mod UBound(Wheel) + 1).Color
        picWindow.CurrentX = (WinW - picWindow.TextWidth(Wheel(i).Kind)) / 2 + px
        picWindow.CurrentY = (WinH - picWindow.TextHeight(Wheel(i).Kind)) / 2 + py
        picWindow.Print Wheel(i).Kind
    Next i
    
    picWindow.ForeColor = vbBlack
    picWindow.FillColor = vbMagenta
    picWindow.FillStyle = vbFSSolid
    picWindow.Circle (cx, cy), 5
    
    picWindow.DrawWidth = 5
    picWindow.Line (cx, cy - 95)-(cx, cy - 115), &H80FF&
    picWindow.DrawWidth = 1
    
    lblDetails(0).Caption = lblDetails(0).Tag & "None"
    lblDetails(1).Caption = lblDetails(1).Tag & RADIUS
    lblDetails(2).Caption = lblDetails(2).Tag & "Divisible by 30 (ex. 0°,30°,60°,90°...)"
    lblDetails(3).Caption = lblDetails(3).Tag & "Starting Angle + 30°"
    lblDetails(4).Caption = lblDetails(4).Tag & Speed
    lblDetails(5).Caption = lblDetails(5).Tag & "it's up to you!"
    lblDetails(6).Caption = lblDetails(6).Tag & "it's up to you!"
End Sub
