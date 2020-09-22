VERSION 5.00
Begin VB.Form frmOne 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test #1"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDetails 
      Caption         =   "&Details"
      Height          =   315
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4620
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   315
      Left            =   3420
      TabIndex        =   2
      Top             =   4320
      Width           =   735
   End
   Begin VB.Timer tmrTimer 
      Interval        =   1
      Left            =   1260
      Top             =   2220
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
      ScaleWidth      =   273
      TabIndex        =   0
      Top             =   0
      Width           =   4155
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
      Begin VB.Label lblDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Angle = "
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   9
         Tag             =   "Starting Angle = "
         Top             =   360
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lblDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Angle  = "
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   8
         Tag             =   "Ending Angle  = "
         Top             =   540
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lblDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step                = "
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   7
         Tag             =   "Step                = "
         Top             =   720
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
         TabIndex        =   6
         Tag             =   "Coefficient      = "
         Top             =   0
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit                 = "
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   5
         Tag             =   "Unit                 = "
         Top             =   900
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lblDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "note: open the POLAR GRAPH to plot the details."
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   7
         Left            =   60
         TabIndex        =   4
         Tag             =   "Unit                 = "
         Top             =   3960
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.Label lblDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Radius            = "
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   3
         Tag             =   "Radius            = "
         Top             =   1080
         Visible         =   0   'False
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type AtomInfo
    Angle  As Integer
    XPos   As Single
    YPos   As Single
    Width  As Integer
    Height As Integer
End Type

Dim WinW As Long
Dim WinH As Long
Dim Atom(0 To 3) As AtomInfo

Private Sub chkDetails_Click()
    Dim i As Integer
    
    For i = lblDetails.LBound To lblDetails.UBound
        lblDetails(i).Visible = CBool(chkDetails.Value)
    Next i
    
    picWindow.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    With frmMain.picAtomSprite
        For i = LBound(Atom()) To UBound(Atom())
            Atom(i).Angle = i * 90 ' 0°, 90°, 180°, 270°
            Atom(i).Width = ScaleX(.Item(i).Picture.Width, vbHimetric, vbPixels)
            Atom(i).Height = ScaleY(.Item(i).Picture.Height, vbHimetric, vbPixels)
        Next i
    End With
    
    WinW = picWindow.ScaleWidth
    WinH = picWindow.ScaleHeight
    
    Call DrawPath
End Sub

Private Sub Render()
    Dim i As Integer
    
    picWindow.Cls
    
    For i = LBound(Atom()) To UBound(Atom())
        Call AtomMovement(i)
    
        Atom(i).Angle = Atom(i).Angle + 1
        If Atom(i).Angle > 360 Then
            Atom(i).Angle = Atom(i).Angle Mod 360
        End If
    Next i
    
    lblDetails(0).Caption = lblDetails(0).Tag & "None"
    lblDetails(1).Caption = lblDetails(1).Tag & "120*Cos(2*t)"
    lblDetails(2).Caption = lblDetails(2).Tag & "0°"
    lblDetails(3).Caption = lblDetails(3).Tag & "360°"
    lblDetails(4).Caption = lblDetails(4).Tag & "1"
    lblDetails(5).Caption = lblDetails(5).Tag & "it's up to you!"
    lblDetails(6).Caption = lblDetails(6).Tag & "it's up to you!"
End Sub

Private Sub AtomMovement(ByVal Index As Integer)
    Dim px     As Single
    Dim py     As Single
    Dim Radian As Single
    
    ' 1° = PI/180 (radian)
    Radian = Atom(Index).Angle * PI / 180 ' convert this to radian
    
    ' polar graph formula (r = radius)
    ' (Geometery)
    '    x = r CosØ
    '    y = r SinØ
    ' (Computer)
    '    x = r *  CosØ
    '    y = r * -SinØ
    px = 120 * Cos(2 * Radian) * Cos(Radian)
    py = 120 * Cos(2 * Radian) * -Sin(Radian)
    
    ' get the position of the atom and make sure that
    ' the atom is in the right position
    Atom(Index).XPos = (WinW - Atom(Index).Width) / 2 + px
    Atom(Index).YPos = (WinH - Atom(Index).Height) / 2 + py
        
    ' draw the atom
    BitBlt picWindow.hdc, Atom(Index).XPos, Atom(Index).YPos, _
                          Atom(Index).Width, Atom(Index).Height, _
           frmMain.picAtomMask(Index).hdc, 0, 0, vbSrcAnd
    BitBlt picWindow.hdc, Atom(Index).XPos, Atom(Index).YPos, _
                          Atom(Index).Width, Atom(Index).Height, _
           frmMain.picAtomSprite(Index).hdc, 0, 0, vbSrcInvert
End Sub

Private Sub DrawPath()
    Dim i      As Single
    Dim px     As Single
    Dim py     As Single
    Dim Radian As Single
    
    For i = 0 To 360 Step 0.5
        Radian = i * PI / 180
        
        px = 120 * Cos(2 * Radian) * Cos(Radian)
        py = 120 * Cos(2 * Radian) * -Sin(Radian)
        px = WinW / 2 + px
        py = WinH / 2 + py
        
        If i = 0 Then
            picWindow.CurrentX = px
            picWindow.CurrentY = py
        End If
        
        picWindow.Line (picWindow.CurrentX, picWindow.CurrentY)- _
                       (px, py), &H808080
    Next i
    
    Set picWindow.Picture = picWindow.Image
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If tmrTimer.Enabled Then tmrTimer.Enabled = False
End Sub

Private Sub tmrTimer_Timer()
    Call Render
End Sub
