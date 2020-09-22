VERSION 5.00
Begin VB.Form frmFive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test #5"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4170
   ControlBox      =   0   'False
   FillColor       =   &H000080FF&
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
      Left            =   1620
      Top             =   2040
   End
   Begin VB.PictureBox picWindow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
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
Attribute VB_Name = "frmFive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type AtomInfo
    Width  As Integer
    Height As Integer
End Type

Private Type SphereInfo
    Angle  As Single
    Width  As Integer
    Height As Integer
End Type

Dim WinW  As Long
Dim WinH  As Long
Dim Atom(0 To 5)     As AtomInfo
Dim Sphere(0 To 240) As SphereInfo

Private Sub chkDetails_Click()
    Dim i As Integer
    
    For i = lblDetails.LBound To lblDetails.UBound
        lblDetails(i).Visible = CBool(chkDetails.Value)
    Next i
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = LBound(Sphere()) To UBound(Sphere())
        Sphere(i).Angle = i * 3
        Sphere(i).Width = ScaleX(frmMain.picSphereSprite(0).Picture.Width, vbHimetric, vbPixels)
        Sphere(i).Height = ScaleY(frmMain.picSphereSprite(0).Picture.Height, vbHimetric, vbPixels)
    Next i
    
    For i = 0 To frmMain.picAtomSprite.Count - 1
        Atom(i).Width = ScaleX(frmMain.picAtomSprite(i).Picture.Width, vbHimetric, vbPixels)
        Atom(i).Height = ScaleY(frmMain.picAtomSprite(i).Picture.Height, vbHimetric, vbPixels)
    Next i
    
    WinW = picWindow.ScaleWidth
    WinH = picWindow.ScaleHeight
End Sub

Private Sub Render()
    Dim i As Integer
    
    picWindow.Cls
    For i = LBound(Sphere()) To UBound(Sphere())
        Call SphereMovement(i)
    
        Sphere(i).Angle = Sphere(i).Angle + 1
        If Sphere(i).Angle > 720 Then
            Sphere(i).Angle = 0
        End If
    Next i
    
    lblDetails(0).Caption = lblDetails(0).Tag & "None"
    lblDetails(1).Caption = lblDetails(1).Tag & "120*Cos(3/2*t)"
    lblDetails(2).Caption = lblDetails(2).Tag & "0°"
    lblDetails(3).Caption = lblDetails(3).Tag & "720°"
    lblDetails(4).Caption = lblDetails(4).Tag & "1"
    lblDetails(5).Caption = lblDetails(5).Tag & "it's up to you!"
    lblDetails(6).Caption = lblDetails(6).Tag & "it's up to you!"
End Sub

Private Sub SphereMovement(ByVal Index As Integer)
    Dim x  As Integer
    Dim y  As Integer
    Dim px As Single
    Dim py As Single
    Dim curidx As Integer
    Dim Radian As Single
    
    Radian = Sphere(Index).Angle * PI / 180
    
    px = 120 * Cos(3 / 2 * Radian) * Cos(Radian)
    py = 120 * Cos(3 / 2 * Radian) * -Sin(Radian)

    x = (WinW - Sphere(Index).Width) / 2 + px
    y = (WinH - Sphere(Index).Height) / 2 + py
    
    BitBlt picWindow.hdc, x, y, _
                          Sphere(Index).Width, Sphere(Index).Height, _
           frmMain.picSphereMask(0).hdc, 0, 0, vbSrcAnd
    BitBlt picWindow.hdc, x, y, _
                          Sphere(Index).Width, Sphere(Index).Height, _
           frmMain.picSphereSprite(0).hdc, 0, 0, vbSrcInvert
    
    If Index = 40 Then
        curidx = 0
    ElseIf Index = 80 Then
        curidx = 1
    ElseIf Index = 120 Then
        curidx = 2
    ElseIf Index = 160 Then
        curidx = 3
    ElseIf Index = 200 Then
        curidx = 4
    ElseIf Index = 240 Then
        curidx = 5
    Else
        curidx = -1
    End If
    
    If (Index Mod 40 = 0) And (Index <> 0) Then
        x = (WinW - Atom(curidx).Width) / 2 + px
        y = (WinH - Atom(curidx).Height) / 2 + py
    End If
    
    If curidx <> -1 Then Call DrawAtom(curidx, x, y)
End Sub

Private Sub DrawAtom(ByVal Index As Integer, ByVal x As Single, ByVal y As Single)
    BitBlt picWindow.hdc, x, y, _
                          Atom(Index).Width, Atom(Index).Height, _
           frmMain.picAtomMask(Index).hdc, 0, 0, vbSrcAnd
    BitBlt picWindow.hdc, x, y, _
                          Atom(Index).Width, Atom(Index).Height, _
           frmMain.picAtomSprite(Index).hdc, 0, 0, vbSrcInvert
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If tmrTimer.Enabled Then tmrTimer.Enabled = False
End Sub

Private Sub tmrTimer_Timer()
    Call Render
End Sub
