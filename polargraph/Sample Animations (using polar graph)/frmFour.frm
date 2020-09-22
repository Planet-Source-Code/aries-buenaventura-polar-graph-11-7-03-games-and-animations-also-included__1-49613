VERSION 5.00
Begin VB.Form frmFour 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test #4"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDetails 
      Caption         =   "&Details"
      Height          =   315
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4620
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   315
      Left            =   3420
      TabIndex        =   10
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
      DrawWidth       =   3
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
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
         Caption         =   "Radius            = "
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   1
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

Dim A    As Single
Dim WinW As Long
Dim WinH As Long

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
    WinW = picWindow.ScaleWidth
    WinH = picWindow.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If tmrTimer.Enabled Then tmrTimer.Enabled = False
    Set frmFour = Nothing
End Sub

Private Sub Render()
    Dim i  As Integer
    Dim px As Single
    Dim py As Single
    Dim Radian As Single
    
    picWindow.Cls
    For i = 0 To 1440
        Radian = i * PI / 180
        
        px = 100 * Sin(A * Radian) * Cos(Radian)
        py = 100 * Sin(A * Radian) * -Sin(Radian)
        px = 100 * Sin(A * Radian) * Cos(Radian)
        py = 100 * Sin(A * Radian) * -Sin(Radian)

        If i Mod 90 = 0 Then
            picWindow.ForeColor = vbRed
        ElseIf i Mod 90 = 15 Then
            picWindow.ForeColor = vbGreen
        ElseIf i Mod 90 = 30 Then
            picWindow.ForeColor = vbBlue
        ElseIf i Mod 90 = 60 Then
            picWindow.ForeColor = vbYellow
        End If
        
        picWindow.PSet (WinW / 2 + px, WinH / 2 + py)
    Next i
    
    picWindow.ForeColor = vbBlack
    picWindow.FillColor = vbCyan
    picWindow.Circle (WinW / 2 + px, WinH / 2 + py), 5
    
    A = IIf(A > 10, 0, A + 0.01)
    
    lblDetails(0).Caption = lblDetails(0).Tag & "A=" & Round(A, 4)
    lblDetails(1).Caption = lblDetails(1).Tag & "100*Sin(A*t)"
    lblDetails(2).Caption = lblDetails(2).Tag & "0°"
    lblDetails(3).Caption = lblDetails(3).Tag & "1440°"
    lblDetails(4).Caption = lblDetails(4).Tag & "1"
    lblDetails(5).Caption = lblDetails(5).Tag & "it's up to you!"
    lblDetails(6).Caption = lblDetails(6).Tag & "it's up to you!"
End Sub

Private Sub tmrTimer_Timer()
    Call Render
End Sub
