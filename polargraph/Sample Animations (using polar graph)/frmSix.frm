VERSION 5.00
Begin VB.Form frmSix 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test #6"
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
      TabIndex        =   10
      Top             =   4620
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   315
      Left            =   3420
      TabIndex        =   11
      Top             =   4320
      Width           =   735
   End
   Begin VB.CheckBox chkInOut 
      Caption         =   "Out"
      Height          =   315
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   855
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
         Tag             =   "Radius            = "
         Top             =   1080
         Visible         =   0   'False
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmSix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Msg = "POLAR GRAPH 1.0 by Aris Buenaventura •••"

Dim OffsetAngle As Integer
Dim WinW        As Long
Dim WinH        As Long

Private Sub Render()
    Dim i  As Long
    Dim s  As String * 1
    Dim px As Single
    Dim py As Single
    Dim Angle      As Integer
    Dim Radian     As Single
    Dim Escapement As Integer
    Dim new_font   As Long
    Dim old_font   As Long
    Dim Radius     As Integer
    
    On Error Resume Next
    
    picWindow.Cls
    For i = 1 To 40
        If CBool(chkInOut.Value) Then
            Radius = 120
            Angle = 360 - i * 9 + OffsetAngle
            Escapement = 2700 + (Angle * 10) Mod 3600
        Else
            Radius = 90
            Angle = i * 9 + OffsetAngle
            Escapement = 900 + (Angle * 10) Mod 3600
        End If
        
        Radian = Angle * PI / 180
        
        px = Radius * Cos(Radian)
        py = Radius * -Sin(Radian)
           
        new_font = CustomFont(30, 0, Escapement, 0, _
                              600, False, False, False, _
                              "Microsoft Sans Serif")
        
        old_font = SelectObject(picWindow.hdc, new_font)
        
        If s <= Len(Msg) Then
            s = StrReverse(Mid$(Msg, i, 1))
            TextOut picWindow.hdc, WinW / 2 + px, WinH / 2 + py, s, Len(s)
        End If
        
        SelectObject picWindow.hdc, old_font
        DeleteObject new_font
    Next i
    
    lblDetails(0).Caption = lblDetails(0).Tag & "None"
    lblDetails(1).Caption = lblDetails(1).Tag & Radius
    lblDetails(2).Caption = lblDetails(2).Tag & "0°"
    lblDetails(3).Caption = lblDetails(3).Tag & "360°"
    lblDetails(4).Caption = lblDetails(4).Tag & "2"
    lblDetails(5).Caption = lblDetails(5).Tag & "it's up to you!"
    lblDetails(6).Caption = lblDetails(6).Tag & "it's up to you!"
End Sub

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
End Sub

Private Sub tmrTimer_Timer()
    Call Render
    
    If CBool(chkInOut.Value) Then
        OffsetAngle = OffsetAngle - 2
        If OffsetAngle < 360 Then OffsetAngle = OffsetAngle Mod 360
    Else
        OffsetAngle = OffsetAngle + 2
        If OffsetAngle > 360 Then OffsetAngle = OffsetAngle Mod 360
    End If
End Sub
