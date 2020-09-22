VERSION 5.00
Begin VB.Form frmTwo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test #2"
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
      Left            =   120
      Top             =   3720
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
Attribute VB_Name = "frmTwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ObjectInfo
    Angle As Integer
    Kind  As String
End Type

Dim WinW As Long
Dim WinH As Long
Dim Object(0 To 72) As ObjectInfo

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
    
    For i = LBound(Object()) To UBound(Object())
        Object(i).Angle = i * 20
        Object(i).Kind = Chr$(i + 120)
    Next i
    
    WinW = picWindow.ScaleWidth
    WinH = picWindow.ScaleHeight
End Sub

Private Sub Render()
    Dim i      As Integer
    Dim px     As Single
    Dim py     As Single
    Dim Radian As Single
    
    picWindow.Cls
    For i = LBound(Object()) To UBound(Object())
        Radian = Object(i).Angle * PI / 180
    
        px = 120 * Sin(2 / 3 * Radian) * Cos(Radian)
        py = 120 * Sin(2 / 3 * Radian) * -Sin(Radian)
        
        picWindow.CurrentX = (WinW - picWindow.TextWidth(Object(i).Kind)) / 2 + px
        picWindow.CurrentY = (WinH - picWindow.TextHeight(Object(i).Kind)) / 2 + py
        picWindow.Print Object(i).Kind
        
        Object(i).Angle = Object(i).Angle + 1
        If Object(i).Angle > 1440 Then
            Object(i).Angle = Object(i).Angle Mod 1440
        End If
    Next i
    
    lblDetails(0).Caption = lblDetails(0).Tag & "None"
    lblDetails(1).Caption = lblDetails(1).Tag & "120*Sin(2/3*t)"
    lblDetails(2).Caption = lblDetails(2).Tag & "0°"
    lblDetails(3).Caption = lblDetails(3).Tag & "1440°"
    lblDetails(4).Caption = lblDetails(4).Tag & "1"
    lblDetails(5).Caption = lblDetails(5).Tag & "it's up to you!"
    lblDetails(6).Caption = lblDetails(6).Tag & "it's up to you!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If tmrTimer.Enabled Then tmrTimer.Enabled = False
End Sub

Private Sub tmrTimer_Timer()
    Call Render
End Sub
