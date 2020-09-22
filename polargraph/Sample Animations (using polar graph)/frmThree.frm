VERSION 5.00
Begin VB.Form frmThree 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test #3"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   4155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDetails 
      Caption         =   "&Details"
      Height          =   315
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4620
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   315
      Left            =   3420
      TabIndex        =   5
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdDirection 
      Caption         =   "Next"
      Height          =   315
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   4320
      Width           =   795
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   315
      Left            =   780
      TabIndex        =   3
      Top             =   4320
      Width           =   795
   End
   Begin VB.CommandButton cmdDirection 
      Caption         =   "Previous"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   4320
      Width           =   795
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Tag             =   "Radius            = "
         Top             =   1080
         Visible         =   0   'False
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmThree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type SelectionInfo
    Angle     As Integer
    Kind      As String
End Type

Dim WinW      As Long
Dim WinH      As Long
Dim Direction As Integer
Dim Selected  As String
Dim TrackSelection     As Integer
Dim Selection(0 To 23) As SelectionInfo

Private Sub chkDetails_Click()
    Dim i As Integer
    
    For i = lblDetails.LBound To lblDetails.UBound
        lblDetails(i).Visible = CBool(chkDetails.Value)
    Next i
End Sub

Private Sub cmdDirection_Click(Index As Integer)
    If tmrTimer.Enabled Then Exit Sub
    
    Select Case Index
    Case Is = 0
        TrackSelection = TrackSelection - 1
    Case Is = 1
        TrackSelection = TrackSelection + 1
    End Select
    
    If TrackSelection = 23 Then
        cmdDirection(1).Enabled = False
    ElseIf TrackSelection = 0 Then
        cmdDirection(0).Enabled = False
    Else
        If Not cmdDirection(0).Enabled Then
            cmdDirection(0).Enabled = True
        End If
        
        If Not cmdDirection(1).Enabled Then
            cmdDirection(1).Enabled = True
        End If
    End If
    
    Direction = Index: tmrTimer.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    Dim i As Integer
    
    For i = LBound(Selection) To UBound(Selection)
        If Selection(i).Angle = 90 Then
            Selected = Selection(i).Kind
            Call Render
        End If
    Next i
End Sub

Private Sub Form_Load()
    Dim i     As Integer
    Dim AtomW As Integer
    Dim AtomH As Integer
    
    For i = LBound(Selection()) To UBound(Selection())
        Selection(i).Angle = (i * 15) + 90
        Selection(i).Kind = Chr$(i + 65)
        Debug.Print Selection(i).Angle
    Next i
    
    WinW = picWindow.ScaleWidth
    WinH = picWindow.ScaleHeight
    
    AtomW = ScaleX(frmMain.picAtomSprite(0).Picture.Width, vbHimetric, vbPixels)
    AtomH = ScaleX(frmMain.picAtomSprite(0).Picture.Height, vbHimetric, vbPixels)
    
    BitBlt picWindow.hdc, (WinW - AtomW) / 2, (WinW - AtomH) / 3, _
                          AtomW, AtomH, _
           frmMain.picAtomMask(2).hdc, 0, 0, vbSrcAnd
    BitBlt picWindow.hdc, (WinW - AtomW) / 2, (WinW - AtomH) / 3, _
                          AtomW, AtomH, _
           frmMain.picAtomSprite(2).hdc, 0, 0, vbSrcInvert
    Set picWindow.Picture = picWindow.Image
    
    Selected = Selection(0).Kind
    
    Call Render
End Sub

Private Sub Render()
    Dim i      As Integer
    Dim px     As Single
    Dim py     As Single
    Dim Radian As Single
    
    picWindow.Cls
    For i = LBound(Selection()) To UBound(Selection())
        Radian = Selection(i).Angle * PI / 180
    
        px = (35 - 210 * Sin(Radian)) * Cos(Radian)
        py = (35 - 210 * Sin(Radian)) * -Sin(Radian)
        
        If Selection(i).Angle = 90 Then
            picWindow.FontSize = 30
            picWindow.ForeColor = vbRed
        Else
            picWindow.FontSize = 10
            picWindow.ForeColor = vbBlack
        End If
        
        picWindow.CurrentX = (WinW - picWindow.TextWidth(Selection(i).Kind)) / 2 + px
        picWindow.CurrentY = WinH / 2 - (WinH - picWindow.TextHeight(Selection(i).Kind)) / 2 + py
        picWindow.Print Selection(i).Kind
        
        Call ShowSelected
    Next i
    
    lblDetails(0).Caption = lblDetails(0).Tag & "None"
    lblDetails(1).Caption = lblDetails(1).Tag & "35-210*sin(t)"
    lblDetails(2).Caption = lblDetails(2).Tag & "0°"
    lblDetails(3).Caption = lblDetails(3).Tag & "360°"
    lblDetails(4).Caption = lblDetails(4).Tag & "1"
    lblDetails(5).Caption = lblDetails(5).Tag & "it's up to you!"
    lblDetails(6).Caption = lblDetails(6).Tag & "it's up to you!"
End Sub

Private Sub ShowSelected()
    picWindow.FontSize = 16
    picWindow.ForeColor = vbRed
    picWindow.CurrentX = (WinW - picWindow.TextWidth(Selected)) / 2
    picWindow.CurrentY = (WinH - picWindow.TextHeight(Selected)) / 3
    picWindow.Print Selected
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If tmrTimer.Enabled Then tmrTimer.Enabled = False
End Sub

Private Sub tmrTimer_Timer()
    Dim i As Integer
    
    Static Counter As Integer
    
    Call Render

    If Counter > 14 Then
        Counter = 0: tmrTimer.Enabled = False
    Else
        Counter = Counter + 1
        
        For i = LBound(Selection()) To UBound(Selection())
            If Direction = 0 Then
                Selection(i).Angle = Selection(i).Angle + 1
            Else
                Selection(i).Angle = Selection(i).Angle - 1
            End If
        Next i
    End If
End Sub
