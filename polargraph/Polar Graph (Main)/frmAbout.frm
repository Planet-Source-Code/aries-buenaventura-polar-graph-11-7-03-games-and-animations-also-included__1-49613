VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Polar Graph 1.0"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   335
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4740
      TabIndex        =   8
      Top             =   4620
      Width           =   1155
   End
   Begin VB.PictureBox picTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   5910
      TabIndex        =   3
      Top             =   0
      Width           =   5910
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Polar Graph"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   780
         Index           =   2
         Left            =   1005
         TabIndex        =   11
         Top             =   30
         Width           =   3750
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Polar Graph"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   780
         Index           =   0
         Left            =   1035
         TabIndex        =   9
         Top             =   15
         Width           =   3750
      End
      Begin VB.Label lblPrompt 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "or"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   1
         Left            =   3420
         TabIndex        =   7
         Top             =   720
         Width           =   165
      End
      Begin VB.Label lblEmail 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ajb2001lg@yahoo.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   1
         Left            =   3600
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   720
         Width           =   1665
      End
      Begin VB.Label lblEmail 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ravemasterharuglory@yahoo.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   0
         Left            =   1020
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   720
         Width           =   2385
      End
      Begin VB.Label lblPrompt 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "email:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Index           =   0
         Left            =   540
         TabIndex        =   4
         Top             =   720
         Width           =   435
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Polar Graph"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   780
         Index           =   1
         Left            =   1080
         TabIndex        =   10
         Top             =   0
         Width           =   3750
      End
   End
   Begin VB.Timer tmrSnow 
      Interval        =   1
      Left            =   60
      Top             =   900
   End
   Begin VB.PictureBox picViewer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   237
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   390
      TabIndex        =   0
      Top             =   960
      Width           =   5910
      Begin VB.Label lblAuthor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Programmed by: Aris Buenaventura"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   1500
         TabIndex        =   2
         Top             =   3120
         Width           =   2490
      End
      Begin VB.Label lblSchool 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Datapro Computer College, Inc. (Recto, Manila)"
         ForeColor       =   &H00C0FFC0&
         Height          =   195
         Left            =   1080
         TabIndex        =   1
         Top             =   3300
         Width           =   3345
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_SNOW_BALL = 200

Private Type SnowInfo
    curX   As Integer
    curY   As Integer
    Color  As Long
    Radius As Integer
    Speed  As Integer
    Weight As Integer
End Type

Dim SnowBallColor()      As Variant
Dim Snow(MAX_SNOW_BALL)  As SnowInfo
Dim IsWingdingsInstalled As Boolean

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim PBuffer  As New StdPicture
    Dim pw       As Long
    Dim ph       As Long
    Dim FilePath As String
    On Error Resume Next
    
    FilePath = App.Path & "\Bitmap\Background.JPG"
    If Dir$(FilePath) <> "" Then
        Set PBuffer = LoadPicture(FilePath)
        pw = ScaleX(PBuffer.Width, vbHimetric, vbPixels)
        ph = ScaleY(PBuffer.Height, vbHimetric, vbPixels)
    
        picViewer.PaintPicture PBuffer, 0, 0, picViewer.ScaleWidth, picViewer.ScaleHeight, _
                                        0, 0, pw, ph, vbSrcCopy
                                           
        Set picViewer.Picture = picViewer.Image
    End If
    
    Set lblEmail(0).MouseIcon = LoadResPicture(101, vbResCursor)
    Set lblEmail(1).MouseIcon = LoadResPicture(101, vbResCursor)
    
    Dim i As Integer
    
    IsWingdingsInstalled = False
    For i = 0 To Screen.FontCount
        If UCase$(Screen.Fonts(i)) = "WINGDINGS" Then
            IsWingdingsInstalled = True
        End If
    Next i
    
    Call DrawGradient
End Sub

Private Sub Form_Resize()
    Dim xmid As Integer
    
    xmid = (picTitle.ScaleWidth - (lblPrompt(0).Width + lblPrompt(1).Width + _
            lblEmail(0).Width + lblEmail(1).Width)) / 2
    lblPrompt(0).Left = xmid
    lblEmail(0).Left = xmid + lblPrompt(0).Width
    lblPrompt(1).Left = xmid + lblPrompt(0).Width + lblEmail(0).Width
    lblEmail(1).Left = xmid + lblPrompt(0).Width + lblEmail(0).Width + _
                       lblPrompt(1).Width
                       
    lblAuthor.Move (picViewer.ScaleWidth - lblAuthor.Width) / 2, _
                    picViewer.ScaleHeight - lblAuthor.Height - lblSchool.Height - 2
    lblSchool.Move (picViewer.ScaleWidth - lblSchool.Width) / 2, _
                    picViewer.ScaleHeight - lblSchool.Height - 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAbout = Nothing
End Sub

Private Sub lblEmail_Click(Index As Integer)
    On Error Resume Next
    
    Call ShellExecute(0, "open", "mailto:" & lblEmail(Index).Caption, 0, 0, 0)
End Sub

Private Sub lblEmail_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set lblEmail(Index).MouseIcon = LoadResPicture(102, vbResCursor)
End Sub

Private Sub lblEmail_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set lblEmail(Index).MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub PlaySnowAni(ByVal MaxSnowBalls As Integer, _
                        ByVal MaxRadius As Integer, _
                        ByVal MaxSpeed As Integer, _
                        ByVal MaxWeight As Integer, _
                        ByVal MaxWindVelocity As Integer, _
                        ByVal MaxWindLength As Integer)
                
    Static bInit As Boolean
    
    Dim i      As Integer
    Dim WinSzX As Long
    Dim WinSzY As Long
    
    Static WindVel    As Integer ' Wind Velocity
    Static WindLen    As Integer ' Wind Length
    Static OldWindLen As Integer ' Old Wind Length
    
    WinSzX = picViewer.ScaleWidth
    WinSzY = picViewer.ScaleHeight
    
    If Not bInit Then
        WindVel = 0
        WindLen = 0
        OldWindLen = CInt(Rnd * MaxWindLength)
        SnowBallColor = Array(&HFFFFFF, &HF5F5F5, &HEBEBEB, &HE1E1E1, &HCDCDCD)
        
        For i = LBound(Snow()) To UBound(Snow())
            Snow(i).curX = CInt(Rnd * WinSzX)
            Snow(i).curY = CInt(Rnd * WinSzY)
            Snow(i).Color = CLng(SnowBallColor(Rnd * UBound(SnowBallColor)))
            Snow(i).Radius = CInt(Rnd * MaxRadius) + 1
            Snow(i).Speed = CInt(Rnd * MaxSpeed) + MaxSpeed / 2
            Snow(i).Weight = CInt(Rnd * MaxWeight) + 1
            DrawSnowBall Snow(i).curX, Snow(i).curY, _
                         Snow(i).Radius, Snow(i).Color
        Next i
        bInit = True
    Else
        Dim nStat   As Integer
        Dim OffSetX As Integer
        Dim OffsetY As Integer
        
        Static bVal    As Integer
        
        picViewer.Cls
        For i = 0 To MaxSnowBalls
            Snow(i).curX = Snow(i).curX + WindVel - Snow(i).Weight
            Snow(i).curY = Snow(i).curY + Snow(i).Speed + Snow(i).Weight
            
            If Snow(i).curX < -Snow(i).Radius Then
                nStat = 0
            ElseIf Snow(i).curX > WinSzX + Snow(i).Radius Then
                nStat = 1
            ElseIf Snow(i).curY < -Snow(i).Radius Then
                nStat = 2
            ElseIf Snow(i).curY > WinSzY + Snow(i).Radius Then
                nStat = 3
            Else
                nStat = -1
            End If
            
            If nStat <> -1 Then
                Snow(i).Radius = CInt(Rnd * MaxRadius) + 1
            End If
            
            Select Case nStat
            Case Is = 0
                OffSetX = WinSzX + Snow(i).Radius
                OffsetY = CInt(Rnd * WinSzY)
            Case Is = 1
                OffSetX = -Snow(i).Radius
                OffsetY = CInt(Rnd * WinSzY)
            Case Is = 2
                OffSetX = CInt(Rnd * WinSzX)
                OffsetY = WinSzY + Snow(i).Radius
            Case Is = 3
                OffSetX = CInt(Rnd * WinSzX)
                OffsetY = -Snow(i).Radius
            End Select
        
            If nStat <> -1 Then
                Snow(i).curX = OffSetX
                Snow(i).curY = OffsetY
                Snow(i).Color = CLng(SnowBallColor(Rnd * UBound(SnowBallColor)))
                Snow(i).Speed = CInt(Rnd * MaxSpeed) + MaxSpeed / 2
                Snow(i).Weight = CInt(Rnd * MaxWeight) + 1
            End If
            
            If Snow(i).Radius = 3 Then
                If IsWingdingsInstalled Then
                    picViewer.FontSize = 5
                    picViewer.CurrentX = Snow(i).curX
                    picViewer.CurrentY = Snow(i).curY
                    picViewer.FontName = "Wingdings"
                    picViewer.Print "X"
                Else
                    DrawSnowBall Snow(i).curX, Snow(i).curY, _
                                 Snow(i).Radius, Snow(i).Color
                End If
            Else
                DrawSnowBall Snow(i).curX, Snow(i).curY, _
                             Snow(i).Radius, Snow(i).Color
            End If
        Next i
        
        If bVal Then
            If OldWindLen = WindLen Then WindVel = WindVel + 1
        Else
            If OldWindLen = WindLen Then WindVel = WindVel - 1
        End If
        
        If OldWindLen = WindLen Then
            WindLen = 0
            OldWindLen = CInt(Rnd * MaxWindLength)
        Else
            WindLen = WindLen + 1
        End If
        
        If (WindVel = MaxWindVelocity) And bVal Then
            bVal = False
        ElseIf (WindVel = -MaxWindVelocity) And Not bVal Then
            bVal = True
        End If
    End If
End Sub

Private Sub DrawGradient()
    Dim i     As Integer
    Dim OldSM As Integer
    
    OldSM = picTitle.ScaleMode
    picTitle.ScaleMode = vbUser
    picTitle.ScaleWidth = 255
    picTitle.ScaleHeight = 255
    
    picTitle.Cls
    For i = 0 To 255
        picTitle.Line (0, i)-(picTitle.ScaleWidth, i), _
                      RGB(0, 0, 255 - i)
    Next i
    
    picTitle.ScaleMode = OldSM
End Sub

Private Sub DrawSnowBall(ByVal X As Long, ByVal Y As Long, ByVal Radius, Color As Long)
    picViewer.ForeColor = Color
    picViewer.DrawWidth = IIf(Radius = 0, 1, Radius)
    picViewer.PSet (X, Y)
End Sub

Private Sub tmrSnow_Timer()
    Call PlaySnowAni(MAX_SNOW_BALL, 3, 2, 3, 10, 50)
End Sub
