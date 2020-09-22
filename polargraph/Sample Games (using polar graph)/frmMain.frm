VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample games"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraFrame 
      Height          =   3795
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test  Five"
         Height          =   435
         Index           =   4
         Left            =   420
         TabIndex        =   19
         Top             =   2880
         Width           =   2235
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test  Four"
         Height          =   435
         Index           =   3
         Left            =   420
         TabIndex        =   18
         Top             =   2400
         Width           =   2235
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   420
         TabIndex        =   17
         Top             =   3360
         Width           =   2235
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test  Three"
         Height          =   435
         Index           =   2
         Left            =   420
         TabIndex        =   16
         Top             =   1920
         Width           =   2235
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test Two"
         Height          =   435
         Index           =   1
         Left            =   420
         TabIndex        =   15
         Top             =   1440
         Width           =   2235
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test One"
         Height          =   435
         Index           =   0
         Left            =   420
         TabIndex        =   14
         Top             =   960
         Width           =   2235
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         Caption         =   " My only purpose here is to show everyone on how to use POLAR GRAPH, so don't expect too much. :)"
         Height          =   675
         Left            =   180
         TabIndex        =   13
         Top             =   240
         Width           =   2820
      End
   End
   Begin VB.PictureBox picBallMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2280
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox picBallSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   1800
      Picture         =   "frmMain.frx":0413
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox picAsteroidMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2340
      Picture         =   "frmMain.frx":0829
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   9
      Top             =   180
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picAsteroidSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1800
      Picture         =   "frmMain.frx":0C3E
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   8
      Top             =   180
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picLShipMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   600
      Picture         =   "frmMain.frx":1088
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   7
      Top             =   780
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picLShipSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   600
      Picture         =   "frmMain.frx":1575
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picRShipSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   1200
      Picture         =   "frmMain.frx":1AB9
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picRShipMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   1200
      Picture         =   "frmMain.frx":202F
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   4
      Top             =   780
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picTShipMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      Picture         =   "frmMain.frx":251B
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   3
      Top             =   780
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picTShipSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      Picture         =   "frmMain.frx":29CF
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picBulletMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   1920
      Picture         =   "frmMain.frx":2F10
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   1
      Top             =   60
      Width           =   90
   End
   Begin VB.PictureBox picBulletSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   1800
      Picture         =   "frmMain.frx":31B2
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   0
      Top             =   60
      Width           =   90
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdTest_Click(Index As Integer)
    Select Case Index
    Case Is = 0
        frmOne.Show vbModal
    Case Is = 1
        frmTwo.Show vbModal
    Case Is = 2
        frmThree.Show vbModal
    Case Is = 3
        frmFour.Show vbModal
    Case Is = 4
        frmFive.Show vbModal
    End Select
End Sub
