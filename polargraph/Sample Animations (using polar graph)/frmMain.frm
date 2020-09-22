VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample animations"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   3225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTray 
      Height          =   4335
      Left            =   60
      TabIndex        =   14
      Top             =   0
      Width           =   3075
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   435
         Left            =   480
         TabIndex        =   22
         Top             =   3780
         Width           =   2055
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test Six"
         Height          =   435
         Index           =   5
         Left            =   480
         TabIndex        =   20
         Top             =   3300
         Width           =   2055
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test Five"
         Height          =   435
         Index           =   4
         Left            =   480
         TabIndex        =   19
         Top             =   2820
         Width           =   2055
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test Four"
         Height          =   435
         Index           =   3
         Left            =   480
         TabIndex        =   18
         Top             =   2340
         Width           =   2055
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test Three"
         Height          =   435
         Index           =   2
         Left            =   480
         TabIndex        =   17
         Top             =   1860
         Width           =   2055
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test One"
         Height          =   435
         Index           =   0
         Left            =   480
         TabIndex        =   16
         Top             =   900
         Width           =   2055
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test Two"
         Height          =   435
         Index           =   1
         Left            =   480
         TabIndex        =   15
         Top             =   1380
         Width           =   2055
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         Caption         =   " My only purpose here is to show everyone on how to use POLAR GRAPH, so don't expect too much. :)"
         Height          =   675
         Left            =   60
         TabIndex        =   21
         Top             =   180
         Width           =   2820
      End
   End
   Begin VB.PictureBox picAtomMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   2100
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   13
      Top             =   420
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAtomMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   1680
      Picture         =   "frmMain.frx":036E
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   12
      Top             =   420
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAtomSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   2100
      Picture         =   "frmMain.frx":06F6
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAtomSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   1680
      Picture         =   "frmMain.frx":0AAF
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picSphereMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   0
      Left            =   2640
      Picture         =   "frmMain.frx":0E85
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   9
      Top             =   300
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picSphereSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   0
      Left            =   2640
      Picture         =   "frmMain.frx":0EF3
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picAtomMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   1260
      Picture         =   "frmMain.frx":1055
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAtomSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   1260
      Picture         =   "frmMain.frx":13CD
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAtomMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   2
      Left            =   840
      Picture         =   "frmMain.frx":178E
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   5
      Top             =   420
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAtomSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   2
      Left            =   840
      Picture         =   "frmMain.frx":1B4F
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAtomMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   1
      Left            =   420
      Picture         =   "frmMain.frx":1F35
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   3
      Top             =   420
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picAtomSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   1
      Left            =   420
      Picture         =   "frmMain.frx":232A
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picAtomMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":2745
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAtomSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":2AC2
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
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
    Case Is = 5
        frmSix.Show vbModal
    End Select
End Sub

