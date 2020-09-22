VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   467
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   6060
      TabIndex        =   60
      Top             =   5340
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   5100
      TabIndex        =   41
      Top             =   5340
      Width           =   915
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   4140
      TabIndex        =   40
      Top             =   5340
      Width           =   915
   End
   Begin MSComDlg.CommonDialog dlgGraph 
      Left            =   360
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList imlLweight 
      Left            =   960
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   10
      Top             =   4440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSComctlLib.ImageList imlLstyle 
      Left            =   300
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame fraPatterns 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   6795
      Begin VB.Frame fraSample 
         Caption         =   "Sample"
         Height          =   1635
         Left            =   120
         TabIndex        =   4
         Top             =   3060
         Width           =   2595
         Begin VB.PictureBox picPattern 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Wingdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1155
            Left            =   120
            ScaleHeight     =   77
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   153
            TabIndex        =   14
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame fraLine 
         Caption         =   "Line"
         Height          =   2895
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   2595
         Begin VB.CheckBox chkLshading 
            Caption         =   "Allow shading"
            Height          =   315
            Left            =   300
            TabIndex        =   53
            Top             =   2460
            Width           =   1455
         End
         Begin VB.CheckBox chkLconnect 
            Caption         =   "Connect vertices"
            Height          =   195
            Left            =   300
            TabIndex        =   52
            Top             =   2220
            Width           =   1575
         End
         Begin VB.PictureBox picLcolor 
            BackColor       =   &H000000FF&
            Height          =   315
            Left            =   1200
            ScaleHeight     =   255
            ScaleWidth      =   1215
            TabIndex        =   11
            Top             =   1260
            Width           =   1275
            Begin VB.CommandButton cmdLcolor 
               Caption         =   "..."
               Height          =   255
               Left            =   960
               TabIndex        =   12
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.OptionButton optLcustom 
            Caption         =   "Custom"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   600
            Width           =   2115
         End
         Begin VB.OptionButton optLnone 
            Caption         =   "None"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   300
            Width           =   1875
         End
         Begin MSComctlLib.ImageCombo imcLweight 
            Height          =   330
            Left            =   1200
            TabIndex        =   13
            Top             =   1620
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
         End
         Begin MSComctlLib.ImageCombo imcLstyle 
            Height          =   330
            Left            =   1200
            TabIndex        =   30
            Top             =   900
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
         End
         Begin VB.Label lblLine 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Weight: "
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   9
            Top             =   1680
            Width           =   600
         End
         Begin VB.Label lblLine 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color: "
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   8
            Top             =   1320
            Width           =   450
         End
         Begin VB.Label lblLine 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Style: "
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   7
            Top             =   960
            Width           =   435
         End
      End
      Begin VB.Frame fraMarker 
         Caption         =   "Marker"
         Height          =   4575
         Left            =   2820
         TabIndex        =   2
         Top             =   120
         Width           =   3855
         Begin VB.ComboBox cmbMalignv 
            Height          =   315
            ItemData        =   "frmProperties.frx":0000
            Left            =   1800
            List            =   "frmProperties.frx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   4080
            Width           =   1515
         End
         Begin VB.ComboBox cmbMalignh 
            Height          =   315
            ItemData        =   "frmProperties.frx":0026
            Left            =   1800
            List            =   "frmProperties.frx":0033
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   3780
            Width           =   1515
         End
         Begin VB.OptionButton optMcustom 
            Caption         =   "Custom"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   600
            Width           =   1875
         End
         Begin VB.OptionButton optMnone 
            Caption         =   "None"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   300
            Width           =   1875
         End
         Begin VB.Frame fraTray 
            BorderStyle     =   0  'None
            Height          =   2775
            Left            =   300
            TabIndex        =   17
            Top             =   780
            Width           =   3315
            Begin VB.CheckBox chkMautosize 
               Caption         =   "Autosize"
               Height          =   195
               Left            =   420
               TabIndex        =   55
               Top             =   2520
               Width           =   1035
            End
            Begin VB.ComboBox cmbMstyle 
               BeginProperty Font 
                  Name            =   "Wingdings"
                  Size            =   9.75
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   345
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   54
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox txtMsize 
               Height          =   285
               Index           =   1
               Left            =   2880
               TabIndex        =   38
               Text            =   "16"
               Top             =   2460
               Width           =   375
            End
            Begin VB.TextBox txtMsize 
               Height          =   285
               Index           =   0
               Left            =   1980
               TabIndex        =   37
               Text            =   "16"
               Top             =   2460
               Width           =   375
            End
            Begin VB.CommandButton cmdMpicloc 
               Caption         =   "..."
               Height          =   255
               Index           =   1
               Left            =   3000
               TabIndex        =   34
               Top             =   2160
               Width           =   255
            End
            Begin VB.CommandButton cmdMpicloc 
               Caption         =   "..."
               Height          =   255
               Index           =   0
               Left            =   3000
               TabIndex        =   33
               Top             =   1860
               Width           =   255
            End
            Begin VB.TextBox txtMpicloc 
               Height          =   285
               Index           =   1
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   32
               Top             =   2160
               Width           =   1935
            End
            Begin VB.TextBox txtMpicloc 
               Height          =   285
               Index           =   0
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   31
               Top             =   1860
               Width           =   1935
            End
            Begin VB.CheckBox chkMtrans 
               Caption         =   "Transparent"
               Height          =   195
               Left            =   360
               TabIndex        =   25
               Top             =   1260
               Width           =   1935
            End
            Begin VB.OptionButton optMtype 
               Caption         =   "Picture effect"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   24
               Top             =   1620
               Width           =   1275
            End
            Begin VB.OptionButton optMtype 
               Caption         =   "Font effect"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   23
               Top             =   240
               Width           =   1275
            End
            Begin VB.PictureBox picMcolor 
               BackColor       =   &H000000FF&
               Height          =   315
               Left            =   960
               ScaleHeight     =   255
               ScaleWidth      =   1395
               TabIndex        =   19
               Top             =   900
               Width           =   1455
               Begin VB.CommandButton cmdMcolor 
                  Caption         =   "..."
                  Height          =   255
                  Left            =   1140
                  TabIndex        =   20
                  Top             =   0
                  Width           =   255
               End
            End
            Begin VB.CommandButton cmdMfont 
               Caption         =   "..."
               Height          =   255
               Left            =   2460
               TabIndex        =   18
               Top             =   495
               Width           =   255
            End
            Begin VB.Label lblMarker 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Picture:"
               Height          =   195
               Index           =   7
               Left            =   420
               TabIndex        =   42
               Top             =   1920
               Width           =   540
            End
            Begin VB.Label lblMarker 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mask:"
               Height          =   195
               Index           =   6
               Left            =   420
               TabIndex        =   39
               Top             =   2220
               Width           =   435
            End
            Begin VB.Label lblMarker 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Height:"
               Height          =   195
               Index           =   5
               Left            =   2400
               TabIndex        =   36
               Top             =   2520
               Width           =   510
            End
            Begin VB.Label lblMarker 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Width:"
               Height          =   195
               Index           =   4
               Left            =   1500
               TabIndex        =   35
               Top             =   2520
               Width           =   465
            End
            Begin VB.Label lblMarker 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Style: "
               Height          =   195
               Index           =   0
               Left            =   360
               TabIndex        =   22
               Top             =   540
               Width           =   435
            End
            Begin VB.Label lblMarker 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Color: "
               Height          =   195
               Index           =   1
               Left            =   360
               TabIndex        =   21
               Top             =   960
               Width           =   450
            End
         End
         Begin VB.Label lblMarker 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vertical Alignment: "
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   29
            Top             =   4140
            Width           =   1350
         End
         Begin VB.Label lblMarker 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Horizontal Alignment: "
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   28
            Top             =   3840
            Width           =   1530
         End
      End
   End
   Begin VB.Frame fraOptions 
      Height          =   4755
      Left            =   120
      TabIndex        =   43
      Top             =   480
      Visible         =   0   'False
      Width           =   6795
      Begin VB.TextBox txtStep 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1620
         TabIndex        =   51
         Text            =   "1.0"
         Top             =   1020
         Width           =   1455
      End
      Begin VB.TextBox txtDelay 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1620
         TabIndex        =   50
         Text            =   "0"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtAngle 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1620
         TabIndex        =   49
         Text            =   "360"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtAngle 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1620
         TabIndex        =   47
         Text            =   "0"
         Top             =   420
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "To turn on animation set delay > 0. At the main menu, click Graph->Play."
         Height          =   435
         Left            =   1620
         TabIndex        =   61
         Top             =   1800
         Width           =   3765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Range: 0...1,000,000"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   6
         Left            =   3180
         TabIndex        =   59
         Top             =   480
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Range: 0...1,000,000"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   5
         Left            =   3180
         TabIndex        =   58
         Top             =   780
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Range: 0.0001...10,000"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   4
         Left            =   3195
         TabIndex        =   57
         Top             =   1140
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Range: 0...1,000,000"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   3
         Left            =   3180
         TabIndex        =   56
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Delay             :"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   435
         TabIndex        =   48
         Top             =   1500
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Step               :"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   46
         Top             =   1020
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ending angle  : "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   45
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Starting angle : "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   420
         TabIndex        =   44
         Top             =   420
         Width           =   1110
      End
   End
   Begin MSComctlLib.TabStrip tbsProp 
      Height          =   5235
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   9234
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Patterns"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsLoading As Boolean

Private Sub chkLconnect_Click()
    If optLnone.Value Then optLcustom.Value = True
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub chkLshading_Click()
    If optLnone.Value Then optLcustom.Value = True
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub chkMautosize_Click()
    txtMsize(0).Enabled = Not CBool(chkMautosize.Value)
    txtMsize(1).Enabled = Not CBool(chkMautosize.Value)
    If optMnone.Value Then optMcustom.Value = True
    If Not IsLoading Then Call DrawPattern
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub chkMtrans_Click()
    If optMnone.Value Then optMcustom.Value = True
    If Not IsLoading Then Call DrawPattern
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub cmbMalignh_Click()
    If optMnone.Value Then optMcustom.Value = True
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub cmbMalignv_Click()
    If optMnone.Value Then optMcustom.Value = True
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub cmbMstyle_Click()
    If optMnone.Value Then optMcustom.Value = True
    If Not IsLoading Then Call DrawPattern
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLcolor_Click()
    Dim sIndex As Integer
    Dim wIndex As Integer
    On Error GoTo ErrHandler
    
    dlgGraph.ShowColor
    If dlgGraph.Color <> picLcolor.BackColor Then
        picLcolor.BackColor = dlgGraph.Color
    End If
    
    sIndex = imcLstyle.SelectedItem.Index
    wIndex = imcLweight.SelectedItem.Index
    
    Call PopulateLineStyle
    Call PopulateLineWeight
    
    imcLstyle.ComboItems(sIndex).Selected = True
    imcLweight.ComboItems(wIndex).Selected = True
    picLcolor.SetFocus
    
    If optLnone.Value Then optLcustom.Value = True
        
    Call DrawPattern
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
    Exit Sub
    
ErrHandler:
End Sub

Private Sub cmdMcolor_Click()
    On Error GoTo ErrHandler
    
    dlgGraph.ShowColor
    If dlgGraph.Color <> cmbMstyle.ForeColor Then
        cmbMstyle.ForeColor = dlgGraph.Color
        picMcolor.BackColor = dlgGraph.Color
    End If
    
    If optMnone.Value Then optMcustom.Value = True
    
    Call DrawPattern
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
    Exit Sub
    
ErrHandler:
End Sub

Private Sub cmdMfont_Click()
    Dim sIndex As Integer
    On Error Resume Next
    
    dlgGraph.Flags = 1
    dlgGraph.ShowFont
    
    sIndex = cmbMstyle.ListIndex
    
    cmbMstyle.Font.Name = dlgGraph.FontName
    cmbMstyle.Font.Italic = dlgGraph.FontItalic
    picPattern.FontName = dlgGraph.FontName
    picPattern.FontSize = dlgGraph.FontSize
    picPattern.FontBold = dlgGraph.FontBold
    picPattern.FontItalic = dlgGraph.FontItalic
    Call PopulateMarkerStyle
    
    cmbMstyle.ListIndex = sIndex
    cmbMstyle.SetFocus
    
    If optMnone.Value Then optMcustom.Value = True
        
    Call DrawPattern
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub cmdMpicloc_Click(Index As Integer)
    On Error GoTo ErrHandler
    
    With dlgGraph
        .Filter = "Bitmap Files (*.bmp) | *.bmp; |" _
                  & "JPEG (*.JPG,*.JPEG) | *.jpg; *.jpeg; |" _
                  & "GIF (*.GIF) | *.GIF; |" _
                  & "All Picture Files | *.bmp; *.gif; *.jpg; *.jpeg; |" _
                  & "All Files (*.*) | *.*"
        .FilterIndex = 1
        .InitDir = ""
        .Filename = ""
        .ShowOpen
            
        If .Filename <> "" Then
            txtMpicloc(Index).Text = .Filename
        End If
        
        If optMnone.Value Then optMcustom.Value = True
            
        Call DrawPattern
    End With
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
    Exit Sub
    
ErrHandler:
    If Err.Number = 32755 Then ' Cancel Selected
    Else
        MsgBox Err.Description, vbOKOnly Or vbInformation, "Error"
    End If
End Sub

Private Sub cmdApply_Click()
    Dim curpos As Integer
    Dim Temp   As Integer
    Dim oGraph As Object
    
    curpos = frmMain.lvwTable.SelectedItem.Index
    Set oGraph = frmMain.pgPolarGraph.Plot(curpos)
    
    With oGraph
        .Series.AllowPen = Not optLnone.Value
        .Series.AllowMarker = Not optMnone.Value
        
        ' line option
        .Series.Pen.AllowConVertices = CBool(chkLconnect.Value)
        .Series.Pen.AllowShading = CBool(chkLshading.Value)
        .Series.Pen.FillColor = picLcolor.BackColor
        .Series.Pen.Style = imcLstyle.SelectedItem.Index - 1
        .Series.Pen.Weight = imcLweight.SelectedItem.Index
        
        ' marker option
        If cmbMalignv.ListIndex = 0 Then
            Temp = 4
        ElseIf cmbMalignv.ListIndex = 1 Then
            Temp = 8
        Else
            Temp = 16
        End If
        
        .Series.Marker.Alignment = cmbMalignh.ListIndex Or Temp
        .Series.Marker.AutoSize = CBool(chkMautosize.Value)
        .Series.Marker.FillColor = picMcolor.BackColor
        Set .Series.Marker.Font = picPattern.Font
        .Series.Marker.Style = Trim$(cmbMstyle.List(cmbMstyle.ListIndex))
        .Series.Marker.Transparent = CBool(chkMtrans.Value)
        .Series.Marker.UsePicture = Not optMtype(0).Value
        .Series.Marker.SetPictureSize Val(txtMsize(0).Text), _
                                      Val(txtMsize(1).Text)
                          
        If Dir$(txtMpicloc(0).Text) <> "" Then
            .Series.Marker.PicturePath = txtMpicloc(0).Text
        End If
        
        If Dir$(txtMpicloc(1).Text) <> "" Then
            .Series.Marker.MaskPicturePath = txtMpicloc(1).Text
        End If
        
        ' options
        .StartingAngle = Val(txtAngle(0).Text)
        .EndingAngle = Val(txtAngle(1).Text)
        .Step = Val(txtStep.Text)
        .Delay = Val(txtDelay.Text)
    End With
    
    With frmMain
        Dim i As Integer
        
        .lvwTable.ListItems.Clear
        Set .lvwTable.SmallIcons = Nothing
        .imlColors.ListImages.Clear
            
        For i = 1 To .pgPolarGraph.Plot.Count
            Call .AddGraph(i)
        Next i
            
        If .lvwTable.ListItems.Count > 0 Then
            .lvwTable.ListItems(1).EnsureVisible
        End If
            
        If oGraph.Visible Then
            .pgPolarGraph.DrawGraph
        End If
        
        For i = 1 To .lvwTable.ListItems.Count
            If i = curpos Then
                .lvwTable.ListItems(i).Selected = True
            Else
                .lvwTable.ListItems(i).Selected = False
            End If
        Next i
        
        .cmdReplace.Enabled = True
        
        .picColor(0).BackColor = picLcolor.BackColor
        .picColor(1).BackColor = picMcolor.BackColor
    End With
    
    cmdApply.Enabled = False
End Sub

Private Sub cmdOk_Click()
    If cmdApply.Enabled Then Call cmdApply_Click
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oGraph As Object
    
    IsLoading = True
    Set oGraph = frmMain.pgPolarGraph.Plot(frmMain.lvwTable.SelectedItem.Index)
        
    With oGraph
        picLcolor.BackColor = .Series.Pen.FillColor
        picMcolor.BackColor = .Series.Marker.FillColor
        cmbMstyle.Font.Name = .Series.Marker.Font.Name
        cmbMstyle.Font.Italic = .Series.Marker.Font.Italic
        
        If Not .Series.AllowPen Then
            optLnone.Value = True
        Else
            optLcustom.Value = True
        End If
            
        If Not .Series.AllowMarker Then
            optMnone.Value = True
        Else
            optMcustom.Value = True
        End If
            
        If Not .Series.Marker.UsePicture Then
            optMtype(0).Value = True
        Else
            optMtype(1).Value = True
        End If
        
        Call PopulateLineStyle
        Call PopulateLineWeight
        Call PopulateMarkerStyle
        
        imcLstyle.ComboItems(.Series.Pen.Style + 1).Selected = True
        imcLweight.ComboItems(.Series.Pen.Weight).Selected = True
        cmbMstyle.ListIndex = Asc(oGraph.Series.Marker.Style) - 33
        cmbMalignh.ListIndex = oGraph.Series.Marker.GetHorizAlign
        If oGraph.Series.Marker.GetVertAlign = 4 Then
            cmbMalignv.ListIndex = 0
        ElseIf oGraph.Series.Marker.GetVertAlign = 8 Then
            cmbMalignv.ListIndex = 1
        Else
            cmbMalignv.ListIndex = 2
        End If
        
        chkLconnect.Value = Abs(CInt(.Series.Pen.AllowConVertices))
        chkLshading.Value = Abs(CInt(.Series.Pen.AllowShading))
        chkMautosize.Value = Abs(CInt(.Series.Marker.AutoSize))
        chkMtrans.Value = Abs(CInt(.Series.Marker.Transparent))
        txtMsize(0).Text = .Series.Marker.PictureWidth
        txtMsize(1).Text = .Series.Marker.PictureHeight
        txtMpicloc(0).Text = .Series.Marker.PicturePath
        txtMpicloc(1).Text = .Series.Marker.MaskPicturePath
        Set picPattern.Font = .Series.Marker.Font
        
        txtAngle(0).Text = .StartingAngle
        txtAngle(1).Text = .EndingAngle
        txtStep.Text = .Step
        txtDelay.Text = .Delay
        
        If Not .Series.AllowPen Then
            optLnone.Value = True
        Else
            optLcustom.Value = True
        End If
            
        If Not .Series.AllowMarker Then
            optMnone.Value = True
        Else
            optMcustom.Value = True
        End If
            
        If Not .Series.Marker.UsePicture Then
            optMtype(0).Value = True
        Else
            optMtype(1).Value = True
        End If
    End With
    
    Call DrawPattern
    IsLoading = False
    cmdApply.Enabled = False
End Sub

Private Sub PopulateLineStyle()
    Dim i  As Integer
    Dim sw As Integer
    Dim sh As Integer
    Dim ds As Integer
    Dim dw As Integer
    
    sw = picBuffer.ScaleWidth
    sh = picBuffer.ScaleHeight
    ds = picBuffer.DrawStyle
    dw = picBuffer.DrawWidth
    
    imcLstyle.ComboItems.Clear
    imlLstyle.ListImages.Clear
    
    For i = 0 To 4
        picBuffer.Cls
        Set picBuffer.Picture = Nothing
        picBuffer.DrawStyle = i
        picBuffer.Line (0, sh / 2)-(sw, sh / 2), picLcolor.BackColor
        Set picBuffer.Picture = picBuffer.Image
    
        imlLstyle.ImageWidth = sw
        imlLstyle.ImageHeight = sh
        imlLstyle.ListImages.Add , , picBuffer.Picture
        
        Set imcLstyle.ImageList = imlLstyle
        imcLstyle.ComboItems.Add , , , imlLstyle.ListImages.Count
    Next i
    
    picBuffer.DrawStyle = ds
    picBuffer.DrawWidth = dw
End Sub

Private Sub PopulateLineWeight()
    Dim i  As Integer
    Dim sw As Integer
    Dim sh As Integer
    Dim ds As Integer
    Dim dw As Integer
    
    sw = picBuffer.ScaleWidth
    sh = picBuffer.ScaleHeight
    ds = picBuffer.DrawStyle
    dw = picBuffer.DrawWidth
    
    imcLweight.ComboItems.Clear
    imlLweight.ListImages.Clear
    
    For i = 1 To 7
        picBuffer.Cls
        Set picBuffer.Picture = Nothing
        picBuffer.DrawWidth = i
        picBuffer.Line (0, sh / 2)-(sw, sh / 2), picLcolor.BackColor
        Set picBuffer.Picture = picBuffer.Image
    
        imlLweight.ImageWidth = sw
        imlLweight.ImageHeight = sh
        imlLweight.ListImages.Add , , picBuffer.Picture
        
        Set imcLweight.ImageList = imlLweight
        imcLweight.ComboItems.Add , , , imlLweight.ListImages.Count
    Next i
    
    picBuffer.DrawStyle = ds
    picBuffer.DrawWidth = dw
End Sub

Private Sub PopulateMarkerStyle()
    Dim i  As Integer
    Dim tw As Integer
    Dim ns As Integer
    
    cmbMstyle.Clear
    Me.FontName = cmbMstyle.Font.Name
    Me.FontSize = cmbMstyle.Font.Size
    Me.FontBold = cmbMstyle.Font.Bold
    Me.FontItalic = cmbMstyle.Font.Italic
    
    ns = 66 / Me.TextWidth(Space$(1))
    
    For i = 33 To 255
        cmbMstyle.AddItem Space$(ns / 2) & Chr$(i) & _
                          Space$(ns / 2)
    Next i
    
    cmbMstyle.ForeColor = picMcolor.BackColor
End Sub

Private Sub imcLstyle_Click()
    If optLnone.Value Then optLcustom.Value = True
    If Not IsLoading Then Call DrawPattern
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub imcLweight_Click()
    If optLnone.Value Then optLcustom.Value = True
    If Not IsLoading Then Call DrawPattern
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub optLcustom_Click()
    If Not IsLoading Then Call DrawPattern
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub optLnone_Click()
    If Not IsLoading Then Call DrawPattern
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub optMcustom_Click()
    If Not IsLoading Then Call DrawPattern
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub optMnone_Click()
    If Not IsLoading Then Call DrawPattern
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub optMtype_Click(Index As Integer)
    If optMnone.Value Then optMcustom.Value = True
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
    If Not IsLoading Then Call DrawPattern
End Sub

Private Sub tbsProp_Click()
    Select Case tbsProp.SelectedItem.Index
    Case Is = 1 ' Patterns
        fraPatterns.Visible = True
        fraOptions.Visible = False
    Case Is = 2 ' Options
        fraPatterns.Visible = False
        fraOptions.Visible = True
    End Select
End Sub

Private Sub DrawPattern()
    Dim sw As Integer
    Dim sh As Integer
    Dim ch As String
    
    On Error GoTo ErrHandler
    
    sw = picPattern.ScaleWidth
    sh = picPattern.ScaleHeight
    
    picPattern.Cls
    picPattern.FontTransparent = CBool(chkMtrans.Value)
    
    If optLcustom.Value Then
        picPattern.DrawStyle = imcLstyle.SelectedItem.Index - 1
        picPattern.DrawWidth = imcLweight.SelectedItem.Index
        picPattern.Line (0, sh / 2)-(sw, sh / 2), picLcolor.BackColor
    End If
    
    If optMcustom.Value And optMtype(0).Value Then
        ch = Trim$(cmbMstyle.List(cmbMstyle.ListIndex))
        picPattern.ForeColor = cmbMstyle.ForeColor
        picPattern.CurrentX = (sw - picPattern.TextWidth(ch)) / 2
        picPattern.CurrentY = (sh - picPattern.TextHeight(ch)) / 2
        picPattern.Print ch
    ElseIf optMcustom.Value And optMtype(1).Value Then
        Dim pw     As Long
        Dim ph     As Long
        Dim Sprite As New StdPicture
        Dim Mask   As New StdPicture
    
        If Dir$(txtMpicloc(1).Text) <> "" Then
            Set Mask = LoadPicture(txtMpicloc(1).Text)
        Else
            Set Mask = Nothing
        End If
       
        If Dir$(txtMpicloc(0).Text) <> "" Then
            Set Sprite = LoadPicture(txtMpicloc(0).Text)
            If Sprite Then
                If chkMautosize.Value = vbChecked Then
                    pw = ScaleX(Sprite.Width, vbHimetric, vbPixels)
                    ph = ScaleY(Sprite.Height, vbHimetric, vbPixels)
                   
                    txtMsize(0).Text = pw ' picture width
                    txtMsize(1).Text = ph ' picture height
                Else
                    pw = Val(Trim$(txtMsize(0).Text)) ' set picture width
                    ph = Val(Trim$(txtMsize(1).Text)) ' set picture height
                End If
               
                If Mask Then
                    picPattern.PaintPicture Mask, (picPattern.ScaleWidth - pw) / 2, _
                                                  (picPattern.ScaleHeight - ph) / 2, pw, ph, , , , , vbSrcAnd
                    picPattern.PaintPicture Sprite, (picPattern.ScaleWidth - pw) / 2, _
                                                    (picPattern.ScaleHeight - ph) / 2, pw, ph, , , , , vbSrcInvert
                Else
                    picPattern.PaintPicture Sprite, (picPattern.ScaleWidth - pw) / 2, _
                                                    (picPattern.ScaleHeight - ph) / 2, pw, ph, , , , , vbSrcCopy
                End If
            End If
        Else
            Set Sprite = Nothing
        End If
    End If
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbOKOnly Or vbInformation, "Polar Graph 1.0"
    Resume Next
End Sub

Private Sub txtAngle_Change(Index As Integer)
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub txtAngle_Validate(Index As Integer, Cancel As Boolean)
    Dim Temp As Single
    
    If Trim$(txtAngle(Index).Text) <> "" Then
        Temp = Val(txtAngle(Index).Text)
        If (Temp < 0) Then
            If Index = 0 Then
                txtAngle(0).Text = _
                    frmMain.pgPolarGraph.Plot(frmMain.lvwTable.SelectedItem.Index).StartingAngle
            Else
                txtAngle(1).Text = _
                    frmMain.pgPolarGraph.Plot(frmMain.lvwTable.SelectedItem.Index).EndingAngle
            End If
        ElseIf (Temp > 1000000) Then
            If Index = 0 Then
                txtAngle(0).Text = _
                    frmMain.pgPolarGraph.Plot(frmMain.lvwTable.SelectedItem.Index).StartingAngle
            Else
                txtAngle(1).Text = _
                    frmMain.pgPolarGraph.Plot(frmMain.lvwTable.SelectedItem.Index).EndingAngle
            End If
        Else
            ' accept input
        End If
    Else
        If Index = 0 Then
            txtAngle(0).Text = _
                frmMain.pgPolarGraph.Plot(frmMain.lvwTable.SelectedItem.Index).StartingAngle
        Else
            txtAngle(1).Text = _
                frmMain.pgPolarGraph.Plot(frmMain.lvwTable.SelectedItem.Index).EndingAngle
        End If
    End If
End Sub

Private Sub txtDelay_Change()
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub txtDelay_Validate(Cancel As Boolean)
    Dim Temp As Single
    
    If Trim$(txtDelay.Text) <> "" Then
        Temp = Val(txtDelay.Text)
        If (Temp < 0) Then
            txtDelay.Text = _
                frmMain.pgPolarGraph.Plot(frmMain.lvwTable.SelectedItem.Index).Delay
        ElseIf (Temp > 1000000) Then
            txtDelay.Text = _
                frmMain.pgPolarGraph.Plot(frmMain.lvwTable.SelectedItem.Index).Delay
        Else
            ' accept input
        End If
    Else
        txtDelay.Text = _
            frmMain.pgPolarGraph.Plot(frmMain.lvwTable.SelectedItem.Index).Delay
    End If
End Sub

Private Sub txtMsize_Change(Index As Integer)
    If Trim$(txtMsize(Index).Text) <> "" Then
        Dim Temp As Integer
        On Error GoTo ErrHandler
        
        Temp = CInt(Val(txtMsize(Index).Text))
        
        If (Temp > 0) Then
            If Not IsLoading Then Call DrawPattern
        End If
    End If
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
    Exit Sub
ErrHandler:
End Sub

Private Sub txtMsize_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If (KeyAscii >= vbKey0) And (KeyAscii <= vbKey9) Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtMsize_Validate(Index As Integer, Cancel As Boolean)
    If Trim$(txtMsize(Index).Text = "") Then
        With frmMain
            Select Case Index
            Case Is = 0 ' picture width
                txtMsize(0).Text = _
                    .pgPolarGraph.Plot(.lvwTable.SelectedItem.Index).Series.Marker.PictureWidth
            Case Is = 1 ' picture height
                txtMsize(1).Text = _
                    .pgPolarGraph.Plot(.lvwTable.SelectedItem.Index).Series.Marker.PictureHeight
            End Select
        End With
    End If
End Sub

Private Sub txtAngle_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If (KeyAscii >= vbKey0) And (KeyAscii <= vbKey9) Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtDelay_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If (KeyAscii >= vbKey0) And (KeyAscii <= vbKey9) Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtStep_Change()
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub txtStep_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If (KeyAscii >= vbKey0) And (KeyAscii <= vbKey9) Then
    ElseIf KeyAscii = &H2E Then ' Decimal
        If InStr(Me.ActiveControl.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtStep_Validate(Cancel As Boolean)
    Dim Temp As Single
    
    If Trim$(txtStep.Text) <> "" Then
        Temp = Round(Val(txtStep.Text), 4)
        If (Temp <= 0) Then
            txtStep.Text = _
                frmMain.pgPolarGraph.Plot(frmMain.lvwTable.SelectedItem.Index).Step
        ElseIf (Temp > 10000) Then
            txtStep.Text = _
                frmMain.pgPolarGraph.Plot(frmMain.lvwTable.SelectedItem.Index).Step
        Else
            ' accept input
        End If
    Else
        txtStep.Text = _
            frmMain.pgPolarGraph.Plot(frmMain.lvwTable.SelectedItem.Index).Step
    End If
End Sub
