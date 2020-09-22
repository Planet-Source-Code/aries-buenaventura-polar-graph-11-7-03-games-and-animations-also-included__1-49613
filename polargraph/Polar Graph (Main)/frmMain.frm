VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Polar Graph 1.0"
   ClientHeight    =   6270
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8730
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   720
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   23
      Top             =   4620
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picTray 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      ScaleHeight     =   180
      ScaleWidth      =   1860
      TabIndex        =   21
      Top             =   6060
      Width           =   1860
      Begin MSComctlLib.ProgressBar pgbProgressBar 
         Height          =   180
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   14
      Top             =   6000
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3316
            MinWidth        =   3316
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5371
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   60
      Top             =   5340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0894
            Key             =   "new"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CE6
            Key             =   "save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1138
            Key             =   "print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D0A
            Key             =   "copy"
         EndProperty
      EndProperty
   End
   Begin prjPolarGraph.PolarGraph pgPolarGraph 
      Height          =   3735
      Left            =   1140
      TabIndex        =   13
      Top             =   660
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6588
      OffsetDrawingAreaX=   35
      OffsetDrawingAreaY=   35
      MousePointer    =   99
      ShowRuler       =   -1  'True
   End
   Begin MSComDlg.CommonDialog dlgGraph 
      Left            =   720
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList imlTools 
      Left            =   60
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   14674159
      ImageWidth      =   23
      ImageHeight     =   23
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2024
            Key             =   "arrow"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26EE
            Key             =   "select"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DB8
            Key             =   "cross"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3482
            Key             =   "circular"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B4C
            Key             =   "magnify"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4216
            Key             =   "pan"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":48E0
            Key             =   "aspect"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlColors 
      Left            =   60
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   46
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwTable 
      Height          =   1635
      Left            =   4800
      TabIndex        =   12
      ToolTipText     =   "Left click to select: Right click to show popup menu"
      Top             =   4380
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   2884
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Legend"
         Text            =   "Legend"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Coefficient"
         Text            =   "Coefficient"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Function"
         Text            =   "Function"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tblToolbar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "New"
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "open"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "save"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "copy"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraToolbox 
      Height          =   5415
      Left            =   15
      TabIndex        =   15
      Top             =   600
      Width           =   1095
      Begin MSComctlLib.Toolbar tblToolbox 
         Height          =   1740
         Left            =   60
         TabIndex        =   16
         Top             =   180
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   3069
         ButtonWidth     =   794
         ButtonHeight    =   767
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "imlTools"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Arrow"
               Object.ToolTipText     =   "Pointer"
               ImageKey        =   "arrow"
               Style           =   1
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Select"
               Object.ToolTipText     =   "Select"
               ImageKey        =   "select"
               Style           =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cross"
               Object.ToolTipText     =   "Cross"
               ImageKey        =   "cross"
               Style           =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Circular"
               Object.ToolTipText     =   "Circular"
               ImageKey        =   "circular"
               Style           =   1
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "pan"
               Object.ToolTipText     =   "Pan"
               ImageKey        =   "pan"
               Style           =   1
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Aspect"
               Object.ToolTipText     =   "Aspect Ratio"
               ImageKey        =   "aspect"
               Style           =   1
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Zoom"
               Object.ToolTipText     =   "Zoom"
               ImageKey        =   "magnify"
               Style           =   1
            EndProperty
         EndProperty
      End
      Begin VB.ListBox lstAspect 
         BackColor       =   &H8000000F&
         Height          =   1035
         ItemData        =   "frmMain.frx":4FAA
         Left            =   135
         List            =   "frmMain.frx":4FBD
         TabIndex        =   20
         Top             =   2280
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.ListBox lstZoom 
         BackColor       =   &H8000000F&
         Height          =   1035
         ItemData        =   "frmMain.frx":4FE0
         Left            =   135
         List            =   "frmMain.frx":4FF3
         TabIndex        =   19
         Top             =   2280
         Visible         =   0   'False
         Width           =   795
      End
   End
   Begin VB.Frame fraHolder 
      Height          =   1680
      Left            =   1140
      TabIndex        =   18
      Top             =   4320
      Width           =   3675
      Begin VB.CommandButton cmdExpr 
         Caption         =   "..."
         Height          =   315
         Left            =   3180
         TabIndex        =   4
         Top             =   540
         Width           =   315
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   315
         Left            =   2520
         TabIndex        =   7
         Top             =   900
         Width           =   795
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "Replace"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1740
         TabIndex        =   6
         Top             =   900
         Width           =   795
      End
      Begin VB.PictureBox picColor 
         BackColor       =   &H000000FF&
         Height          =   315
         Index           =   1
         Left            =   2760
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   315
         TabIndex        =   11
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cmdPlot 
         Caption         =   "Plot"
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   900
         Width           =   795
      End
      Begin VB.PictureBox picColor 
         BackColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   1140
         MousePointer    =   99  'Custom
         ScaleHeight     =   255
         ScaleWidth      =   315
         TabIndex        =   9
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtCoefficient 
         Height          =   285
         Left            =   1140
         TabIndex        =   1
         Top             =   180
         Width           =   2055
      End
      Begin VB.ComboBox cmbFunction 
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Top             =   540
         Width           =   2055
      End
      Begin VB.Line linLine 
         X1              =   120
         X2              =   3540
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marker color :"
         Height          =   195
         Index           =   1
         Left            =   1740
         TabIndex        =   10
         Top             =   1380
         Width           =   975
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line color  :"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   1380
         Width           =   825
      End
      Begin VB.Label lblCoefficient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "coefficient : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   0
         Top             =   215
         Width           =   1050
      End
      Begin VB.Label lblFunction 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "f(t) = "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   660
         TabIndex        =   2
         Top             =   540
         Width           =   465
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOption 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "&Open"
         Index           =   1
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "&Save"
         Index           =   2
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "&Add Comment"
         Index           =   4
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "&Print"
         Index           =   6
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "E&xit"
         Index           =   8
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditOption 
         Caption         =   "Copy"
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditOption 
         Caption         =   "Select &All"
         Index           =   1
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditOption 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEditOption 
         Caption         =   "Copy To..."
         Index           =   3
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewToolbox 
         Caption         =   "Toolbox"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRuler 
         Caption         =   "Ruler"
         Begin VB.Menu mnuViewRulerShow 
            Caption         =   "Show"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewRulerMode 
            Caption         =   "Mode"
            Begin VB.Menu mnuViewRulerModeOption 
               Caption         =   "Axis"
               Index           =   0
            End
            Begin VB.Menu mnuViewRulerModeOption 
               Caption         =   "Unit"
               Checked         =   -1  'True
               Index           =   1
            End
         End
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewAspectRatio 
         Caption         =   "Aspect &Ratio"
         Begin VB.Menu mnuViewAspectOption 
            Caption         =   "1.5"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewAspectOption 
            Caption         =   "2.5"
            Index           =   1
         End
         Begin VB.Menu mnuViewAspectOption 
            Caption         =   "3.5"
            Index           =   2
         End
         Begin VB.Menu mnuViewAspectOption 
            Caption         =   "4.5"
            Index           =   3
         End
         Begin VB.Menu mnuViewAspectOption 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuViewAspectOption 
            Caption         =   "&Custom..."
            Index           =   5
         End
      End
      Begin VB.Menu mnuViewZoom 
         Caption         =   "Zoom"
         Begin VB.Menu mnuViewZoomSize 
            Caption         =   "Normal Size"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewZoomSize 
            Caption         =   "Large Size"
            Index           =   1
         End
         Begin VB.Menu mnuViewZoomSize 
            Caption         =   "Custom..."
            Index           =   2
         End
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewDA 
         Caption         =   "Reset Drawing Area"
      End
      Begin VB.Menu mnuViewBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
   End
   Begin VB.Menu mnuPolar 
      Caption         =   "&Polar"
      Begin VB.Menu mnuPolarOption 
         Caption         =   "Grid"
         Index           =   0
         Begin VB.Menu mnuPolarGridOption 
            Caption         =   "None"
            Index           =   0
         End
         Begin VB.Menu mnuPolarGridOption 
            Caption         =   "Minor"
            Index           =   1
         End
         Begin VB.Menu mnuPolarGridOption 
            Caption         =   "Major"
            Checked         =   -1  'True
            Index           =   2
         End
      End
      Begin VB.Menu mnuPolarOption 
         Caption         =   "Label"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuPolarOption 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPolarOption 
         Caption         =   "Coordinate System"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuPolarOption 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuPolarOption 
         Caption         =   "Options..."
         Index           =   5
      End
   End
   Begin VB.Menu mnuGraph 
      Caption         =   "&Graph"
      Begin VB.Menu mnuGraphOption 
         Caption         =   "&Hide"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuGraphOption 
         Caption         =   "&Show"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuGraphOption 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnuGraphOption 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuGraphOption 
         Caption         =   "Select &All"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mnuGraphOption 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuGraphOption 
         Caption         =   "&Play"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu mnuGraphOption 
         Caption         =   "Sto&p"
         Enabled         =   0   'False
         Index           =   7
      End
      Begin VB.Menu mnuGraphOption 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuGraphOption 
         Caption         =   "&Table..."
         Enabled         =   0   'False
         Index           =   9
      End
      Begin VB.Menu mnuGraphOption 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuGraphOption 
         Caption         =   "&Properties"
         Enabled         =   0   'False
         Index           =   11
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CUR_ASPECT = 103
Private Const CUR_CIRCULAR = 104
Private Const CUR_HAND_FLAT = 105
Private Const CUR_HAND_GRAB = 106
Private Const CUR_SELECT = 107
Private Const CUR_ZOOM = 108

Dim IsHandGrab   As Boolean
Dim SelectedTool As Integer

Private Sub cmbFunction_Change()
    If pgPolarGraph.StopAni Then
        cmdPlot.Enabled = CBool(Len(cmbFunction.Text))
        If lvwTable.ListItems.Count > 0 Then
            cmdReplace.Enabled = cmdPlot.Enabled
        End If
    End If
End Sub

Private Sub cmbFunction_Click()
    Dim curpos As Integer
    
    With lvwTable
        For curpos = 1 To .ListItems.Count
            If curpos <> cmbFunction.ListIndex + 1 Then
                .ListItems(curpos).Selected = True
            Else
                .ListItems(curpos).Selected = False
            End If
        Next curpos
        
        txtCoefficient.Text = .ListItems(.SelectedItem.Index).SubItems(1)
    End With
End Sub

Private Sub cmbFunction_KeyPress(KeyAscii As Integer)
    If Len(cmbFunction.Text) > 0 Then
        If KeyAscii = vbKeyReturn Then Call cmdPlot_Click
    End If
End Sub

Private Sub cmdClear_Click()
    txtCoefficient.Text = ""
    cmbFunction.Text = ""
End Sub

Private Sub cmdExpr_Click()
    frmExpr.Show vbModal, Me
    If frmExpr.RetVal <> "" Then
        cmbFunction.Text = frmExpr.RetVal
    End If
End Sub

Private Sub cmdPlot_Click()
    Dim Script  As New MSScriptControl.ScriptControl
    Dim MathLib As New MathLibrary
    On Error Resume Next
    
    Call tblToolbox_ButtonClick(tblToolbox.Buttons(1))
    
    Script.Language = "VBScript"
    Script.Timeout = NoTimeout
    Script.AddObject "MathLib", MathLib, True
    Script.AddCode Trim$(txtCoefficient.Text)
    
    If Err Then
        txtCoefficient.SelStart = Script.Error.Column
        txtCoefficient.SelLength = 1
        MsgBox Script.Error.Description & " : '" & _
               txtCoefficient.SelText & "'", vbInformation, "Equation"
        txtCoefficient.SetFocus
        Exit Sub
    End If
    
    Script.Eval Trim$(cmbFunction.Text)
    
    If Err Then
        If Script.Error.Number <> 11 Then
            cmbFunction.SelStart = Script.Error.Column
            cmbFunction.SelLength = 1
            MsgBox Script.Error.Description & " : '" & _
                   cmbFunction.SelText & "'", vbInformation, "Equation"
            cmbFunction.SelStart = Script.Error.Column
            cmbFunction.SelLength = 1
            Exit Sub
        End If
    End If
        
    pgPolarGraph.Plot.Add txtCoefficient.Text, cmbFunction.Text, _
                        picColor(0).BackColor, picColor(1).BackColor
                        
    pgPolarGraph.Plot(pgPolarGraph.Plot.Count).Series.Pen.FillColor = picColor(0).BackColor
    pgPolarGraph.Plot(pgPolarGraph.Plot.Count).Series.Marker.FillColor = picColor(1).BackColor
    pgPolarGraph.DrawGraph pgPolarGraph.Plot.Count
    Call AddGraph(pgPolarGraph.Plot.Count)
    
    cmbFunction.AddItem cmbFunction.Text
End Sub

Private Sub cmdReplace_Click()
    Dim oGraph         As Object
    Dim curpos         As Integer
    Dim bIsUpdateTable As Boolean
    
    With lvwTable
        bIsUpdateTable = False
        curpos = .SelectedItem.Index
        Set oGraph = pgPolarGraph.Plot(curpos)
        
        .ListItems(curpos).ListSubItems(1).Text = Trim$(txtCoefficient.Text)
        .ListItems(curpos).ListSubItems(2).Text = Trim$(cmbFunction.Text)
        oGraph.Coefficient = Trim$(txtCoefficient.Text)
        oGraph.Equation = Trim$(cmbFunction.Text)
        
        If oGraph.Series.Pen.FillColor <> picColor(0).BackColor Or _
           oGraph.Series.Marker.FillColor <> picColor(1).BackColor Then
           
            oGraph.Series.Pen.FillColor = picColor(0).BackColor
            oGraph.Series.Marker.FillColor = picColor(1).BackColor
            bIsUpdateTable = True
        End If
        
        pgPolarGraph.DrawGraph
        
        If bIsUpdateTable Then
            .ListItems.Clear
            Set .SmallIcons = Nothing
            imlColors.ListImages.Clear
            
            cmbFunction.Clear
            For curpos = 1 To pgPolarGraph.Plot.Count
                Call AddGraph(curpos)
                cmbFunction.AddItem pgPolarGraph.Plot(curpos).Equation
            Next curpos
            
            txtCoefficient.Text = .ListItems(.SelectedItem.Index).SubItems(1)
        End If
    End With
End Sub

Private Sub Form_Load()
    Call InitImageList
    Set picColor(0).MouseIcon = LoadResPicture(101, vbResCursor)
    Set picColor(1).MouseIcon = LoadResPicture(101, vbResCursor)
    
    If LoadWallPaper("WP1.ajb") Then
        pgPolarGraph.PictureStyle = ajb_PSNone
    ElseIf LoadWallPaper("WP2.ajb") Then
        pgPolarGraph.PictureStyle = ajb_PSCenter
    ElseIf LoadWallPaper("WP3.ajb") Then
        pgPolarGraph.PictureStyle = ajb_PSStretch
    ElseIf LoadWallPaper("WP4.ajb") Then
        pgPolarGraph.PictureStyle = ajb_PSTile
    End If
    
    lstAspect.ListIndex = lstAspect.TopIndex ' 1.5
    lstZoom.ListIndex = lstZoom.TopIndex + 1 ' 50%
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        On Error Resume Next
        
        Dim sw As Long
        Dim sh As Long
        
        sw = Me.ScaleWidth
        sh = Me.ScaleHeight
        
        fraToolbox.Move 15, IIf(tblToolbar.Visible, tblToolbar.Height, 0) - 90, fraToolbox.Width, _
                        sh - sbStatusBar.Height - tblToolbar.Height + _
                        IIf(tblToolbar.Visible, 0, tblToolbar.Height) + 90
        pgPolarGraph.Move IIf(fraToolbox.Visible, fraToolbox.Width + 30, 0), _
                          IIf(tblToolbar.Visible, tblToolbar.Height, 0), _
                          sw - IIf(fraToolbox.Visible, fraToolbox.Width + 45, 0), _
                          sh - sbStatusBar.Height - tblToolbar.Height - lvwTable.Height + _
                          IIf(tblToolbar.Visible, 0, tblToolbar.Height)
        lvwTable.Move IIf(fraToolbox.Visible, fraToolbox.Width + 30, 0) + fraHolder.Width + 30, pgPolarGraph.Top + _
                    pgPolarGraph.Height + 30, sw - IIf(fraToolbox.Visible, fraToolbox.Width + 45, 0) - 30 - _
                      fraHolder.Width, lvwTable.Height
        fraHolder.Move IIf(fraToolbox.Visible, fraToolbox.Width + 45, 0), lvwTable.Top - 75, _
                       fraHolder.Width, fraHolder.Height
        picTray.Move sbStatusBar.Left + 15, Me.ScaleHeight - sbStatusBar.Height + 60
        
        With lvwTable.ColumnHeaders
            ' Legend
            .Item(1).Width = 800
            ' Coefficient
            .Item(2).Width = lvwTable.Width * 0.35
            ' Function
            .Item(3).Width = lvwTable.Width - .Item(1).Width - .Item(2).Width - 60
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not pgPolarGraph.StopAni Then
        pgPolarGraph.StopAni = True
    End If
End Sub

Private Sub lstAspect_Click()
    If lstAspect.ListIndex <> 4 Then
        Dim Temp As Single
        
        Temp = Val(lstAspect.List(lstAspect.ListIndex))
        If Temp <> pgPolarGraph.AspectRatio Then
            Call mnuViewAspectOption_Click(lstAspect.ListIndex)
        End If
    Else
        Call mnuViewAspectOption_Click(lstAspect.ListIndex)
    End If
End Sub

Private Sub lstAspect_KeyPress(KeyAscii As Integer)
    If lstAspect.ListIndex = 4 Then
        If KeyAscii = vbKeyReturn Then
            frmAspect.Show vbModal, Me
        End If
    End If
End Sub

Private Sub lstZoom_Click()
    If lstZoom.ListIndex <> 4 Then
        Dim Temp As Integer
        
        Temp = Val(lstZoom.List(lstZoom.ListIndex))
        If pgPolarGraph.Zoom <> Temp Then
            pgPolarGraph.Zoom = Temp
        End If
    Else
        frmZoom.Show vbModal, Me
    End If
End Sub

Private Sub lstZoom_KeyPress(KeyAscii As Integer)
    If lstZoom.ListIndex = 4 Then
        If KeyAscii = vbKeyReturn Then
            frmZoom.Show vbModal, Me
        End If
    End If
End Sub

Private Sub lvwTable_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim curpos  As Integer
    Dim counter As Integer
    
    counter = 0
    For curpos = 1 To lvwTable.ListItems.Count
        If lvwTable.ListItems(curpos).Selected Then
            counter = counter + 1
            If counter > 1 Then Exit For
        End If
    Next curpos
    
    If counter = 1 Then
        With pgPolarGraph.Plot
            txtCoefficient.Text = Item.ListSubItems(1).Text
            cmbFunction.Text = Item.ListSubItems(2).Text
            picColor(0).BackColor = .Item(lvwTable.SelectedItem.Index).Series.Pen.FillColor
            picColor(1).BackColor = .Item(lvwTable.SelectedItem.Index).Series.Marker.FillColor
        End With
    End If
End Sub

Private Sub lvwTable_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbRightButton Then
        If lvwTable.ListItems.Count > 0 Then
           Me.PopupMenu mnuGraph
        End If
    End If
End Sub

Private Sub mnuEdit_Click()
    mnuEditOption(0).Enabled = pgPolarGraph.ActiveSelection ' Copy
    mnuEditOption(3).Enabled = pgPolarGraph.ActiveSelection ' Copy To...
End Sub

Private Sub mnuEditOption_Click(Index As Integer)
    Select Case Index
    Case Is = 0 ' Copy
        pgPolarGraph.EditCopy
    Case Is = 1 ' Select All
        tblToolbox_ButtonClick tblToolbox.Buttons("Select")
        pgPolarGraph.EditSelectAll
    Case Is = 2 ' ...
    Case Is = 3 ' Copy To...
        pgPolarGraph.EditCopyTo
    End Select
End Sub

Private Sub mnuFileOption_Click(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case Is = 0 ' New
        pgPolarGraph.ClearGraph
        Set pgPolarGraph.Plot = Nothing
        lvwTable.ListItems.Clear
        Set lvwTable.SmallIcons = Nothing
        imlColors.ListImages.Clear
        
        cmbFunction.Clear
        txtCoefficient.Text = ""
        cmdPlot.Enabled = False
        cmdReplace.Enabled = False
    Case Is = 1 ' Open
        With dlgGraph
            .Filter = "Polar Graph (*.pgp) | *.pgp; |" & _
                      "All Files (*.*) | *.*"
            .FilterIndex = 1
            .InitDir = App.Path & "\Save"
            .Filename = ""
            .ShowOpen
            
            If .Filename <> "" Then
                If pgPolarGraph.OpenGraph(.Filename) Then
                    Dim curpos As Integer
                    
                    lvwTable.ListItems.Clear
                    Set lvwTable.SmallIcons = Nothing
                    imlColors.ListImages.Clear
          
                    cmbFunction.Clear
                    For curpos = 1 To pgPolarGraph.Plot.Count
                        Call AddGraph(curpos)
                        cmbFunction.AddItem pgPolarGraph.Plot(curpos).Equation
                    Next curpos
    
                    If lvwTable.ListItems.Count > 0 Then
                        lvwTable.ListItems(1).EnsureVisible
                        cmbFunction.ListIndex = cmbFunction.TopIndex
                        txtCoefficient.Text = lvwTable.ListItems(1).SubItems(1)
                        cmdPlot.Enabled = True
                        cmdReplace.Enabled = True
                    Else
                        cmdPlot.Enabled = False
                        cmdReplace.Enabled = False
                    End If
                    
                    pgPolarGraph.DrawGraph
                    
                    If Trim$(pgPolarGraph.Comment) <> "" Then
                        frmComment.txtComment = pgPolarGraph.Comment
                        frmComment.Show vbModal, Me
                    End If
                    
                    For curpos = 1 To lvwTable.ListItems.Count
                        lvwTable.ListItems(curpos).Selected = False
                    Next curpos
                    
                    If lvwTable.ListItems.Count > 0 Then
                        lvwTable.ListItems(1).Selected = True
                    End If
                End If
            End If
        End With
    Case Is = 2 ' Save
        With dlgGraph
            .Filter = "Polar Graph (*.pgp) | *.pgp"
            .FilterIndex = 1
            .Filename = ""
            .InitDir = App.Path & "\Save"
            .Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist
            .ShowSave
        
            If .Filename = "" Then Exit Sub

            pgPolarGraph.SaveGraph .Filename
        End With
    Case Is = 3 ' ...
    Case Is = 4 ' Add Comment
        frmComment.txtComment.Text = pgPolarGraph.Comment
        frmComment.Show vbModal, Me
    Case Is = 5 ' ...
    Case Is = 6 ' Print
        pgPolarGraph.SentToPrinter
    Case Is = 7 ' ...
    Case Is = 8 ' Exit
        Unload Me
    End Select
End Sub

Private Sub mnuGraph_Click()
    Dim curpos  As Integer
    Dim counter As Integer
    
    With lvwTable
        For curpos = mnuGraphOption.LBound To mnuGraphOption.UBound
            If mnuGraphOption(curpos).Caption <> "-" Then
                mnuGraphOption(curpos).Enabled = False
            End If
        Next curpos
        
        If pgPolarGraph.StopAni Then
            For curpos = 1 To .ListItems.Count
                If .ListItems(curpos).Selected Then
                    ' Hide
                    If Not mnuGraphOption(0).Enabled Then
                        If Not .ListItems(curpos).Ghosted Then
                            mnuGraphOption(0).Enabled = True
                        End If
                    End If
                      
                    ' Show
                    If Not mnuGraphOption(1).Enabled Then
                        If .ListItems(curpos).Ghosted Then
                            mnuGraphOption(1).Enabled = True
                        End If
                    End If
                      
                    ' Delete
                    If Not mnuGraphOption(2).Enabled Then
                        If .ListItems(curpos).Selected Then
                            mnuGraphOption(2).Enabled = True
                        End If
                    End If
                    Exit For
                End If
            Next curpos
        End If
        
        counter = 0
        For curpos = 1 To .ListItems.Count
            If .ListItems(curpos).Selected Then
                counter = counter + 1
                If counter <> 1 Then Exit For
            End If
        Next curpos
        
        If counter = 1 Then
            ' Play
            curpos = lvwTable.SelectedItem.Index
            If (pgPolarGraph.Plot(curpos).Delay > 0) And _
               (pgPolarGraph.StopAni) Then
               
                If Not mnuGraphOption(6).Enabled Then
                    mnuGraphOption(6).Enabled = True
                End If
            End If
            
            ' Stop
            If Not pgPolarGraph.StopAni Then
                If Not mnuGraphOption(7).Enabled Then
                    mnuGraphOption(7).Enabled = True
                End If
            End If
            
            If pgPolarGraph.StopAni Then
                ' Table...
                If Not mnuGraphOption(9).Enabled Then
                    mnuGraphOption(9).Enabled = True
                End If
                
                ' Properties
                If Not mnuGraphOption(11).Enabled Then
                    mnuGraphOption(11).Enabled = True
                End If
            End If
        End If
        
        If .ListItems.Count > 0 Then
            ' Select All
            If Not mnuGraphOption(4).Enabled Then
                mnuGraphOption(4).Enabled = True
            End If
        End If
    End With
End Sub

Private Sub mnuGraphOption_Click(Index As Integer)
    Dim curpos   As Integer
    Dim IsUpdate As Boolean
    
    Select Case Index
    Case Is = 0 ' Hide
        IsUpdate = False
        
        With lvwTable
            For curpos = 1 To .ListItems.Count
                If .ListItems(curpos).Selected Then
                    If Not .ListItems(curpos).Ghosted Then
                        .ListItems(curpos).Ghosted = True
                        .ListItems(curpos).ListSubItems(1).ForeColor = &H808080
                        .ListItems(curpos).ListSubItems(2).ForeColor = &H808080
                        pgPolarGraph.Plot(curpos).Visible = False
                        If Not IsUpdate Then IsUpdate = True
                    End If
                End If
            Next curpos
        End With
        
        If IsUpdate Then pgPolarGraph.DrawGraph
    Case Is = 1  ' Show
        IsUpdate = False
        
        With lvwTable
            For curpos = 1 To .ListItems.Count
                If .ListItems(curpos).Selected Then
                    If .ListItems(curpos).Ghosted Then
                        .ListItems(curpos).Ghosted = False
                        .ListItems(curpos).ListSubItems(1).ForeColor = &H0
                        .ListItems(curpos).ListSubItems(2).ForeColor = &H0
                        pgPolarGraph.Plot(curpos).Visible = True
                        If Not IsUpdate Then IsUpdate = True
                    End If
                End If
            Next curpos
        End With
        
        If IsUpdate Then pgPolarGraph.DrawGraph
    Case Is = 2  ' Delete
        IsUpdate = False: curpos = 1
        
        With lvwTable
            Do
                If .ListItems.Count > 0 Then
                    If .ListItems(curpos).Selected Then
                        .ListItems.Remove curpos
                        pgPolarGraph.Plot.Remove curpos
                        
                        If Not IsUpdate Then IsUpdate = True
                    Else
                        curpos = curpos + 1
                    End If
                Else
                    Exit Do
                End If
            Loop While curpos <= .ListItems.Count
            
            .ListItems.Clear
            Set .SmallIcons = Nothing
            imlColors.ListImages.Clear
            
            cmbFunction.Clear
            For curpos = 1 To pgPolarGraph.Plot.Count
                Call AddGraph(curpos)
                cmbFunction.AddItem pgPolarGraph.Plot(curpos).Equation
            Next curpos
            
            .MultiSelect = False
            For curpos = 1 To .ListItems.Count
                .ListItems(curpos).Selected = False
            Next curpos
            
            If .ListItems.Count > 0 Then
                .ListItems(1).EnsureVisible
                .ListItems(1).Selected = True
                
                cmbFunction.ListIndex = cmbFunction.TopIndex
                txtCoefficient.Text = .ListItems(.SelectedItem.Index).SubItems(1)
            Else
                cmdReplace.Enabled = False
                txtCoefficient.Text = ""
            End If
            
            .MultiSelect = True
        End With
        
        If IsUpdate Then pgPolarGraph.DrawGraph
    Case Is = 3  ' ...
    Case Is = 4  ' Select All
        For curpos = 1 To lvwTable.ListItems.Count
            lvwTable.ListItems(curpos).Selected = True
        Next curpos
    Case Is = 5  ' ...
    Case Is = 6  ' Play
        Dim bIsAni As Boolean
        
        Call tblToolbox_ButtonClick(tblToolbox.Buttons(1))
        
        pgPolarGraph.ClearGraph
        For curpos = 1 To lvwTable.ListItems.Count
            If curpos = lvwTable.SelectedItem.Index Then
                bIsAni = True
            Else
                bIsAni = False
            End If
            
            pgPolarGraph.DrawGraph curpos, bIsAni
        Next curpos
    Case Is = 7  ' Stop
        pgPolarGraph.StopAni = True
    Case Is = 8  ' ....
    Case Is = 9  ' Table...
        Load frmTable
        frmTable.Show vbModal
    Case Is = 10 ' ...
    Case Is = 11 ' Properties
        frmProperties.Show vbModal
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuPolarGridOption_Click(Index As Integer)
    Dim i As Integer
    
    With mnuPolarGridOption
        For i = .LBound To .UBound
            If i <> Index Then
                .Item(i).Checked = False
            Else
                .Item(i).Checked = True
            End If
        Next i
    End With
    
    pgPolarGraph.ShowGrid = Index
End Sub

Private Sub mnuPolarOption_Click(Index As Integer)
    Select Case Index
    Case Is = 0 ' Grid
    Case Is = 1 ' Label
        mnuPolarOption(1).Checked = Not mnuPolarOption(1).Checked
        pgPolarGraph.ShowLabel = mnuPolarOption(1).Checked
        pgPolarGraph.DrawGraph
    Case Is = 2 ' ...
    Case Is = 3 ' Coordinate System
        mnuPolarOption(3).Checked = Not mnuPolarOption(3).Checked
        pgPolarGraph.ShowCoordinateSystem = mnuPolarOption(3).Checked
    Case Is = 4 ' ...
    Case Is = 5 ' Options...
        frmOptions.Show vbModal
    End Select
End Sub

Private Sub mnuViewAspectOption_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
    Case 0 To 3
        For i = 0 To 3
            If i <> Index Then
                mnuViewAspectOption(i).Checked = False
            Else
                mnuViewAspectOption(i).Checked = True
            End If
        Next i
        
        If pgPolarGraph.AspectRatio <> Val(mnuViewAspectOption(Index).Caption) Then
            lstAspect.ListIndex = Index
            pgPolarGraph.AspectRatio = Val(mnuViewAspectOption(Index).Caption)
        End If
    Case Else
        frmAspect.Show vbModal, Me
    End Select
    
    If pgPolarGraph.SetActiveTool = ajb_TSelect Then
        Dim OldAT As ActiveToolConstants
        
        OldAT = pgPolarGraph.SetActiveTool
        pgPolarGraph.SetActiveTool = ajb_TArrow
        pgPolarGraph.SetActiveTool = OldAT
        tblToolbar.Buttons("copy").Enabled = False
    End If
End Sub

Private Sub mnuViewDA_Click()
    pgPolarGraph.ResetDrawingArea
End Sub

Private Sub mnuViewRefresh_Click()
    pgPolarGraph.Refresh
End Sub

Private Sub mnuViewRulerModeOption_Click(Index As Integer)
    mnuViewRulerModeOption(0).Checked = False
    mnuViewRulerModeOption(1).Checked = False
    mnuViewRulerModeOption(Index).Checked = True
    
    pgPolarGraph.SetRulerMode = Index
End Sub

Private Sub mnuViewRulerShow_Click()
    mnuViewRulerShow.Checked = Not mnuViewRulerShow.Checked
    pgPolarGraph.ShowRuler = mnuViewRulerShow.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tblToolbar.Visible = mnuViewToolbar.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolbox_Click()
    mnuViewToolbox.Checked = Not mnuViewToolbox.Checked
    fraToolbox.Visible = mnuViewToolbox.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewZoomSize_Click(Index As Integer)
    Select Case Index
    Case Is = 0 ' Normal Size
        pgPolarGraph.Zoom = 50
    Case Is = 1 ' Large Size
        pgPolarGraph.Zoom = 100
    Case Is = 2 ' Custom
        frmZoom.Show vbModal, Me
    End Select
    
    If pgPolarGraph.SetActiveTool = ajb_TSelect Then
        Dim OldAT As ActiveToolConstants
        
        OldAT = pgPolarGraph.SetActiveTool
        pgPolarGraph.SetActiveTool = ajb_TArrow
        pgPolarGraph.SetActiveTool = OldAT
        tblToolbar.Buttons("copy").Enabled = False
    End If
End Sub

Private Sub pgPolarGraph_AniStat(bPlay As Boolean)
    Dim i As Integer
    
    For i = 1 To tblToolbar.Buttons.Count
        tblToolbar.Buttons(i).Enabled = bPlay
    Next i
    
    For i = mnuFileOption.LBound To mnuFileOption.UBound - 1
        If mnuFileOption(i).Caption <> "-" Then
            mnuFileOption(i).Enabled = bPlay
        End If
    Next i
    
    For i = mnuViewAspectOption.LBound To mnuViewAspectOption.UBound
        If mnuViewAspectOption(i).Caption <> "-" Then
            mnuViewAspectOption(i).Enabled = bPlay
        End If
    Next i
    
    For i = mnuViewZoomSize.LBound To mnuViewZoomSize.UBound
        mnuViewZoomSize(i).Enabled = bPlay
    Next i
    
    For i = mnuPolarOption.LBound To mnuPolarOption.UBound
        If i Mod 2 = 1 Then
            mnuPolarOption(i).Enabled = bPlay
        End If
    Next i
    
    For i = mnuPolarGridOption.LBound To mnuPolarGridOption.UBound
        mnuPolarGridOption(i).Enabled = bPlay
    Next i
    
    If Len(Trim$(cmbFunction.Text)) > 0 Then
        cmdPlot.Enabled = bPlay
        cmdReplace.Enabled = cmdPlot.Enabled
    Else
        cmdPlot.Enabled = False
        cmdReplace.Enabled = False
    End If
    
    tblToolbox.Buttons("Select").Enabled = bPlay
    tblToolbox.Buttons("Aspect").Enabled = bPlay
    tblToolbox.Buttons("Zoom").Enabled = bPlay
    
    mnuViewRefresh.Enabled = bPlay       ' Refresh
    mnuEditOption(1).Enabled = bPlay     ' Select All
End Sub

Private Sub pgPolarGraph_Click()
    Select Case SelectedTool
    Case Is = 1 ' Arrow
    Case Is = 2 ' Select
    Case Is = 3 ' Cross
    Case Is = 4 ' Circular
    Case Is = 5 ' Pan
    Case Is = 6 ' Aspect ratio
        If lstAspect.ListIndex = 0 Then
            lstAspect.Selected(1) = True ' 2.5
        ElseIf lstAspect.ListIndex = 1 Then
            lstAspect.Selected(2) = True ' 3.5
        ElseIf lstAspect.ListIndex = 2 Then
            lstAspect.Selected(3) = True ' 4.5
        Else
            lstAspect.Selected(0) = True ' 1.5, Custom...
        End If
    Case Is = 7 ' Zoom
        If lstZoom.ListIndex = 0 Then
            lstZoom.Selected(1) = True ' 50%
        ElseIf lstZoom.ListIndex = 1 Then
            lstZoom.Selected(2) = True ' 75%
        ElseIf lstZoom.ListIndex = 2 Then
            lstZoom.Selected(3) = True ' 100%
        Else
            lstZoom.Selected(0) = True ' 25%, Custom...
        End If
    End Select
End Sub

Private Sub pgPolarGraph_Graph(IndexKey As Variant)
    pgbProgressBar.Min = 0
    pgbProgressBar.Max = (pgPolarGraph.Plot(IndexKey).EndingAngle - _
                          pgPolarGraph.Plot(IndexKey).StartingAngle) / _
                          pgPolarGraph.Plot(IndexKey).Step + 1
    pgbProgressBar.Value = 0
End Sub

Private Sub pgPolarGraph_Location(ByVal X As Single, ByVal Y As Single, ByVal Angle As Single, _
                                ByVal Area As Single, ByVal Circumference As Variant, _
                                ByVal Diameter As Single, ByVal Radius As Single)
                                
    sbStatusBar.Panels(2).Text = "X=" & Round(X, 2) & "; Y=" & Round(Y, 2) & _
                                 "; Angle=" & Round(Angle, 2) & Chr(&HB0) & _
                                 "; Radius=" & Round(Radius, 2)
    sbStatusBar.Panels(3).Text = "Area=" & Round(Area, 2) & _
                                 "; Circumference=" & Round(Circumference, 2) & _
                                 "; Diameter=" & Round(Diameter, 2)
End Sub

Private Sub pgPolarGraph_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then
        Select Case SelectedTool
        Case Is = 1 ' Arrow
        Case Is = 2 ' Select
        Case Is = 3 ' Cross
        Case Is = 4 ' Circular
        Case Is = 5 ' Pan
            IsHandGrab = True
            Call SetGraphCursor(CUR_HAND_GRAB)
        Case Is = 6 ' Aspect ratio
        Case Is = 7 ' Zoom
        End Select
    End If
End Sub

Private Sub pgPolarGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then
        Select Case SelectedTool
        Case Is = 1 ' Arrow
        Case Is = 2 ' Select
        Case Is = 3 ' Cross
        Case Is = 4 ' Circular
        Case Is = 5 ' Pan
            If Not IsHandGrab Then
                Call SetGraphCursor(CUR_HAND_GRAB)
            End If
        Case Is = 6 ' Aspect ratio
        Case Is = 7 ' Zoom
        End Select
    End If
    
    ' copy button
    tblToolbar.Buttons(7).Enabled = pgPolarGraph.ActiveSelection
End Sub

Private Sub pgPolarGraph_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then
        Select Case SelectedTool
        Case Is = 1 ' Arrow
        Case Is = 2 ' Select
        Case Is = 3 ' Cross
        Case Is = 4 ' Circular
        Case Is = 5 ' Pan
            IsHandGrab = False
            Call SetGraphCursor(CUR_HAND_FLAT)
        Case Is = 6 ' Aspect ratio
        Case Is = 7 ' Zoom
        End Select
    ElseIf Button And vbRightButton Then
        Me.PopupMenu mnuEdit
    End If
End Sub

Private Sub pgPolarGraph_Status(ByVal Coefficient As String, _
                                ByVal Equation As String, ByVal Color As Long, _
                                ByVal Degrees As Single, ByVal Data As Variant)
                                
    pgbProgressBar.Value = pgbProgressBar.Value + 1
End Sub

Private Sub picColor_Click(Index As Integer)
    On Error GoTo ErrHandler

    dlgGraph.ShowColor
    If dlgGraph.Color <> picColor(Index).BackColor Then
        picColor(Index).BackColor = dlgGraph.Color
    End If
    Exit Sub
    
ErrHandler:
End Sub

Private Sub picColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set picColor(Index).MouseIcon = LoadResPicture(102, vbResCursor)
End Sub

Private Sub picColor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set picColor(Index).MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub tblToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case Is = "new"
        Call mnuFileOption_Click(0)
    Case Is = "open"
        Call mnuFileOption_Click(1)
    Case Is = "save"
        Call mnuFileOption_Click(2)
    Case Is = "print"
        Call mnuFileOption_Click(4)
    Case Is = "copy"
        Call mnuEditOption_Click(0)
    End Select
End Sub

Private Sub tblToolbox_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim i As Integer
    
    For i = 1 To tblToolbox.Buttons.Count
        If i <> Button.Index Then
            tblToolbox.Buttons(i).Value = tbrUnpressed
        Else
            tblToolbox.Buttons(i).Value = tbrPressed
        End If
    Next i
    
    If lstAspect.Visible Then lstAspect.Visible = False
    If lstZoom.Visible Then lstZoom.Visible = False
    
    Select Case Button.Index
    Case Is = 1 ' Arrow
        pgPolarGraph.SetActiveTool = ajb_TArrow
        pgPolarGraph.MousePointer = vbArrow
    Case Is = 2 ' Select
        pgPolarGraph.SetActiveTool = ajb_TSelect
        Call SetGraphCursor(CUR_SELECT)
    Case Is = 3 ' Cross
        pgPolarGraph.SetActiveTool = ajb_TCross
        pgPolarGraph.MousePointer = vbCrosshair
    Case Is = 4 ' Circular
        pgPolarGraph.SetActiveTool = ajb_TCircular
        Call SetGraphCursor(CUR_CIRCULAR)
    Case Is = 5 ' Pan
        pgPolarGraph.SetActiveTool = ajb_TPan
        Call SetGraphCursor(CUR_HAND_FLAT)
    Case Is = 6 ' Aspect ratio
        pgPolarGraph.SetActiveTool = ajb_TArrow
        If Not lstAspect.Visible Then lstAspect.Visible = True
        Call SetGraphCursor(CUR_ASPECT)
    Case Is = 7 ' Zoom
        pgPolarGraph.SetActiveTool = ajb_TArrow
        If Not lstZoom.Visible Then lstZoom.Visible = True
        Call SetGraphCursor(CUR_ZOOM)
    End Select
    
    ' copy button
    tblToolbar.Buttons(7).Enabled = pgPolarGraph.ActiveSelection
    
    SelectedTool = Button.Index
End Sub

Public Sub AddGraph(GraphIndexKey As Variant)
    Dim curpos As Integer
    Dim oGraph As Object
    
    On Error GoTo ErrHandler
    
    With lvwTable
        Set oGraph = pgPolarGraph.Plot(GraphIndexKey).Series
          
        picTemp.Cls
        Set picTemp.Picture = Nothing
        picTemp.BackColor = oGraph.Pen.FillColor
        
        If oGraph.AllowMarker Then
            If oGraph.Marker.UsePicture Then
                Dim pw     As Long
                Dim ph     As Long
                Dim Sprite As New StdPicture
                Dim Mask   As New StdPicture
                
                pw = 15: ph = 15
                
                If Dir$(oGraph.Marker.PicturePath) <> "" Then
                    Set Sprite = LoadPicture(oGraph.Marker.PicturePath)
                End If
                
                If Dir$(oGraph.Marker.MaskPicturePath) <> "" Then
                    Set Mask = LoadPicture(oGraph.Marker.MaskPicturePath)
                End If
                
                If Sprite Then
                    If Mask Then
                        picTemp.PaintPicture Mask, _
                                         (picTemp.ScaleWidth - pw) / 2, _
                                         (picTemp.ScaleHeight - ph) / 2, _
                                         pw, ph, , , , , vbSrcAnd
                        picTemp.PaintPicture Sprite, _
                                             (picTemp.ScaleWidth - pw) / 2, _
                                             (picTemp.ScaleHeight - ph) / 2, _
                                             pw, ph, , , , , vbSrcInvert
    
                    Else
                        picTemp.PaintPicture Sprite, _
                                             (picTemp.ScaleWidth - pw) / 2, _
                                             (picTemp.ScaleHeight - ph) / 2, _
                                             pw, ph, , , , , vbSrcCopy
                    End If
                End If
            Else
                picTemp.ForeColor = oGraph.Marker.FillColor
                picTemp.FontName = oGraph.Marker.Font.Name
                picTemp.CurrentX = (picTemp.ScaleWidth - picTemp.TextWidth(oGraph.Marker.Style)) / 2
                picTemp.CurrentY = (picTemp.ScaleHeight - picTemp.TextHeight(oGraph.Marker.Style)) / 2
                picTemp.Print oGraph.Marker.Style
            End If
        End If
        
        Set picTemp.Picture = picTemp.Image
        
        imlColors.ListImages.Add , , picTemp.Picture
        Set .SmallIcons = imlColors
        
        .ListItems.Add , , , , imlColors.ListImages.Count
        .ListItems(.ListItems.Count).SubItems(1) = _
            pgPolarGraph.Plot(GraphIndexKey).Coefficient
        .ListItems(.ListItems.Count).SubItems(2) = _
            pgPolarGraph.Plot(GraphIndexKey).Equation
        .ListItems(.ListItems.Count).EnsureVisible
        
        If Not pgPolarGraph.Plot(.ListItems.Count).Visible Then
            .ListItems(.ListItems.Count).Ghosted = True
        End If
                
        For curpos = 1 To .ListItems.Count
            If curpos <> .ListItems.Count Then
                .ListItems(curpos).Selected = False
            Else
                .ListItems(curpos).Selected = True
            End If
        Next curpos
        
        .Refresh
    End With
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbOKOnly Or vbInformation, "Polar Graph 1.0"
    Resume Next
End Sub

Private Sub InitImageList()
    picTemp.Move 0, 0, 46 * Screen.TwipsPerPixelX, _
                       18 * Screen.TwipsPerPixelY
End Sub

Private Sub SetGraphCursor(ByVal ResID As Long)
    pgPolarGraph.MousePointer = vbCustom
    Set pgPolarGraph.MouseIcon = LoadResPicture(ResID, vbResCursor)
End Sub

Private Sub txtCoefficient_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmbFunction.SetFocus
    End If
End Sub

Private Function LoadWallPaper(ByVal WallPaperFile As String) As Boolean
    Dim WallPaper As String
    On Error Resume Next
    
    WallPaper = App.Path & "\Bitmap\" & WallPaperFile
    
    If Dir$(WallPaper) <> "" Then
        Set pgPolarGraph.Picture = LoadPicture(WallPaper)
        LoadWallPaper = True
    Else
        LoadWallPaper = False
    End If
End Function
