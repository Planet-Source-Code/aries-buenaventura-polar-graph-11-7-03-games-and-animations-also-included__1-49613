VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   3180
      TabIndex        =   50
      Top             =   4560
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   4140
      TabIndex        =   49
      Top             =   4560
      Width           =   915
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "&Default"
      Height          =   315
      Left            =   60
      TabIndex        =   48
      Top             =   4560
      Width           =   915
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5100
      TabIndex        =   5
      Top             =   4560
      Width           =   915
   End
   Begin prjPolarGraph.PolarGraph pgPreview 
      Height          =   3015
      Left            =   420
      TabIndex        =   4
      Top             =   840
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5318
      OffsetDrawingAreaX=   0
      OffsetDrawingAreaY=   0
      ScrollBars      =   0
      Zoom            =   32
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   60
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame fraMain 
      Height          =   3975
      Left            =   180
      TabIndex        =   1
      Top             =   420
      Width           =   5715
      Begin VB.Frame fraPreview 
         Caption         =   "Preview"
         Height          =   3675
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   3255
         Begin VB.Frame fraTray 
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   120
            TabIndex        =   6
            Top             =   3180
            Width           =   3015
            Begin VB.TextBox txtUnitIn 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1980
               TabIndex        =   10
               Top             =   120
               Width           =   615
            End
            Begin VB.TextBox txtRadiusIn 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   840
               TabIndex        =   9
               Top             =   120
               Width           =   615
            End
            Begin VB.Label lblLabel 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Unit : "
               Height          =   195
               Index           =   8
               Left            =   1530
               TabIndex        =   8
               Top             =   180
               Width           =   420
            End
            Begin VB.Label lblLabel 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Radius : "
               Height          =   195
               Index           =   7
               Left            =   180
               TabIndex        =   7
               Top             =   180
               Width           =   630
            End
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Bitmap"
            Enabled         =   0   'False
            Height          =   315
            Left            =   840
            TabIndex        =   3
            Top             =   3300
            Visible         =   0   'False
            Width           =   1635
         End
      End
      Begin VB.Frame fraMisc 
         Height          =   3675
         Left            =   3480
         TabIndex        =   22
         Top             =   180
         Visible         =   0   'False
         Width           =   2115
         Begin VB.PictureBox picColor 
            Height          =   255
            Index           =   0
            Left            =   120
            ScaleHeight     =   195
            ScaleWidth      =   1455
            TabIndex        =   36
            Top             =   540
            Width           =   1515
         End
         Begin VB.PictureBox picColor 
            Height          =   255
            Index           =   1
            Left            =   120
            ScaleHeight     =   195
            ScaleWidth      =   1455
            TabIndex        =   35
            Top             =   1020
            Width           =   1515
         End
         Begin VB.CommandButton cmdColorOp 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   34
            Top             =   480
            Width           =   315
         End
         Begin VB.CommandButton cmdColorOp 
            Caption         =   "..."
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   33
            Top             =   960
            Width           =   315
         End
         Begin VB.CommandButton cmdColorOp 
            Caption         =   "..."
            Height          =   315
            Index           =   2
            Left            =   1680
            TabIndex        =   32
            Top             =   1440
            Width           =   315
         End
         Begin VB.PictureBox picColor 
            Height          =   255
            Index           =   2
            Left            =   120
            ScaleHeight     =   195
            ScaleWidth      =   1455
            TabIndex        =   31
            Top             =   1500
            Width           =   1515
         End
         Begin VB.Frame fraOrigin 
            Caption         =   "Origin"
            Height          =   1635
            Left            =   120
            TabIndex        =   23
            Top             =   1920
            Width           =   1875
            Begin VB.CheckBox chkAuto 
               Caption         =   "Auto"
               Height          =   315
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Value           =   1  'Checked
               Width           =   1515
            End
            Begin VB.TextBox txtAxis 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   420
               TabIndex        =   25
               Text            =   "0"
               Top             =   960
               Width           =   555
            End
            Begin VB.TextBox txtAxis 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   420
               TabIndex        =   24
               Text            =   "0"
               Top             =   600
               Width           =   555
            End
            Begin VB.Label lblY 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y : "
               Enabled         =   0   'False
               Height          =   195
               Left            =   180
               TabIndex        =   30
               Top             =   960
               Width           =   240
            End
            Begin VB.Label lblX 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "X : "
               Enabled         =   0   'False
               Height          =   195
               Left            =   180
               TabIndex        =   29
               Top             =   660
               Width           =   240
            End
            Begin VB.Label lblCX 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Center-X"
               Enabled         =   0   'False
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   1020
               TabIndex        =   28
               Top             =   660
               Width           =   615
            End
            Begin VB.Label lblCY 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Center-Y"
               Enabled         =   0   'False
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   1020
               TabIndex        =   27
               Top             =   1020
               Width           =   615
            End
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grid color"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   675
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fill color"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   38
            Top             =   840
            Width           =   570
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label color"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Top             =   1320
            Width           =   780
         End
      End
      Begin VB.Frame fraDetails 
         Height          =   3675
         Left            =   3480
         TabIndex        =   11
         Top             =   180
         Width           =   2115
         Begin VB.TextBox txtArea 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   660
            Width           =   1575
         End
         Begin VB.TextBox txtCircumference 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtDiameter 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1740
            Width           =   1575
         End
         Begin VB.TextBox txtRadius 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox txtUnit 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Area : "
            Height          =   195
            Index           =   9
            Left            =   180
            TabIndex        =   21
            Top             =   420
            Width           =   465
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Circumference : "
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   20
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Diameter : "
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   19
            Top             =   1500
            Width           =   765
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Radius : (radius * unit)"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   18
            Top             =   2040
            Width           =   1545
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Unit : "
            Height          =   195
            Index           =   13
            Left            =   180
            TabIndex        =   17
            Top             =   2640
            Width           =   420
         End
      End
      Begin VB.Frame fraBk 
         Height          =   3675
         Left            =   3480
         TabIndex        =   40
         Top             =   180
         Visible         =   0   'False
         Width           =   2115
         Begin VB.CommandButton cmdBitmap 
            Caption         =   "Bitmap"
            Height          =   315
            Left            =   240
            TabIndex        =   47
            Top             =   720
            Width           =   1635
         End
         Begin VB.CommandButton cmdBkColor 
            Caption         =   "Color"
            Height          =   315
            Left            =   240
            TabIndex        =   46
            Top             =   360
            Width           =   1635
         End
         Begin VB.Frame fraBmpOp 
            Caption         =   "Bitmap Options"
            Height          =   2355
            Left            =   120
            TabIndex        =   41
            Top             =   1140
            Width           =   1875
            Begin VB.OptionButton optBmp 
               Caption         =   "None"
               Height          =   255
               Index           =   0
               Left            =   180
               TabIndex        =   45
               Top             =   360
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.OptionButton optBmp 
               Caption         =   "Center"
               Height          =   255
               Index           =   1
               Left            =   180
               TabIndex        =   44
               Top             =   660
               Width           =   1455
            End
            Begin VB.OptionButton optBmp 
               Caption         =   "Stretch"
               Height          =   255
               Index           =   2
               Left            =   180
               TabIndex        =   43
               Top             =   960
               Width           =   1455
            End
            Begin VB.OptionButton optBmp 
               Caption         =   "Tile"
               Height          =   255
               Index           =   3
               Left            =   180
               TabIndex        =   42
               Top             =   1260
               Width           =   1455
            End
         End
      End
   End
   Begin MSComctlLib.TabStrip tbsOption 
      Height          =   4455
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   7858
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Details"
            Key             =   "Details"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Miscellaneous"
            Key             =   "Misc"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Background"
            Key             =   "Bk"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAuto_Click()
    Dim bVal As Boolean
    
    If CBool(chkAuto.Value) Then
        pgPreview.Polar.Origin.Auto = True
    ElseIf chkAuto.Value = vbUnchecked Then
        pgPreview.Polar.Origin.Auto = False
        pgPreview.Polar.Origin.SetPos Val(txtAxis(0).Text), _
                                      Val(txtAxis(1).Text)
    End If
    
    bVal = Not CBool(chkAuto.Value)
    
    lblCX.Enabled = bVal
    lblCY.Enabled = bVal
    lblX.Enabled = bVal
    lblY.Enabled = bVal
    txtAxis(0).Enabled = bVal
    txtAxis(1).Enabled = bVal
    
    pgPreview.Refresh
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
    On Error GoTo ErrHandler
    
    With frmMain.pgPolarGraph
        .Polar.Radius = pgPreview.Polar.Radius
        .Polar.Unit = pgPreview.Polar.Unit
        
        .Polar.Origin.Auto = pgPreview.Polar.Origin.Auto
        .Polar.Origin.X = pgPreview.Polar.Origin.X
        .Polar.Origin.Y = pgPreview.Polar.Origin.Y
        .Polar.GridColor = pgPreview.Polar.GridColor
        .Polar.FillColor = pgPreview.Polar.FillColor
        .Polar.LabelColor = pgPreview.Polar.LabelColor
        
        .AutoUpdate = False
        .BackColor = pgPreview.BackColor
        .PictureStyle = pgPreview.PictureStyle
        Set .Picture = pgPreview.Picture
        .AutoUpdate = True
        .Refresh
        
        Dim Temp As New StdPicture
        Dim WallPaperFile As String
        
        If .PictureStyle = ajb_PSNone Then
            WallPaperFile = App.Path & "\Bitmap\WP1.ajb"
        ElseIf .PictureStyle = ajb_PSCenter Then
            WallPaperFile = App.Path & "\Bitmap\WP2.ajb"
        ElseIf .PictureStyle = ajb_PSStretch Then
            WallPaperFile = App.Path & "\Bitmap\WP3.ajb"
        Else
            WallPaperFile = App.Path & "\Bitmap\WP4.ajb"
        End If
        
        If Dir$(App.Path & "\Bitmap\*.ajb") <> "" Then
            Kill App.Path & "\Bitmap\*.ajb"
        End If
        
        Set Temp = .Picture
        If Temp Then SavePicture Temp, WallPaperFile
        
        cmdApply.Enabled = False
    End With
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbOKOnly Or vbCritical, "Polar Graph 1.0"
End Sub

Private Sub cmdBitmap_Click()
    On Error GoTo ErrHandler
    
    With dlgDialog
        .Filter = "Bitmap Files (*.bmp) | *.bmp; |" _
                  & "JPEG (*.JPG,*.JPEG) | *.jpg; *.jpeg; |" _
                  & "GIF (*.GIF) | *.GIF; |" _
                  & "All Picture Files | *.bmp; *.gif; *.jpg; *.jpeg; |" _
                  & "All Files (*.*) | *.*"
        .FilterIndex = 4
        .InitDir = ""
        .Filename = ""
        .ShowOpen
            
        If .Filename <> "" Then
            Set pgPreview.Picture = LoadPicture(.Filename)
            cmdClear.Enabled = True
            
            If Not cmdApply.Enabled Then cmdApply.Enabled = True
        End If
    End With
    Exit Sub
    
ErrHandler:
    If Err.Number = 32755 Then ' Cancel Selected
    Else
        MsgBox Err.Description, vbOKOnly Or vbInformation, "Error"
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Set pgPreview.Picture = Nothing
    cmdClear.Enabled = False
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub cmdBkColor_Click()
    On Error GoTo ErrHandler
    
    dlgDialog.ShowColor
    pgPreview.BackColor = dlgDialog.Color
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
    Exit Sub

ErrHandler:
End Sub

Private Sub cmdColorOp_Click(Index As Integer)
    On Error GoTo ErrHandler
    
    dlgDialog.ShowColor
    picColor(Index).BackColor = dlgDialog.Color
    
    Select Case Index
    Case Is = 0 ' Grid color
        pgPreview.Polar.GridColor = dlgDialog.Color
    Case Is = 1 ' Fill color
        pgPreview.Polar.FillColor = dlgDialog.Color
    Case Is = 2 ' Label color
        pgPreview.Polar.LabelColor = dlgDialog.Color
    End Select
    
    pgPreview.Refresh
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
    Exit Sub

ErrHandler:
End Sub

Private Sub cmdDefault_Click()
    With pgPreview
        .Polar.Radius = 5
        .Polar.Unit = 0.5
    
        .Polar.Origin.Auto = True
        .Polar.Origin.X = 0
        .Polar.Origin.Y = 0
        .Polar.GridColor = &HE0E0E0
        .Polar.FillColor = vbYellow
        .Polar.LabelColor = &H808080
    
        .AutoUpdate = False
        .BackColor = &HFFFFFF
        Set .Picture = Nothing
        .PictureStyle = ajb_PSNone
        .AutoUpdate = True
        .Refresh
    
        txtRadiusIn.Text = .Polar.Radius
        txtUnitIn.Text = .Polar.Unit
        Call UpdatePolarDetails
        
        txtAxis(0).Text = .Polar.Origin.X
        txtAxis(1).Text = .Polar.Origin.Y
        chkAuto.Value = Abs(CInt(.Polar.Origin.Auto))
                                                
        picColor(0).BackColor = .Polar.GridColor
        picColor(1).BackColor = .Polar.FillColor
        picColor(2).BackColor = .Polar.LabelColor
        
        optBmp(.PictureStyle).Value = True
        cmdClear.Enabled = False
        If Not cmdApply.Enabled Then cmdApply.Enabled = True
    End With
End Sub

Private Sub cmdOk_Click()
    If cmdApply.Enabled Then Call cmdApply_Click
    Unload Me
End Sub

Private Sub Form_Load()
    With frmMain.pgPolarGraph
        pgPreview.Polar.Radius = .Polar.Radius
        pgPreview.Polar.Unit = .Polar.Unit
        
        pgPreview.Polar.Origin.Auto = .Polar.Origin.Auto
        pgPreview.Polar.Origin.X = .Polar.Origin.X
        pgPreview.Polar.Origin.Y = .Polar.Origin.Y
        pgPreview.Polar.GridColor = .Polar.GridColor
        pgPreview.Polar.FillColor = .Polar.FillColor
        pgPreview.Polar.LabelColor = .Polar.LabelColor
        
        pgPreview.AutoUpdate = False
        pgPreview.AspectRatio = .AspectRatio
        pgPreview.BackColor = .BackColor
        Set pgPreview.Picture = .Picture
        pgPreview.PictureStyle = .PictureStyle
        pgPreview.AutoUpdate = True
        pgPreview.Refresh
    End With
    
    With pgPreview
        txtRadiusIn.Text = .Polar.Radius
        txtUnitIn.Text = .Polar.Unit
        Call UpdatePolarDetails
        
        txtAxis(0).Text = .Polar.Origin.X
        txtAxis(1).Text = .Polar.Origin.Y
        chkAuto.Value = Abs(CInt(.Polar.Origin.Auto))
                                                
        picColor(0).BackColor = .Polar.GridColor
        picColor(1).BackColor = .Polar.FillColor
        picColor(2).BackColor = .Polar.LabelColor
        
        optBmp(.PictureStyle).Value = True
        
        Dim Temp As New StdPicture
        
        Set Temp = .Picture
        If Temp Then cmdClear.Enabled = True
    End With
End Sub

Private Sub optBmp_Click(Index As Integer)
    pgPreview.PictureStyle = Index
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub tbsOption_Click()
    fraBk.Visible = False
    fraDetails.Visible = False
    fraMisc.Visible = False
    fraTray.Visible = False
    cmdClear.Visible = False
       
    Select Case tbsOption.SelectedItem.Index
    Case Is = 1 ' Details
        fraDetails.Visible = True
        fraTray.Visible = True
    Case Is = 2 ' Miscellaneous
        fraMisc.Visible = True
        fraTray.Visible = True
    Case Is = 3 ' Background
        fraBk.Visible = True
        cmdClear.Visible = True
    End Select
End Sub

Private Sub txtAxis_Change(Index As Integer)
    If Trim$(txtAxis(Index).Text) <> "" Then
        If IsNumeric(txtAxis(Index).Text) Then
            Dim Temp As Integer
            On Error GoTo ErrHandler
            
            Temp = Val(txtAxis(Index).Text)
            With pgPreview.Polar
                If Index = 0 Then ' X-Axis
                    If Temp = .Origin.X Then Exit Sub
                Else              ' Y-Axis
                    If Temp = .Origin.Y Then Exit Sub
                End If
                
                If (Temp >= -10000) And (Temp <= 10000) Then
                    .Origin.Auto = False
                    .Origin.SetPos Val(txtAxis(0).Text), _
                                   Val(txtAxis(1).Text)
                    pgPreview.Refresh
                End If
            End With
        End If
    End If
    
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
    Exit Sub
    
ErrHandler:
End Sub

Private Sub txtAxis_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If (KeyAscii >= vbKey0) And (KeyAscii <= vbKey9) Then
    ElseIf KeyAscii = &H2D Then ' minus sign
        If InStr(Me.ActiveControl.Text, "-") > 0 Then
            Me.ActiveControl.Text = Abs(Val(Me.ActiveControl.Text))
            SendKeys "{End}", False
        Else
            Me.ActiveControl.Text = "-" & Me.ActiveControl.Text
            SendKeys "{End}", False
        End If
        
        KeyAscii = 0
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtAxis_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
    Case Is = 0 ' X-Axis
        If Trim$(Me.ActiveControl.Text) = "" Then
            Me.ActiveControl.Text = pgPreview.Polar.Origin.X
        End If
    Case Is = 1 ' Y-Axis
        If Trim$(Me.ActiveControl.Text) = "" Then
            Me.ActiveControl.Text = pgPreview.Polar.Origin.Y
        End If
    End Select
End Sub

Private Sub txtRadiusIn_Change()
    If Trim$(txtRadiusIn.Text) <> "" Then
        Dim Temp As Integer
        On Error GoTo ErrHandler
        
        Temp = CInt(Val(txtRadiusIn.Text))
        If Temp = pgPreview.Polar.Radius Then Exit Sub
        
        If (Temp > 0) And (Temp <= 500) Then
            pgPreview.Polar.Radius = Temp
            pgPreview.Refresh
            Call UpdatePolarDetails
        End If
    End If
    
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
    Exit Sub
    
ErrHandler:
End Sub

Private Sub txtRadiusIn_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If (KeyAscii >= vbKey0) And (KeyAscii <= vbKey9) Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtRadiusIn_Validate(Cancel As Boolean)
    If Trim$(txtRadiusIn.Text) = "" Then
        txtRadiusIn.Text = pgPreview.Polar.Radius
    End If
End Sub

Private Sub txtUnitIn_Change()
    If Trim$(txtUnitIn.Text) <> "" Then
        If IsNumeric(txtUnitIn.Text) Then
            Dim Temp As Single
            On Error GoTo ErrHandler
            
            Temp = Val(txtUnitIn.Text)
            If Temp = pgPreview.Polar.Unit Then Exit Sub
            
            If (Temp >= -10000) And (Temp <= 10000) Then
                pgPreview.Polar.Unit = Temp
                Call UpdatePolarDetails
            End If
        End If
    End If
    
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
    Exit Sub
    
ErrHandler:
End Sub

Private Sub UpdatePolarDetails()
    With pgPreview
        txtArea.Text = Round(.Polar.Area, 4)
        txtCircumference.Text = Round(.Polar.Circumference, 4)
        txtDiameter.Text = Round(.Polar.Diameter, 4)
        txtRadius.Text = Round(.Polar.Radius * .Polar.Unit, 4)
        txtUnit.Text = .Polar.Unit
    End With
End Sub

Private Sub txtUnitIn_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If (KeyAscii >= vbKey0) And (KeyAscii <= vbKey9) Then
    ElseIf KeyAscii = &H2E Then ' Decimal
        If InStr(Me.ActiveControl.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    ElseIf KeyAscii = &H2D Then ' minus sign
        If InStr(txtUnitIn.Text, "-") > 0 Then
            txtUnitIn.Text = Abs(Val(txtUnitIn.Text))
            SendKeys "{End}", False
        Else
            txtUnitIn.Text = "-" & txtUnitIn.Text
            SendKeys "{End}", False
        End If
        
        KeyAscii = 0
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtUnitIn_Validate(Cancel As Boolean)
    If Trim$(txtUnitIn.Text) = "" Then
        txtUnitIn.Text = pgPreview.Polar.Unit
    End If
End Sub
 
