VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl PolarGraph 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0099A8AC&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2850
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   190
   Begin MSComDlg.CommonDialog dlgPolar 
      Left            =   1200
      Top             =   1260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar hsScroll 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2700
      Width           =   2595
   End
   Begin VB.VScrollBar vsScroll 
      Height          =   2715
      Left            =   2580
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picScroll 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2640
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   2
      Top             =   2700
      Width           =   195
   End
   Begin VB.PictureBox picHolderRV 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   435
      Begin VB.PictureBox picRulerV 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   60
         MousePointer    =   99  'Custom
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   9
         Top             =   60
         Width           =   315
         Begin VB.Line linCrossV 
            BorderColor     =   &H00404040&
            BorderStyle     =   3  'Dot
            Index           =   2
            Visible         =   0   'False
            X1              =   20
            X2              =   20
            Y1              =   0
            Y2              =   80
         End
         Begin VB.Line linCrossV 
            BorderColor     =   &H00404040&
            BorderStyle     =   3  'Dot
            Index           =   1
            Visible         =   0   'False
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   80
         End
      End
   End
   Begin VB.PictureBox picHolderRH 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   480
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   435
      Begin VB.PictureBox picRulerH 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   60
         MousePointer    =   99  'Custom
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   8
         Top             =   60
         Width           =   315
         Begin VB.Line linCrossH 
            BorderColor     =   &H00404040&
            BorderStyle     =   3  'Dot
            Index           =   2
            Visible         =   0   'False
            X1              =   20
            X2              =   20
            Y1              =   0
            Y2              =   80
         End
         Begin VB.Line linCrossH 
            BorderColor     =   &H00404040&
            BorderStyle     =   3  'Dot
            Index           =   1
            Visible         =   0   'False
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   80
         End
      End
   End
   Begin VB.PictureBox picRulerJ 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   960
      MousePointer    =   99  'Custom
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   6
      ToolTipText     =   "Left click to reset the drawing area: Right click to change the ruler mode."
      Top             =   2160
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5250
      Left            =   0
      ScaleHeight     =   348
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   348
      TabIndex        =   0
      Top             =   0
      Width           =   5250
      Begin VB.Label lblTrackAngle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angle?"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   60
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape shpSelect 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  'Dot
         Height          =   435
         Left            =   360
         Top             =   1020
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Shape shpCircle 
         BorderStyle     =   3  'Dot
         Height          =   615
         Left            =   300
         Shape           =   3  'Circle
         Top             =   420
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line linCrossH 
         BorderColor     =   &H00404040&
         BorderStyle     =   3  'Dot
         Index           =   0
         Visible         =   0   'False
         X1              =   16
         X2              =   16
         Y1              =   28
         Y2              =   108
      End
      Begin VB.Line linCrossV 
         BorderColor     =   &H00404040&
         BorderStyle     =   3  'Dot
         Index           =   0
         Visible         =   0   'False
         X1              =   8
         X2              =   8
         Y1              =   28
         Y2              =   108
      End
   End
   Begin VB.PictureBox picShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   7
      Top             =   0
      Width           =   315
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   795
   End
End
Attribute VB_Name = "PolarGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' use to identify the file
Private Const SIGNATURE = "POLAR GRAPH 1.0"

Private Const MAX_DAREA_WIDTH = 600
Private Const MAX_DAREA_HEIGHT = 600

Private Const OffsetShadowX = 3
Private Const OffsetShadowY = 3

Public Enum ActiveToolConstants
    ajb_TArrow
    ajb_TCircular
    ajb_TCross
    ajb_TPan
    ajb_TSelect
End Enum

Public Enum GridConstants
    ajb_GCNone
    ajb_GCMinor
    ajb_GCMajor
End Enum

Public Enum PictureStyleConstants
    ajb_PSNone
    ajb_PSCenter
    ajb_PSStretch
    ajb_PSTile
End Enum

Public Enum RulerConstants
    ajb_RAxis
    ajb_RUnit
End Enum

Public Enum ScrollBarsConstants
    ajb_SBNone
    ajb_SBHorizontal
    ajb_SBVertical
    ajb_SBBoth
End Enum

Private Type PolarGraphProperties
    AspectRatio          As Single
    AutoUpdate           As Boolean
    BackColor            As OLE_COLOR
    Comment              As String
    OffsetDrawingAreaX   As Integer
    OffsetDrawingAreaY   As Integer
    PictureHolder        As Picture
    PictureStyle         As PictureStyleConstants
    ScrollBars           As ScrollBarsConstants
    SetActiveTool        As ActiveToolConstants
    SetRulerMode         As RulerConstants
    ShowCoordinateSystem As Boolean
    ShowGrid             As GridConstants
    ShowLabel            As Boolean
    ShowRuler            As Boolean
    StopAni        As Boolean
    Zoom                 As Integer
End Type

Public Event AniStat(bPlay As Boolean)
Public Event Click()
Public Event DblClick()
Public Event Graph(IndexKey As Variant)
Public Event KeyDown(KeyCode As Integer, _
                     Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, _
                   Shift As Integer)
Public Event Location(ByVal X As Single, _
                      ByVal Y As Single, _
                      ByVal Angle As Single, _
                      ByVal Area As Single, _
                      ByVal Circumference, _
                      ByVal Diameter As Single, _
                      ByVal Radius As Single)
Public Event MouseDown(Button As Integer, _
                       Shift As Integer, _
                       X As Single, _
                       Y As Single)
Public Event MouseMove(Button As Integer, _
                       Shift As Integer, _
                       X As Single, _
                       Y As Single)
Public Event MouseUp(Button As Integer, _
                     Shift As Integer, _
                     X As Single, _
                     Y As Single)
Public Event Status(ByVal Coefficient As String, _
                    ByVal Equation As String, _
                    ByVal Color As Long, _
                    ByVal Degrees As Single, _
                    ByVal Data As Variant)

Dim mvarPlot       As New Plot
Dim mvarPolar      As New Polar
Dim GraphHolder    As New GrapInfo
Dim MyProp         As PolarGraphProperties
Dim pox            As Single
Dim poy            As Single
Dim xPos           As Single
Dim yPos           As Single
Dim PolarScale     As Single
Dim MousePos       As POINTAPI
Dim DAreaMousePos  As POINTAPI
Dim hRulerMousePos As POINTAPI
Dim vRulerMousePos As POINTAPI
Dim SelectX1       As Single
Dim SelectY1       As Single
Dim SelectX2       As Single
Dim SelectY2       As Single

Dim FSys           As New Scripting.FileSystemObject
Dim scEval         As New MSScriptControl.ScriptControl

Public Property Get AspectRatio() As Single
    AspectRatio = MyProp.AspectRatio
End Property

Public Property Let AspectRatio(Value As Single)
    If Value <= 0 Then
        MsgBox "The AspectRatio value must be greater than 0.", _
               vbExclamation, "Polar Graph"
    ElseIf Value > 5 Then
        MsgBox "The AspectRatio value must be less than or equal to 5.", _
               vbExclamation, "Polar Graph"
    Else
        Dim Temp As Single
        
        Temp = Round(Value, 4)
        MyProp.AspectRatio = IIf(Temp = 0, 0.0001, Temp)
        PropertyChanged "AspectRatio"
        Call RedrawPolar(3)
    End If
End Property

Public Property Get AutoUpdate() As Boolean
Attribute AutoUpdate.VB_MemberFlags = "400"
    AutoUpdate = MyProp.AutoUpdate
End Property

Public Property Let AutoUpdate(bVal As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    MyProp.AutoUpdate = bVal
    PropertyChanged "AutoUpdate"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = MyProp.BackColor
End Property

Public Property Let BackColor(BkColor As OLE_COLOR)
    MyProp.BackColor = BkColor
    PropertyChanged "BackColor"
    Call RedrawPolar(1)
End Property

Public Property Get Comment() As String
    Comment = MyProp.Comment
End Property

Public Property Let Comment(Data As String)
    MyProp.Comment = Data
    PropertyChanged "Comment"
End Property

Public Property Get OffsetDrawingAreaX() As Integer
    OffsetDrawingAreaX = MyProp.OffsetDrawingAreaX
End Property

Public Property Let OffsetDrawingAreaX(X As Integer)
    MyProp.OffsetDrawingAreaX = X
    PropertyChanged "OffsetDrawingAreaX"
    picGraph.Left = OffsetDrawingAreaX
    Call picGraph_Resize
End Property

Public Property Get OffsetDrawingAreaY() As Integer
    OffsetDrawingAreaY = MyProp.OffsetDrawingAreaY
End Property

Public Property Let OffsetDrawingAreaY(Y As Integer)
    MyProp.OffsetDrawingAreaY = Y
    PropertyChanged "OffsetDrawingAreaY"
    
    picGraph.Top = OffsetDrawingAreaY
    Call picGraph_Resize
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = picGraph.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set picGraph.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = picGraph.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    picGraph.MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Picture() As Picture
    Set Picture = MyProp.PictureHolder
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set MyProp.PictureHolder = New_Picture
    PropertyChanged "Picture"
    Call RedrawPolar(1)
End Property

Public Property Get PictureStyle() As PictureStyleConstants
    PictureStyle = MyProp.PictureStyle
End Property

Public Property Let PictureStyle(bVal As PictureStyleConstants)
    MyProp.PictureStyle = bVal
    PropertyChanged "PictureStyle"
    
    Dim PBuffer As New StdPicture
    
    Set PBuffer = MyProp.PictureHolder
    If PBuffer Then
        Call RedrawPolar(1)
    End If
End Property

Public Property Get Polar() As Polar
    Set Polar = mvarPolar
End Property

Public Property Set Polar(vData As Polar)
    Set mvarPolar = vData
End Property

Public Property Get Plot() As Plot
    Set Plot = mvarPlot
End Property

Public Property Set Plot(vData As Plot)
    Set mvarPlot = vData
End Property

Public Property Get ScrollBars() As ScrollBarsConstants
    ScrollBars = MyProp.ScrollBars
End Property

Public Property Let ScrollBars(ScrollBarOp As ScrollBarsConstants)
    MyProp.ScrollBars = ScrollBarOp
    PropertyChanged "ScrollBars"
    
    hsScroll.Visible = False
    vsScroll.Visible = False
    picScroll.Visible = False
        
    Call picGraph_Resize
        
    If MyProp.ScrollBars = ajb_SBNone Then
    ElseIf MyProp.ScrollBars = ajb_SBHorizontal Then
        hsScroll.Visible = True
        vsScroll.Visible = False
        picScroll.Visible = False
    ElseIf MyProp.ScrollBars = ajb_SBVertical Then
        hsScroll.Visible = False
        vsScroll.Visible = True
        picScroll.Visible = False
    Else
        hsScroll.Visible = True
        vsScroll.Visible = True
        picScroll.Visible = True
    End If
End Property

Public Property Get SetActiveTool() As ActiveToolConstants
    SetActiveTool = MyProp.SetActiveTool
End Property

Public Property Let SetActiveTool(ToolOp As ActiveToolConstants)
    MyProp.SetActiveTool = ToolOp
    PropertyChanged "SetActiveTool"
    
    If shpCircle.Visible Then shpCircle.Visible = False
    If shpSelect.Visible Then shpSelect.Visible = False
    If linCrossH(0).Visible Then linCrossH(0).Visible = False
    If linCrossV(0).Visible Then linCrossV(0).Visible = False
    If linCrossH(1).Visible Then linCrossH(1).Visible = False
    If linCrossV(1).Visible Then linCrossV(1).Visible = False
    If linCrossH(2).Visible Then linCrossH(2).Visible = False
    If linCrossV(2).Visible Then linCrossV(2).Visible = False
    
    Select Case MyProp.SetActiveTool
    Case Is = ajb_TArrow
    Case Is = ajb_TCircular
        If Not shpCircle.Visible Then shpCircle.Visible = True
        If Not linCrossH(1).Visible Then linCrossH(1).Visible = True
        If Not linCrossV(1).Visible Then linCrossV(1).Visible = True
        If Not linCrossH(2).Visible Then linCrossH(2).Visible = True
        If Not linCrossV(2).Visible Then linCrossV(2).Visible = True
        
        shpCircle.Move picGraph.ScaleWidth / 2, picGraph.ScaleHeight / 2, 0, 0
        
        Dim i As Integer
        
        For i = 1 To 2
            linCrossH(i).X1 = picRulerH.ScaleWidth / 2
            linCrossH(i).Y1 = 0
            linCrossH(i).X2 = picRulerH.ScaleWidth / 2
            linCrossH(i).Y2 = picRulerH.ScaleHeight
            
            linCrossV(i).X1 = 0
            linCrossV(i).Y1 = picRulerV.ScaleHeight / 2
            linCrossV(i).X2 = picRulerV.ScaleWidth
            linCrossV(i).Y2 = picRulerV.ScaleHeight / 2
        Next i
    Case Is = ajb_TCross
        If Not linCrossH(0).Visible Then linCrossH(0).Visible = True
        If Not linCrossV(0).Visible Then linCrossV(0).Visible = True
        If Not linCrossH(1).Visible Then linCrossH(1).Visible = True
        If Not linCrossV(1).Visible Then linCrossV(1).Visible = True
        
        linCrossH(0).X1 = 0
        linCrossH(0).Y1 = picGraph.ScaleHeight / 2
        linCrossH(0).X2 = picGraph.ScaleWidth
        linCrossH(0).Y2 = picGraph.ScaleHeight / 2
        
        linCrossV(0).X1 = picGraph.ScaleWidth / 2
        linCrossV(0).Y1 = 0
        linCrossV(0).X2 = picGraph.ScaleWidth / 2
        linCrossV(0).Y2 = picGraph.ScaleHeight
        
        linCrossH(1).X1 = picRulerH.ScaleWidth / 2
        linCrossH(1).Y1 = 0
        linCrossH(1).X2 = picRulerH.ScaleWidth / 2
        linCrossH(1).Y2 = picRulerH.ScaleHeight
        
        linCrossV(1).X1 = 0
        linCrossV(1).Y1 = picRulerV.ScaleHeight / 2
        linCrossV(1).X2 = picRulerV.ScaleWidth
        linCrossV(1).Y2 = picRulerV.ScaleHeight / 2
    Case Is = ajb_TPan
    Case Is = ajb_TSelect
    End Select
End Property

Public Property Get SetRulerMode() As RulerConstants
    SetRulerMode = MyProp.SetRulerMode
End Property

Public Property Let SetRulerMode(RulerOp As RulerConstants)
    MyProp.SetRulerMode = RulerOp
    PropertyChanged "SetRulerMode"
    Call RedrawRuler
End Property

Public Property Get ShowCoordinateSystem() As Boolean
    ShowCoordinateSystem = MyProp.ShowCoordinateSystem
End Property

Public Property Let ShowCoordinateSystem(bVal As Boolean)
    MyProp.ShowCoordinateSystem = bVal
    PropertyChanged "ShowCoordinateSystem"
    Call RedrawPolar(1)
End Property

Public Property Get ShowGrid() As GridConstants
    ShowGrid = MyProp.ShowGrid
End Property

Public Property Let ShowGrid(GridOp As GridConstants)
    MyProp.ShowGrid = GridOp
    PropertyChanged "ShowGrid"
    Call RedrawPolar(1)
End Property

Public Property Get ShowLabel() As Boolean
    ShowLabel = MyProp.ShowLabel
End Property

Public Property Let ShowLabel(bVal As Boolean)
    MyProp.ShowLabel = bVal
    PropertyChanged "ShowLabel"
    Call RedrawPolar(1)
End Property

Public Property Get ShowRuler() As Boolean
    ShowRuler = MyProp.ShowRuler
End Property

Public Property Let ShowRuler(bShow As Boolean)
    MyProp.ShowRuler = bShow
    PropertyChanged "ShowRuler"
    
    picRulerJ.Visible = MyProp.ShowRuler
    picHolderRH.Visible = MyProp.ShowRuler
    picHolderRV.Visible = MyProp.ShowRuler
End Property

Public Property Get StopAni() As Boolean
    StopAni = MyProp.StopAni
End Property

Public Property Let StopAni(bStop As Boolean)
    MyProp.StopAni = bStop
    PropertyChanged "StopAni"
End Property

Public Property Get Zoom() As Integer
    Zoom = MyProp.Zoom
End Property

Public Property Let Zoom(Percent As Integer)
    MyProp.Zoom = Percent
    PropertyChanged "Zoom"
    
    hsScroll.Value = 0
    vsScroll.Value = 0
    picGraph.Move picGraph.Left, picGraph.Top, _
                  MAX_DAREA_WIDTH * (Percent * 0.01), _
                  MAX_DAREA_HEIGHT * (Percent * 0.01)
    
    Call RedrawPolar(3)
End Property

Private Sub picGraph_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub picRulerH_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set picRulerH.MouseIcon = LoadResPicture(105, vbResCursor)
End Sub

Private Sub picRulerV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set picRulerV.MouseIcon = LoadResPicture(105, vbResCursor)
End Sub

Private Sub UserControl_InitProperties()
    MyProp.AspectRatio = 1.5
    MyProp.AutoUpdate = True
    MyProp.BackColor = &HFFFFFF
    MyProp.OffsetDrawingAreaX = 0
    MyProp.OffsetDrawingAreaY = 0
    Set MyProp.PictureHolder = Nothing
    MyProp.PictureStyle = ajb_PSNone
    MyProp.ScrollBars = ajb_SBBoth
    MyProp.SetActiveTool = ajb_TArrow
    MyProp.SetRulerMode = ajb_RUnit
    MyProp.ShowCoordinateSystem = True
    MyProp.ShowGrid = ajb_GCMajor
    MyProp.ShowLabel = True
    MyProp.ShowRuler = False
    MyProp.StopAni = True
    MyProp.Zoom = 50
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    MyProp.AspectRatio = PropBag.ReadProperty("AspectRatio", 1.5)
    MyProp.AutoUpdate = PropBag.ReadProperty("AutoUpdate", True)
    MyProp.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    MyProp.Comment = PropBag.ReadProperty("Comment", "")
    MyProp.OffsetDrawingAreaX = PropBag.ReadProperty("OffsetDrawingAreaX", 0)
    MyProp.OffsetDrawingAreaY = PropBag.ReadProperty("OffsetDrawingAreaY", 0)
    Set picGraph.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    picGraph.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MyProp.PictureHolder = PropBag.ReadProperty("Picture", Nothing)
    MyProp.PictureStyle = PropBag.ReadProperty("PictureStyle", ajb_PSNone)
    MyProp.ScrollBars = PropBag.ReadProperty("ScrollBars", ajb_SBBoth)
    MyProp.SetActiveTool = PropBag.ReadProperty("SetActiveTool", ajb_TArrow)
    MyProp.SetRulerMode = PropBag.ReadProperty("SetRulerMode", ajb_RUnit)
    MyProp.ShowCoordinateSystem = PropBag.ReadProperty("ShowCoordinateSystem", True)
    MyProp.ShowGrid = PropBag.ReadProperty("ShowGrid", ajb_GCMajor)
    MyProp.ShowLabel = PropBag.ReadProperty("ShowLabel", True)
    MyProp.ShowRuler = PropBag.ReadProperty("ShowRuler", False)
    MyProp.StopAni = PropBag.ReadProperty("StopAni", True)
    MyProp.Zoom = PropBag.ReadProperty("Zoom", 50)
End Sub

Private Sub UserControl_Terminate()
    If Not MyProp.StopAni Then
        MyProp.StopAni = True
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AspectRatio", MyProp.AspectRatio, 1.5)
    Call PropBag.WriteProperty("AutoUpdate", MyProp.AutoUpdate, True)
    Call PropBag.WriteProperty("BackColor", MyProp.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Comment", MyProp.Comment, "")
    Call PropBag.WriteProperty("OffsetDrawingAreaX", MyProp.OffsetDrawingAreaX, True)
    Call PropBag.WriteProperty("OffsetDrawingAreaY", MyProp.OffsetDrawingAreaY, True)
    Call PropBag.WriteProperty("MouseIcon", picGraph.MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", picGraph.MousePointer, 0)
    Call PropBag.WriteProperty("Picture", MyProp.PictureHolder, Nothing)
    Call PropBag.WriteProperty("PictureStyle", MyProp.PictureStyle, ajb_PSNone)
    Call PropBag.WriteProperty("ScrollBars", MyProp.ScrollBars, ajb_SBBoth)
    Call PropBag.WriteProperty("SetActiveTool", MyProp.SetActiveTool, ajb_TArrow)
    Call PropBag.WriteProperty("SetRulerMode", MyProp.SetRulerMode, ajb_RUnit)
    Call PropBag.WriteProperty("ShowCoordinateSystem", MyProp.ShowCoordinateSystem, True)
    Call PropBag.WriteProperty("ShowGrid", MyProp.ShowGrid, ajb_GCMajor)
    Call PropBag.WriteProperty("ShowLabel", MyProp.ShowLabel, True)
    Call PropBag.WriteProperty("ShowRuler", MyProp.ShowRuler, False)
    Call PropBag.WriteProperty("StopAni", MyProp.StopAni, True)
    Call PropBag.WriteProperty("Zoom", MyProp.Zoom, 50)
End Sub

Private Sub picRulerJ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then
        hsScroll.Value = 0
        vsScroll.Value = 0
        OffsetDrawingAreaX = picRulerJ.Width + 5
        OffsetDrawingAreaY = picRulerJ.Height + 5
        Call picGraph_Resize
    ElseIf Button And vbRightButton Then
        If MyProp.SetRulerMode = ajb_RUnit Then
            SetRulerMode = ajb_RAxis
        Else
            SetRulerMode = ajb_RUnit
        End If
        Call RedrawRuler
    End If
    
    Set picRulerJ.MouseIcon = LoadResPicture(102, vbResCursor)
End Sub

Private Sub picRulerJ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set picRulerJ.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub picRulerH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    hRulerMousePos.X = X
    hRulerMousePos.Y = Y
    
    Set picRulerH.MouseIcon = LoadResPicture(106, vbResCursor)
End Sub

Private Sub picRulerH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then
        Dim NewPos As POINTAPI
        
        NewPos.X = picGraph.Left - (hRulerMousePos.X - X)
        NewPos.Y = picGraph.Top - (hRulerMousePos.Y - Y)
        
        OffsetDrawingAreaX = NewPos.X
    End If
End Sub

Private Sub picRulerV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    vRulerMousePos.X = X
    vRulerMousePos.Y = Y
    
    Set picRulerV.MouseIcon = LoadResPicture(106, vbResCursor)
End Sub

Private Sub picRulerV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbLeftButton Then
        Dim NewPos As POINTAPI
        
        NewPos.X = picGraph.Left - (vRulerMousePos.X - X)
        NewPos.Y = picGraph.Top - (vRulerMousePos.Y - Y)
        
        OffsetDrawingAreaY = NewPos.Y
    End If
End Sub

Private Sub picGraph_Click()
    RaiseEvent Click
End Sub

Private Sub picGraph_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub picGraph_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub picGraph_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub picGraph_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
    Select Case MyProp.SetActiveTool
    Case Is = ajb_TArrow
    Case Is = ajb_TCircular
    Case Is = ajb_TCross
    Case Is = ajb_TPan
        If Button And vbLeftButton Then
            DAreaMousePos.X = X
            DAreaMousePos.Y = Y
        End If
    Case Is = ajb_TSelect
        If Button And vbLeftButton Then
            If shpSelect.Visible Then shpSelect.Visible = False
        
            picGraph.DrawMode = vbXorPen
            picGraph.DrawStyle = vbDot
        
            SelectX1 = X
            SelectY1 = Y
            SelectX2 = SelectX1
            SelectY2 = SelectY1
        Else
        End If
    End Select
End Sub

Private Sub picGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Radius As Single
    
    On Error Resume Next

    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    Select Case MyProp.SetActiveTool
    Case Is = ajb_TArrow
    Case Is = ajb_TCircular
        Radius = MathLib.Radius(X, X)
        
        shpCircle.Width = Radius
        shpCircle.Height = Radius
    
        shpCircle.Move pox + (picGraph.ScaleWidth - shpCircle.Width) / 2, _
                       poy + (picGraph.ScaleHeight - shpCircle.Height) / 2
        
        linCrossH(1).X1 = xPos + Radius / 2
        linCrossH(1).Y1 = 0
        linCrossH(1).X2 = xPos + Radius / 2
        linCrossH(1).Y2 = picRulerH.ScaleHeight
    
        linCrossH(2).X1 = xPos - Radius / 2 + 2
        linCrossH(2).Y1 = 0
        linCrossH(2).X2 = xPos - Radius / 2 + 2
        linCrossH(2).Y2 = picRulerH.ScaleHeight
    
        linCrossV(1).X1 = 0
        linCrossV(1).Y1 = yPos + Radius / 2
        linCrossV(1).X2 = picRulerV.ScaleWidth
        linCrossV(1).Y2 = yPos + Radius / 2
            
        linCrossV(2).X1 = 0
        linCrossV(2).Y1 = yPos - Radius / 2 + 2
        linCrossV(2).X2 = picRulerV.ScaleWidth
        linCrossV(2).Y2 = yPos - Radius / 2 + 2
    Case Is = ajb_TCross
        linCrossH(0).X1 = 0
        linCrossH(0).Y1 = Y
        linCrossH(0).X2 = picGraph.ScaleWidth
        linCrossH(0).Y2 = Y
        
        linCrossV(0).X1 = X
        linCrossV(0).Y1 = 0
        linCrossV(0).X2 = X
        linCrossV(0).Y2 = picGraph.ScaleHeight
        
        linCrossH(1).X1 = X + 1
        linCrossH(1).Y1 = 0
        linCrossH(1).X2 = X + 1
        linCrossH(1).Y2 = picRulerH.ScaleHeight
        
        linCrossV(1).X1 = 0
        linCrossV(1).Y1 = Y + 1
        linCrossV(1).X2 = picRulerV.ScaleHeight
        linCrossV(1).Y2 = Y + 1
    Case Is = ajb_TPan
        If Button And vbLeftButton Then
            Dim NewPos As POINTAPI
        
            NewPos.X = picGraph.Left - (DAreaMousePos.X - X)
            NewPos.Y = picGraph.Top - (DAreaMousePos.Y - Y)
        
            OffsetDrawingAreaX = NewPos.X
            OffsetDrawingAreaY = NewPos.Y
            Exit Sub
        End If
    Case Is = ajb_TSelect
        If Button And vbLeftButton Then
            Dim px As Single
            Dim py As Single
            
            If X < 0 Then
                px = 0
            ElseIf X > picGraph.ScaleWidth Then
                px = picGraph.ScaleWidth - 1
            Else
                px = X
            End If
            
            If Y < 0 Then
                py = 0
            ElseIf Y > picGraph.ScaleHeight Then
                py = picGraph.ScaleHeight - 1
            Else
                py = Y
            End If
            
            picGraph.Line (SelectX1, SelectY1)-(SelectX2, SelectY2), _
                          &H0 Xor picGraph.BackColor, B
            picGraph.Line (SelectX1, SelectY1)-(px, py), _
                          &H0 Xor picGraph.BackColor, B
            
            SelectX2 = px
            SelectY2 = py
        End If
    End Select
    
    X = (xPos - X) / PolarScale
    Y = (yPos - Y) / PolarScale
    
    If MyProp.SetActiveTool = ajb_TCircular Then
        Radius = Radius * Polar.Unit / PolarScale / 2
    Else
        Radius = MathLib.Radius(X, Y) * Polar.Unit
    End If
    
    Dim Angle    As Single
    Dim Area     As Single
    Dim Circum   As Single
    Dim Crest    As Single
    Dim Diameter As Single
     
    ' tan Ø = (Y/X)
    Angle = MathLib.Degrees(Atn(Y / IIf(X = 0, 0.00000001, X)))
    
    If (X < 0) And (Y >= 0) Then
        Angle = Abs(Angle)
    ElseIf (X >= 0) And (Y >= 0) Then
        Angle = 180 - Angle
    ElseIf (X >= 0) And (Y < 0) Then
        Angle = 180 - Angle
    ElseIf (X < 0) And (Y < 0) Then
        Angle = 360 - Angle
    End If
   
    Area = MathLib.Area(Radius)
    Circum = MathLib.Circumference(Radius)
    Diameter = MathLib.Diameter(Radius)

    RaiseEvent Location(-X, Y, Angle, Area, Circum, Diameter, Radius)
End Sub

Private Sub picGraph_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
    Select Case MyProp.SetActiveTool
    Case Is = ajb_TArrow
    Case Is = ajb_TCircular
    Case Is = ajb_TCross
    Case Is = ajb_TPan
    Case Is = ajb_TSelect
        If Button And vbLeftButton Then
            picGraph.Line (SelectX1, SelectY1)-(SelectX2, SelectY2), _
                          &H0 Xor picGraph.BackColor, B
            
            If SelectX2 - SelectX1 < 0 Then
                shpSelect.Left = SelectX2
                shpSelect.Width = SelectX1 - SelectX2 + 1
            Else
                shpSelect.Left = SelectX1
                shpSelect.Width = SelectX2 - shpSelect.Left + 1
            End If
        
            If SelectY2 - SelectY1 < 0 Then
                shpSelect.Top = SelectY2
                shpSelect.Height = SelectY1 - SelectY2 + 1
            Else
                shpSelect.Top = SelectY1
                shpSelect.Height = SelectY2 - shpSelect.Top + 1
            End If
            
            picGraph.DrawMode = vbCopyPen
            picGraph.DrawStyle = vbSolid
            If Not shpSelect.Visible Then shpSelect.Visible = True
        End If
    End Select
End Sub

Private Sub picGraph_Resize()
    On Error Resume Next

    picRulerH.Left = picGraph.Left - picRulerJ.Width
    picRulerV.Top = picGraph.Top - picRulerJ.Height
    picShadow.Move picGraph.Left + OffsetShadowX, picGraph.Top + OffsetShadowY, _
                   picGraph.Width, picGraph.Height
    
    If MyProp.ScrollBars = ajb_SBBoth Then
        picScroll.Move UserControl.ScaleWidth - vsScroll.Width, _
                       UserControl.ScaleHeight - hsScroll.Height, _
                       hsScroll.Height, vsScroll.Width
    End If
    
    If (MyProp.ScrollBars = ajb_SBHorizontal) Or (MyProp.ScrollBars = ajb_SBBoth) Then
        hsScroll.Move 0, UserControl.ScaleHeight - hsScroll.Height, _
                      UserControl.ScaleWidth - IIf(MyProp.ScrollBars = ajb_SBHorizontal, 0, picScroll.Width), _
                      hsScroll.Height
        hsScroll.Enabled = UserControl.ScaleWidth < picGraph.Width + IIf(MyProp.ScrollBars = ajb_SBHorizontal, 0, vsScroll.Width) + _
                           OffsetDrawingAreaX * 2 + OffsetShadowX
        
        If hsScroll.Enabled Then
            hsScroll.Max = picGraph.Width - UserControl.ScaleWidth + IIf(MyProp.ScrollBars = ajb_SBHorizontal, 0, vsScroll.Width) + _
                           OffsetDrawingAreaX * 2 + OffsetShadowX
            hsScroll.SmallChange = IIf(hsScroll.Max * 0.1 < 1, 1, hsScroll.Max * 0.1)
            hsScroll.LargeChange = IIf(hsScroll.Max * 0.2 < 1, 1, hsScroll.Max * 0.2)
        End If
    End If
    
    If (MyProp.ScrollBars = ajb_SBVertical) Or (MyProp.ScrollBars = ajb_SBBoth) Then
        vsScroll.Move UserControl.ScaleWidth - vsScroll.Width, 0, _
                      vsScroll.Width, UserControl.ScaleHeight - IIf(MyProp.ScrollBars = ajb_SBVertical, _
                      0, picScroll.Height)
        vsScroll.Enabled = UserControl.ScaleHeight < picGraph.Height + IIf(MyProp.ScrollBars = ajb_SBVertical, 0, hsScroll.Height) + _
                           OffsetDrawingAreaY * 2 + OffsetShadowY
                           
        If vsScroll.Enabled Then
            vsScroll.Max = picGraph.Height - UserControl.ScaleHeight + IIf(MyProp.ScrollBars = ajb_SBVertical, 0, hsScroll.Height) + _
                           OffsetDrawingAreaY * 2 + OffsetShadowY
            vsScroll.SmallChange = IIf(vsScroll.Max * 0.1 < 1, 1, vsScroll.Max * 0.1)
            vsScroll.LargeChange = IIf(vsScroll.Max * 0.2 < 1, 1, vsScroll.Max * 0.2)
        End If
    End If
End Sub

Private Sub hsScroll_Change()
    picGraph.Left = -hsScroll.Value + OffsetDrawingAreaX
    picShadow.Left = picGraph.Left + OffsetShadowX
    picRulerH.Left = picGraph.Left - picRulerJ.Width
End Sub

Private Sub vsScroll_Change()
    picGraph.Top = -vsScroll.Value + OffsetDrawingAreaY
    picShadow.Top = picGraph.Top + OffsetShadowY
    picRulerV.Top = picGraph.Top - picRulerJ.Height
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    Call picGraph_Resize
End Sub

Private Sub UserControl_Show()
    picGraph.Visible = False
            
    scEval.Language = "VBScript"
    scEval.Timeout = NoTimeout
    scEval.AddObject "MathLib", MathLib, True
    ShowRuler = UserControl.Extender.ShowRuler
    ScrollBars = UserControl.Extender.ScrollBars
    picGraph.Move OffsetDrawingAreaX, OffsetDrawingAreaY, _
                  MAX_DAREA_WIDTH * (MyProp.Zoom * 0.01), _
                  MAX_DAREA_HEIGHT * (MyProp.Zoom * 0.01)
    picShadow.Move OffsetDrawingAreaX + OffsetShadowX, _
                   OffsetDrawingAreaY + OffsetShadowY
    Call RedrawPolar(2)
    
    Set picRulerJ.MouseIcon = LoadResPicture(101, vbResCursor)
    Set picRulerH.MouseIcon = LoadResPicture(105, vbResCursor)
    Set picRulerV.MouseIcon = LoadResPicture(105, vbResCursor)
    
    picGraph.Visible = True
End Sub

Private Sub RedrawPolar(Op As Integer)
    If Not MyProp.AutoUpdate Then Exit Sub
    
    Dim xmid      As Single
    Dim ymid      As Single
    Dim PBuffer   As New StdPicture
    Dim ValHolder As Integer
    
    picGraph.Cls
    picGraph.BackColor = MyProp.BackColor
    Set picGraph.Picture = Nothing
    Set PBuffer = MyProp.PictureHolder
    
    If PBuffer Then
        Dim pw As Integer
        Dim ph As Integer
        
        pw = ScaleX(PBuffer.Width, vbHimetric, vbPixels)
        ph = ScaleY(PBuffer.Height, vbHimetric, vbPixels)
            
        If MyProp.PictureStyle = ajb_PSNone Then
            picGraph.PaintPicture PBuffer, 0, 0, pw, ph, 0, 0, pw, ph, vbSrcCopy
        ElseIf MyProp.PictureStyle = ajb_PSCenter Then
            picGraph.PaintPicture PBuffer, (picGraph.ScaleWidth - pw) / 2, (picGraph.ScaleHeight - ph) / 2, _
                                           pw, ph, 0, 0, pw, ph, vbSrcCopy
        ElseIf MyProp.PictureStyle = ajb_PSStretch Then
            picGraph.PaintPicture PBuffer, 0, 0, picGraph.ScaleWidth, picGraph.ScaleHeight, _
                                           0, 0, pw, ph, vbSrcCopy
        Else
            Dim cx As Integer
            Dim cy As Integer
            
            For cy = 0 To picGraph.ScaleWidth / pw + 1
                For cx = 0 To picGraph.ScaleHeight / ph + 1
                    picGraph.PaintPicture PBuffer, cx * pw, cy * ph, pw, ph, _
                                                   0, 0, pw, ph, vbSrcCopy
                Next cx
            Next cy
        End If
    End If
    
    xmid = picGraph.ScaleWidth / 2
    ymid = picGraph.ScaleHeight / 2
        
    If Polar.Origin.Auto Then
        xPos = xmid
        yPos = ymid
            
        Polar.Origin.X = 0
        Polar.Origin.Y = 0
        
        pox = 0
        poy = 0
    Else
        pox = Polar.Origin.X * (MyProp.Zoom * 0.01)
        poy = -Polar.Origin.Y * (MyProp.Zoom * 0.01)
        
        xPos = xmid + pox
        yPos = ymid + poy
            
        pox = IIf(Polar.Origin.Auto, 0, pox)
        poy = IIf(Polar.Origin.Auto, 0, poy)
    End If
        
    PolarScale = IIf(xmid > ymid, ymid, xmid) / _
                 Polar.Radius / MyProp.AspectRatio
    
    If MyProp.ShowCoordinateSystem Then
        Dim i     As Integer
        Dim px    As Single
        Dim py    As Single
        Dim Deg   As Single
        Dim sTemp As String
        
        picGraph.FillStyle = vbFSSolid
        picGraph.FillColor = Polar.FillColor
        picGraph.ForeColor = Polar.FillColor
        picGraph.Circle (xPos, yPos), PolarScale * Polar.Radius
        picGraph.FillStyle = vbFSTransparent
        
        If MyProp.ShowGrid <> ajb_GCNone Then
            For i = 1 To Polar.Radius
                picGraph.Circle (xPos, yPos), PolarScale * i, Polar.GridColor
            Next i
        End If
        
        picGraph.ForeColor = Polar.LabelColor
             
        If MyProp.ShowGrid = ajb_GCNone Then
            ValHolder = 1
        ElseIf MyProp.ShowGrid = ajb_GCMinor Then
            ValHolder = 2
        Else
            ValHolder = 1
        End If
        
        For i = 0 To 24 Step ValHolder
            px = Polar.Radius * PolarScale * Cos(MathLib.Radians(Deg))
            py = Polar.Radius * PolarScale * Sin(MathLib.Radians(Deg))
            
            If MyProp.ShowLabel Then
                sTemp = CStr(Deg) & Chr(&HB0)
                picGraph.CurrentX = pox + (picGraph.ScaleWidth - picGraph.TextWidth(sTemp)) / 2 + px + _
                                    picGraph.TextWidth(sTemp) * 0.75 * Cos(MathLib.Radians(Deg))
                picGraph.CurrentY = poy + (picGraph.ScaleHeight - picGraph.TextHeight(sTemp)) / 2 - py - _
                                    picGraph.TextHeight(sTemp) * 0.75 * Sin(MathLib.Radians(Deg))
                
                picGraph.Print sTemp
            End If
            
            If MyProp.ShowGrid Then
                picGraph.Line (xPos, yPos)-(xPos + px, yPos + py), Polar.GridColor
            End If
            
            Deg = i * 15
        Next i
        
        picGraph.ForeColor = vbBlack
        
        If MyProp.ShowGrid Then
            picGraph.Line (xPos, yPos - Polar.Radius * PolarScale)- _
                          (xPos, yPos + Polar.Radius * PolarScale), &H0
            picGraph.Line (xPos - Polar.Radius * PolarScale, yPos)- _
                          (xPos + Polar.Radius * PolarScale, yPos), &H0
        End If
    End If
    
    Set picGraph.Picture = picGraph.Image
    
    If Op = 0 Then
    ElseIf Op = 1 Then
        Call DrawGraph
    ElseIf Op = 2 Then
        Call RedrawRuler
    Else
        Call DrawGraph
        Call RedrawRuler
    End If
End Sub

Private Sub PlotGraph(Graph As GrapInfo)
    Dim dx       As Single
    Dim dy       As Single
    Dim px       As Single
    Dim py       As Single
    Dim Deg      As Single
    Dim Radius   As Single
    Dim OldTimer As Single
    Dim OldFC    As Long
    Dim OldDS    As Integer
    Dim OldDW    As Integer
    Dim OldFnt   As StdFont
    Dim OldFT    As Boolean
    Dim Temp     As String
    Dim olddx    As Single
    Dim olddy    As Single
    Dim Sprite   As New StdPicture
    Dim Mask     As New StdPicture
                    
    On Error Resume Next
    
    MyProp.StopAni = False
    Set Graph = GraphHolder
    RaiseEvent AniStat(MyProp.StopAni)
    
    With Graph.Series
        If Graph.Delay > 0 Then
            lblTrackAngle.Caption = "Angle : 0°"
            lblTrackAngle.Move 10, picGraph.ScaleHeight - lblTrackAngle.Height - 10
            lblTrackAngle.Visible = True
        End If

        OldFC = picGraph.ForeColor
        
        If .Marker.UsePicture Then
            If Dir$(.Marker.PicturePath) <> "" Then
                Set Sprite = LoadPicture(.Marker.PicturePath)
            End If
            
            If Dir$(.Marker.MaskPicturePath) <> "" Then
                Set Mask = LoadPicture(.Marker.MaskPicturePath)
            End If
        End If
        
        If .AllowPen Then
            OldDS = picGraph.DrawStyle
            OldDW = picGraph.DrawWidth
            picGraph.ForeColor = .Pen.FillColor
            picGraph.DrawStyle = .Pen.Style
            picGraph.DrawWidth = .Pen.Weight
        End If
        
        If .AllowMarker Then
            Set OldFnt = picGraph.Font
            OldFT = picGraph.FontTransparent
            Set picGraph.Font = .Marker.Font
            picGraph.ForeColor = .Marker.FillColor
            picGraph.FontTransparent = .Marker.Transparent
        End If
        
        If Trim$(Graph.Coefficient) <> "" Then
            Dim l As Integer
            
            l = Len(Graph.Coefficient)
            Temp = ":" & IIf(Mid$(Graph.Coefficient, l, 1) = ":", _
                         Left$(Graph.Coefficient, l - 1), Graph.Coefficient)
        End If
        
        scEval.AddCode Temp
        For Deg = Graph.StartingAngle To Graph.EndingAngle Step Graph.Step
            scEval.AddCode "t=" & MathLib.Radians((Deg))
            Radius = scEval.Eval(Graph.Equation)
           
            px = Radius * Cos(MathLib.Radians(Deg))
            py = -(Radius * Sin(MathLib.Radians(Deg)))
        
            dx = xPos + PolarScale * px / Polar.Unit
            dy = yPos + PolarScale * py / Polar.Unit
            
            If .AllowPen Then
                picGraph.ForeColor = .Pen.FillColor
                
                If .Pen.AllowConVertices Then
                    If Deg = Graph.StartingAngle Then
                        olddx = dx
                        olddy = dy
                        picGraph.CurrentX = olddx
                        picGraph.CurrentY = olddy
                    End If
                
                    picGraph.Line (olddx, olddy)-(dx, dy)
                End If
            
                If .Pen.AllowShading Then
                    picGraph.Line (xPos, yPos)-(dx, dy)
                Else
                    picGraph.PSet (dx, dy)
                End If
            
                olddx = dx
                olddy = dy
            End If
            
            If .AllowMarker Then
                Dim varW As Integer
                Dim varH As Integer
                Dim curX As Integer
                Dim curY As Integer
                    
                picGraph.ForeColor = .Marker.FillColor
                
                If .Marker.UsePicture Then
                    If Sprite Then
                        If .Marker.AutoSize Then
                            varW = ScaleX(Sprite.Width, vbHimetric, vbPixels)
                            varH = ScaleY(Sprite.Height, vbHimetric, vbPixels)
                        Else
                            If .Marker.PictureWidth <= 0 Then
                                varW = ScaleX(Sprite.Width, vbHimetric, vbPixels)
                            Else
                                varW = .Marker.PictureWidth
                            End If
                                
                            If .Marker.PictureHeight <= 0 Then
                                varH = ScaleY(Sprite.Height, vbHimetric, vbPixels)
                            Else
                                varH = .Marker.PictureHeight
                            End If
                        End If
                    Else
                        varW = picGraph.TextWidth(Graph.Series.Marker.Style)
                        varH = picGraph.TextHeight(Graph.Series.Marker.Style)
                    End If
                Else
                    varW = picGraph.TextWidth(Graph.Series.Marker.Style)
                    varH = picGraph.TextHeight(Graph.Series.Marker.Style)
                End If
                    
                Select Case .Marker.Alignment
                Case Is = 0  ' (0 #)  ajb_HORIZ_LEFT
                    curX = dx: curY = dy
                    .Marker.GetHorizAlign = ajb_HORIZ_LEFT
                    .Marker.GetVertAlign = ajb_VERT_TOP
                Case Is = 1 '  (1 #)  ajb_HORIZ_RIGHT
                    curX = dx - varW: curY = dy
                    .Marker.GetHorizAlign = ajb_HORIZ_RIGHT
                    .Marker.GetVertAlign = ajb_VERT_TOP
                Case Is = 2 '  (2 #)  ajb_HORIZ_CENTER
                    curX = dx - varW / 2: curY = dy
                    .Marker.GetHorizAlign = ajb_HORIZ_CENTER
                    .Marker.GetVertAlign = ajb_VERT_TOP
                Case Is = 4  ' (0 4)  ajb_HORIZ_LEFT   OR ajb_VERT_TOP
                    curX = dx: curY = dy
                    .Marker.GetHorizAlign = ajb_HORIZ_LEFT
                    .Marker.GetVertAlign = ajb_VERT_TOP
                Case Is = 5  ' (1 4)  ajb_HORIZ_RIGHT  OR ajb_VERT_TOP
                    curX = dx - varW: curY = dy
                    .Marker.GetHorizAlign = ajb_HORIZ_RIGHT
                    .Marker.GetVertAlign = ajb_VERT_TOP
                Case Is = 6  ' (2 4)  ajb_HORIZ_CENTER OR ajb_VERT_TOP
                    curX = dx - varW / 2: curY = dy
                    .Marker.GetHorizAlign = ajb_HORIZ_CENTER
                    .Marker.GetVertAlign = ajb_VERT_TOP
                Case Is = 8  ' (0 8)  ajb_HORIZ_LEFT   OR ajb_VERT_BOTTOM
                    curX = dx: curY = dy - varH
                    .Marker.GetHorizAlign = ajb_HORIZ_LEFT
                    .Marker.GetVertAlign = ajb_VERT_BOTTOM
                Case Is = 9  ' (1 8)  ajb_HORIZ_RIGHT  OR ajb_VERT_BOTTOM
                    curX = dx - varW: curY = dy - varH
                    .Marker.GetHorizAlign = ajb_HORIZ_RIGHT
                    .Marker.GetVertAlign = ajb_VERT_BOTTOM
                Case Is = 10 ' (2 8)  ajb_HORIZ_CENTER OR ajb_VERT_BOTTOM
                    curX = dx - varW / 2: curY = dy - varH
                    .Marker.GetHorizAlign = ajb_HORIZ_CENTER
                    .Marker.GetVertAlign = ajb_VERT_BOTTOM
                Case Is = 16 ' (0 16) ajb_HORIZ_LEFT   OR ajb_VERT_CENTER
                    curX = dx: curY = dy - varH / 2
                    .Marker.GetHorizAlign = ajb_HORIZ_LEFT
                    .Marker.GetVertAlign = ajb_VERT_CENTER
                Case Is = 17 ' (1 16) ajb_HORIZ_RIGHT  OR ajb_VERT_CENTER
                    curX = dx - varW: curY = dy - varH / 2
                    .Marker.GetHorizAlign = ajb_HORIZ_RIGHT
                    .Marker.GetVertAlign = ajb_VERT_CENTER
                Case Is = 18 ' (2 16) ajb_HORIZ_CENTER OR ajb_VERT_CENTER
                    curX = dx - varW / 2: curY = dy - varH / 2
                    .Marker.GetHorizAlign = ajb_HORIZ_CENTER
                    .Marker.GetVertAlign = ajb_VERT_CENTER
                Case Else
                    curX = dx: curY = dy
                    .Marker.GetHorizAlign = ajb_HORIZ_LEFT
                    .Marker.GetVertAlign = ajb_VERT_TOP
                End Select
    
                If .Marker.UsePicture Then
                    If Sprite Then
                        If .Marker.AutoSize Then
                            If Mask Then
                                picGraph.PaintPicture Mask, curX, curY, , , , , , , vbSrcAnd
                                picGraph.PaintPicture Sprite, curX, curY, , , , , , , vbSrcInvert
                            Else
                                picGraph.PaintPicture Sprite, curX, curY
                            End If
                        Else
                            .Marker.PictureWidth = varW
                            .Marker.PictureHeight = varH
                                
                            If Mask Then
                                picGraph.PaintPicture Mask, curX, curY, _
                                                      .Marker.PictureWidth, .Marker.PictureHeight, , , , , vbSrcAnd
                                picGraph.PaintPicture Sprite, curX, curY, _
                                                      .Marker.PictureWidth, .Marker.PictureHeight, , , , , vbSrcInvert
                            Else
                                picGraph.PaintPicture Sprite, curX, curY, _
                                                      .Marker.PictureWidth, .Marker.PictureHeight
                            End If
                        End If
                    End If
                End If
                    
                If (Not .Marker.UsePicture) Or (Sprite = 0) Then
                    picGraph.CurrentX = curX
                    picGraph.CurrentY = curY
                    picGraph.Print Graph.Series.Marker.Style
                End If
            End If
            
            RaiseEvent Status(Graph.Coefficient, Graph.Equation, 1, Deg, Radius)
            OldTimer = Timer
                
            Do While (Timer < OldTimer + MathLib.Milliseconds(Graph.Delay)) And (Not MyProp.StopAni)
                DoEvents
            Loop
            
            If Graph.Delay > 0 Then
                lblTrackAngle.Caption = "Angle : " & Deg & "°"
            End If
        Next Deg
        
        If .AllowPen Then
            picGraph.DrawStyle = OldDS
            picGraph.DrawWidth = OldDW
        End If
        
        If .AllowMarker Then
            Set picGraph.Font = OldFnt
            picGraph.FontTransparent = OldFT
        End If

        picGraph.ForeColor = OldFC
        
        If Graph.Delay > 0 Then
            If lblTrackAngle.Visible Then lblTrackAngle.Visible = False
        End If
        
        If Not MyProp.StopAni Then MyProp.StopAni = True
        RaiseEvent AniStat(MyProp.StopAni)
    End With
End Sub

Private Sub RedrawRuler()
    Dim i           As Single
    Dim cntr        As Integer
    Dim fw          As Single
    Dim fz          As Single
    Dim PolarHeight As Single
    
    picRulerJ.Cls
    picRulerH.Cls
    picRulerV.Cls
    picHolderRH.Cls
    picHolderRV.Cls
    
    For i = -Polar.Radius To Polar.Radius
        If MyProp.SetRulerMode = ajb_RUnit Then
            fw = picRulerJ.TextWidth(CStr(i * Polar.Unit)) * 0.8
        Else
            fw = picRulerJ.TextWidth(CStr(i)) * 0.8
        End If
        
        If fz < fw Then
            fz = fw
        End If
    Next i
    
    fz = fz + picRulerJ.TextWidth("W") + 4
    PolarHeight = PolarScale * Polar.Radius
    
    picRulerJ.Move 0, 0, fz, fz
    picHolderRH.Move picRulerJ.Width, 0, UserControl.Width - _
                     hsScroll.Width - picRulerJ.Width, picRulerJ.Height
    picHolderRV.Move 0, picRulerJ.Height, picRulerJ.Width, _
                     UserControl.Height - vsScroll.Height - picRulerJ.Width
    picRulerH.Move picGraph.Left - picRulerJ.Width, 4, picGraph.Width, picHolderRH.ScaleHeight - 8
    picRulerV.Move 4, picGraph.Top - picRulerJ.Height, picHolderRV.ScaleWidth - 8, picGraph.Height
    
    picRulerJ.Line (5, 5)-(picRulerJ.ScaleWidth - 5, picRulerJ.ScaleHeight - 5), &HFFFFFF, B
    picRulerJ.Line (4, 4)-(picRulerJ.ScaleWidth - 5, picRulerJ.ScaleWidth - 5), &H808080, B
    picRulerH.Line (0, 0)-(xPos - PolarHeight - picRulerH.TextWidth("W") / 2, picRulerH.ScaleWidth), &H8000000F, BF
    picRulerH.Line (xPos + PolarHeight + picRulerH.TextWidth("W") / 2, 0)-(picRulerH.ScaleWidth, picRulerH.ScaleHeight), &H8000000F, BF
    picRulerH.Line (0, 0)-(xPos - PolarHeight - 10, picRulerH.ScaleWidth), &H99A8AC, BF
    picRulerH.Line (xPos + PolarHeight + 10, 0)-(picRulerH.ScaleWidth, picRulerH.ScaleHeight), &H99A8AC, BF
    picRulerV.Line (0, 0)-(picRulerV.ScaleWidth, yPos - PolarHeight - picRulerV.TextHeight("H") / 2), &H8000000F, BF
    picRulerV.Line (0, yPos + PolarHeight + picRulerV.TextHeight("H") / 2)-(picRulerV.ScaleWidth, picRulerV.ScaleHeight), &H8000000F, BF
    picRulerV.Line (0, 0)-(picRulerV.ScaleWidth, yPos - PolarHeight - 10), &H99A8AC, BF
    picRulerV.Line (0, yPos + PolarHeight + 10)-(picRulerV.ScaleWidth, picRulerV.ScaleHeight), &H99A8AC, BF
    picHolderRH.Line (-1, -1)-(picHolderRH.ScaleWidth - 1, picHolderRH.ScaleHeight - 1), &H808080, B
    picHolderRV.Line (-1, -1)-(picHolderRV.ScaleWidth - 1, picHolderRV.ScaleHeight - 1), &H808080, B
    
    Dim j_new_font As Long
    Dim h_new_font As Long
    Dim v_new_font As Long
    Dim j_old_font As Long
    Dim h_old_font As Long
    Dim v_old_font As Long
    Dim j_tSZ      As Size
    Dim h_tSZ      As Size
    Dim v_tSZ      As Size
    Dim j_sTemp    As String
    Dim h_sTemp    As String
    Dim v_sTemp    As String
    
    j_new_font = CustomFont(14, 0, 0, 0, _
                            0, False, False, False, _
                            "Microsoft Sans Serif")
    j_old_font = SelectObject(picRulerJ.hdc, j_new_font)
    
    h_new_font = CustomFont(9, 0, 900, 0, _
                            0, False, False, False, _
                            "Microsoft Sans Serif")
    h_old_font = SelectObject(picRulerH.hdc, h_new_font)
    
    v_new_font = CustomFont(10, 0, 0, 0, _
                            0, False, False, False, _
                            "Microsoft Sans Serif")
    v_old_font = SelectObject(picRulerV.hdc, v_new_font)
    
    j_sTemp = IIf(MyProp.SetRulerMode = ajb_RUnit, "U", "A")
    picRulerJ.ForeColor = &HFFFFFF
    picRulerJ.CurrentX = (picRulerJ.ScaleWidth - picRulerJ.TextWidth(j_sTemp)) / 2 + 1
    picRulerJ.CurrentY = (picRulerJ.ScaleHeight - picRulerJ.TextHeight(j_sTemp)) / 2 + 1
    picRulerJ.Print j_sTemp
    picRulerJ.ForeColor = &H0
    picRulerJ.CurrentX = (picRulerJ.ScaleWidth - picRulerJ.TextWidth(j_sTemp)) / 2
    picRulerJ.CurrentY = (picRulerJ.ScaleHeight - picRulerJ.TextHeight(j_sTemp)) / 2
    picRulerJ.Print j_sTemp
    
    cntr = 0
    For i = -Polar.Radius To Polar.Radius Step 0.5
        If cntr Mod 2 = 0 Then
            If MyProp.SetRulerMode = ajb_RUnit Then
                h_sTemp = "-" & CStr(Abs(CSng(i)) * Polar.Unit)
            Else
                h_sTemp = "-" & CStr(Abs(CSng(i)))
            End If
            
            GetTextExtentPoint picRulerH.hdc, h_sTemp, Len(h_sTemp), h_tSZ
            picRulerH.CurrentX = xPos - PolarScale * i - h_tSZ.cy / 2 + 1
            picRulerH.CurrentY = picRulerH.ScaleHeight - 5
            picRulerH.Print IIf(MyProp.SetRulerMode = ajb_RUnit, _
                            CSng(i) * Polar.Unit, CSng(i))
            picRulerH.Line (xPos - PolarScale * i + 1, picRulerH.ScaleHeight)- _
                           (xPos - PolarScale * i + 1, picRulerH.ScaleHeight - 5), &H99A8AC
    
            If MyProp.SetRulerMode = ajb_RUnit Then
                v_sTemp = "-" & CStr(Abs(CSng(i) * Polar.Unit))
            Else
                v_sTemp = "-" & CStr(Abs(CSng(i)))
            End If
            
            GetTextExtentPoint picRulerV.hdc, v_sTemp, Len(v_sTemp), v_tSZ
            picRulerV.CurrentX = picRulerV.ScaleWidth - v_tSZ.cx - 5
            picRulerV.CurrentY = yPos - PolarScale * i - v_tSZ.cy / 2 + 1
            picRulerV.Print IIf(MyProp.SetRulerMode = ajb_RUnit, _
                            CSng(i) * Polar.Unit, CSng(i))
            picRulerV.Line (picRulerV.ScaleWidth, yPos - PolarScale * i + 1)- _
                           (picRulerV.ScaleWidth - 5, yPos - PolarScale * i + 1), &H99A8AC
        Else
            picRulerH.Line (xPos - PolarScale * i, picRulerH.ScaleHeight)- _
                           (xPos - PolarScale * i, picRulerH.ScaleHeight - 3), &H99A8AC
            picRulerV.Line (picRulerV.ScaleWidth, yPos - PolarScale * i)- _
                           (picRulerV.ScaleWidth - 3, yPos - PolarScale * i), &H99A8AC
        End If
        cntr = cntr + 1
    Next i
    
    SelectObject picRulerJ.hdc, j_old_font
    SelectObject picRulerH.hdc, h_old_font
    SelectObject picRulerV.hdc, v_old_font
    
    DeleteObject j_new_font
    DeleteObject h_new_font
    DeleteObject v_new_font
End Sub

Private Sub CopyToBuffer()
    Dim Temp   As Picture
        
    picBuffer.Move 0, 0, picGraph.Width, picGraph.Height
    BitBlt picBuffer.hdc, 0, 0, picGraph.ScaleWidth, picGraph.ScaleHeight, _
           picGraph.hdc, 0, 0, vbSrcCopy
    picBuffer.Refresh
        
    Set Temp = picBuffer.Image
        
    picBuffer.Cls
    Set picBuffer.Picture = Nothing
    picBuffer.Move 0, 0, shpSelect.Width, shpSelect.Height
    picBuffer.PaintPicture Temp, 0, 0, _
                           shpSelect.Width, shpSelect.Height, _
                           shpSelect.Left, shpSelect.Top, _
                           shpSelect.Width, shpSelect.Height
                           
    Set picBuffer.Picture = picBuffer.Image
End Sub

Private Sub EmptyBuffer()
    picBuffer.Cls
    Set picBuffer.Picture = Nothing
    picBuffer.Move 0, 0, 0, 0
End Sub

Public Function ActiveSelection() As Boolean
    If (shpSelect.Width = 1) And (shpSelect.Height = 1) Then
        ActiveSelection = False
    Else
        ActiveSelection = shpSelect.Visible
    End If
End Function

Public Sub ClearGraph()
    picGraph.Cls
End Sub

Public Sub DrawGraph(Optional vntIndexKey As Variant = 0, Optional TurnAniOn As Boolean = False)
    Dim OldDelay As Long
    
    If (vntIndexKey = 0) Or (Trim$(vntIndexKey) = "") Then
        Dim i As Integer
        
        picGraph.Cls
        For i = 1 To Plot.Count
            If Plot(i).Visible Then
                RaiseEvent Graph(i)
                Set GraphHolder = Plot(i)
                If Not TurnAniOn Then
                    OldDelay = GraphHolder.Delay
                    GraphHolder.Delay = 0
                End If
                Call PlotGraph(GraphHolder)
                If Not TurnAniOn Then GraphHolder.Delay = OldDelay
            End If
        Next i
    Else
        If Plot(vntIndexKey).Visible Then
            RaiseEvent Graph(vntIndexKey)
            Set GraphHolder = Plot(vntIndexKey)
            If Not TurnAniOn Then
                OldDelay = GraphHolder.Delay
                GraphHolder.Delay = 0
            End If
            Call PlotGraph(GraphHolder)
            If Not TurnAniOn Then GraphHolder.Delay = OldDelay
        End If
    End If
    
    Set GraphHolder = Nothing
End Sub

Public Sub EditCopy()
    If shpSelect.Visible Then
        Call CopyToBuffer
        
        Clipboard.Clear
        Clipboard.SetData picBuffer.Picture
        
        Call EmptyBuffer
    End If
End Sub

Public Sub EditCopyTo()
    On Error GoTo ErrHandler
    
    If shpSelect.Visible Then
        Call CopyToBuffer
        
        With dlgPolar
            .Filter = "Bitmap Files (*.bmp) | *.bmp;"
            .Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist
            .FilterIndex = 1
            .Filename = ""
            .ShowSave
            
            If Trim$(.Filename) <> "" Then
                SavePicture picBuffer.Picture, .Filename
            End If
        End With
        
        Call EmptyBuffer
    End If
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbOKOnly Or vbCritical, "Error"
End Sub

Public Sub EditSelectAll()
    MyProp.SetActiveTool = ajb_TSelect
    If Not ActiveSelection Then shpSelect.Visible = True
    shpSelect.Move 0, 0, picGraph.ScaleWidth, picGraph.ScaleHeight
End Sub

Public Function OpenGraph(ByVal Filename As String) As Boolean
    Dim arrData() As String
    Dim Data      As String
    Dim InFile    As Long
    Dim curpos    As Integer
    
    On Error GoTo ErrHandler
    
    OpenGraph = False
    
    InFile = FreeFile
    Open Filename For Input As InFile
        Input #InFile, Data ' File signature
        
        If Data = SIGNATURE Then
            Set Plot = Nothing
            
            Do While Not EOF(InFile)
                Input #InFile, Data
                
                arrData = Split(Data, "//")
                Plot.Add arrData(0), CStr(arrData(1)), CLng(arrData(11)), CLng(arrData(16))
                curpos = Plot.Count
                
                Plot(curpos).StartingAngle = CSng(arrData(2))
                Plot(curpos).EndingAngle = CSng(arrData(3))
                Plot(curpos).Step = CSng(arrData(4))
                Plot(curpos).Delay = CInt(arrData(5))
                Plot(curpos).Visible = CBool(arrData(6))
                Plot(curpos).Series.AllowMarker = CBool(arrData(7))
                Plot(curpos).Series.AllowPen = CBool(arrData(8))
                Plot(curpos).Series.Pen.AllowConVertices = CBool(arrData(9))
                Plot(curpos).Series.Pen.AllowShading = CBool(arrData(10))
                Plot(curpos).Series.Pen.Style = CInt(arrData(12))
                Plot(curpos).Series.Pen.Weight = CInt(arrData(13))
                Plot(curpos).Series.Marker.Alignment = CInt(arrData(14))
                Plot(curpos).Series.Marker.AutoSize = CBool(arrData(15))
                Plot(curpos).Series.Marker.Font.Name = CStr(arrData(17))
                Plot(curpos).Series.Marker.Font.Size = CInt(arrData(18))
                Plot(curpos).Series.Marker.Font.Bold = CBool(arrData(19))
                Plot(curpos).Series.Marker.Font.Italic = CBool(arrData(20))
                Plot(curpos).Series.Marker.GetHorizAlign = CInt(arrData(21))
                Plot(curpos).Series.Marker.GetVertAlign = CInt(arrData(22))
                Plot(curpos).Series.Marker.MaskPicturePath = CStr(arrData(23))
                Plot(curpos).Series.Marker.PicturePath = CStr(arrData(24))
                Plot(curpos).Series.Marker.PictureHeight = CLng(arrData(25))
                Plot(curpos).Series.Marker.PictureWidth = CLng(arrData(26))
                Plot(curpos).Series.Marker.Style = CStr(arrData(27))
                Plot(curpos).Series.Marker.Transparent = CBool(arrData(28))
                Plot(curpos).Series.Marker.UsePicture = CBool(arrData(29))
                MyProp.Comment = CStr(arrData(30))
                
                Dim i           As Integer
                Dim FileNamePic As String
                Dim FilePathPic As String
                Dim ThisFile()  As Variant
                
                ThisFile() = Array(Plot(curpos).Series.Marker.PicturePath, _
                                   Plot(curpos).Series.Marker.MaskPicturePath)
                                   
                For i = LBound(ThisFile) To UBound(ThisFile)
                    FileNamePic = FSys.GetFileName(ThisFile(i))
                    FilePathPic = FSys.GetParentFolderName(Filename)
                    
                    If Mid$(FilePathPic, Len(FilePathPic), 1) = "\" Then
                    Else
                        FilePathPic = FilePathPic & "\"
                    End If
                                        
                    If FSys.FileExists(ThisFile(i)) Then
                    ElseIf FSys.FileExists(App.Path & "\" & FileNamePic) Then
                        If i = 0 Then
                            Plot(curpos).Series.Marker.PicturePath = App.Path & "\" & FileNamePic
                        Else
                            Plot(curpos).Series.Marker.MaskPicturePath = App.Path & "\" & FileNamePic
                        End If
                    ElseIf FSys.FileExists(FilePathPic & FileNamePic) Then
                        If i = 0 Then
                            Plot(curpos).Series.Marker.PicturePath = FilePathPic & FileNamePic
                        Else
                            Plot(curpos).Series.Marker.MaskPicturePath = FilePathPic & FileNamePic
                        End If
                    Else
                        If i = 0 Then
                            Plot(curpos).Series.Marker.PicturePath = ""
                        Else
                            Plot(curpos).Series.Marker.MaskPicturePath = ""
                        End If
                    End If
                Next i
            Loop
            
            OpenGraph = True
        Else
            MsgBox "File format error!", vbCritical Or vbOKOnly, "Polar Graph 1.0"
        End If
    Close #InFile
    Exit Function
    
ErrHandler:
    If InFile > 0 Then Close #InFile
    MsgBox Err.Description, vbCritical Or vbOKOnly, "Polar Graph 1.0"
End Function

Public Sub SaveGraph(ByVal Filename As String)
    Dim i      As Integer
    Dim InFile As Long
    Dim Data   As String
    
    On Error GoTo ErrHandler
    
    InFile = FreeFile
    
    Open Filename For Output As InFile
        Write #InFile, SIGNATURE
        
        For i = 1 To Plot.Count
            Data = ""
            Data = Data & Plot(i).Coefficient & "//"
            Data = Data & Plot(i).Equation & "//"
            Data = Data & Plot(i).StartingAngle & "//"
            Data = Data & Plot(i).EndingAngle & "//"
            Data = Data & Plot(i).Step & "//"
            Data = Data & Plot(i).Delay & "//"
            Data = Data & Plot(i).Visible & "//"
            Data = Data & Plot(i).Series.AllowMarker & "//"
            Data = Data & Plot(i).Series.AllowPen & "//"
            Data = Data & Plot(i).Series.Pen.AllowConVertices & "//"
            Data = Data & Plot(i).Series.Pen.AllowShading & "//"
            Data = Data & Plot(i).Series.Pen.FillColor & "//"
            Data = Data & Plot(i).Series.Pen.Style & "//"
            Data = Data & Plot(i).Series.Pen.Weight & "//"
            Data = Data & Plot(i).Series.Marker.Alignment & "//"
            Data = Data & Plot(i).Series.Marker.AutoSize & "//"
            Data = Data & Plot(i).Series.Marker.FillColor & "//"
            Data = Data & Plot(i).Series.Marker.Font.Name & "//"
            Data = Data & Plot(i).Series.Marker.Font.Size & "//"
            Data = Data & Plot(i).Series.Marker.Font.Bold & "//"
            Data = Data & Plot(i).Series.Marker.Font.Italic & "//"
            Data = Data & Plot(i).Series.Marker.GetHorizAlign & "//"
            Data = Data & Plot(i).Series.Marker.GetVertAlign & "//"
            Data = Data & Plot(i).Series.Marker.MaskPicturePath & "//"
            Data = Data & Plot(i).Series.Marker.PicturePath & "//"
            Data = Data & Plot(i).Series.Marker.PictureHeight & "//"
            Data = Data & Plot(i).Series.Marker.PictureWidth & "//"
            Data = Data & Plot(i).Series.Marker.Style & "//"
            Data = Data & Plot(i).Series.Marker.Transparent & "//"
            Data = Data & Plot(i).Series.Marker.UsePicture & "//"
            Data = Data & MyProp.Comment
            
            Write #InFile, Data
        Next i
    Close #InFile
    Exit Sub
    
ErrHandler:
    If InFile > 0 Then Close #InFile
    MsgBox Err.Description, vbCritical Or vbOKOnly, "Polar Graph 1.0"
End Sub

Public Sub Refresh()
    Call RedrawPolar(3)
    UserControl.Refresh
End Sub

Public Sub ResetDrawingArea()
    Call picRulerJ_MouseDown(vbLeftButton, 0, 0, 0)
End Sub

Public Sub SentToPrinter()
    Dim TMargin As Long ' top margin
    Dim LMargin As Long ' left margin
    On Error GoTo PrintError
    
    If Not ActiveSelection Then
        shpSelect.Move 0, 0, picGraph.ScaleWidth, picGraph.ScaleHeight
    End If
    
    Call CopyToBuffer
    LMargin = ScaleX(1, vbInches, vbTwips)
    TMargin = ScaleY(1, vbInches, vbTwips)
    Printer.ScaleMode = vbTwips
    Printer.PaintPicture picBuffer.Picture, LMargin, TMargin
    Printer.EndDoc ' start printing...
    Call EmptyBuffer
    Exit Sub
    
PrintError:
    MsgBox Err.Description, vbCritical, "Polar Graph 1.0"
End Sub

Public Sub ShowAbout()
Attribute ShowAbout.VB_Description = "Display the copyright dialog."
Attribute ShowAbout.VB_UserMemId = -552
    MsgBox "Polar Graph 1.0" & Chr(13) & "Programmed by: Aris Buenaventura" _
        & Chr(13) & "Email : ravemasterharuglory@yahoo.com", , "About"
End Sub

