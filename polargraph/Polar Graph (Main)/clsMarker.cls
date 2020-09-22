VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Marker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Enum MarkerConstants
    ajb_HORIZ_LEFT = 0
    ajb_HORIZ_RIGHT = 1
    ajb_HORIZ_CENTER = 2
    
    ajb_VERT_TOP = 4
    ajb_VERT_BOTTOM = 8
    ajb_VERT_CENTER = 16
End Enum

Private mvarAlignment        As MarkerConstants
Private mvarAutoSize         As Boolean
Private mvarFillColor        As Long
Private mvarFont             As New StdFont
Private mvarGetHorizAlign    As MarkerConstants
Private mvarGetVertAlign     As MarkerConstants
Private mvarMaskPicturePath  As String
Private mvarPicturePath      As String
Private mvarPictureWidth     As Long
Private mvarPictureHeight    As Long
Private mvarStyle            As String
Private mvarTransparent      As Boolean
Private mvarUsePicture       As Boolean

Public Property Get Alignment() As MarkerConstants
    Alignment = mvarAlignment
End Property

Public Property Let Alignment(vData As MarkerConstants)
    mvarAlignment = vData
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = mvarAutoSize
End Property

Public Property Let AutoSize(vData As Boolean)
    mvarAutoSize = vData
End Property

Public Property Get FillColor() As OLE_COLOR
    FillColor = mvarFillColor
End Property

Public Property Let FillColor(vData As OLE_COLOR)
    mvarFillColor = vData
End Property

Public Property Get Font() As StdFont
    Set Font = mvarFont
End Property

Public Property Set Font(ByVal vData As StdFont)
    Set mvarFont = vData
End Property

Public Property Get GetHorizAlign() As MarkerConstants
     GetHorizAlign = mvarGetHorizAlign
End Property

Public Property Let GetHorizAlign(ByVal vData As MarkerConstants)
    mvarGetHorizAlign = vData
End Property

Public Property Get GetVertAlign() As MarkerConstants
    GetVertAlign = mvarGetVertAlign
End Property

Public Property Let GetVertAlign(ByVal vData As MarkerConstants)
    mvarGetVertAlign = vData
End Property

Public Property Get MaskPicturePath() As String
    MaskPicturePath = mvarMaskPicturePath
End Property

Public Property Let MaskPicturePath(ByVal vData As String)
    mvarMaskPicturePath = vData
End Property

Public Property Get PicturePath() As String
    PicturePath = mvarPicturePath
End Property

Public Property Let PicturePath(ByVal vData As String)
    mvarPicturePath = vData
End Property

Public Property Get PictureWidth() As Long
    PictureWidth = mvarPictureWidth
End Property

Public Property Let PictureWidth(vData As Long)
    mvarPictureWidth = vData
End Property

Public Property Get PictureHeight() As Long
    PictureHeight = mvarPictureHeight
End Property

Public Property Let PictureHeight(vData As Long)
    mvarPictureHeight = vData
End Property

Public Property Get Style() As String
    Style = mvarStyle
End Property

Public Property Let Style(ByVal vData As String)
Attribute Style.VB_UserMemId = 0
    mvarStyle = vData
End Property

Public Property Get Transparent() As Boolean
    Transparent = mvarTransparent
End Property

Public Property Let Transparent(ByVal vData As Boolean)
    mvarTransparent = vData
End Property

Public Property Get UsePicture() As Boolean
    UsePicture = mvarUsePicture
End Property

Public Property Let UsePicture(ByVal vData As Boolean)
    mvarUsePicture = vData
End Property

Public Sub SetPictureSize(ByVal vWidth As Long, ByVal vHeight As Long)
    mvarPictureWidth = vWidth
    mvarPictureHeight = vHeight
End Sub

Private Sub Class_Initialize()
    mvarAlignment = ajb_HORIZ_CENTER Or ajb_VERT_CENTER
    mvarAutoSize = True
    mvarFillColor = &HFF0000
    mvarFont.Name = "Wingdings"
    mvarGetHorizAlign = ajb_HORIZ_CENTER
    mvarGetVertAlign = ajb_VERT_CENTER
    mvarMaskPicturePath = ""
    mvarPicturePath = ""
    mvarPictureWidth = 16
    mvarPictureHeight = 16
    mvarStyle = "A"
    mvarTransparent = True
    mvarUsePicture = False
End Sub
