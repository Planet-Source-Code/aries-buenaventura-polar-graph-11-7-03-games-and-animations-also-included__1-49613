VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mvarAllowConVertices As Boolean
Private mvarAllowShading     As Boolean
Private mvarFillColor        As Long
Private mvarStyle            As DrawStyleConstants
Private mvarWeight           As Integer

Public Property Get AllowConVertices() As Boolean
    AllowConVertices = mvarAllowConVertices
End Property

Public Property Let AllowConVertices(vData As Boolean)
    mvarAllowConVertices = vData
End Property

Public Property Let AllowShading(ByVal vData As Boolean)
    mvarAllowShading = vData
End Property

Public Property Get AllowShading() As Boolean
    AllowShading = mvarAllowShading
End Property

Public Property Get FillColor() As OLE_COLOR
    FillColor = mvarFillColor
End Property

Public Property Let FillColor(vData As OLE_COLOR)
    mvarFillColor = vData
End Property

Public Property Get Style() As DrawStyleConstants
    Style = mvarStyle
End Property

Public Property Let Style(ByVal vData As DrawStyleConstants)
    mvarStyle = vData
End Property

Public Property Let Weight(ByVal vData As Integer)
    mvarWeight = vData
End Property

Public Property Get Weight() As Integer
Attribute Weight.VB_UserMemId = 0
    Weight = mvarWeight
End Property

Private Sub Class_Initialize()
    mvarAllowConVertices = True
    mvarAllowShading = False
    mvarFillColor = &HFF0000
    mvarStyle = vbSolid
    mvarWeight = 1
End Sub
