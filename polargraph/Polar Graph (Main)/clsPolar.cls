VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Polar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Origin"
Option Explicit

Private mvarFillColor  As Long
Private mvarGridColor  As Long
Private mvarLabelColor As Long
Private mvarOrigin     As Origin
Private mvarRadius     As Single
Private mvarUnit       As Single

Public Property Get Area() As Single
    Area = MathLib.Area(mvarRadius * mvarUnit)
End Property

Public Property Get Circumference() As Single
    Circumference = MathLib.Circumference(mvarRadius * mvarUnit)
End Property

Public Property Get Diameter() As Single
    Diameter = MathLib.Diameter(mvarRadius * mvarUnit)
End Property

Public Property Get GridColor() As Long
    GridColor = mvarGridColor
End Property

Public Property Let GridColor(ByVal vData As Long)
    mvarGridColor = vData
End Property

Public Property Get FillColor() As Long
    FillColor = mvarFillColor
End Property

Public Property Let FillColor(ByVal vData As Long)
    mvarFillColor = vData
End Property

Public Property Get LabelColor() As Long
    LabelColor = mvarLabelColor
End Property

Public Property Let LabelColor(ByVal vData As Long)
    mvarLabelColor = vData
End Property

Public Property Let Radius(ByVal vData As Single)
    mvarRadius = vData
End Property

Public Property Get Radius() As Single
    Radius = mvarRadius
End Property

Public Property Get Origin() As Origin
    Set Origin = mvarOrigin
End Property

Public Property Set Origin(vData As Origin)
    Set mvarOrigin = vData
End Property

Public Property Get Unit() As Single
    Unit = mvarUnit
End Property

Public Property Let Unit(ByVal vData As Single)
    mvarUnit = vData
End Property

Private Sub Class_Initialize()
    Set mvarOrigin = New Origin
    
    mvarFillColor = vbYellow
    mvarGridColor = &HE0E0E0
    mvarLabelColor = &H808080
    mvarUnit = 0.5
    mvarRadius = 6
End Sub

Private Sub Class_Terminate()
    Set mvarOrigin = Nothing
End Sub
