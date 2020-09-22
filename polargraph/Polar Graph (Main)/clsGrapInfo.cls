VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GrapInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Series"
Option Explicit

Private mvarCoefficient   As String
Private mvarDelay         As Long
Private mvarEndingAngle   As Single
Private mvarEquation      As String
Private mvarVisible       As Boolean
Private mvarStartingAngle As Single
Private mvarStep          As Single
Private mvarSeries        As Series

Public Property Get Coefficient() As String
    Coefficient = mvarCoefficient
End Property

Public Property Let Coefficient(ByVal vData As String)
    mvarCoefficient = vData
End Property

Public Property Get Equation() As String
    Equation = mvarEquation
End Property

Public Property Let Equation(ByVal vData As String)
    mvarEquation = vData
End Property

Public Property Get Delay() As Integer
    Delay = mvarDelay
End Property

Public Property Let Delay(ByVal vData As Integer)
    mvarDelay = vData
End Property

Public Property Get EndingAngle() As Single
    EndingAngle = mvarEndingAngle
End Property

Public Property Let EndingAngle(ByVal vData As Single)
    mvarEndingAngle = vData
End Property

Public Property Get Series() As Series
    Set Series = mvarSeries
End Property

Public Property Set Series(vData As Series)
    Set mvarSeries = vData
End Property

Public Property Get StartingAngle() As Single
    StartingAngle = mvarStartingAngle
End Property

Public Property Let StartingAngle(ByVal vData As Single)
    mvarStartingAngle = vData
End Property

Public Property Get Step() As Single
    Step = mvarStep
End Property

Public Property Let Step(ByVal vData As Single)
    mvarStep = vData
End Property

Public Property Get Visible() As Boolean
    Visible = mvarVisible
End Property

Public Property Let Visible(ByVal vData As Boolean)
    mvarVisible = vData
End Property

Private Sub Class_Initialize()
    Set mvarSeries = New Series
    
    mvarDelay = 0
    mvarEndingAngle = 360
    mvarStartingAngle = 0
    mvarStep = 1
    mvarVisible = True
End Sub

