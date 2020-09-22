VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Origin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mvarAuto As Boolean
Private mvarX    As Single
Private mvarY    As Single

Public Property Let Auto(ByVal vData As Boolean)
    mvarAuto = vData
End Property

Public Property Get Auto() As Boolean
    Auto = mvarAuto
End Property

Public Property Let X(ByVal vData As Single)
    mvarX = vData
End Property

Public Property Get X() As Single
    X = mvarX
End Property

Public Property Let Y(ByVal vData As Single)
    mvarY = vData
End Property

Public Property Get Y() As Single
    Y = mvarY
End Property

Public Sub SetPos(ByVal vDataX As Single, ByVal vDataY As Single)
    mvarX = vDataX
    mvarY = vDataY
End Sub

Private Sub Class_Initialize()
    Auto = True
    X = 0
    Y = 0
End Sub
