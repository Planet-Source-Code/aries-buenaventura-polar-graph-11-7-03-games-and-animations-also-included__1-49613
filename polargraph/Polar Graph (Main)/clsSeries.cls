VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Series"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Marker"
Attribute VB_Ext_KEY = "Member1" ,"Pen"
Option Explicit

Private mvarAllowMarker As Boolean
Private mvarAllowPen    As Boolean
Private mvarMarker      As Marker
Private mvarPen         As Pen

Public Property Get AllowMarker() As Boolean
    AllowMarker = mvarAllowMarker
End Property

Public Property Let AllowMarker(vData As Boolean)
    mvarAllowMarker = vData
End Property

Public Property Get AllowPen() As Boolean
    AllowPen = mvarAllowPen
End Property

Public Property Let AllowPen(vData As Boolean)
    mvarAllowPen = vData
End Property

Public Property Get Marker() As Marker
    Set Marker = mvarMarker
End Property

Public Property Set Marker(vData As Marker)
    Set mvarMarker = vData
End Property

Public Property Get Pen() As Pen
    Set Pen = mvarPen
End Property

Public Property Set Pen(vData As Pen)
    Set mvarPen = vData
End Property

Private Sub Class_Initialize()
    mvarAllowMarker = False
    mvarAllowPen = True
    Set mvarMarker = New Marker
    Set mvarPen = New Pen
End Sub

Private Sub Class_Terminate()
    Set mvarMarker = Nothing
    Set mvarPen = Nothing
End Sub

