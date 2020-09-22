VERSION 5.00
Begin VB.Form frmZoom 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Custom Zoom"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   915
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   180
      Width           =   915
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Zoom to (%)"
      Height          =   975
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   2355
      Begin VB.ComboBox cmbZoom 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   1995
      End
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current zoom: "
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1020
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbZoom_Change()
    If Len(cmbZoom.Text) > 0 Then
        If Not cmdOk.Enabled Then cmdOk.Enabled = True
    Else
        If cmdOk.Enabled Then cmdOk.Enabled = False
    End If
End Sub

Private Sub cmbZoom_Click()
    If Not cmdOk.Enabled Then cmdOk.Enabled = True
End Sub

Private Sub cmbZoom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If (KeyAscii >= vbKey0) And (KeyAscii <= vbKey9) Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If (Val(cmbZoom.Text) < 20) Or (Val(cmbZoom.Text) > 120) Then
        MsgBox "Enter a value between 20 to 120.", vbInformation Or vbOKOnly, "Zoom"
    Else
        With frmMain.pgPolarGraph
            If .Zoom <> Val(cmbZoom.Text) Then
                .Zoom = Val(cmbZoom.Text)
            End If
        End With
        
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    For i = 20 To 120
        cmbZoom.AddItem i
    Next i
    
    cmbZoom.Text = frmMain.pgPolarGraph.Zoom
    lblLabel.Caption = "Current zoom: " & cmbZoom.Text & "%"
    cmdOk.Enabled = False
    
    cmbZoom.SelStart = 0
    cmbZoom.SelLength = Len(cmbZoom.Text)
End Sub
