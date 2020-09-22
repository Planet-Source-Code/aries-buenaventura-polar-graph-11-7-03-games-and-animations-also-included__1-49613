VERSION 5.00
Begin VB.Form frmAspect 
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
   Begin VB.Frame fraAspect 
      Caption         =   "Set aspect to"
      Height          =   975
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   2355
      Begin VB.TextBox txtAspect 
         Height          =   285
         Left            =   120
         MaxLength       =   6
         TabIndex        =   2
         Top             =   420
         Width           =   2055
      End
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current aspect: "
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1125
   End
End
Attribute VB_Name = "frmAspect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtAspect_Change()
    If Len(txtAspect.Text) > 0 Then
        If Not cmdOk.Enabled Then cmdOk.Enabled = True
    Else
        If cmdOk.Enabled Then cmdOk.Enabled = False
    End If
End Sub

Private Sub txtAspect_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If (KeyAscii >= vbKey0) And (KeyAscii <= vbKey9) Then
    ElseIf KeyAscii = &H2E Then ' Decimal
        If InStr(Me.ActiveControl.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If Val(txtAspect.Text) <= 0 Then
        MsgBox "The aspect ratio  must be greater than 0.", _
               vbExclamation, "Aspect Ratio"
    ElseIf Val(txtAspect.Text) > 5 Then
        MsgBox "The aspect ratio must be less than or equal to 5", _
            vbExclamation, "Aspect Ratio"
    Else
        With frmMain.pgPolarGraph
            If .AspectRatio <> Val(txtAspect.Text) Then
                .AspectRatio = Val(txtAspect.Text)
            End If
        End With
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    txtAspect.Text = frmMain.pgPolarGraph.AspectRatio
    lblLabel.Caption = "Current Aspect: " & txtAspect.Text
    cmdOk.Enabled = False
    
    txtAspect.SelStart = 0
    txtAspect.SelLength = Len(txtAspect.Text)
End Sub

