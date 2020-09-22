VERSION 5.00
Begin VB.Form frmExpr 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Custom Expression"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   540
      Width           =   975
   End
   Begin VB.Frame fraFrame 
      Height          =   3435
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3675
      Begin VB.PictureBox picTray 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   180
         ScaleHeight     =   375
         ScaleWidth      =   3375
         TabIndex        =   8
         Top             =   3000
         Width           =   3375
         Begin VB.CommandButton cmdRemov 
            Caption         =   "&Remove"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            TabIndex        =   13
            Top             =   0
            Width           =   795
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   375
            Left            =   840
            TabIndex        =   12
            Top             =   0
            Width           =   795
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   795
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2520
            TabIndex        =   10
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.ListBox lstExpr 
         Height          =   2400
         ItemData        =   "frmExpr.frx":0000
         Left            =   60
         List            =   "frmExpr.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   540
         Width           =   3555
      End
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Select the expression would you like to used, then click OK."
         Height          =   495
         Left            =   180
         TabIndex        =   7
         Top             =   120
         Width           =   3285
      End
   End
   Begin VB.Frame fraTray 
      Caption         =   "Enter your expression"
      Enabled         =   0   'False
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   3480
      Width           =   4695
      Begin VB.CommandButton cmdCancelExpr 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   3960
         TabIndex        =   3
         Top             =   225
         Width           =   675
      End
      Begin VB.CommandButton cmdOkExpr 
         Caption         =   "Ok"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3240
         TabIndex        =   2
         Top             =   225
         Width           =   675
      End
      Begin VB.TextBox txtExpr 
         Height          =   285
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FileSignature = "Polar Graph 1.0"

Public RetVal As String

Dim IsInitAni           As Boolean
Dim IsChanged           As Boolean
Dim Filename            As String

Private Sub cmdAdd_Click()
    If Not fraTray.Enabled Then
        fraTray.Enabled = True
        fraTray.Tag = "Add"
        fraTray.Caption = "Enter your expression:"
        txtExpr.Text = ""
        txtExpr.SetFocus
        
        Me.Height = Me.Height + fraTray.Height + _
                    Screen.TwipsPerPixelY * 5
        Call EnabledControl(False)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelExpr_Click()
    If fraTray.Enabled Then
        fraTray.Enabled = False
        Me.Height = Me.Height - fraTray.Height - _
                    Screen.TwipsPerPixelY * 5
        Call EnabledControl(True)
    End If
End Sub

Private Sub cmdEdit_Click()
    If Not fraTray.Enabled Then
        If lstExpr.ListIndex = -1 Then
            If lstExpr.ListCount > 0 Then
                lstExpr.Selected(lstExpr.TopIndex) = True
            End If
        End If
        
        fraTray.Enabled = True
        fraTray.Tag = "Edit"
        fraTray.Caption = "Enter your new expression:"
        txtExpr.Text = lstExpr.List(lstExpr.ListIndex)
        txtExpr.SetFocus
        
        Me.Height = Me.Height + fraTray.Height + _
                    Screen.TwipsPerPixelY * 5
        Call EnabledControl(False)
    End If
End Sub

Private Sub cmdOk_Click()
    If lstExpr.ListIndex <> -1 Then
        RetVal = lstExpr.List(lstExpr.ListIndex)
    Else
        RetVal = ""
    End If

    Unload Me
End Sub

Private Sub cmdOkExpr_Click()
    If Trim$(txtExpr.Text) = "" Then Exit Sub
        
    Dim MathLib As New MathLibrary
    Dim Script  As New MSScriptControl.ScriptControl
    On Error Resume Next
        
    Script.Language = "VBScript"
    Script.Timeout = NoTimeout
    Script.AddObject "MathLib", MathLib, True
    Script.Eval Trim$(txtExpr.Text)
        
    If Err Then
        MsgBox Err.Description, vbOKOnly Or vbInformation, "Error"
    Else
        Dim i  As Integer
        Dim s1 As String
        Dim s2 As String
        
        For i = 0 To lstExpr.ListCount - 1
            s1 = RemovSpace(LCase(Trim$(txtExpr.Text)))
            s2 = RemovSpace(LCase(lstExpr.List(i)))
            
            If s1 = s2 Then
                MsgBox "Expression already exist!", _
                        vbInformation Or vbOKOnly, fraTray.Tag
                Exit Sub
            End If
        Next i
        
        If fraTray.Tag = "Add" Then
            lstExpr.AddItem Trim$(txtExpr.Text)
        ElseIf fraTray.Tag = "Edit" Then
            Dim curIndex As Integer
            
            curIndex = lstExpr.ListIndex
            lstExpr.RemoveItem curIndex
            lstExpr.AddItem Trim$(txtExpr.Text), curIndex
            lstExpr.ListIndex = curIndex
        End If
        
        txtExpr.Text = ""
        IsChanged = True
    End If
End Sub

Private Sub cmdRemov_Click()
    If lstExpr.ListIndex <> -1 Then
        If MsgBox("Are you sure?", vbYesNo Or _
                   vbQuestion, "Remove") = vbYes Then
            lstExpr.RemoveItem lstExpr.ListIndex
            lstExpr.SetFocus
            If lstExpr.ListCount = 0 Then
                cmdEdit.Enabled = False
                cmdRemov.Enabled = False
            End If
            
            IsChanged = True
            cmdSave.Enabled = True
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    Call FileSave
    
    IsChanged = False
    cmdSave.Enabled = False
End Sub

Private Sub Form_Load()
    Filename = App.Path & "\expr.dat"
    
    Call FileOpen
    If lstExpr.ListCount > 0 Then
        cmdEdit.Enabled = True
        cmdRemov.Enabled = True
    End If
    
    IsChanged = False
    IsInitAni = False
End Sub

Private Sub FileOpen()
    On Error GoTo OpenErr
    
    If Dir$(Filename) <> "" Then
        Dim InFile As Integer
        Dim Buffer As String
        
        lstExpr.Clear
        
        InFile = FreeFile
        Open Filename For Input As InFile
            Input #InFile, Buffer
            If CStr(Buffer) = FileSignature Then
                Do While Not EOF(InFile)
                    Input #InFile, Buffer
                    lstExpr.AddItem CStr(Buffer)
                Loop
            End If
        Close InFile
    End If
    Exit Sub

OpenErr:
End Sub

Private Sub FileSave()
    On Error GoTo SaveErr
    
    Dim i      As Integer
    Dim InFile As Integer
    
    InFile = FreeFile
    Open Filename For Output As InFile
        Write #InFile, FileSignature
        For i = 0 To lstExpr.ListCount - 1
            Write #InFile, lstExpr.List(i)
        Next i
    Close InFile
    Exit Sub
    
SaveErr:
End Sub

Private Sub EnabledControl(bVal As Boolean)
    cmdAdd.Enabled = bVal
    
    If IsChanged And Not fraTray.Enabled Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
    
    If lstExpr.ListCount = 0 Then
        cmdEdit.Enabled = False
        cmdRemov.Enabled = False
    Else
        cmdEdit.Enabled = bVal
        cmdRemov.Enabled = bVal
    End If
End Sub

Private Sub lstExpr_Click()
    If fraTray.Enabled Then
        If lstExpr.ListIndex <> -1 Then
            txtExpr.Text = lstExpr.List(lstExpr.ListIndex)
        End If
    End If
End Sub

Private Sub txtExpr_Change()
    If Len(txtExpr.Text) > 0 Then
        cmdOkExpr.Enabled = True
    Else
        cmdOkExpr.Enabled = False
    End If
End Sub

Private Function RemovSpace(Data As String) As String
    If Data = "" Then Exit Function
    
    Dim i      As Integer
    Dim Buffer As String
    
    Buffer = ""
    For i = 1 To Len(Data)
        If Mid$(Data, i, 1) <> " " Then
            Buffer = Buffer & Mid$(Data, i, 1)
        End If
    Next
    
    RemovSpace = Buffer
End Function

