VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTable 
   Caption         =   "Table"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTable 
      Height          =   4395
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Select &All"
         Height          =   315
         Left            =   5040
         TabIndex        =   10
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   5040
         TabIndex        =   9
         Top             =   900
         Width           =   975
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "&Copy"
         Height          =   315
         Left            =   5040
         TabIndex        =   8
         Top             =   540
         Width           =   975
      End
      Begin VB.PictureBox picLegend 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         ScaleHeight     =   225
         ScaleWidth      =   765
         TabIndex        =   5
         Top             =   180
         Width           =   795
      End
      Begin MSComctlLib.ListView lvwData 
         Height          =   3075
         Left            =   60
         TabIndex        =   1
         Top             =   1260
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   5424
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Angle"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "X"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Y"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Radius"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblFunction 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         Height          =   195
         Left            =   1080
         TabIndex        =   7
         Top             =   780
         Width           =   105
      End
      Begin VB.Label lblCoefficient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         Height          =   195
         Left            =   1080
         TabIndex        =   6
         Top             =   480
         Width           =   105
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Function    : "
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   780
         Width           =   885
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coefficient : "
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   885
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Legend     : "
         Height          =   195
         Index           =   0
         Left            =   155
         TabIndex        =   2
         Top             =   180
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCopy_Click()
    Dim i      As Long
    Dim Buffer As String
    
    Buffer = ""
    
    With lvwData
        For i = 1 To .ListItems.Count
            If .ListItems(i).Selected Then
                Buffer = Buffer & .ListItems(i).Text & vbTab & _
                                  .ListItems(i).ListSubItems(1).Text & vbTab & _
                                  .ListItems(i).ListSubItems(2).Text & vbTab & _
                                  .ListItems(i).ListSubItems(3).Text & vbTab & _
                                  .ListItems(i).ListSubItems(4).Text & vbCrLf
            End If
        Next i
    End With
    
    Clipboard.Clear
    Clipboard.SetText Buffer
    
    lvwData.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSelectAll_Click()
    Dim i As Long
        
    For i = 1 To lvwData.ListItems.Count
        lvwData.ListItems(i).Selected = True
    Next i
    
    lvwData.SetFocus
End Sub

Private Sub Form_Load()
    With frmMain
        Dim Color  As Long
        Dim curpos As Integer
        Dim oPlot  As Object

        On Error Resume Next
        
        curpos = .lvwTable.SelectedItem.Index
        Set oPlot = .pgPolarGraph.Plot(curpos)
        
        Set picLegend.Picture = .imlColors.ListImages(curpos).Picture
        lblCoefficient.Caption = oPlot.Coefficient
        lblFunction.Caption = oPlot.Equation
        
        Dim Angle   As Single
        Dim Radius  As Single
        Dim X       As Single
        Dim Y       As Single
        Dim MathLib As New MathLibrary
        Dim Script  As New MSScriptControl.ScriptControl
        
        Script.Language = "VBScript"
        Script.Timeout = NoTimeout
        Script.AddCode oPlot.Coefficient
        Script.AddObject "MathLib", MathLib, True
        
        .pgbProgressBar.Min = 0
        .pgbProgressBar.Max = (oPlot.EndingAngle - oPlot.StartingAngle) / oPlot.Step + 1
        .pgbProgressBar.Value = 0
        
        For Angle = oPlot.StartingAngle To oPlot.EndingAngle Step oPlot.Step
            Script.AddCode "t=" & MathLib.Radians(Angle)
            X = Script.Eval(oPlot.Equation) * Cos(MathLib.Radians(Angle))
            Y = Script.Eval(oPlot.Equation) * Sin(MathLib.Radians(Angle))
            Radius = MathLib.Radius(X, Y)
            
            lvwData.ListItems.Add , , .pgbProgressBar.Value + 1
            
            curpos = lvwData.ListItems.Count
            lvwData.ListItems(curpos).SubItems(1) = Angle & Chr(&HB0)
            lvwData.ListItems(curpos).SubItems(2) = Round(X, 4)
            lvwData.ListItems(curpos).SubItems(3) = Round(Y, 4)
            lvwData.ListItems(curpos).SubItems(4) = Round(Radius, 4)
            .pgbProgressBar.Value = .pgbProgressBar.Value + 1
        Next Angle
    End With
End Sub

Private Sub Form_Resize()
    fraTable.Move 30, 30, Me.ScaleWidth - 60, Me.ScaleHeight - 60
    lvwData.Move 60, lvwData.Top, Me.ScaleWidth - 190, Me.ScaleHeight - lvwData.Top - 140
    cmdCopy.Move Me.ScaleWidth - cmdCopy.Width - 190, cmdCopy.Top
    cmdExit.Move Me.ScaleWidth - cmdExit.Width - 190, cmdExit.Top
    cmdSelectAll.Move Me.ScaleWidth - cmdSelectAll.Width - 190, cmdSelectAll.Top
    
    lvwData.ColumnHeaders(1).Width = lvwData.Width * 0.1 ' ...0
    lvwData.ColumnHeaders(2).Width = lvwData.Width * 0.2 ' Angle
    lvwData.ColumnHeaders(3).Width = lvwData.Width * 0.2 ' X
    lvwData.ColumnHeaders(4).Width = lvwData.Width * 0.2 ' Y
    lvwData.ColumnHeaders(5).Width = lvwData.Width * 0.2 ' Radius
End Sub

Private Sub lvwData_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Index = 1 Then
        Call cmdSelectAll_Click
    End If
End Sub
