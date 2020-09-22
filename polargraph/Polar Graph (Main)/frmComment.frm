VERSION 5.00
Begin VB.Form frmComment 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Comment"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      Top             =   2940
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   4140
      TabIndex        =   2
      Top             =   2940
      Width           =   855
   End
   Begin VB.TextBox txtComment 
      Height          =   2895
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4995
   End
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    frmMain.pgPolarGraph.Comment = Trim$(txtComment.Text)
    Unload Me
End Sub
