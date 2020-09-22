VERSION 5.00
Begin VB.Form frmDocument 
   Caption         =   "Document #"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3720
   ScaleWidth      =   5085
   Begin VB.TextBox txtMain 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Resize()
    Call txtMain.Move(0, 0, Me.ScaleWidth, Me.ScaleHeight)
End Sub
