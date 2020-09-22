VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Boody Menus Example"
   ClientHeight    =   2625
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   2775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   2775
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Effects"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.CheckBox chkFont 
         Caption         =   "Custom Font ."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CheckBox chkCustomColors 
         Caption         =   "Custom colors ."
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox chkRightToLeft 
         Caption         =   "Right To Left ."
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox chkBackground 
         Caption         =   "Background picture ."
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cboDrawStyle 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   1200
         List            =   "frmMain.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin MSComctlLib.ImageList imlMenus 
         Left            =   1680
         Top             =   1680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0016
               Key             =   "OPEN"
            EndProperty
         EndProperty
      End
      Begin VB.Image imgBack 
         Height          =   330
         Left            =   960
         Picture         =   "frmMain.frx":03B0
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Draw Style :"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "C&ut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewNormalCheck 
         Caption         =   "Normal Check"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewRadioCheck 
         Caption         =   "Radio Check"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewIconCheck 
         Caption         =   "Icon Check"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_Menus As cMenus
Attribute m_Menus.VB_VarHelpID = -1

Private Sub cboDrawStyle_Click()
    m_Menus.DrawStyle = cboDrawStyle.ListIndex
End Sub

Private Sub chkBackground_Click()
    If (chkBackground.Value = 1) Then
        m_Menus.BackGroundPicture = imgBack.Picture
    Else
        m_Menus.BackGroundPicture = Nothing
    End If
End Sub
Private Sub chkCustomColors_Click()
    If (chkCustomColors.Value = 1) Then
        With m_Menus
            .BackColor = RGB(250, 235, 215)
            .HighlightColor = RGB(216, 191, 216)
            .ForeColor = vbRed
            .HighlightForeColor = vbBlack
        End With
    Else
    
    End If
End Sub

Private Sub chkFont_Click()
    If (chkFont.Value = 1) Then
        Set m_Menus.Font = chkFont.Font
    Else
        Set m_Menus.Font = Nothing 'default
    End If
End Sub

Private Sub chkRightToLeft_Click()
    Me.RightToLeft = (chkRightToLeft.Value = 1)
    m_Menus.RightToLeft = Me.RightToLeft
End Sub

Private Sub Form_Load()
    Set m_Menus = New cMenus
    With m_Menus
        cboDrawStyle.ListIndex = 0
        Set .ImageList = imlMenus
        Call .CreateFromForm(Me)
        .ItemCheckedStyle("mnuViewRadioCheck") = mcs_Radio
        .ItemCheckedStyle("mnuViewIconCheck") = mcs_Icon
        .ItemImage("mnuViewIconCheck") = "OPEN"
    End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set m_Menus = Nothing
End Sub
