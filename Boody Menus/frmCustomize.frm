VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customize"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Frame framTabs 
      Height          =   3615
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   5655
      Begin VB.TextBox txtNewShortcut 
         Height          =   405
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Frame framAssignedTo 
         Caption         =   "Assigned To"
         Height          =   855
         Left            =   2400
         TabIndex        =   18
         Top             =   2520
         Visible         =   0   'False
         Width           =   1935
         Begin VB.Label lblOwner 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "(Unassigned)"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1665
         End
      End
      Begin VB.CommandButton cmdResetAll 
         Caption         =   "&Reset All"
         Height          =   375
         Left            =   4560
         TabIndex        =   16
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton cmdRemoveKey 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4560
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdAssignKey 
         Caption         =   "&Assign"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4560
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox lstKeys 
         Height          =   1230
         Left            =   2400
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
      Begin VB.ListBox lstCommands 
         Height          =   1035
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   2055
      End
      Begin VB.ComboBox cboCategories 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned To :"
         Height          =   195
         Left            =   4560
         TabIndex        =   17
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Press new shortcut key :"
         Height          =   195
         Index           =   4
         Left            =   2400
         TabIndex        =   13
         Top             =   1800
         Width           =   1785
      End
      Begin VB.Label lblDescription 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Keys :"
         Height          =   195
         Index           =   2
         Left            =   2400
         TabIndex        =   9
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Commands :"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   885
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category :"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.Frame framTabs 
      Height          =   3615
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5655
      Begin VB.Label lblUnused 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmCustomize.frx":000C
         Height          =   1095
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.Label lbl0 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "There is not items available to show in this tab ."
         ForeColor       =   &H80000010&
         Height          =   195
         Left            =   1080
         TabIndex        =   3
         Top             =   1680
         Width           =   3420
      End
   End
   Begin MSComctlLib.TabStrip tbsMain 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7223
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "(General)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Keyboard"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCustomize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== Variabkes
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
Private m_oldTabIndex As Long
Private m_Extender    As cMenus
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== Main Functions
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
Public Function ShowDialog(ByRef Extender As Object, Optional OwnerForm)
    Dim I As Long
    For I = 0 To framTabs.UBound
        With framTabs(I)
            .Visible = (I = 0)
            .BorderStyle = 0
        End With
    Next
    Set m_Extender = Extender
    With m_Extender
        For I = 1 To .ItemCount
            If (.ItemTopMenu(I)) Then
                Call cboCategories.AddItem(Mid$(.ItemKey(I), 4))
            End If
        Next
    End With
    cboCategories.ListIndex = 0
    '//fill commands
    Call Me.Show(, OwnerForm)
End Function
Private Sub cmdClose_Click()
    Call Unload(Me)
End Sub
Private Sub cmdResetAll_Click()
    '//read all shortcuts from the original file
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set m_Extender = Nothing
End Sub
Private Sub tbsMain_Click()
    framTabs(m_oldTabIndex).Visible = False
    m_oldTabIndex = (tbsMain.SelectedItem.Index - 1)
    framTabs(m_oldTabIndex).Visible = True
End Sub
Private Sub tbsMain_GotFocus()
    On Error Resume Next
    Call cmdClose.SetFocus
End Sub
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== Menus-Shortcuts
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
Private Sub cboCategories_Click()
    On Error Resume Next
    Call lstCommands.Clear
    Call lstKeys.Clear
    Call pvFillComamnds(m_Extender.ItemID("mnu" & cboCategories.Text))
    lstCommands.ListIndex = 0
End Sub
Private Sub pvFillComamnds(ByVal ID As Long)
    Dim I As Long
    Dim currID As Long
    Dim cKey As String
    'Exit Sub
    For I = 1 To m_Extender.ItemCount
        currID = m_Extender.ItemParentID(I)
        If (currID = ID) Then
            cKey = m_Extender.ItemKey(I)
            If (cKey <> vbNullString) Then
                Call lstCommands.AddItem(cKey)
                'Call pvFillComamnds(currID)
            End If
        End If
    Next
End Sub
Private Sub pvFillKeys(ByVal Key As String)
    Dim cKeys As String
    Dim keyArr() As String
    Dim I As Long
    Call lstKeys.Clear
    cKeys = m_Extender.ItemKeyAccel(Key)
    keyArr = Split(cKeys, "|")
    For I = 0 To UBound(keyArr)
        Call lstKeys.AddItem(keyArr(I))
    Next
    Erase keyArr
    cmdRemoveKey.Enabled = False
End Sub
Private Sub cmdAssignKey_Click()
    lstKeys.AddItem (txtNewShortcut.Text)
    Call pvSetCurrentShortcut
    txtNewShortcut.Text = vbNullString
End Sub
Private Sub cmdRemoveKey_Click()
    If (MsgBox("Are you sure you want to delete this key ?", vbYesNo Or vbQuestion, "Prompt") = vbYes) Then
        Call lstKeys.RemoveItem(lstKeys.ListIndex)
        If (lstKeys.ListCount = 0) Then
            m_Extender.ItemKeyAccel(lstCommands.Text) = vbNullString
        Else
            Call pvSetCurrentShortcut
        End If
    End If
End Sub
Private Sub pvSetCurrentShortcut()
    Dim I As Long
    Dim rKey As String
    For I = 0 To (lstKeys.ListCount - 1)
        If (I = 0) Then
            rKey = lstKeys.List(I)
        Else
            rKey = rKey & "|" & lstKeys.List(I)
        End If
    Next
    m_Extender.ItemKeyAccel(lstCommands.Text) = rKey
End Sub
Private Sub lstCommands_Click()
    On Error Resume Next
    Call pvFillKeys(lstCommands.Text)
    lblDescription.Caption = m_Extender.ItemDescription(lstCommands.Text)
    lstKeys.ListIndex = 0
End Sub
Private Sub lstKeys_Click()
    If (lstKeys.ListIndex < 0) Then
        cmdRemoveKey.Enabled = False
    Else
        cmdRemoveKey.Enabled = True
    End If
End Sub
Private Sub txtNewShortcut_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ownerItem As Long
    txtNewShortcut.Text = GetKeyName(KeyCode, ((Shift And vbAltMask) = vbAltMask), ((Shift And vbCtrlMask) = vbCtrlMask), ((Shift And vbShiftMask) = vbShiftMask))
    ownerItem = m_Extender.IndexForKeyAccel(txtNewShortcut.Text)
    If (ownerItem = 0) Then
        lblOwner.Caption = "(Unassigned)"
    Else
        lblOwner.Caption = m_Extender.ItemKey(ownerItem)
    End If
    cmdAssignKey.Enabled = ((txtNewShortcut.Text <> vbNullString) And (lstCommands.Text <> vbNullString) And (ownerItem = 0))
    framAssignedTo.Visible = (txtNewShortcut.Text <> vbNullString)
End Sub
Private Sub txtNewShortcut_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
