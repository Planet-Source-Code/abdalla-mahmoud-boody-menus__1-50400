VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Boody Menus (R.1.0) Test"
   ClientHeight    =   6090
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5760
      Left            =   0
      ScaleHeight     =   5700
      ScaleWidth      =   2235
      TabIndex        =   1
      Top             =   0
      Width           =   2295
      Begin VB.Frame Frame3 
         Caption         =   "Complete Example"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   9
         Top             =   3720
         Width           =   1935
         Begin VB.CommandButton cmdSDI 
            Caption         =   "SDI Form"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Popup Menu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   1935
         Begin VB.CommandButton cmdPopupNew 
            Caption         =   "(New Menu)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdPopupEdit 
            Caption         =   "Edit Menu"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Shortcuts"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1935
         Begin VB.CommandButton cmdSL 
            Caption         =   "&Load"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton cmdSS 
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton cmdSC 
            Caption         =   "&Customize"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   1335
         End
      End
   End
   Begin MSComctlLib.ImageList imlMenus 
      Left            =   5880
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0000
            Key             =   "NEW"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":039A
            Key             =   "OPEN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0734
            Key             =   "SAVE"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0ACE
            Key             =   "CUT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0E68
            Key             =   "COPY"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1202
            Key             =   "PASTE"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11853
            Text            =   "For Help Press F1"
            TextSave        =   "For Help Press F1"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "MDImain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== Variables
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
Private WithEvents m_Menus As cMenus
Attribute m_Menus.VB_VarHelpID = -1
Private Const ODID_COLOR As Long = 1&

Private Sub cmdPopupEdit_Click()
    Call m_Menus.PopUpMenu("mnuEdit")
End Sub
Private Sub cmdPopupNew_Click()
    '// example of using Menu From Nothing
    Dim NewMenu As cMenus
    Dim lParentIndex As Long
    Dim rID As Long
    Set NewMenu = New cMenus
    With NewMenu
        .DrawStyle = mds_XP
        Call .CreateFromNothing(Me.hWnd)
        lParentIndex = .AddItem(0, Key:="MainMenu")
        Call .AddItem(lParentIndex, "This menu is", , , "a")
        Call .AddItem(lParentIndex, "Created without form", "Ctrl+N|Ctrl+O", , "b")
        rID = .PopUpMenu("MainMenu")
        If (rID <> 0) Then
            .CurrentMenuIndex = .IndexForID(rID)
            MsgBox .ItemCaption
        End If
    End With
    Set NewMenu = Nothing
End Sub
Private Sub cmdSC_Click()
    Call frmCustomize.ShowDialog(m_Menus, Me)
End Sub

Private Sub cmdSL_Click()
    If (m_Menus.LoadShortcuts(App.Path & "\Shortcuts.non")) Then
        Call MsgBox("File loaded successfully .", vbInformation, "Success")
    Else
        Call MsgBox("Error while loading file .", vbCritical, "Failed")
    End If
End Sub
Private Sub cmdSS_Click()
    If (m_Menus.SaveShortcuts(App.Path & "\Shortcuts.non")) Then
        Call MsgBox("File saved successfully .", vbInformation, "Success")
    Else
        Call MsgBox("Error while saving file .", vbCritical, "Failed")
    End If
End Sub

Private Sub cmdSDI_Click()
    Call frmMain.Show
End Sub

Private Sub m_Menus_Click(ByVal Index As Long)
    Call MsgBox("Menu Clicked" & vbCrLf _
              & "Index = " & Index & vbCrLf _
              & "Key   = " & m_Menus.ItemKey(Index) & vbCrLf _
              & "Caption = " & m_Menus.ItemCaption(Index) & vbCrLf _
              , vbInformation, "Clicked")
End Sub
Private Sub m_Menus_DrawItem(Cancel As Boolean, ByVal Index As Long, ByVal hDC As Long, ByVal bSelected As Boolean, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
    'draw the color
    Dim DrawID   As Long
    Dim DrawData As Long
    DrawID = m_Menus.ItemOwnerDrawID(Index)
    DrawData = m_Menus.ItemOwnerDrawData(Index)
    If (DrawID = ODID_COLOR) Then
        Dim xMemDC As New cMemDC
        With xMemDC
            Call .Init(X2 - X1, Y2 - Y1)
            If (bSelected) Then
                Call .Rectangle(0, 0, .Width, .Height, SC_XPHighlight, , SC_XPBorder)
            Else
                Call .FillRect(0, 0, .Width, .Height, SC_XPBack)
            End If
            Call .Rectangle(5, 5, .Width - 5, .Height - 5, DrawData, , vbBlack)
            Call .BitBlt(hDC, X1, Y1, .Width, .Height, 0, 0)
        End With
        Set xMemDC = Nothing
    Else
        Exit Sub
    End If
    Cancel = True
End Sub

Private Sub m_Menus_MeasureItem(Cancel As Boolean, ByVal Index As Long, rWidth As Long, rHeight As Long)
    'measure the size of the color
    Dim DrawID   As Long
    Dim DrawData As Long
    DrawID = m_Menus.ItemOwnerDrawID(Index)
    DrawData = m_Menus.ItemOwnerDrawData(Index)
    If (DrawID = ODID_COLOR) Then
        rWidth = 15
        rHeight = 25
    Else
        Exit Sub
    End If
    Cancel = True
End Sub
Private Sub m_Menus_MenuExit()
    sbrMain.Panels(1).Text = "For Help Press F1"
End Sub
Private Sub m_Menus_MenuHighlight(ByVal Index As Long)
    sbrMain.Panels(1).Text = m_Menus.ItemHelpText(Index)
End Sub

'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== Form Standard Events
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
Private Sub MDIForm_Load()
    Set m_Menus = New cMenus
    Call pvSetupMenus
    Call frmDocument.Show
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set m_Menus = Nothing
End Sub
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
'==== Menus
'*=*==*=*==*=*==*=*==*=*==*=*==*=*==*=*==*
Private Sub pvSetupMenus()
    Dim lParentIndex  As Long
    Dim lParentIndex2 As Long
    With m_Menus
        '// Set the draw style
        .DrawStyle = mds_XP
        '// Set the image list
        Set .ImageList = imlMenus
        '// Subclass
        Call .CreateFromForm(Me)
        '// Fille File Menu
        lParentIndex = .IndexForKey("mnuFile")
        Call .AddItem(lParentIndex, "&New", "Ctrl+N", , "New", , , , "NEW", "Start a new document")
        Call .AddItem(lParentIndex, "&Open", "Ctrl+O", , "Open", , , , "OPEN", "Open an existing document")
        Call .AddItem(lParentIndex, "&Save", "Ctrl+S", , "Save", , , , "SAVE", "Save the current document")
        Call .AddItem(lParentIndex, "Save As", "Ctrl+Shift+S", , "Save As", False)
        Call .AddItem(lParentIndex, "-")
        Call .AddItem(lParentIndex, "E&xit", , , "Exit", , , , , "Exit from the program")
        '// Fille Edit Menu
        lParentIndex = .IndexForKey("mnuEdit")
        Call .AddItem(lParentIndex, "C&ut", "Ctrl+X", , "mnuEditCut", , , , "CUT", "Cut selected text")
        Call .AddItem(lParentIndex, "&Copy", "Ctrl+C", , "mnuEditCopy", , , , "COPY", "Copy selected text")
        Call .AddItem(lParentIndex, "&Paste", "Ctrl+V", , "mnuEditPaste", , , , "PASTE", "Paste text from the clipboard")
        '// Fille View Menu
        lParentIndex = .IndexForKey("mnuView")
        Call .AddItem(lParentIndex, "Check Boxes", , , "mnuViewCheckBoxes")
        Call .AddItem(lParentIndex, "Option Boxes", , , "mnuViewOptionBoxes")
        Call .AddItem(lParentIndex, "Colors", , , "mnuViewColors")
        '// Fille View Check Boxes Menu
        lParentIndex2 = .IndexForKey("mnuViewCheckBoxes")
        Call .AddItem(lParentIndex2, "Item 1", , , "Check1", , True)
        Call .AddItem(lParentIndex2, "Item 2", , , "Check2")
        Call .AddItem(lParentIndex2, "Item 3", , , "Check2")
        '// Fille View Option Boxes Menu
        lParentIndex2 = .IndexForKey("mnuViewOptionBoxes")
        Call .AddItem(lParentIndex2, "Item 1", , , "Option1", , True, mcs_Radio)
        Call .AddItem(lParentIndex2, "Item 2", , , "Option2", , , mcs_Radio)
        Call .AddItem(lParentIndex2, "Item 3", , , "Option3", , , mcs_Radio)
        '// Fille View Colors Menu
        lParentIndex2 = .IndexForKey("mnuViewColors")
        Dim I As Long
        For I = 1 To 16
            Call .AddItem(lParentIndex2, OwnerDraw:=True, OwnerDrawID:=ODID_COLOR, OwnerDrawData:=QBColor(I - 1), Break:=(((I - 1) Mod 4) = 0))
        Next
    End With
End Sub
