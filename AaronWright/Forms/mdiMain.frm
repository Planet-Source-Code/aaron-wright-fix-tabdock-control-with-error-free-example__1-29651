VERSION 5.00
Object = "{DE2F0821-EA7E-11D5-A5C7-DF4677A50515}#1.0#0"; "TABDOCKOCX.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H00808080&
   Caption         =   "Advanced Code Editor"
   ClientHeight    =   5625
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8025
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin TabDock.TTabDock tDock 
      Left            =   3840
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu fileSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
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
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Window"
      Begin VB.Menu mnuTWins 
         Caption         =   "Toolbars"
         Begin VB.Menu mnuViewDock 
            Caption         =   "Toolbars"
            Index           =   0
         End
         Begin VB.Menu mnuViewDock 
            Caption         =   "Debug"
            Index           =   1
         End
         Begin VB.Menu mnuViewDock 
            Caption         =   "Global Tools"
            Index           =   2
         End
         Begin VB.Menu mnuViewDock 
            Caption         =   "Project Explorer"
            Index           =   3
         End
         Begin VB.Menu mnuViewDock 
            Caption         =   "Document Properties"
            Index           =   4
         End
         Begin VB.Menu toolsbarsSep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUnDockAll 
            Caption         =   "Undock All..."
         End
         Begin VB.Menu mnuDockAll 
            Caption         =   "Dock All..."
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHlpCntr 
         Caption         =   "Help Center"
         Shortcut        =   {F1}
      End
      Begin VB.Menu hlpSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOTW 
         Caption         =   "On The Web"
         Begin VB.Menu mnuChkForUpdates 
            Caption         =   "Check for updates..."
         End
         Begin VB.Menu mnuWBIHome 
            Caption         =   "White Blotter, Inc. Home"
         End
      End
      Begin VB.Menu mnuUpgrade 
         Caption         =   "Upgrade"
         Begin VB.Menu mnuUpgradeSt 
            Caption         =   "Standard"
         End
         Begin VB.Menu mnuUpgradePro 
            Caption         =   "Professional"
         End
         Begin VB.Menu mnuUpgradeEnt 
            Caption         =   "Enterprise"
         End
         Begin VB.Menu mnuUpgradeCorp 
            Caption         =   "Corporate"
         End
      End
      Begin VB.Menu mnuReg 
         Caption         =   "Register..."
      End
      Begin VB.Menu hlpSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "Advanced Code Editor"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    Dock
    setupMenus
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbRightButton _
    Then PopupMenu mnuFile, , , , mnuFileNew
End Sub

Private Sub mnuDockAll_Click()
        dockall
        mnuDockAll.Visible = False
        mnuUnDockAll.Visible = True
End Sub

Private Sub mnuExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuFileNew_Click()
    frmDoc.Show
End Sub

Private Sub mnuUnDockAll_Click()
    mnuUnDockAll.Visible = False
    mnuDockAll.Visible = True
    UnDock
End Sub

Private Sub mnuViewDock_Click(Index As Integer)
    Dim Key As String
    
    ' This is a simple use of the TabDock Host.
    ' Based on the menu clicked item we will hide or
    ' show the selected form
    mnuViewDock(Index).Checked = Not mnuViewDock(Index).Checked
    
    ' Select the form you wish to operate with
    Select Case Index
        Case 0: Key = "frmToolbar"
        Case 1: Key = "frmDebug"
        Case 2: Key = "frmTools"
        Case 3: Key = "frmPExplorer"
        Case 4: Key = "frmHTMProps"
    End Select
    
    ' Now toggle visibility
    If mnuViewDock(Index).Checked Then
        tDock.FormShow Key
    Else
        tDock.FormHide Key
    End If
End Sub
