Attribute VB_Name = "modMDIMain"
''''''''''''''''''''''''''''''''''''''''''''''''''
'   THIS MODULES MAIN PURPOSE IS FOR DOCKING AND '
'   UNDOCKING ALL THE FORMS AND WORKING WITH THE '
'   MENUS ASSOCIATED WITH SHOWING HIDING FORMS.  '
'	White Blotter, Inc.			 '
''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Dock()
    'Setup Docked Forms
    mdiMain.tDock.GrabMain mdiMain.hWnd
    mdiMain.tDock.AddForm frmTools, tdDocked, tdAlignLeft, "frmTools", tdDockLeft
    mdiMain.tDock.AddForm frmToolbar, tdDocked, tdAlignTop, "frmToolbar", tdDockTop
    mdiMain.tDock.AddForm frmDebug, tdDocked, tdAlignBottom, "frmDebug", tdDockBottom
    mdiMain.tDock.AddForm frmHTMProps, tdDocked, tdAlignRight, "frmHTMProps", tdDockRight 'tdDockFloat Or tdDockRight Or tdDockBottom
    mdiMain.tDock.AddForm frmPExplorer, tdDocked, tdAlignRight, "frmPExplorer", tdDockRight 'tdDockFloat Or tdDockRight Or tdDockTop
    
    'Set Docked Form Height
    frmToolbar.Height = 600
    frmTools.Height = mdiMain.ScaleHeight
    frmHTMProps.Height = mdiMain.ScaleHeight / 2
    frmPExplorer.Height = mdiMain.ScaleHeight / 2
    
    'Make Sure the forms are not resizable!
    mdiMain.tDock.BorderStyle = bdrRaisedOuter
    mdiMain.tDock.Panels(tdAlignTop).Height = 1000
    
    mdiMain.tDock.Panels(tdAlignLeft).Resizable = False
    mdiMain.tDock.Panels(tdAlignTop).Resizable = False
    mdiMain.tDock.Panels(tdAlignRight).Resizable = False
    mdiMain.tDock.Panels(tdAlignBottom).Resizable = False
    
    'Finalize the settings and show forms
    mdiMain.tDock.Show
    mdiMain.tDock.Panels(tdAlignTop).Refresh
    mdiMain.tDock.Panels(tdAlignLeft).Refresh
    mdiMain.tDock.Panels(tdAlignRight).Refresh
    mdiMain.tDock.Panels(tdAlignBottom).Refresh
End Sub


Public Sub setupMenus()
    mdiMain.mnuViewDock(0).Checked = True
    mdiMain.mnuViewDock(1).Checked = True
    mdiMain.mnuViewDock(2).Checked = True
    mdiMain.mnuViewDock(3).Checked = True
    mdiMain.mnuViewDock(4).Checked = True
End Sub

Public Sub UnDock()
    mdiMain.mnuViewDock(0).Checked = Flase
    mdiMain.mnuViewDock(1).Checked = Flase
    mdiMain.mnuViewDock(2).Checked = Flase
    mdiMain.mnuViewDock(3).Checked = Flase
    mdiMain.mnuViewDock(4).Checked = Flase
    UndockAll
End Sub

Public Sub UndockAll()
    mdiMain.tDock.FormHide "frmDebug"
    mdiMain.tDock.FormHide "frmHTMProps"
    mdiMain.tDock.FormHide "frmToolbar"
    mdiMain.tDock.FormHide "frmTools"
    mdiMain.tDock.FormHide "frmPExplorer"
End Sub

Public Sub dockall()
    mdiMain.tDock.FormShow "frmDebug"
    mdiMain.tDock.FormShow "frmHTMProps"
    mdiMain.tDock.FormShow "frmToolbar"
    mdiMain.tDock.FormShow "frmTools"
    mdiMain.tDock.FormShow "frmPExplorer"
    setupMenus
End Sub
