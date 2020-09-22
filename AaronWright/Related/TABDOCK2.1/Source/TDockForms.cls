VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TDockForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"DockForm"
Attribute VB_Ext_KEY = "Member0" ,"DockForm"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ******************************************************************************
' Class:             DockedForm
' Description:       Collection of Docked Forms, used by TabDock
'                    to keep track of the forms that are part of the
'                    engine.
' Created by:        Marclei V Silva
' Machine:           ZEUS
' Date-Time:         09/05/2000 3:07:33
' Last modification: 09/05/2000 3:07:33
' ******************************************************************************
Option Explicit

' local variable to hold collection
Private mCol As Collection

Friend Function Add( _
    ByVal Item As Form, _
    ByVal Panel As TTabDockHost, _
    Style As tdDockStyles, _
    State As tdDockedState, _
    sKey As String, _
    hWnd As Long) As TDockForm
    
    Dim objNewMember As TDockForm
    Dim i As Integer
    Dim Index As Integer
    Dim lWidth As Long, lHeight As Long, lLeft As Long, lTop As Long, vis As Boolean
    
    ' create a new object
    Set objNewMember = New TDockForm

    ' although we are passing the object references here, only
    ' a pointer to them will be saved for further tracking of the object
    ' this avoid vb creation of multiple references to the object
    ' and minimizes the chances of crashing when we finishing the
    ' application
    Set objNewMember.Panel = Panel
    Set objNewMember.Extender = Item
    
    ' set the properties passed into the method
    vis = True
    If Persist = True Then
      lWidth = (GetSetting(App.Title, "Docking", sKey & "FloatWidth", Item.Width / Screen.TwipsPerPixelX))
      lHeight = (GetSetting(App.Title, "Docking", sKey & "FloatHeight", Item.Height / Screen.TwipsPerPixelY))
      lLeft = GetSetting(App.Title, "Docking", sKey & "FloatLeft", Item.Left / Screen.TwipsPerPixelX)
      lTop = GetSetting(App.Title, "Docking", sKey & "FloatTop", Item.Top / Screen.TwipsPerPixelY)
      State = GetSetting(App.Title, "Docking", sKey & "State", State)
      Style = GetSetting(App.Title, "Docking", sKey & "Style", Style)
      vis = GetSetting(App.Title, "Docking", sKey & "Visible", True)
    Else
      lWidth = Item.Width / Screen.TwipsPerPixelX
      lHeight = Item.Height / Screen.TwipsPerPixelY
      lLeft = Item.Left / Screen.TwipsPerPixelX
      lTop = Item.Top / Screen.TwipsPerPixelY
    End If
    objNewMember.Style = Style
    objNewMember.State = State
    objNewMember.FloatingWidth = lWidth
    objNewMember.FloatingHeight = lHeight
    objNewMember.FloatingLeft = lLeft
    objNewMember.FloatingTop = lTop
    objNewMember.ParentHWND = hWnd
    objNewMember.Key = sKey
    objNewMember.StoreVis = vis
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    ' return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    
End Function

Public Property Get Item(Index As Variant) As TDockForm
Attribute Item.VB_Description = "Returns a docked form from the list"
Attribute Item.VB_UserMemId = 0
    Set Item = mCol(Index)
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the collection"
    Count = mCol.Count
End Property

Public Sub Remove(Index As Variant)
Attribute Remove.VB_Description = "Removes a item (DockForm) from the collection list"
    mCol.Remove Index
End Sub

' ******************************************************************************
' Routine       : RemoveByHandle
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 23/12/0020:20:03
' Inputs        : hWndA : The form handle
' Outputs       : nothing
' Credits       :
' Modifications :
' Description   : This procedure was create because
'                 we can not reference the key attribute once
'                 we do not know wether the user supplied one
'                 This come to rescue us from this.
' ******************************************************************************
Friend Sub RemoveByHandle(hWndA As Long)
    Dim Index As Integer
    
    For Index = 1 To mCol.Count
        If mCol(Index).hWnd = hWndA Then
            mCol.Remove Index
            Exit Sub
        End If
    Next
End Sub

Friend Function ItemByHandle(hWndA As Long) As TDockForm
    Dim Index As Integer
    
    For Index = 1 To mCol.Count
        If mCol(Index).hWnd = hWndA Then
            Set ItemByHandle = mCol(Index)
            Exit Function
        End If
    Next
End Function

Friend Function HandleExists(hWndA As Long) As Boolean
    Dim Index As Integer
    
    For Index = 1 To mCol.Count
        If mCol(Index).hWnd = hWndA Then
            HandleExists = True
            Exit Function
        End If
    Next
    HandleExists = False
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    ' creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    ' destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Friend Sub Append(DockForm As TDockForm, Index As Integer)
    If Index <= mCol.Count Then
        mCol.Add DockForm, , Index
    Else
        mCol.Add DockForm
    End If
End Sub

'Friend Function GetIndex(DockForm As TDockForm) As Integer
'    Dim Index As Integer
'    Dim Rc As RECT
'    Dim Rw As RECT
'
'    ' let's DockArrange window list when a new window is added to
'    ' the list
'    GetWindowRect DockForm.hWnd, Rc
'    For Index = 1 To mCol.Count
'        GetWindowRect mCol(Index).hWnd, Rw
'        If Rw.Top > Rc.Top Then
'            Exit For
'        End If
'    Next
'    GetIndex = Index
'End Function

Friend Function VisibleCount() As Integer
    Dim Count As Integer
    Dim i As Integer
    Dim w As TDockForm

    Count = 0
    For Each w In mCol
        If w.Visible = True Then
            Count = Count + 1
        End If
    Next
    VisibleCount = Count
End Function
'-- end code
