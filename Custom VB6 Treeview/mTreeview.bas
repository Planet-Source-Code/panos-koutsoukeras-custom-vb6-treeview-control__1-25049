Attribute VB_Name = "mTreeview"
Option Explicit

Public TVhwnd As Long   ' TV.hwnd

' Treeview messages
Public Const WM_NOTIFY = &H4E
Public Const WM_PAINT = &HF
Public Const WM_ERASEBKGND = &H14
Public Const WM_DESTROY = &H2

Public Const GWL_STYLE = (-16)
Public Const TV_FIRST = &H1100

' TVMessages
Public Const TVM_GETIMAGELIST = (TV_FIRST + 8)
Public Const TVM_SETIMAGELIST = (TV_FIRST + 9)
Public Const TVM_SETBKCOLOR As Long = (TV_FIRST + 29)
Public Const TVM_SETTEXTCOLOR As Long = (TV_FIRST + 30)
Public Const TVM_GETBKCOLOR As Long = (TV_FIRST + 31)
Public Const TVM_GETTEXTCOLOR As Long = (TV_FIRST + 32)
Public Const TVM_GETITEM = (TV_FIRST + 12)
Public Const TVM_SETITEM = (TV_FIRST + 13)
Public Const TVM_GETNEXTITEM = (TV_FIRST + 10)
Public Const TVM_GETITEMRECT = (TV_FIRST + 4)
Public Const TVM_HITTEST = (TV_FIRST + 17)
Public Const TVM_SELECTITEM = (TV_FIRST + 11)
Public Const TVM_INSERTITEM = (TV_FIRST + 0)
Public Const TVM_SORTCHILDRENCB = (TV_FIRST + 21)
Public Const TVM_EXPAND = (TV_FIRST + 2)

' TVN_Notifications in WM_NOTIFY
Public Const TVN_FIRST = -400&   ' (0U-400U)
Public Const TVN_LAST = -499&
Public Const TVN_ITEMEXPANDINGA = (TVN_FIRST - 5)
Public Const TVN_ITEMEXPANDEDA = (TVN_FIRST - 6)
Public Const TVN_DELETEITEM = (TVN_FIRST - 9)  ' lParam = NMTREEVIEW
Public Const TVN_SELCHANGING = (TVN_FIRST - 1) ' lParam = NMTREEVIEW
Public Const TVN_SELCHANGED = (TVN_FIRST - 2)  ' lParam = NMTREEVIEW

' TVS_Styles
Public Const TVS_NOTOOLTIPS = &H80

' TVSIL_Imagelist types (wParam)
Public Const TVSIL_NORMAL = 0
Public Const TVSIL_STATE = 2

' TVGN_Item Relationships (wParam)
Public Const TVGN_ROOT = &H0
Public Const TVGN_CHILD = &H4
Public Const TVGN_NEXT = &H1
Public Const TVGN_NEXTVISIBLE As Long = &H6
Public Const TVGN_CARET = &H9

' TVIF_Mask for (TVITEM.mask)
'Public Const TVIF_TEXT = &H1
Public Const TVIF_IMAGE = &H2
Public Const TVIF_PARAM = &H4
Public Const TVIF_STATE = &H8
Public Const TVIF_SELECTEDIMAGE = &H20
Public Const TVIF_CHILDREN = &H40

' TVIS_State for (TVITEM.stateMask)
Public Const TVIS_SELECTED = &H2
Public Const TVIS_EXPANDED = &H20
Public Const TVIS_STATEIMAGEMASK = &HF000
Public Const TVIS_EXPANDEDONCE = &H40
Public Const TVIS_OVERLAYMASK = &HF00
Public Const TVIS_CUT = &H4
Public Const TVIS_BOLD  As Long = &H10

' TVHT_Hittest flags
Public Const TVHT_NOWHERE = &H1   ' In the client area, but below the last item
Public Const TVHT_ONITEMICON = &H2
Public Const TVHT_ONITEMLABEL = &H4
Public Const TVHT_ONITEMINDENT = &H8
Public Const TVHT_ONITEMBUTTON = &H10
Public Const TVHT_ONITEMRIGHT = &H20
Public Const TVHT_ONITEMSTATEICON = &H40
Public Const TVHT_ONITEM = (TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON)
' User-Defined
Public Const TVHT_ONITEMLINE = (TVHT_ONITEM Or TVHT_ONITEMINDENT Or TVHT_ONITEMBUTTON Or TVHT_ONITEMRIGHT)


' User-defined Enums
Public Enum TVItemStateImages   ' Must be Public
    tvisNoButton = 0
    tvisCollapsed = 1
    tvisExpanded = 2
End Enum

Public Enum CBoolean
    CFalse = 0
    CTrue = 1
End Enum

' TVM_GET/SETITEM lParam
' Specifies or receives attributes of a tree view item
Public Type TVITEM
    mask As Long
    hItem As Long
    state As Long
    stateMask As Long
    pszText As Long    ' if a string, must be pre-allocated!!
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
End Type

Public Type TVINSERTSTRUCT
    hParent As Long
    hInsertAfter As Long
    item As TVITEM
End Type

' TVM_HITTEST lParam
' Contains information used to determine the location of a point relative to a tree view control.
Public Type TVHITTESTINFO
    pt As POINTAPI
    flags As Long
    hItem As Long
End Type




'################## Treeview Functions #################

' Sets the specified Node's state image.
' Called from TV_Collapse(UserControl) and TV_Expand(UserControl)
Public Sub SetTVItemStateImage(hwndTV As Long, Nod As Node, dwImage As TVItemStateImages)

  Dim hItem As Long

    hItem = GetTVItemFromNode(hwndTV, Nod)
    If hItem Then
        Call Treeview_SetItemStateImage(hwndTV, hItem, dwImage)
    End If

End Sub

' If successful, returns the treeview item handle represented by
' the specified Node, returns 0 otherwise.
' Called from TV_NodeClick(UserControl) and SetTVItemStateImage
Public Function GetTVItemFromNode(hwndTV As Long, _
                                Nod As Node) As Long
  
  Dim nod1 As Node
  Dim anSiblingPos() As Integer  ' contains the sibling position of the node and all it's parents
  Dim nLevel As Integer              ' hierarchical level of the node
  Dim hItem As Long
  Dim i As Integer
  Dim nPos As Integer

    Set nod1 = Nod

    ' Continually work backwards from the current node to the current node's
    ' first sibling, caching the current node's sibling position in the one-based
    ' array. Then get the first sibling's parent node and start over. Keep going
    ' until the postion of the specified node's top level parent item is obtained...
    Do While (nod1 Is Nothing) = False
        nLevel = nLevel + 1
      ReDim Preserve anSiblingPos(nLevel)
        anSiblingPos(nLevel) = GetNodeSiblingPos(nod1)
        Set nod1 = nod1.Parent
    Loop

    ' Get the hItem of the first item in the treeview
    hItem = TreeView_GetRoot(hwndTV)
    If hItem Then
        ' Now work backwards through the cached node positions in the array
        ' (from the first treeview node to the specified node), obtaining the respective
        ' item handle for each node at the cached position. When we get to the
        ' specified node's position (the value of the first element in the array), we
        ' got it's hItem...
        For i = nLevel To 1 Step -1
            nPos = anSiblingPos(i)
            Do While nPos > 1
                hItem = TreeView_GetNextSibling(hwndTV, hItem)
                nPos = nPos - 1
            Loop
            If (i > 1) Then hItem = TreeView_GetChild(hwndTV, hItem)
        Next
        GetTVItemFromNode = hItem
    End If

End Function

' Returns the one-base position of the specified node
' with respect to it's sibling order.
' Called from GetTVItemFromNode
Private Function GetNodeSiblingPos(Nod As Node) As Integer
  
  Dim nod1 As Node
  Dim nPos As Integer
  
    Set nod1 = Nod
    
    ' Keep counting up from one until the node has no more previous siblings
    Do While (nod1 Is Nothing) = False
        nPos = nPos + 1
        Set nod1 = nod1.Previous
    Loop
    
    GetNodeSiblingPos = nPos
  
End Function

' Removes the root folder and all of its subfolders from the specified TreeView,
' Called from UserControl_ReadProperties
Public Sub EmptyTreeView(TV As TreeView)
  
    Do While TV.Nodes.Count > 0
        ' If there are any nodes
        If TV.Nodes.Count Then
            ' Collapse the root folder
            TV.Nodes(1).Root.Expanded = False
            ' Remove it, Invoking a DoUCNotify/TVN_DELETEITEM freeing the pidls
            ' we stored in InsertFolder below (we do not want to iterate the
            ' Folders collection here and free all pidls since
            ' RefreshTreeview may remove some, but not all, treeview items).
            Call TV.Nodes.Remove(TV.Nodes(1).Root.Index)
        End If
    Loop

End Sub




'###################### Treeview Macros ####################

' Retrieves the tree-view item that bears the specified relationship to a specified item.
' Returns the handle to the item if successful or 0 otherwise.
Public Function TreeView_GetNextItem(hWnd As Long, hItem As Long, flag As Long) As Long
  
    TreeView_GetNextItem = SendMessage(hWnd, TVM_GETNEXTITEM, ByVal flag, ByVal hItem)

End Function

' Retrieves the first child item. The hitem parameter must be NULL.
' Returns the handle to the item if successful or 0 otherwise.
Public Function TreeView_GetChild(hWnd As Long, hItem As Long) As Long
  
    TreeView_GetChild = TreeView_GetNextItem(hWnd, hItem, TVGN_CHILD)

End Function

' Retrieves the next sibling item.
' Returns the handle to the item if successful or 0 otherwise.
Public Function TreeView_GetNextSibling(hWnd As Long, hItem As Long) As Long
  
    TreeView_GetNextSibling = TreeView_GetNextItem(hWnd, hItem, TVGN_NEXT)

End Function

' Retrieves the topmost or very first item of the tree-view control.
' Returns the handle to the item if successful or 0 otherwise.
Public Function TreeView_GetRoot(hWnd As Long) As Long
  
    TreeView_GetRoot = TreeView_GetNextItem(hWnd, 0, TVGN_ROOT)

End Function

' Retrieves the bounding rectangle for a tree-view item and indicates whether the item is visible.
' If the item is visible and retrieves the bounding rectangle, the return value is TRUE.
' Otherwise, the TVM_GETITEMRECT message returns FALSE and does not retrieve
' the bounding rectangle.
Public Function TreeView_GetItemRect(hWnd As Long, hItem As Long, prc As RECT, fItemRect As CBoolean) As Boolean

    prc.Left = hItem
    TreeView_GetItemRect = SendMessage(hWnd, TVM_GETITEMRECT, ByVal fItemRect, prc)

End Function


' Sets some or all of a tree-view item's attributes.
' Old docs say returns zero if successful or - 1 otherwise.
' New docs say returns TRUE if successful, or FALSE otherwise
Public Function TreeView_SetItem(hWnd As Long, pitem As TVITEM) As Boolean
  
    TreeView_SetItem = SendMessage(hWnd, TVM_SETITEM, 0, pitem)

End Function

' Sets the normal or state image list for a tree-view control and redraws the control using the new images.
' Returns the handle to the previous image list, if any, or 0 otherwise.
Public Function TreeView_SetImageList(hWnd As Long, himl As Long, iImage As Long) As Long
  
    TreeView_SetImageList = SendMessage(hWnd, TVM_SETIMAGELIST, ByVal iImage, ByVal himl)

End Function

' Determines the location of the specified point relative to the client area of a tree-view control.
' Returns the handle to the tree-view item that occupies the specified point or NULL if no item
' occupies the point.
Public Function TreeView_HitTest(hWnd As Long, lpht As TVHITTESTINFO) As Long
  
    TreeView_HitTest = SendMessage(hWnd, TVM_HITTEST, 0, lpht)

End Function

Public Function Treeview_SetItemStateImage(hwndTV As Long, hItem As Long, dwState As TVItemStateImages) As Boolean

  Dim TVI As TVITEM

    TVI.hItem = hItem
    TVI.mask = TVIF_STATE
    TVI.state = INDEXTOSTATEIMAGEMASK(dwState)
    TVI.stateMask = TVIS_STATEIMAGEMASK

    Treeview_SetItemStateImage = TreeView_SetItem(hwndTV, TVI)

End Function




'###################### Imagelist Macro ####################

' Returns the one-based index of the specifed state image mask, shifted
' left twelve bits.
' Prepares the index of a state image so that a tree view control or list
' view control can use the index to retrieve the state image for an item.
' Called from Treeview_SetItemStateImage
Public Function INDEXTOSTATEIMAGEMASK(iIndex As Long) As Long
    
    INDEXTOSTATEIMAGEMASK = iIndex * (2 ^ 12)

End Function


'
'' Inserts a new item in a tree-view control.
'' Returns the handle to the new item if successful or 0 otherwise.
'Public Function TreeView_InsertItem(hWnd As Long, lpis As TVINSERTSTRUCT) As Long
'
'    TreeView_InsertItem = SendMessage(hWnd, TVM_INSERTITEM, 0, lpis)
'
'End Function



