VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl icTreeview1 
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   LockControls    =   -1  'True
   ScaleHeight     =   1905
   ScaleWidth      =   2685
   ToolboxBitmap   =   "TreeView.ctx":0000
   Begin MSComctlLib.TreeView TV 
      Height          =   1860
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   3281
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "icTreeview1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'#########################################################

' Treeview.ocx
' A UserControl made with a VB6 Treeview control (MSCOMCTL.OCX - 6.00.8862)
' with customized Background (Color, GradientRectHor, GradientRectHor
' GradientTri, tiled Picture), Backcolor, Forecolor, Buttons, Tooltips etc.

' Copyright © 2001 Panos Koutsoukeras
' Company:  Inspired Creations
' Web:      http://globalinspired.com
' Mail:     software@globalinspired.com
' Date:     16 July 2001

' Credits:
' Ben Baird, http://www.vbthunder.com
' Brad Martinez, http://www.mvps.org
' http://vbaccelerator.com/

' Limitations:
' The UserControl does not repaint correctly if the
' ClipControls property of the Form1 has been set to False
'
' Some of the Background options behave strangely, because
' of the Treeview control, but this is only a demo, showing
' some ways to draw on a VB6 Treeview background.
'
' If Background = fvGrdntRectVer or fvGrdntTri
' or fvPicturedTiled (m_BackScroll = True) then the control
' does not show Tooltips and Scrolls one at a time

'#########################################################

' In all other cases we have tooltips and Scrolls Ok
Option Explicit

Implements cSubclass

Private TTStd As cToolTip    ' Tooltip standard

'Property Variables:
Private m_Background As efvBackground
Private m_BackScroll As Boolean
Private m_BackColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_BorderStyle As efvBordersStyle

Private m_HasLinesAtRoot As Boolean
Private m_HasLines As Boolean
Private m_HasButtons As Boolean
Private m_ButtonSet As efvButtonSet

Private m_Indent As Single
Private m_ToolTips As efvTooltips
Private m_Picture As Picture

' Paint variables
Private memDC(1 To 3) As MemoryDC
Private RC As RECT
Private TVWidth As Long         ' Treeview width
Private TVHeight As Long        ' Treeview height
Private m_lHdc As Long          ' Stores a copy of the Picture in the memory
Private lBitmapW As Long        ' Picture's Width
Private lBitmapH As Long        ' Picture's Height
Private XOriginOffset As Long   ' The x coord of the first Item's(Root) Rect
Private YOriginOffset As Long   ' The y coord of the first Item's(Root) Rect
Private m_lXOffset As Long
Private bStopPaint As Boolean   ' Flag to prevent painting
Private prevItem As Node        ' Stores the previous node (MouseOver)

Private gRect As GRADIENT_RECT
Private gTri(1) As GRADIENT_TRIANGLE
Private vert(3) As TRIVERTEX

Private bCustButtons As Boolean ' Custom buttons
Private UChwnd As Long          ' User Control's hwnd

'Default Property Values:
Private Const m_def_Background = 0
Private Const m_def_BackScroll = False
Private Const m_def_BackColor = &H80000005  ' Window background
Private Const m_def_ForeColor = &H80000008  ' Window Text
Private Const m_def_BorderStyle = 1

Private Const m_def_HasLines = True
Private Const m_def_HasButtons = True
Private Const m_def_HasLinesAtRoot = True

Private Const m_def_Indent = 19 ' Pixels
Private Const m_def_Tooltips = False
Private Const m_def_ButtonSet = 0

' User Control's Background
Public Enum efvBackground
    fvgNone
    fvColor
    fvGrdntRectHor
    fvGrdntRectVer
    fvGrdntTri
    fvPicturedTiled
End Enum

' User Control's Borders
Public Enum efvBordersStyle
    fvbNone
    fvSingle
End Enum

' User Control's ButtonSet (These are the checkboxes)
Public Enum efvButtonSet
    fvNormal
    fvCustom
End Enum

' User Control's Tooltips
Public Enum efvTooltips
    fvtNone
    fvName  ' Works OK only if BackScroll = True or if bStopPaint = True, if we have a background
    fvTag   ' Works OK only if BackScroll = True or if bStopPaint = True, if we have a background
End Enum

'Event Declarations:
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=TV,TV,-1,MouseDown
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=TV,TV,-1,MouseMove
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=TV,TV,-1,MouseUp
Public Event Click() 'MappingInfo=TV,TV,-1,Click
Public Event DblClick() 'MappingInfo=TV,TV,-1,DblClick

Public Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=TV,TV,-1,KeyDown
Public Event KeyPress(KeyAscii As Integer) 'MappingInfo=TV,TV,-1,KeyPress
Public Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=TV,TV,-1,KeyUp

Public Event Expand(ByVal Node As Node) 'MappingInfo=TV,TV,-1,Expand
Public Event Collapse(ByVal Node As Node) 'MappingInfo=TV,TV,-1,Collapse
Public Event NodeCheck(ByVal Node As Node) 'MappingInfo=TV,TV,-1,NodeCheck
Public Event NodeClick(ByVal Node As Node) 'MappingInfo=TV,TV,-1,NodeClick

Private Sub UserControl_InitProperties()

    m_BackColor = m_def_BackColor
    m_Background = m_def_Background
    m_BackScroll = m_def_BackScroll
    m_ForeColor = m_def_ForeColor
    m_BorderStyle = m_def_BorderStyle
    
    m_HasLinesAtRoot = m_def_HasLinesAtRoot
    m_HasLines = m_def_HasLines
    m_HasButtons = m_def_HasButtons
    m_ButtonSet = m_def_ButtonSet

    m_Indent = m_def_Indent
    m_ToolTips = m_def_Tooltips
    
    Set m_Picture = LoadPicture("")

End Sub

Private Sub UserControl_Initialize()
  
  Dim i As Integer
  Dim Node1 As Node
  Dim Node2 As Node
    
    ' Fill up the treeview with Sample Nodes ala original
    Set Node1 = TV.Nodes.Add(, , , "Sample Node")
    For i = 1 To 2
        Set Node2 = TV.Nodes.Add(Node1.Index, tvwChild, , "Sample Node")
    Next
    Node1.Expanded = True
    Set Node1 = TV.Nodes.Add(, , , "Sample Node")
    
    ' We need to subclass the TreeView to catch notification messages
    Subclass

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  
  Dim i As Integer
  Dim lColor As Long
  Dim Style As Long
    
    ' Clear the Samples
    If Ambient.UserMode Then
        EmptyTreeView TV
    End If

    m_Background = PropBag.ReadProperty("Background", m_def_Background)
    m_BackScroll = PropBag.ReadProperty("BackScroll", m_def_BackScroll)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)

    m_HasLinesAtRoot = PropBag.ReadProperty("HasLinesAtRoot", m_def_HasLinesAtRoot)
    m_HasLines = PropBag.ReadProperty("HasLines", m_def_HasLines)
    m_HasButtons = PropBag.ReadProperty("HasButtons", m_def_HasButtons)
    m_ButtonSet = PropBag.ReadProperty("ButtonSet", m_def_ButtonSet)

    m_Indent = PropBag.ReadProperty("Indent", m_def_Indent)
    m_ToolTips = PropBag.ReadProperty("Tooltips", m_def_Tooltips)
    
    Appearance = PropBag.ReadProperty("Appearance", 1)
    CheckBoxes = PropBag.ReadProperty("Checkboxes", False)
    Enabled = PropBag.ReadProperty("Enabled", True)
    HideSelection = PropBag.ReadProperty("HideSelection", False)
    HotTracking = PropBag.ReadProperty("HotTracking", False)
    FullRowSelect = PropBag.ReadProperty("FullRowSelect", False)
    LabelEdit = PropBag.ReadProperty("LabelEdit", 1)
    OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
    OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    Scroll = PropBag.ReadProperty("Scroll", True)
    
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    Set ImageList = PropBag.ReadProperty("ImageList", Nothing)
    
    ' If Background = fvGrdntRectVer or fvGrdntTri
    ' or fvPicturedTiled (BackScroll = True) then the
    ' control does not show Tooltips
    Select Case m_Background
        Case fvgNone, fvGrdntRectVer, fvGrdntTri
            bStopPaint = False
        Case fvColor, fvGrdntRectHor
            bStopPaint = True
        Case fvPicturedTiled
            If m_BackScroll Then
                bStopPaint = True
                Else
                bStopPaint = False
            End If
    End Select
    
    If Not m_Picture Is Nothing Then
        GetBitmapIntoDC
    End If

    If m_HasLinesAtRoot Then
        TV.LineStyle = tvwRootLines
        Else
        TV.LineStyle = tvwTreeLines
    End If
    
    If m_ButtonSet = fvCustom Then
        ' Turn off the NORMAL buttons
        bCustButtons = True
        Else
        bCustButtons = False
    End If

    ' HasLines = True
    If m_HasLines Then
        ' HasButtons = True
        If m_HasButtons Then
            ' ButtonSet = fvCustom
            If bCustButtons Then
                TV.Style = tvwTreelinesPictureText
                ' ButtonSet = fvNormal
                Else
                TV.Style = tvwTreelinesPlusMinusPictureText
            End If
            ' HasButtons = False
            Else
            TV.Style = tvwTreelinesPictureText
        End If
        ' HasLines = False
        Else
        ' HasButtons = True
        If m_HasButtons Then
            ' ButtonSet = fvCustom
            If bCustButtons Then
                TV.Style = tvwPictureText
                ' ButtonSet = fvNormal
                Else
                TV.Style = tvwPlusPictureText
            End If
            ' HasButtons = False
            Else
            TV.Style = tvwPictureText
        End If
    End If

    ' Get the TV style (For Tooltips)
    Style = GetWindowLong(TVhwnd, GWL_STYLE)
  
    ' Remove the Tooltips from the treeview because they are bugy
    Call SetWindowLong(TVhwnd, GWL_STYLE, Style Or TVS_NOTOOLTIPS)

    ' If we want tooltips we will set our own
    If m_ToolTips > fvtNone Then
        If TTStd Is Nothing Then
            Set TTStd = New cToolTip
            ' Initialize TTstd tooltip object
            With TTStd
                '.BkColor = &H0&
                '.TxtColor = &HFFFF&
                .DelayTime = 500
                .VisibleTime = 3000
                .TipWidth = 300
                .Style = ttStyleStandard
            End With
        End If
    End If
    
    ' If in Design mode then Unsubclass
    If Ambient.UserMode = False Then Unsubclass

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Appearance", TV.Appearance, 1)
    Call PropBag.WriteProperty("Background", m_Background, m_def_Background)
    Call PropBag.WriteProperty("BackScroll", m_BackScroll, m_def_BackScroll)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    
    Call PropBag.WriteProperty("HasLinesAtRoot", m_HasLinesAtRoot, m_def_HasLinesAtRoot)
    Call PropBag.WriteProperty("HasLines", m_HasLines, m_def_HasLines)
    Call PropBag.WriteProperty("HasButtons", m_HasButtons, m_def_HasButtons)
    Call PropBag.WriteProperty("ButtonSet", m_ButtonSet, m_def_ButtonSet)
    
    Call PropBag.WriteProperty("Checkboxes", TV.CheckBoxes, False)
    Call PropBag.WriteProperty("Indent", m_Indent, m_def_Indent)
    Call PropBag.WriteProperty("Tooltips", m_ToolTips, m_def_Tooltips)
    
    Call PropBag.WriteProperty("HideSelection", TV.HideSelection, False)
    Call PropBag.WriteProperty("HotTracking", TV.HotTracking, False)
    Call PropBag.WriteProperty("FullRowSelect", TV.FullRowSelect, False)
    Call PropBag.WriteProperty("LabelEdit", TV.LabelEdit, 1)
    Call PropBag.WriteProperty("OLEDragMode", TV.OLEDragMode, 0)
    Call PropBag.WriteProperty("OLEDropMode", TV.OLEDropMode, 0)
    Call PropBag.WriteProperty("Scroll", TV.Scroll, True)
    
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("ImageList", ImageList, Nothing)

End Sub

Private Sub UserControl_Terminate()

  Dim i As Long

    ' Unsubclass the TreeView
    Call Unsubclass
 
    ' Clear up the objects
    Set m_Picture = Nothing
    Set TTStd = Nothing
    
    For i = 1 To 3
        ClearMemDC i
    Next i

End Sub

Private Sub UserControl_Resize()

  On Error Resume Next
    
    TV.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    InvalidateRect TVhwnd, 0, 0

End Sub




























'########################## EVENTS ##########################

Private Sub TV_Click()
    
    RaiseEvent Click

End Sub

Private Sub TV_DblClick()
    
    RaiseEvent DblClick

End Sub

Private Sub TV_KeyDown(KeyCode As Integer, Shift As Integer)
    
    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub TV_KeyPress(KeyAscii As Integer)
    
    RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub TV_KeyUp(KeyCode As Integer, Shift As Integer)
    
    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

' If the selected Node has no child Nodes, and the Node's button is clicked,
' remove the Button

' Invoked after the TreeView sends any TVN_GETDISPINFOs, but before
' it sends any NM_CLICK, NM_DBLCLK, TVN_SELCHANGING or
' TVN_ITEMEXPANDING.
Private Sub TV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    RaiseEvent MouseDown(Button, Shift, x, y)
    
    ' Exit if the control uses Checkboxes
    If TV.CheckBoxes Then Exit Sub
  
  Dim TVHTI As TVHITTESTINFO
  Dim Nod As Node
  
    If (Button = vbLeftButton) Then
        ' Get the item under the cursor (if any).
        TVHTI.pt.x = x / Screen.TwipsPerPixelX
        TVHTI.pt.y = y / Screen.TwipsPerPixelY
        ' If there is an Item
        If TreeView_HitTest(TVhwnd, TVHTI) Then
            ' If the item's state icon was left-clicked...
            If (TVHTI.flags = TVHT_ONITEMSTATEICON) Then
                ' Get the Node under the cursor, and toggle its expanded state.
                Set Nod = TV.HitTest(x, y)
                If (Nod Is Nothing) = False Then
                    Nod.Expanded = Not Nod.Expanded
                End If
            End If
        End If
    End If

End Sub

Private Sub TV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  Dim TVHTI As TVHITTESTINFO
  Dim Nod As Node
  Dim RCItem As RECT
    
    RaiseEvent MouseMove(Button, Shift, x, y)
    
'------------------- Tooltips -------------------------
    ' If we want Tooltips we already have set one
    If Not TTStd Is Nothing Then
        ' Get the item under the cursor (if any).
        TVHTI.pt.x = x / Screen.TwipsPerPixelX
        TVHTI.pt.y = y / Screen.TwipsPerPixelY

        ' If the mouse is over an Item
        If TreeView_HitTest(TVhwnd, TVHTI) Then
            ' Get the Node under the cursor
            Set Nod = TV.HitTest(x, y)
            ' If we got the Node
            If (Nod Is Nothing) = False Then
                ' If this is not the first time we set a tooltip
                If (prevItem Is Nothing) = False Then
                    ' If it is different than prevItem
                    If Not prevItem = Nod Then
                        ' Set the Tooltip
                        TTStd.DelToolTip TVhwnd
                        ' If we want the Name
                        If m_ToolTips = fvName Then
                            TTStd.SetToolTip Nod.Text, False
                            ' If we want the Tag
                            Else
                            TTStd.SetToolTip Nod.Tag, False
                        End If
                        Set prevItem = Nod
                    End If
                    ' If this is the first time we set a tooltip
                    Else
                    ' Set the Tooltip
                    TTStd.DelToolTip TVhwnd
                    ' If we want the Name
                    If m_ToolTips = fvName Then
                        TTStd.SetToolTip Nod.Text, False
                        ' If we want the Tag
                        Else
                        TTStd.SetToolTip Nod.Tag, False
                    End If
                    Set prevItem = Nod
                End If
                ' Mouse is moving over the rest area
                Else
                Set prevItem = Nothing
                TTStd.DelToolTip TVhwnd
            End If
            ' Mouse is moving over the rest area
            Else
            Set prevItem = Nothing
            TTStd.DelToolTip TVhwnd
        End If
    End If
'--------------------------------------------------------

End Sub

Private Sub TV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    RaiseEvent MouseUp(Button, Shift, x, y)

End Sub

Private Sub TV_NodeClick(ByVal Node As Node)

    RaiseEvent NodeClick(Node)
    
End Sub

Private Sub TV_Collapse(ByVal Node As Node)
    
    RaiseEvent Collapse(Node)

    ' If we want Custom buttons
    If m_HasButtons And bCustButtons Then
        Call SetTVItemStateImage(TVhwnd, Node, tvisCollapsed)
    End If

End Sub

Private Sub TV_Expand(ByVal Node As Node)
    
    RaiseEvent Expand(Node)
    
    ' If we want Custom buttons
    If m_HasButtons And bCustButtons Then
        Call SetTVItemStateImage(TVhwnd, Node, tvisExpanded)
    End If

End Sub

Private Sub TV_NodeCheck(ByVal Node As Node)
    
    RaiseEvent NodeCheck(Node)

End Sub




























'######################### PROPERTIES #######################

Public Property Get hWnd() As Long

    hWnd = TVhwnd

End Property

Public Property Get Nodes() As Nodes
   
   Set Nodes = TV.Nodes

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,Appearance
Public Property Get Appearance() As AppearanceConstants
    
    Appearance = TV.Appearance

End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    
    TV.Appearance() = New_Appearance
    PropertyChanged "Appearance"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=31,0,0,0
Public Property Get Background() As efvBackground
    
    Background = m_Background

End Property

Public Property Let Background(ByVal New_Background As efvBackground)
    
    ' If m_Background = fvGrdntRectVer or fvGrdntTri
    ' or fvPicturedTiled (m_BackScroll = True) then
    ' does not show Tooltips
    
    ' Set the bStopPaint Flag
    Select Case New_Background
        ' Scroll OK, Tooltips OK
        Case fvgNone, fvColor, fvGrdntRectHor
            m_Background = New_Background
            m_BackScroll = False
            bStopPaint = False
        ' Scroll ONE, Tooltips NO
        Case fvGrdntRectVer, fvGrdntTri
            m_Background = New_Background
            m_BackScroll = False
            m_ToolTips = fvtNone
            bStopPaint = False
            MsgBox "Tooltips = ftvNone and BackScroll = False.", vbCritical, "Properties changed"
        Case fvPicturedTiled
            If m_Picture Is Nothing Then
                Beep
                MsgBox "Please select a Picture first!", vbCritical, "Invalid Picture"
                Else
                m_Background = New_Background
                ' Scroll ONE, Tooltips NO
                If m_BackScroll = False Then
                    m_ToolTips = fvtNone
                    bStopPaint = False
                    MsgBox "Tooltips = ftvNone. Set BackScroll= True for Tooltips.", vbCritical, "Properties changed"
                    ' Scroll OK, Tooltips OK
                    Else
                    bStopPaint = True
                End If
            End If
    End Select
                
    PropertyChanged "Background"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get BackScroll() As Boolean
    
    BackScroll = m_BackScroll

End Property

Public Property Let BackScroll(ByVal New_BackScroll As Boolean)
    
    m_BackScroll = New_BackScroll
    ' Allow Scroll only if Background = fvPicturedTiled
    If Background <> fvPicturedTiled Then
        m_BackScroll = False
        ' Scroll OK, Tooltips OK
        If m_BackScroll Then
            bStopPaint = True
            ' Scroll ONE, Tooltips NO
            Else
            bStopPaint = False
        End If
    End If
    PropertyChanged "BackScroll"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR

    BackColor = m_BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

    m_BackColor = New_BackColor
    If m_Background = fvColor Then
        PrepareRect TV
        SetBackColor memDC(2).MemHDC, m_BackColor
    End If
    PropertyChanged "BackColor"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
    
    ForeColor = m_ForeColor

End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)

  Dim lColor As Long

    m_ForeColor = New_ForeColor

    ' Set the colour in the TreeView:
    lColor = TranslateColor(m_ForeColor)
    SendMessageLong TV.hWnd, TVM_SETTEXTCOLOR, 0, lColor
    ' Request a redraw:
    InvalidateRectAsNull TV.hWnd, 0, 1
    UpdateWindow TV.hWnd

    PropertyChanged "ForeColor"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=29,0,0,0
Public Property Get BorderStyle() As efvBordersStyle
    
    BorderStyle = m_BorderStyle

End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As efvBordersStyle)
    
    m_BorderStyle = New_BorderStyle
    TV.BorderStyle = m_BorderStyle
    PropertyChanged "BorderStyle"

End Property

Public Property Get Enabled() As Boolean
    
    Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal bState As Boolean)

  Dim TVI As TVITEM
  Dim hItem As Long
  Dim lR As Long
  Dim lColor As Long
    
    TVI.mask = TVIF_STATE
    hItem = SendMessageLong(TV.hWnd, TVM_GETNEXTITEM, TVGN_ROOT, 0&)
    Do While hItem <> 0
        With TVI
            .hItem = hItem
            .mask = TVIF_STATE
            .stateMask = TVIS_CUT
            If (bState) Then
                .state = .stateMask And Not TVIS_CUT
                Else
                .state = .stateMask Or TVIS_CUT
            End If
            lR = SendMessage(TV.hWnd, TVM_SETITEM, 0&, TVI)
        End With
        hItem = SendMessageLong(TV.hWnd, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, hItem)
    Loop
    
    ' Set the Forecolor with Send Message because we do not
    ' want to change the ForeColor property
    If (bState) Then
        ' Set the colour in the TreeView:
        lColor = TranslateColor(m_ForeColor)
        SendMessageLong TV.hWnd, TVM_SETTEXTCOLOR, 0, lColor
        ' Request a redraw:
        InvalidateRectAsNull TV.hWnd, 0, 1
        UpdateWindow TV.hWnd
        Else
        ' Set the colour in the TreeView:
        lColor = TranslateColor(vbGrayText)
        SendMessageLong TV.hWnd, TVM_SETTEXTCOLOR, 0, lColor
        ' Request a redraw:
        InvalidateRectAsNull TV.hWnd, 0, 1
        UpdateWindow TV.hWnd
    End If
    
    UserControl.Enabled = bState
    PropertyChanged "Enabled"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get HasLinesAtRoot() As Boolean
    
    HasLinesAtRoot = m_HasLinesAtRoot

End Property

Public Property Let HasLinesAtRoot(ByVal New_HasLinesAtRoot As Boolean)
    
    m_HasLinesAtRoot = New_HasLinesAtRoot
    If m_HasLinesAtRoot Then
        TV.LineStyle = tvwRootLines
        Else
        TV.LineStyle = tvwTreeLines
    End If
    PropertyChanged "HasLinesAtRoot"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get HasLines() As Boolean
    
    HasLines = m_HasLines

End Property

Public Property Let HasLines(ByVal New_HasLines As Boolean)
    
    m_HasLines = New_HasLines
    If m_HasLines Then
        If m_HasButtons Then
            TV.Style = tvwTreelinesPlusMinusPictureText
            Else
            TV.Style = tvwTreelinesPictureText
        End If
        Else
        If m_HasButtons Then
            TV.Style = tvwPlusPictureText
            Else
            TV.Style = tvwPictureText
        End If
    End If
    PropertyChanged "HasLines"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get HasButtons() As Boolean
    
    HasButtons = m_HasButtons

End Property

Public Property Let HasButtons(ByVal New_HasButtons As Boolean)
    
    m_HasButtons = New_HasButtons
    If m_HasLines Then
        If m_HasButtons Then
            TV.Style = tvwTreelinesPlusMinusPictureText
            Else
            TV.Style = tvwTreelinesPictureText
        End If
        Else
        If m_HasButtons Then
            TV.Style = tvwPlusPictureText
            Else
            TV.Style = tvwPictureText
        End If
    End If
    PropertyChanged "HasButtons"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=32,0,0,fvNormal
Public Property Get ButtonSet() As efvButtonSet
    
    ButtonSet = m_ButtonSet

End Property

Public Property Let ButtonSet(ByVal New_ButtonSet As efvButtonSet)
    
    m_ButtonSet = New_ButtonSet
    ' Change the property only if the Treeview Has buttons
    If m_HasButtons = True Then
        ' Normal buttons
        If m_ButtonSet = fvNormal Then
            If m_HasLines Then
                TV.Style = tvwTreelinesPlusMinusPictureText
                Else
                TV.Style = tvwPlusPictureText
            End If
            ' Custom buttons
            Else
            If m_HasLines Then
                TV.Style = tvwTreelinesPictureText
                Else
                TV.Style = tvwPictureText
            End If
        End If
    End If
    
    If m_ButtonSet = fvCustom And CheckBoxes = True Then
        CheckBoxes = False
        MsgBox "Checkboxes = False.", vbCritical, "Property changed"
    End If
    
    ' Set the bCustButtons flag
    If m_ButtonSet = fvNormal Then
        bCustButtons = False
        Else
        bCustButtons = True
    End If
    
    PropertyChanged "ButtonSet"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,CheckBoxes
Public Property Get CheckBoxes() As Boolean
    
    CheckBoxes = TV.CheckBoxes

End Property

Public Property Let CheckBoxes(ByVal New_CheckBoxes As Boolean)

    TV.CheckBoxes() = New_CheckBoxes
    If New_CheckBoxes = True And m_ButtonSet = fvCustom Then
        m_ButtonSet = fvNormal
        MsgBox "ButtonSet = fvNormal.", vbCritical, "Property changed"
    End If
    PropertyChanged "Checkboxes"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,Indentation
Public Property Get Indent() As Single ' In Pixels

    Indent = Int(TV.Indentation / Screen.TwipsPerPixelX)

End Property

Public Property Let Indent(ByVal New_Indent As Single) ' In Pixels

  On Error GoTo errIndent

    m_Indent = New_Indent
    TV.Indentation() = m_Indent * Screen.TwipsPerPixelX
    PropertyChanged "Indent"

  Exit Property

errIndent:
    Indent = 19
    TV.Indentation() = Indent * Screen.TwipsPerPixelX
    PropertyChanged "Indent"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=32,0,0
Public Property Get ToolTips() As efvTooltips
    
    ToolTips = m_ToolTips

End Property

Public Property Let ToolTips(ByVal New_ToolTips As efvTooltips)
    
    m_ToolTips = New_ToolTips
    
    ' Check if tooltips are allowed
    Select Case m_Background
        Case fvGrdntRectVer, fvGrdntTri
            m_ToolTips = fvtNone
        Case fvPicturedTiled
            If m_BackScroll = False Then m_ToolTips = fvtNone
    End Select
    
    PropertyChanged "Tooltips"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,HideSelection
Public Property Get HideSelection() As Boolean
    
    HideSelection = TV.HideSelection

End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    
    TV.HideSelection() = New_HideSelection
    PropertyChanged "HideSelection"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,HotTracking
Public Property Get HotTracking() As Boolean
    
    HotTracking = TV.HotTracking

End Property

Public Property Let HotTracking(ByVal New_HotTracking As Boolean)
    
    TV.HotTracking() = New_HotTracking
    PropertyChanged "HotTracking"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,FullRowSelect
Public Property Get FullRowSelect() As Boolean
    
    FullRowSelect = TV.FullRowSelect

End Property

Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)
    
    TV.FullRowSelect() = New_FullRowSelect
    PropertyChanged "FullRowSelect"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,LabelEdit
Public Property Get LabelEdit() As LabelEditConstants
    
    LabelEdit = TV.LabelEdit

End Property

Public Property Let LabelEdit(ByVal New_LabelEdit As LabelEditConstants)
    
    TV.LabelEdit() = New_LabelEdit
    PropertyChanged "LabelEdit"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,OLEDragMode
Public Property Get OLEDragMode() As OLEDragConstants
    
    OLEDragMode = TV.OLEDragMode

End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As OLEDragConstants)
    
    TV.OLEDragMode() = New_OLEDragMode
    PropertyChanged "OLEDragMode"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,OLEDropMode
Public Property Get OLEDropMode() As OLEDropConstants
    
    OLEDropMode = TV.OLEDropMode

End Property

' There is a problem if OleDropMode = 2 ???
Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
    
  On Error GoTo OLer
    
    TV.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"

  Exit Property

OLer:

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,Scroll
Public Property Get Scroll() As Boolean
    
    Scroll = TV.Scroll

End Property

Public Property Let Scroll(ByVal New_Scroll As Boolean)
    
    TV.Scroll() = New_Scroll
    PropertyChanged "Scroll"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture

    Set Picture = m_Picture

End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    
    Set m_Picture = New_Picture
    
    ' Load the Picture into a memory DC
    ' and get its dimensions
    If Not New_Picture Is Nothing Then
        GetBitmapIntoDC
        ' Picture is cleared
        Else
        m_Background = fvgNone
        m_BackScroll = False
        MsgBox "Background = fbgNone", vbCritical, "Property changed"
    End If
    
    TV.Refresh
    PropertyChanged "Picture"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=TV,TV,-1,ImageList
Public Property Get ImageList() As Object

    Set ImageList = TV.ImageList

End Property

Public Property Set ImageList(ByVal New_ImageList As Object)

    Set TV.ImageList = New_ImageList
    PropertyChanged "ImageList"

End Property






























'###################### SUBCLASSING #####################

Private Sub Subclass()
    
    Unsubclass

    TVhwnd = TV.hWnd
    UChwnd = UserControl.hWnd

    AttachMessage Me, TVhwnd, TVM_INSERTITEM
    AttachMessage Me, TVhwnd, WM_PAINT
    AttachMessage Me, TVhwnd, WM_ERASEBKGND
    AttachMessage Me, TVhwnd, WM_DESTROY
    AttachMessage Me, UChwnd, WM_NOTIFY

End Sub

Private Sub Unsubclass()

    If TVhwnd <> 0 Then
        DetachMessage Me, TVhwnd, TVM_INSERTITEM
        DetachMessage Me, TVhwnd, WM_PAINT
        DetachMessage Me, TVhwnd, WM_ERASEBKGND
        DetachMessage Me, TVhwnd, WM_DESTROY
        DetachMessage Me, UChwnd, WM_NOTIFY
    End If
    
    TVhwnd = 0
    UChwnd = 0

End Sub

Private Property Let cSubclass_MsgResponse(ByVal RHS As EMsgResponse)

   '

End Property

Private Property Get cSubclass_MsgResponse() As EMsgResponse

    If CurrentMessage = WM_PAINT Or CurrentMessage = WM_ERASEBKGND Or CurrentMessage = WM_DESTROY Or CurrentMessage = TVM_INSERTITEM Then
        cSubclass_MsgResponse = emrConsume
        Else
        cSubclass_MsgResponse = emrPreprocess
    End If

End Property

Private Function cSubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  Dim i As Long
  Dim lR As Long
  Dim bR As Boolean
    
    Select Case iMsg
    
        Case WM_PAINT, WM_ERASEBKGND
            ' Do this only if we have selected a background
            If m_Background = fvgNone Then
                cSubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
              Exit Function
            End If
            If m_Background = fvPicturedTiled And m_Picture Is Nothing Then
                cSubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
              Exit Function
            End If
            ' Custom process the WM_PAINT and WM_ERASEBKGND messages
            TvWMPaint hWnd, iMsg, lR, bR
            If bR Then
                cSubclass_WindowProc = lR
                Else
                cSubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            End If
        
        ' Set the Custom buttons (If any)
        ' This is Not the right way but works.
        ' We should better get a TVINSERTSTRUCT structure,
        ' change the state and stateMask members of its
        ' TVITEMEX or TVITEM member, and then use the
        ' TreeView_InsertItem macro to insert the Item.
        Case TVM_INSERTITEM
          On Error Resume Next
            ' If the control has buttons AND CUSTOM buttons
            If m_HasButtons And m_ButtonSet = fvCustom Then
                ' Normal process
                ' We first insert the item to increase the TV.Nodes.Count
                cSubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
                ' Set the counter to the new(last) Item so
                ' i will be the index of the Node
                i = TV.Nodes.Count
                ' If it has Children
                If TV.Nodes(i).Children > 0 Then
                    Call SetTVItemStateImage(TVhwnd, TV.Nodes(i), tvisCollapsed)
                    ' If it has not Children
                    Else
                    Call SetTVItemStateImage(TVhwnd, TV.Nodes(i), tvisNoButton)
                End If
                ' Normal buttons
                Else
                cSubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            End If
          
        ' Unsubclass the window.
        Case WM_DESTROY
            cSubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            Call Unsubclass
          Exit Function
    End Select

End Function
































'####################### FUNCTIONS ########################

Private Sub MakeMemDC(ByVal SourceDC As Long, _
                            ByVal MemDCIndex As Long, _
                            ByVal MemDCWidth As Long, _
                            ByVal MemDCHeight As Long)
      
    With memDC(MemDCIndex)
        If MemDCWidth > .MemWidth Or MemDCHeight > .MemHeight Or .MemHDC = 0 Then
            ClearMemDC MemDCIndex
            ' Make a memory DC like the Treeview DC
            .MemHDC = CreateCompatibleDC(SourceDC)
            If .MemHDC <> 0 Then
                If SourceDC = 0 Then SourceDC = .MemHDC
                ' Specify the height, width, and color for the TempDC
                ' according to the Treeview DC
                .MemBmp = CreateCompatibleBitmap(SourceDC, MemDCWidth, MemDCHeight)
                ' If the Properties of the MemoryDC was set OK
                If .MemBmp <> 0 Then
                    ' Select it
                    .MemBmpOld = SelectObject(.MemHDC, .MemBmp)
                  Dim memDcRect As RECT
                  Dim hBrush As Long
                    ' Dimension the memDcRect
                    memDcRect.Right = MemDCWidth
                    memDcRect.Bottom = MemDCHeight
                    ' Make a Brush
                    hBrush = CreateSolidBrush(TranslateColor(vbWindowBackground))
                    ' and paint it vbWindowBackground (initialize)
                    FillRect .MemHDC, memDcRect, hBrush
                    ' Clear the objects
                    DeleteObject hBrush
                    ' If the Properties of the MemoryDC was not set OK
                    Else
                    ' Clear the objects
                    ClearMemDC MemDCIndex
                End If
            End If
        End If
    End With
   
End Sub

' Clear the memory DC (memDC)
Private Sub ClearMemDC(ByVal MemDCIndex As Long)

   With memDC(MemDCIndex)
      If .MemBmpOld <> 0 Then
         SelectObject .MemHDC, .MemBmpOld
      End If
      If .MemBmp <> 0 Then
         DeleteObject .MemBmp
      End If
      If .MemHDC <> 0 Then
         DeleteDC .MemHDC
      End If
   End With

End Sub
    
' Make a DC to hold the picture bitmap which we can blt from
Private Function GetBitmapIntoDC() As Boolean

  Dim lHwnd As Long
  Dim lHDC As Long
  
  Dim lHDCTemp As Long
  Dim BMP As BITMAP
  Dim lHBmpTempOld As Long
  Dim m_lHBmp As Long
  Dim m_lHBmpOld As Long

    ' Get the Desktop handle
    lHwnd = GetDesktopWindow()
    ' Get the DC of the Desktop
    lHDC = GetDC(lHwnd)
    
    ' Create 2 DCs into the memory
    m_lHdc = CreateCompatibleDC(lHDC)
    lHDCTemp = CreateCompatibleDC(lHDC)
    
    ' If the first DC has been created succesfully
    If (m_lHdc <> 0) Then
        GetObjectAPI m_Picture.Handle, LenB(BMP), BMP
        ' Get the size of the bitmap
        lBitmapW = BMP.bmWidth
        lBitmapH = BMP.bmHeight
        lHBmpTempOld = SelectObject(lHDCTemp, m_Picture.Handle)
        m_lHBmp = CreateCompatibleBitmap(lHDC, lBitmapW, lBitmapH)
        m_lHBmpOld = SelectObject(m_lHdc, m_lHBmp)
        BitBlt m_lHdc, 0, 0, lBitmapW, lBitmapH, lHDCTemp, 0, 0, vbSrcCopy
        SelectObject lHDCTemp, lHBmpTempOld
        DeleteDC lHDCTemp
        If (m_lHBmpOld <> 0) Then
            GetBitmapIntoDC = True
            Else
            pErr 2, "Unable to select bitmap into DC"
        End If
        Else
        pErr 1, "Unable to create compatible DC"
    End If

    ReleaseDC lHwnd, lHDC

End Function

Private Sub Tile(ByRef hdc As Long, _
                        ByVal x As Long, _
                        ByVal y As Long, _
                        ByVal Width As Long, _
                        ByVal Height As Long)
                        
  Dim lSrcX As Long
  Dim lSrcY As Long
  Dim lSrcStartX As Long
  Dim lSrcStartY As Long
  Dim lSrcStartWidth As Long
  Dim lSrcStartHeight As Long
  Dim lDstX As Long
  Dim lDstY As Long
  Dim lDstWidth As Long
  Dim lDstHeight As Long

    If m_Picture Is Nothing Then Exit Sub
    If lBitmapW = 0 Or lBitmapH = 0 Then Exit Sub

    lSrcStartX = ((x + XOriginOffset) Mod lBitmapW)
    lSrcStartY = ((y + YOriginOffset) Mod lBitmapH)
    lSrcStartWidth = (lBitmapW - lSrcStartX)
    lSrcStartHeight = (lBitmapH - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    
    lDstY = y
    lDstHeight = lSrcStartHeight
    
    Do While lDstY < (y + Height)
        If (lDstY + lDstHeight) > (y + Height) Then
            lDstHeight = y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = x
        lSrcX = lSrcStartX
        Do While lDstX < (x + Width)
            If (lDstX + lDstWidth) > (x + Width) Then
                lDstWidth = x + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            ' Copy the image stored to the Memory DC (m_lHdc)
            ' to the desired coordinates on hdc
            BitBlt hdc, lDstX, lDstY, lDstWidth, lDstHeight, m_lHdc, lSrcX, lSrcY, vbSrcCopy
            ' Move next X
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = lBitmapW
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = lBitmapH
    Loop

End Sub

Private Sub pErr(lNumber As Long, smsg As String)

    MsgBox "Error: " & smsg & ", " & lNumber, vbExclamation

End Sub

' Paint the memDC(2) with the selected lColor
Private Sub SetBackColor(hdc As Long, ByVal bgColor As OLE_COLOR)

  Dim lColor As Long
  Dim sBr As Long

    lColor = TranslateColor(bgColor)
    sBr = CreateSolidBrush(lColor)
    FillRect hdc, RC, sBr
    DeleteObject sBr

End Sub

Private Sub PrepareRect(obj As Object)

    GetClientRect obj.hWnd, RC
    ' Store the dimensions to variables
    With RC
        TVWidth = .Right - .Left
        TVHeight = .Bottom - .Top
    End With

End Sub

Private Sub PrepareVertexRect(obj As Object)

    GetClientRect obj.hWnd, RC
    ' Store the dimensions to variables
    With RC
        TVWidth = .Right - .Left
        TVHeight = .Bottom - .Top
    End With

    ' Initialize
    With vert(0)
        .x = 0
        .y = 0
        .Red = 0
        .Green = &HFF&
        .Blue = 0
        .Alpha = 0
    End With

    With vert(1)
        .x = TVWidth
        .y = TVHeight
        .Red = 0
        .Green = LongToUShort(&HFF00&)
        .Blue = LongToUShort(&HFF00&)
        .Alpha = 0
    End With

    gRect.UpperLeft = 1
    gRect.LowerRight = 0

End Sub

Private Sub PrepareVertexTri(obj As Object)

    GetClientRect obj.hWnd, RC
    ' Store the dimensions to variables
    With RC
        TVWidth = .Right - .Left
        TVHeight = .Bottom - .Top
    End With

    ' Initialize
    With vert(0)
        .x = 0
        .y = 0
        .Red = 0&
        .Green = LongToUShort(&HFF00&) '0
        .Blue = 0&
        .Alpha = 0&
    End With
    With vert(1)
        .x = TVWidth
        .y = 0
        .Red = 0 'LongToUShort(&HFF00&)
        .Green = 0&
        .Blue = LongToUShort(&HFF00&)
        .Alpha = 0&
    End With
    With vert(2)
        .x = TVWidth
    '    .x = Me.ScaleWidth
        .y = TVHeight
        .Red = LongToUShort(&HFF00&)
        .Green = 0&
        .Blue = 0 'LongToUShort(&HFF00&)
        .Alpha = 0&
    End With
    With vert(3)
        .x = 0
        .y = TVHeight
        .Red = 0 'LongToUShort(&HFF00&)
        .Green = LongToUShort(&HFF00&)
        .Blue = LongToUShort(&HFF00&)
        .Alpha = 0&
    End With

    gTri(0).Vertex1 = 0
    gTri(0).Vertex2 = 1
    gTri(0).Vertex3 = 2

    gTri(1).Vertex1 = 0
    gTri(1).Vertex2 = 2
    gTri(1).Vertex3 = 3

End Sub

Private Function LongToUShort(ulong As Long) As Integer

   LongToUShort = CInt(ulong - &H10000)

End Function

Private Sub TvWMPaint(ByVal hWnd As Long, _
                            ByVal wMsg As Long, _
                            RetVal As Long, _
                            UseRetVal As Boolean)

  'Prevent recursion with this variable
  Static bPainting As Boolean

  Dim ps As PAINTSTRUCT
  Dim TvDc As Long
  
  Dim hDCC As Long
  Dim rectPS As RECT
  Dim LPTR As Long
      
  Dim rcFirstItem As RECT
  Dim hItemFirst As Long
    
    Select Case wMsg
        
        Case WM_PAINT
        
            ' If working do not disturb
            If bPainting = True Then Exit Sub
            ' Working...
            bPainting = True
    
            ' Get the coordinates of the TV rect
            GetClientRect hWnd, RC
            
            ' Store the dimensions to variables
            With RC
                TVWidth = .Right - .Left
                TVHeight = .Bottom - .Top
            End With
    
            'Prepare a DC for painting
            BeginPaint hWnd, ps
    
            ' Get the Treeview DC
            TvDc = ps.hdc
    
            ' Copy the RECT coordinates of the prepaint surface (visible)
            LSet rectPS = ps.rcPaint
           
            ' Create a MemoryDCs 1 for the Treeview CONTENTS
            MakeMemDC TvDc, 1, TVWidth, TVHeight
            ' Create a MemoryDCs 2 for the Treeview BACKGROUND
            MakeMemDC TvDc, 2, TVWidth, TVHeight
            ' Create a MemoryDC 3, Monochrome (2 colors- Back/Fore)
            ' with the size of the painted area
            MakeMemDC 0, 3, rectPS.Right - rectPS.Left, rectPS.Bottom - rectPS.Top
            
            ' Paint the TreeView CONTENTS on the MemoryDC 1
            CallOldWindowProc hWnd, WM_PAINT, memDC(1).MemHDC, 0&
    
            ' Paint the BACKGROUND on the MemoryDC 2
'---------------------------------------------------------------
            
            Select Case m_Background
                Case fvgNone
                    
                Case fvColor
                    PrepareRect TV
                    SetBackColor memDC(2).MemHDC, m_BackColor
                Case fvGrdntRectHor
                    PrepareVertexRect TV
                    ' Gradient Rectangle Horizontally
                    GradientFillRect memDC(2).MemHDC, vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_H
                Case fvGrdntRectVer
                    PrepareVertexRect TV
                    ' Gradient Rectangle Vertically
                    GradientFillRect memDC(2).MemHDC, vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_V
                Case fvGrdntTri
                    PrepareVertexTri TV
                    ' Gradient Triangle
                    GradientFillTri memDC(2).MemHDC, vert(0), 4, gTri(0), 2, GRADIENT_FILL_TRIANGLE
                ' Tile the picture
                Case fvPicturedTiled
                    If m_BackScroll = True Then
                        ' Get the handle to the Root item
                        hItemFirst = TreeView_GetNextItem(hWnd, TVGN_ROOT, 0)
                        ' Get the RECT of the Root item
                        TreeView_GetItemRect hWnd, hItemFirst, rcFirstItem, CTrue
                        If rcFirstItem.Left > m_lXOffset Then m_lXOffset = rcFirstItem.Left
                        ' If the first item's rectangle has been moved
                        ' by scrolling horizontaly or verticaly set the new
                        ' coordinates of the paint area
                        XOriginOffset = -rcFirstItem.Left + m_lXOffset
                        YOriginOffset = -rcFirstItem.Top
                        Else
                        XOriginOffset = 0
                        YOriginOffset = 0
                    End If
                    Tile memDC(2).MemHDC, rectPS.Left, rectPS.Top, rectPS.Right - rectPS.Left, rectPS.Bottom - rectPS.Top
            End Select
         
'---------------------------------------------------------------

            ' Set BkColor of the MemoryDC 1 to match transparent colour
            SetBkColor memDC(1).MemHDC, TranslateColor(vbWindowBackground)
             
            ' Copy from MemoryDC 1 (CONTENTS) to MemoryDC 3 (monochrome)
            ' When bitblt'ing from color to monochrome, Windows sets to 1
            ' all pixels that match the background color of the source DC.
            ' All other bits are set to 0.
            BitBlt memDC(3).MemHDC, 0, 0, rectPS.Right - rectPS.Left, rectPS.Bottom - rectPS.Top, memDC(1).MemHDC, rectPS.Left, rectPS.Top, vbSrcCopy
            ' Now the MemoryDC 3 is a mask of the Treeview CONTENTS
            ' colored with the Textcolor of the DC (vbButtonText)
    
            SetTextColor memDC(2).MemHDC, vbBlack
            SetBkColor memDC(2).MemHDC, vbWhite
            
            ' AND the mask to the Background so we go white where the
            ' treeview is black and make a BLACK hole of the
            ' Treeview CONTENTS on the MemoryDC 2
            BitBlt memDC(2).MemHDC, rectPS.Left, rectPS.Top, rectPS.Right - rectPS.Left, rectPS.Bottom - rectPS.Top, memDC(3).MemHDC, 0, 0, vbSrcAnd
    
            SetTextColor memDC(1).MemHDC, vbBlack
            SetBkColor memDC(1).MemHDC, vbWhite
           
            ' Copy from the MemoryDC 3 to the MemoryDC 1 Dsna
            ' What we want here is black at the transparent color, and
            ' the original colors everywhere else.  To do this, we first
            ' paint the original onto the cover (which we already did), then we
            ' AND the inverse of the mask onto that using the DSna ternary raster
            ' operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
            ' Operation Codes", "Ternary Raster Operations", or search in MSDN
            ' for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".
            '
            ' When bitblt'ing from monochrome to color, Windows transforms all white
            ' bits (1) to the background color of the destination hdc. All black (0)
            ' bits are transformed to the foreground color.
            BitBlt memDC(1).MemHDC, rectPS.Left, rectPS.Top, rectPS.Right - rectPS.Left, rectPS.Bottom - rectPS.Top, memDC(3).MemHDC, 0, 0, DSna
    
            ' Paint the CONTENTS from the MemoryDC 1 to the MemoryDC 2
            ' and now the Memory DC 2 contains ALL the Treeview
            BitBlt memDC(2).MemHDC, rectPS.Left, rectPS.Top, rectPS.Right - rectPS.Left, rectPS.Bottom - rectPS.Top, memDC(1).MemHDC, rectPS.Left, rectPS.Top, vbSrcPaint
            
            'Draw ALL to the Treeview DC
            BitBlt TvDc, rectPS.Left, rectPS.Top, rectPS.Right - rectPS.Left, rectPS.Bottom - rectPS.Top, _
                        memDC(2).MemHDC, rectPS.Left, rectPS.Top, vbSrcCopy
            
            EndPaint hWnd, ps
            
            ' If an Action does not paint correctly
            ' we do not InvalidateRect (this solves almost all problems)
            If bStopPaint = False Then
                ' Refreshes OK but
                ' Scrolls one page or line each time
                InvalidateRect TVhwnd, 0, 0
            End If
            RetVal = 0
            UseRetVal = True
            bPainting = False
        
        Case WM_ERASEBKGND
        
            'Return TRUE
            RetVal = 1
            UseRetVal = True
           ' Scrolls OK with No animation but
           ' does not Refreshes OK when dragging the thumbtrack
           ' with complex backgrounds
            If bStopPaint = True Then
                InvalidateRect TVhwnd, 0, 0
            End If
    End Select

End Sub
