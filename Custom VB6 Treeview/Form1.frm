VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Treeview"
   ClientHeight    =   5220
   ClientLeft      =   1320
   ClientTop       =   1380
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   3045
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ilsState 
      Left            =   2400
      Top             =   4515
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilSmallIcons 
      Left            =   1740
      Top             =   4515
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VBTreeview.icTreeview1 icTreeview1 
      Height          =   5130
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   9049
      Background      =   2
      BackColor       =   4194304
      ForeColor       =   8454143
      ButtonSet       =   1
      HideSelection   =   -1  'True
      Picture         =   "Form1.frx":0000
   End
   Begin VB.Menu q 
      Caption         =   "q"
      Visible         =   0   'False
      Begin VB.Menu qq 
         Caption         =   "qq"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'#########################################################

' Treeview.ocx Demo
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
' http://vbaccelerator.com/ (sSubTimer code is used)

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

Option Explicit

Private Sub Form_Load()

  Dim i
  Dim x
  Dim C
  Dim Nod As Node
  Dim NodX As Node
  
    ' Initialize the normal small icon Imagelist
    With ilSmallIcons
        .ImageWidth = 16
        .ImageHeight = 16
        .ListImages.Add , , LoadPicture("Icons\cfold16.ico")
        .ListImages.Add , , LoadPicture("Icons\ofold16.ico")
        .ListImages.Add , , LoadPicture("Icons\smpostit.ico")
        .ListImages.Add , , LoadPicture("Icons\grpostit.ico")
    End With

    ' Initialize the State icon Imagelist
    With ilsState
        .ImageWidth = 16
        .ImageHeight = 16
        .ListImages.Add , , LoadPicture("Icons\Nobutton.ico")
        .ListImages.Add , , LoadPicture("Icons\Collapsed.ico")
        .ListImages.Add , , LoadPicture("Icons\Expanded.ico")
    End With

    Set icTreeview1.ImageList = ilSmallIcons
  
    ' Assign the ilsState ImageList as the TreeView's STATE imagelist.
    Call TreeView_SetImageList(icTreeview1.hWnd, ilsState.hImageList, TVSIL_STATE)
    
    For i = 1 To 10
        Set Nod = icTreeview1.Nodes.Add(, , , "Sample Node " & i, 1, 2)
        C = C + 1
        ' This is used for Tooltips
        Nod.Tag = "SampleNode " & C
        For x = 1 To 5
           Set NodX = icTreeview1.Nodes.Add(Nod, TVGN_CHILD, , "Sample Node " & x, 3, 4)
           C = C + 1
           ' This is used for Tooltips
           NodX.Tag = "SampleNode " & C
        Next x
    Next i

End Sub

Private Sub Form_Resize()

  On Error Resume Next
    
    icTreeview1.Move 0, 0, Me.Width - 200, Me.Height - 500

End Sub

Private Sub icTreeview1_NodeClick(ByVal Node As MSComctlLib.Node)

   Caption = Node.Text

End Sub

