Attribute VB_Name = "mSubclass"
Option Explicit

Private Const GWL_WNDPROC = (-4)

Private m_iCurrentMessage As Long
Private m_iProcOld As Long

Public Enum EErrorWindowProc
    eeBaseWindowProc = 13080 ' WindowProc
    eeCantSubclass           ' Can't subclass window
    eeAlreadyAttached        ' Message already handled by another class
    eeInvalidWindow          ' Invalid window
    eeNoExternalWindow       ' Can't modify external window
End Enum

Private Declare Function IsWindow Lib "user32" _
                            (ByVal hWnd As Long) As Long
                            
Private Declare Function GetProp Lib "user32" Alias "GetPropA" _
                            (ByVal hWnd As Long, _
                             ByVal lpString As String) As Long
                             
Private Declare Function SetProp Lib "user32" Alias "SetPropA" _
                            (ByVal hWnd As Long, _
                             ByVal lpString As String, _
                             ByVal hData As Long) As Long
                             
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" _
                            (ByVal hWnd As Long, _
                             ByVal lpString As String) As Long
                             
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
                            (ByVal lpPrevWndFunc As Long, _
                             ByVal hWnd As Long, _
                             ByVal Msg As Long, _
                             ByVal wParam As Long, _
                             ByVal lParam As Long) As Long
                             
Private Declare Function GetWindowThreadProcessId Lib "user32" _
                            (ByVal hWnd As Long, _
                             lpdwProcessId As Long) As Long
                             
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

' Called from WindowProc
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                            (pDest As Any, _
                             pSrc As Any, _
                             ByVal ByteLen As Long)

' Called from cSubclass_MsgResponse(UserControl)
Public Property Get CurrentMessage() As Long
   
   CurrentMessage = m_iCurrentMessage

End Property

' Called from Subclass(UserControl)
Public Sub AttachMessage(iwp As cSubclass, ByVal hWnd As Long, _
                        ByVal iMsg As Long)
    
  Dim procOld As Long
  Dim F As Long
  Dim C As Long
  Dim iC As Long
  Dim bFail As Boolean
    
    ' Validate window
    If IsWindow(hWnd) = False Then ErrRaise eeInvalidWindow
    If IsWindowLocal(hWnd) = False Then ErrRaise eeNoExternalWindow

    ' Get the message count
    C = GetProp(hWnd, "C" & hWnd)
    If C = 0 Then
        ' Subclass window by installing window procecure
        procOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
        If procOld = 0 Then ErrRaise eeCantSubclass
        ' Associate old procedure with handle
        F = SetProp(hWnd, hWnd, procOld)
        Debug.Assert F <> 0
        ' Count this message
        C = 1
        F = SetProp(hWnd, "C" & hWnd, C)
        Else
        ' Count this message
        C = C + 1
        F = SetProp(hWnd, "C" & hWnd, C)
    End If
    Debug.Assert F <> 0
    
    ' SPM - in this version I am allowing more than one class to
    ' make a subclass to the same hWnd and Msg.  Why am I doing
    ' this?  Well say the class in question is a control, and it
    ' wants to subclass its container.  In this case, we want
    ' all instances of the control on the form to receive the
    ' form notification message.
    C = GetProp(hWnd, hWnd & "#" & iMsg & "C")
    If (C > 0) Then
        For iC = 1 To C
            If (GetProp(hWnd, hWnd & "#" & iMsg & "#" & iC) = ObjPtr(iwp)) Then
                ErrRaise eeAlreadyAttached
                bFail = True
              Exit For
            End If
        Next iC
    End If
    If Not (bFail) Then
        C = C + 1
        ' Increase count for hWnd/Msg:
        F = SetProp(hWnd, hWnd & "#" & iMsg & "C", C)
        Debug.Assert F <> 0
        ' Associate object with message at the count:
        F = SetProp(hWnd, hWnd & "#" & iMsg & "#" & C, ObjPtr(iwp))
        Debug.Assert F <> 0
    End If

End Sub

' Called from Unsubclass(UserControl)
Public Sub DetachMessage(iwp As cSubclass, ByVal hWnd As Long, _
                        ByVal iMsg As Long)
    
  Dim procOld As Long
  Dim F As Long
  Dim C As Long
  Dim iC As Long
  Dim iP As Long
  Dim LPTR As Long
    
    ' Get the message count
    C = GetProp(hWnd, "C" & hWnd)
    If C = 1 Then
        ' This is the last message, so unsubclass
        procOld = GetProp(hWnd, hWnd)
        Debug.Assert procOld <> 0
        ' Unsubclass by reassigning old window procedure
        Call SetWindowLong(hWnd, GWL_WNDPROC, procOld)
        ' Remove unneeded handle (oldProc)
        RemoveProp hWnd, hWnd
        ' Remove unneeded count
        RemoveProp hWnd, "C" & hWnd
        Else
        ' Uncount this message
        C = GetProp(hWnd, "C" & hWnd)
        C = C - 1
        F = SetProp(hWnd, "C" & hWnd, C)
    End If
    
    ' SPM - in this version I am allowing more than one class to
    ' make a subclass to the same hWnd and Msg.  Why am I doing
    ' this?  Well say the class in question is a control, and it
    ' wants to subclass its container.  In this case, we want
    ' all instances of the control on the form to receive the
    ' form notification message.
    
    ' How many instances attached to this hwnd/msg?
    C = GetProp(hWnd, hWnd & "#" & iMsg & "C")
    If (C > 0) Then
        ' Find this iwp object amongst the items:
        For iC = 1 To C
            If (GetProp(hWnd, hWnd & "#" & iMsg & "#" & iC) = ObjPtr(iwp)) Then
                iP = iC
              Exit For
            End If
        Next iC
        If (iP <> 0) Then
             ' Remove this item:
             For iC = iP + 1 To C
                LPTR = GetProp(hWnd, hWnd & "#" & iMsg & "#" & iC)
                SetProp hWnd, hWnd & "#" & iMsg & "#" & (iC - 1), LPTR
             Next iC
        End If
        ' Decrement the count
        RemoveProp hWnd, hWnd & "#" & iMsg & "#" & C
        C = C - 1
        SetProp hWnd, hWnd & "#" & iMsg & "C", C
    End If

End Sub

' Called from cSubclass_WindowProc(UserControl)
Private Function WindowProc(ByVal hWnd As Long, _
                            ByVal iMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long
    
  Dim iwp As cSubclass
  Dim iwpT As cSubclass
  Dim procOld As Long
  Dim pSubclass As Long
  Dim F As Long
  Dim iPC As Long
  Dim iP As Long
  Dim bNoProcess As Long
  Dim bCalled As Boolean
    
    ' Get the old procedure from the window's properties list
    procOld = GetProp(hWnd, hWnd)
    
    ' SPM - in this version I am allowing more than one class to
    ' make a subclass to the same hWnd and Msg.  Why am I doing
    ' this?  Well say the class in question is a control, and it
    ' wants to subclass its container.  In this case, we want
    ' all instances of the control on the form to receive the
    ' form notification message.
    
    ' Get the number of instances for this msg/hwnd:
    bCalled = False
    iPC = GetProp(hWnd, hWnd & "#" & iMsg & "C")
    If (iPC > 0) Then
        ' For each instance attached to this msg/hwnd, call the subclass:
        For iP = 1 To iPC
            bNoProcess = False
            ' Get the object pointer from the message
            pSubclass = GetProp(hWnd, hWnd & "#" & iMsg & "#" & iP)
            If pSubclass = 0 Then
                ' This message is not handled, so pass on to old procedure
                WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                            wParam, ByVal lParam)
                bNoProcess = True
            End If
            If Not (bNoProcess) Then
                ' Turn the pointer into an illegal, uncounted interface
                CopyMemory iwpT, pSubclass, 4
                ' Do NOT hit the End button here! You will crash!
                ' Assign to legal reference
                Set iwp = iwpT
                ' Still do NOT hit the End button here! You will still crash!
                ' Destroy the illegal reference
                CopyMemory iwpT, 0&, 4
                ' OK, hit the End button if you must--you'll probably still crash,
                ' but it will be because of the subclass, not the uncounted reference
                
                ' Store the current message, so the client can check it:
                m_iCurrentMessage = iMsg
                m_iProcOld = procOld
                ' Use the interface to call back to the class
                With iwp
                    ' Preprocess (only check this the first time around):
                    If (iP = 1) Then
                        If .MsgResponse = emrPreprocess Then
                           If Not (bCalled) Then
                              WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                                        wParam, ByVal lParam)
                              bCalled = True
                           End If
                        End If
                    End If
                    ' Consume (this message is always passed to all control
                    ' instances regardless of whether any single one of them
                    ' requests to consume it):
                    WindowProc = .WindowProc(hWnd, iMsg, wParam, ByVal lParam)
                    ' PostProcess (only check this the last time around):
                    If (iP = iPC) Then
                        If .MsgResponse = emrPostProcess Then
                           If Not (bCalled) Then
                              WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                                        wParam, ByVal lParam)
                              bCalled = True
                           End If
                        End If
                    End If
                End With
            End If
        Next iP
        Else
        ' This message is not handled, so pass on to old procedure
        WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                    wParam, ByVal lParam)
    End If

End Function

' Called from cSubclass_WindowProc(UserControl) and TvWMPaint(UserControl)
Public Function CallOldWindowProc(ByVal hWnd As Long, _
                        ByVal iMsg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long
   
   CallOldWindowProc = CallWindowProc(m_iProcOld, hWnd, iMsg, wParam, lParam)

End Function

' Cheat! Cut and paste from MWinTool rather than reusing
' file because reusing file would cause many unneeded dependencies
' Called from AttachMessage
Private Function IsWindowLocal(ByVal hWnd As Long) As Boolean
    
  Dim idWnd As Long
    
    Call GetWindowThreadProcessId(hWnd, idWnd)
    IsWindowLocal = (idWnd = GetCurrentProcessId())

End Function

' This Subclasing is independent of VBCore, so it hard codes error handling
' Called from AttachMessage
Private Sub ErrRaise(E As Long)
    
  Dim sText As String
  Dim sSource As String
    
    If E > 1000 Then
        sSource = App.EXEName & ".WindowProc"
        Select Case E
            Case eeCantSubclass
                sText = "Can't subclass window"
            Case eeAlreadyAttached
                sText = "Message already handled by another class"
            Case eeInvalidWindow
                sText = "Invalid window"
            Case eeNoExternalWindow
                sText = "Can't modify external window"
        End Select
        Err.Raise E Or vbObjectError, sSource, sText
        Else
        ' Raise standard Visual Basic error
        Err.Raise E, sSource
    End If

End Sub

