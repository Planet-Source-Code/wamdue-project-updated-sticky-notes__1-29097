Attribute VB_Name = "HotKey"
Option Explicit
'##################################
'This module does all the work necessary to create a system wide hot key
'Some of this code is not original it came from many snippet sources
'there is also plenty of original code to make it work
'Your free to use it
'charles@Lightwave Software.com
'##################################
'NOTE:  **LOOK**
'DO NOT FORGET TO {HotKey.DeleteHotKey hWnd } PUT THIS IN THE FORM UNLOAD EVENT OR CALL IT
'WHEN ENDING A PROJECT OR VB WILL CRASH (hWnd is the hWnd of the form the hot key is assigned to Eg. Me.hWnd)
'**********************************************
'Example
'Look at the form load event and Combo Click and Form Unload on the form for examples
'You can modify the action code in this module to do any action
'It is also possible to handle the events on the form
'The WParam of the window message identifyies each key created for a form with a index number
'If you create more than one hot key you can handle these individualy with the WParam
'Scan the code below and you'll see where.
'###################################



'Do not modify these declares
Public CurrentModifier As String
Public CurrentKey As String


Private Declare Function RegisterHotKey Lib "user32" _
    (ByVal hWnd As Long, ByVal id As Long, _
    ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" _
    (ByVal hWnd As Long, ByVal id As Long) As Long

Public Enum ModConst
    MOD_ALT = &H1          'Alt Key
    MOD_CONTROL = &H2      'Ctrl Key
    MOD_SHIFT = &H4        'Shift Key
    MOD_WIN = &H8
End Enum
Const WM_HOTKEY = &H312
Private m_hkCount As Integer
Private subclassStatus As Boolean

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = (-4)
Private defaultProc As Long
'End do not modify


'#########################################################
'Place The Action to perform in the sub below

Public Sub HotKeyAction()
'Place Code for your action here
'don't forget to make sure that any constants or declares are
'available to this sub that are needed
frmMain.mnuNew_Click

End Sub






'#########################################################
'you do not need to modify code below this line unless
'you want to customize one of the functions
'##########################################################









Public Sub CreateHotKey(ByVal hWnd As Long, Modifier As ModConst, KeyCode As Integer)
    Subclass hWnd
    HotKeyActivate hWnd, Modifier, KeyCode
    Select Case Modifier
        Case MOD_ALT         'Alt Key
            CurrentModifier = "Alt"
        Case MOD_CONTROL     'Ctrl Key
            CurrentModifier = "Ctrl"
        Case MOD_SHIFT        'Shift Key
            CurrentModifier = "Shift"
        Case MOD_WIN
            CurrentModifier = "Win Button"
    End Select
        CurrentKey = Chr(KeyCode)
        
End Sub
Public Sub DeleteHotKey(ByVal hWnd As Long)
    StopSubclassing hWnd
    HotKeyDeactivate hWnd
End Sub


Private Function VirtualProc(ByVal hWnd As Long, ByVal WindowMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    'This sends the message to the Form that = defaultproc hWnd the actions can also be
    'handled there instead of in the HotKeyAction Sub above
    VirtualProc = CallWindowProc(defaultProc, hWnd, WindowMsg, wParam, lParam)
    If WindowMsg = WM_HOTKEY Then
        'The WParam determins the index identifier of the hot
        'Key combonation pressed so this can be used for more than on action
        'with more than one hot key pair
        HotKeyAction
    End If
End Function

Private Function HotKeyActivate(ByVal hWnd As Long, _
    Modifier As ModConst, KeyCode As Integer) As Integer
    
    m_hkCount = m_hkCount + 1
    
    RegisterHotKey hWnd, m_hkCount, Modifier, KeyCode
    
    HotKeyActivate = m_hkCount
End Function

Private Function HotKeyDeactivate(ByVal hWnd As Long)
   
    Dim i As Integer
    For i = 1 To m_hkCount
        UnregisterHotKey hWnd, i
    Next i
    m_hkCount = 0
End Function



Private Sub StopSubclassing(ByVal hWnd As Long)
    If subclassStatus = False Then Exit Sub
    Call SetWindowLong(hWnd, GWL_WNDPROC, defaultProc)
    subclassStatus = False

    
End Sub



Private Sub Subclass(ByVal hWnd As Long)
    If subclassStatus = True Then Exit Sub
    defaultProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf VirtualProc)
   subclassStatus = True
End Sub

