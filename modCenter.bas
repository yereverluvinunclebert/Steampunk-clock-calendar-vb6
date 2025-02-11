Attribute VB_Name = "modCenter"
'---------------------------------------------------------------------------------------
' Module    : modCentre
' Author    : https://www.vbforums.com/member.php?65196-Chris001
' Date      : 11/02/2025
' Purpose   : Intercepts all form WM_CREATE messages, tests for a dialog class placing the form in the middle rather than top left.
'---------------------------------------------------------------------------------------

Option Explicit

Private Type CWPSTRUCT
        lParam As Long
        wParam As Long
        message As Long
        hwnd As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOSIZE = &H1
Private Const WM_CREATE = &H1
Private Const WH_CALLWNDPROC = 4

Private lHook As Long

Private Function WindowHook(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim tCWP As CWPSTRUCT
    Dim tRECT As RECT
    Dim sClass As String
    
    CopyMemory tCWP, ByVal lParam, Len(tCWP)
    Select Case tCWP.message
    Case WM_CREATE
        'Check the Type of Windows Being Created
        sClass = Space(255)
        sClass = Left$(sClass, GetClassName(tCWP.hwnd, ByVal sClass, 255))
        'If it's a Dialog Window, Center it..
        If sClass = "#32770" Then
            Call GetWindowRect(tCWP.hwnd, tRECT)
            Call SetWindowPos(tCWP.hwnd, 0, ((Screen.Width / Screen.TwipsPerPixelX) - (tRECT.Right - tRECT.Left)) / 2, ((Screen.Height / Screen.TwipsPerPixelY) - (tRECT.Bottom - tRECT.Top)) / 2, 0, 0, SWP_NOSIZE Or SWP_FRAMECHANGED)
        End If
    End Select
    WindowHook = CallNextHookEx(lHook, nCode, wParam, lParam)
End Function

Public Sub SetHook()
    'Hook the Threads Messages
    lHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf WindowHook, App.hInstance, App.ThreadID)
End Sub

Public Sub ReleaseHook()
    'UnHook the Threads Messages
    Call UnhookWindowsHookEx(lHook)
End Sub
