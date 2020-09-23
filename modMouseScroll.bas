Attribute VB_Name = "modMouseScroll"
'Added:
'Use: For use with mouse wheel

Option Explicit

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As tagTRACKMOUSEEVENT) As Long

Type tagTRACKMOUSEEVENT
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Public Const TME_HOVER = &H1
Public Const TME_LEAVE = &H2
Public Const TME_CANCEL = &H80000000
Public Const HOVER_DEFAULT = &HFFFFFFFF
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_MOUSELEAVE = &H2A3
Public Const GWL_WNDPROC = (-4)
Private Const WM_MOUSEMOVE = &H200
Private Const WM_ACTIVATE = &H6

Public IsHooked As Boolean
Dim prevProc As Long
Dim PrevProc2 As Long
Dim m_hWnd As Long
Dim trackCol As Collection

Public Sub Hook(trak As clsTracking)
Dim prevProc As Long

If trackCol Is Nothing Then
    Set trackCol = New Collection
End If

trak.prevProc = SetWindowLong(trak.hwnd, GWL_WNDPROC, AddressOf WindowProc)
trak.prevProcScroll = SetWindowLong(trak.ScrollHwnd, GWL_WNDPROC, AddressOf WindowProc)
trak.prevProcText = SetWindowLong(trak.TextHwnd, GWL_WNDPROC, AddressOf WindowProc)
trackCol.Add trak, CStr(trak.hwnd)
trackCol.Add trak, CStr(trak.ScrollHwnd)
trackCol.Add trak, CStr(trak.TextHwnd)

'RequestTracking trak

    
    'prevProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
    'PrevProc2 = SetWindowLong(vscroll, GWL_WNDPROC, AddressOf WindowProc)
    IsHooked = True
End Sub

Public Sub UnHook(trak As clsTracking)
If trackCol Is Nothing Then Exit Sub

Call SetWindowLong(trak.hwnd, GWL_WNDPROC, trak.prevProc)
Call SetWindowLong(trak.ScrollHwnd, GWL_WNDPROC, trak.prevProcScroll)
Call SetWindowLong(trak.TextHwnd, GWL_WNDPROC, trak.prevProcText)

Dim trk As tagTRACKMOUSEEVENT
trk.cbSize = 16
trk.dwFlags = TME_LEAVE Or TME_HOVER Or TME_CANCEL
trk.hwndTrack = trak.hwnd
TrackMouseEvent trk

trackCol.Remove CStr(trak.hwnd)
trackCol.Remove CStr(trak.ScrollHwnd)
trackCol.Remove CStr(trak.TextHwnd)
If trackCol.Count = 0 Then
    Set trackCol = Nothing
End If
End Sub

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim trak As clsTracking
Set trak = trackCol.Item(CStr(hwnd))

    WindowProc = CallWindowProc(trak.prevProc, hwnd, uMsg, wParam, lParam)
    'Debug.Print "WHEEL", Rnd, uMsg
    If uMsg = WM_MOUSEWHEEL Then
    'Debug.Print "WHEEL", Rnd, uMsg
        'Dim ctrl As Control
        'Dim strVar As String
        If (wParam \ 65536) < 0 Then
            trak.RaiseScrollDown
            'Screen.ActiveControl.ScrollDown
'            For Each ctrl In Screen.ActiveForm.Controls
'                If TypeOf ctrl Is Datalist Then
'                    If ctrl.ScrollDown Then Exit For
'                Else
'                    'sometimes loses if the control is an IsPanel or not setting
'                    'modifying this variable prevents that for some unknown reason
'                    strVar = ctrl.Name
'                    DoEvents
'                End If
'            Next
        Else
            trak.RaiseScrollUp
            'Screen.ActiveControl.ScrollUp
'            For Each ctrl In Screen.ActiveForm.Controls
'                If TypeOf ctrl Is Datalist Then
'                    If ctrl.ScrollUp Then Exit For
'                Else
'                    'sometimes loses if the control is an IsPanel or not setting
'                    'modifying this variable prevents that for some unknown reason
'                    strVar = ctrl.Name
'                    DoEvents
'                End If
'            Next
        End If
    ElseIf uMsg = WM_MOUSELEAVE Then
        'Debug.Print "MOUSE LEAVE", hwnd
        'If trak.hwnd = hwnd Then
            trak.RaiseMouseLeaveList
        'End If
    ElseIf uMsg = WM_MOUSEMOVE Then
        'Debug.Print "MOUSE MOVE", hWnd, uMsg, wParam, lParam
        Dim trk As tagTRACKMOUSEEVENT
        trk.cbSize = 16
        trk.dwFlags = TME_LEAVE Or TME_HOVER
        trk.dwHoverTime = 10
        trk.hwndTrack = hwnd
        TrackMouseEvent trk
    ElseIf uMsg = 8 Then
        If wParam = 0 Then
        'Debug.Print wParam, lParam, trak.hwnd
        trak.RaisecLostFocus
        End If
    End If
    
    m_hWnd = hwnd
End Function

