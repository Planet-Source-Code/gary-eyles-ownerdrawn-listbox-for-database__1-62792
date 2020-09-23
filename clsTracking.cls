VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public hwnd As Long
Public ScrollHwnd As Long
Public TextHwnd As Long
Public prevProc As Long
Public prevProcScroll As Long
Public prevProcText As Long

Public Event MouseLeave()
Public Event ScrollUp()
Public Event ScrollDown()
Public Event cLostFocus()

Public Sub RaiseMouseLeaveList()
RaiseEvent MouseLeave
End Sub

Public Sub RaiseScrollUp()
RaiseEvent ScrollUp
End Sub

Public Sub RaiseScrollDown()
RaiseEvent ScrollDown
End Sub

Public Sub RaisecLostFocus()
RaiseEvent cLostFocus
End Sub
