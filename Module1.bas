Attribute VB_Name = "Module1"
Option Explicit

'API Types
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

#If UNICODE Then
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal uMgs As Long, ByVal wParam As Long, lParam As Any) As Long
#Else
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal uMgs As Long, ByVal wParam As Long, lParam As Any) As Long
#End If
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Any) As Long
'the standard version of this API function uses a Point structure, but we cant pass
'that using VB, so it has been modified to accept Long Integers
'Public Declare Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal lLeft As Long, ByVal lTop As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal HwndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long


Public Const SWW_HPARENT = -8
Public Const HTRIGHT = 11
Public Const WM_NCLBUTTONDOWN = &HA1


