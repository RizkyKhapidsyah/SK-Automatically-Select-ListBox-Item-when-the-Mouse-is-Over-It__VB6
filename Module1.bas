Attribute VB_Name = "Module1"
Option Explicit

Public Const LB_SETCURSEL = &H186
Public Const LB_GETCURSEL = &H188
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function ClientToScreen Lib "user32" _
(ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function LBItemFromPt Lib "COMCTL32.DLL" _
(ByVal hLB As Long, ByVal ptX As Long, ByVal ptY As Long, _
ByVal bAutoScroll As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Long) As Long

