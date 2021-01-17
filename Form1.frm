VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox listBox1 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Created by Rizky Khapidsyah

Public Sub HighlightLBItem(ByVal LBHwnd As Long, ByVal X As Single, ByVal Y As Single)

Dim ItemIndex As Long
Dim AtThisPoint As POINTAPI

    AtThisPoint.X = X \ Screen.TwipsPerPixelX
    AtThisPoint.Y = Y \ Screen.TwipsPerPixelY
    
    Call ClientToScreen(LBHwnd, AtThisPoint)
    
    ItemIndex = LBItemFromPt(LBHwnd, AtThisPoint.X, AtThisPoint.Y, False)

    If ItemIndex <> SendMessage(LBHwnd, LB_GETCURSEL, 0, 0) Then
        Call SendMessage(LBHwnd, LB_SETCURSEL, ItemIndex, 0)
    End If
End Sub

Private Sub Form_Load()
    With listBox1
        .Clear
        .AddItem "Item 1", 0
        .AddItem "Item 2", 1
        .AddItem "Item 3", 2
        .AddItem "Item 4", 3
        .AddItem "Item 5", 4
        .AddItem "Item 6", 5
        .AddItem "Item 7", 6
    End With
End Sub
