VERSION 5.00
Begin VB.Form wndFullScreen 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   WindowState     =   2  'Maximiert
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1650
      Top             =   1050
   End
End
Attribute VB_Name = "wndFullScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_BOTTOM = 1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Const playingWallpaper As Boolean = False

Private Sub Form_DblClick()
    wndMain.Visible = True
    wndMain.WindowState = vbNormal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If Not playingWallpaper Then
        wndMain.Visible = True
        wndMain.WindowState = vbNormal
    'End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If playingWallpaper Then
        sendToBottom
    End If
End Sub

Private Sub Form_Resize()
    If playingWallpaper Then
        sendToBottom
    End If
End Sub

Public Sub sendToBottom()
    SetWindowPos Me.hWnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
