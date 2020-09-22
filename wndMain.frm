VERSION 5.00
Begin VB.Form wndMain 
   Appearance      =   0  '2D
   BackColor       =   &H00000000&
   Caption         =   "Eyes without Legs"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   8895
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer tmrProc 
      Interval        =   1
      Left            =   150
      Top             =   150
   End
   Begin VB.PictureBox pctFace 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      FillStyle       =   0  'Ausgefüllt
      Height          =   3150
      Left            =   150
      Picture         =   "wndMain.frx":0000
      ScaleHeight     =   210
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   571
      TabIndex        =   0
      Top             =   150
      Width           =   8565
   End
End
Attribute VB_Name = "wndMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type EYEBALL
        CenterX As Double
        CenterY As Double
        Radius As Double
        PupilX As Double
        PupilY As Double
        PupilDestX As Double
        PupilDestY As Double
        PupilRadius As Double
End Type
Private Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private PI              As Single

Private lpEyeBalls(1)   As EYEBALL
Private Const EYESPEED  As Double = 1

Dim cntFace             As Form         'this is the form where the eyes are

Private Sub Form_Load()
    Dim I               As Integer
    
    Set cntFace = Me
    
    PI = Atn(1) * 4

    lpEyeBalls(0).CenterX = 222
    lpEyeBalls(0).CenterY = 67
    lpEyeBalls(0).Radius = 24
    lpEyeBalls(0).PupilRadius = 24
    lpEyeBalls(1).CenterX = 347
    lpEyeBalls(1).CenterY = 67
    lpEyeBalls(1).Radius = 24
    lpEyeBalls(1).PupilRadius = 24

    For I = 0 To 1
        lpEyeBalls(I).PupilX = lpEyeBalls(I).CenterX
        lpEyeBalls(I).PupilY = lpEyeBalls(I).CenterY
    Next I
    
    Show

End Sub

Public Static Sub Form_Resize()
    Dim lastWindowState     As Integer

    If Me.WindowState <> lastWindowState Then
        Select Case WindowState
            Case vbMaximized
                lastWindowState = vbMaximized
                sendThemAway
                tmrProc.Enabled = True
            Case vbNormal
                lastWindowState = vbNormal
                getThemBack
                tmrProc.Enabled = True
            Case vbMinimized
                lastWindowState = vbMinimized
                getThemBack
                tmrProc.Enabled = False
        End Select
    End If
    Me.pctFace.Move (cntFace.ScaleWidth - Me.pctFace.Width) / 2, (cntFace.ScaleHeight - Me.pctFace.Height) / 2
End Sub

Private Sub pctFace_DblClick()
    Me.Visible = True
    Me.WindowState = vbNormal
End Sub

Private Sub pctFace_KeyPress(KeyAscii As Integer)
    Me.Visible = True
    Me.WindowState = vbNormal
End Sub

Private Sub tmrProc_Timer()
    Dim lpMouse         As POINTAPI
    Dim lpCenter(1)     As POINTAPI
    Dim lpWindow        As RECT
    Dim dblDist         As Double
    Dim dblAngle        As Double
    Dim I               As Integer
    Dim bChanged        As Boolean

    GetCursorPos lpMouse                                  'position of mouse
    GetWindowRect Me.pctFace.hWnd, lpWindow               'position of picture
    
    lpCenter(0).X = lpWindow.Left + lpEyeBalls(0).CenterX 'center of left eye
    lpCenter(0).Y = lpWindow.Top + lpEyeBalls(0).CenterY
    
    lpCenter(1).X = lpWindow.Left + lpEyeBalls(1).CenterX 'center of right eye
    lpCenter(1).Y = lpWindow.Top + lpEyeBalls(1).CenterY
 

    For I = 0 To 1
'first, we calculate the target position of the pupil
    'distance from center of eyeball
        dblDist = Sqr((lpMouse.X - lpCenter(I).X) ^ 2 + (lpMouse.Y - lpCenter(I).Y) ^ 2)
        'must not be more than the radius
        If dblDist > lpEyeBalls(I).Radius Then
            dblDist = lpEyeBalls(I).Radius
        End If
    'angle from the center of the eyeball
        dblAngle = getAngle(lpMouse.X - lpCenter(I).X, lpMouse.Y - lpCenter(I).Y)
    'the coordinates within the picturebox
        lpEyeBalls(I).PupilDestX = lpEyeBalls(I).CenterX + Sin(dblAngle) * dblDist
        lpEyeBalls(I).PupilDestY = lpEyeBalls(I).CenterY + Cos(dblAngle) * dblDist
    
'then, we calculate the new position of the pupil
    'distance from pupil to target
        dblDist = Sqr((lpEyeBalls(I).PupilDestX - lpEyeBalls(I).PupilX) ^ 2 + (lpEyeBalls(I).PupilDestY - lpEyeBalls(I).PupilY) ^ 2)
    'if it didn't reach its target yet
        If dblDist > 0 Then
        'then we limit the speed to the max
            If dblDist > EYESPEED Then
                dblDist = EYESPEED
            End If
        'calculate the angle for its journey
            dblAngle = getAngle(lpEyeBalls(I).PupilDestX - lpEyeBalls(I).PupilX, lpEyeBalls(I).PupilDestY - lpEyeBalls(I).PupilY)
        'and move it
            lpEyeBalls(I).PupilX = lpEyeBalls(I).PupilX + Sin(dblAngle) * dblDist
            lpEyeBalls(I).PupilY = lpEyeBalls(I).PupilY + Cos(dblAngle) * dblDist
            bChanged = True
        End If
    Next I
    
    If bChanged Then
        Me.pctFace.Cls
        For I = 0 To 1
            Me.pctFace.Circle (lpEyeBalls(I).PupilX, lpEyeBalls(I).PupilY), lpEyeBalls(I).PupilRadius
        Next I
    End If
End Sub

Private Function getAngle(X As Double, Y As Double) As Double
    If X = 0 And Y = 0 Then Exit Function
    getAngle = ASin(X / Sqr(X ^ 2 + Y ^ 2))
    If Y < 0 Then getAngle = Atn(1) * 4 - getAngle
End Function

Private Function ASin(X As Double) As Double
    If Abs(X) = 1 Then
        ASin = Sgn(X) * Atn(1) * 2
    Else
        ASin = Atn(X / Sqr(-X * X + 1))
    End If
End Function

Public Sub getThemBack()
    If cntFace Is Me Then
        'you can't get them back, because they aren't away
    Else
        SetParent Me.pctFace.hWnd, Me.hWnd
        Unload cntFace
        Set cntFace = Me
        Me.Visible = True
    End If
End Sub

Public Sub sendThemAway()
    If cntFace Is Me Then
        Set cntFace = New wndFullScreen
        SetParent pctFace.hWnd, cntFace.hWnd
        cntFace.Show
        Me.Visible = False
    Else
        'you can't send them away, because they aren't here
    End If
End Sub
