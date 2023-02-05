VERSION 5.00
Begin VB.Form frmStar 
   BorderStyle     =   0  'None
   Caption         =   "Salat Timings"
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   ScaleHeight     =   2900
   ScaleMode       =   0  'User
   ScaleWidth      =   2900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   2100
      Left            =   397
      Picture         =   "frmStar.frx":0000
      ScaleHeight     =   2027.586
      ScaleMode       =   0  'User
      ScaleWidth      =   2100
      TabIndex        =   0
      Top             =   450
      Width           =   2100
   End
   Begin VB.Timer Timer1 
      Left            =   2520
      Top             =   120
   End
End
Attribute VB_Name = "frmStar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NumCoords As Byte
Dim MyLeft As Long
Dim MyTop As Long
Dim Hr As Byte
Dim Mt As Byte
Dim sec As Byte
Dim wid As Single
Dim hgt As Single

Private Sub Form_Click()
pathclockpic = GetNextPic(App.Path & "\clocks\")

End Sub

Private Sub Form_DblClick()
'Timer1.Enabled = Not (Timer1.Enabled)
MyLeft = Int((Screen.Width - Me.Width) * Rnd)
MyTop = Int((Screen.Height - Me.Height) * Rnd)

Me.Move MyLeft, MyTop
End Sub

Private Sub Form_GotFocus()

starPos

End Sub

Private Sub Form_Load()
Timer1.Interval = 1000
Randomize
Picture1.ScaleMode = 3
'CenterIt fMainForm, Me
SetTransparency Me, Settings1.Transparency
pathclockpic = GetNextPic(App.Path & "\clocks\")

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Result As Long

If Button = 2 Then
    Result = SetForegroundWindow(fMainForm.hwnd)
    fMainForm.PopupMenu fMainForm.mpopsys
End If
End Sub

Private Sub Form_Paint()
starPos
End Sub

Private Sub Form_Resize()
CreateStar
If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
    starPos
End If

End Sub

Private Sub Picture1_Click()
pathclockpic = GetNextPic(App.Path & "\clocks\")
End Sub

Private Sub Timer1_Timer()
'If StopTimer Then
    'Timer1.Enabled = False
'Else
'    Timer1.Enabled = True
'End If

'MyLeft = Int((Screen.Width - Me.Width) * Rnd)
'MyTop = Int((Screen.Height - Me.Height) * Rnd)
'If Not Me.WindowState = vbMinimized Then

'End If
ClockTick


End Sub

Sub ClockTick()
On Local Error GoTo cterr
Hr = hour(Now)
Mt = Minute(Now)
sec = Second(Now)
If pathclockpic = "" Then
    pathclockpic = App.Path & "\clocks\" & "EWHIT140.BMP"
End If
Picture1.Picture = LoadPicture(pathclockpic)
Call ClckHnd_rotate(Picture1, sec * PI / 30, 80, 2, Settings1.SecondHandColor)   ' &H101FF)
Call ClckHnd_rotate(Picture1, (Mt * PI / 30) - PI / 2, 70, 3, Settings1.MinuteHandColor)    ' &HFF0101)
'Call ClckHnd_rotate(Picture1, (Hr * PI / 6) - PI / 3, 40, 5, Settings1.HourHandColor) ' &HFF0101)
Call ClckHnd_rotate(Picture1, (Hr * PI / 6) - PI / 3 - (Mt / 180), 50, 6, Settings1.HourHandColor)   ' &HFF0101)
'Me.Picture1.Circle (wid / 2, hgt / 2), 6, RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Exit Sub
cterr:
If Err.Number = 53 Then
    Err.Clear
    Picture1.Picture = LoadResPicture(101, vbResBitmap)
    fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
Else
    fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End If
End Sub

Sub CreateStar()
Const RGN_DIFF = 4

Dim outer_rgn As Long
Dim inner_rgn As Long
Dim combined_rgn As Long
Dim border_width As Single
Dim title_height As Single

    If WindowState = vbMinimized Then Exit Sub

' Dimension coordinate array.
ReDim Poly(1 To 10)
' Number of vertices in polygon.
NumCoords = 10
' Set scalemode to pixels to set up points of star.
frmStar.ScaleMode = 3
' Assign values to points.
Poly(1).x = frmStar.ScaleWidth * 0.5
Poly(1).y = frmStar.ScaleHeight * 0.001 '0.1

Poly(2).x = frmStar.ScaleWidth * 0.68 '0.63
Poly(2).y = frmStar.ScaleHeight * 0.26 '0.43

Poly(3).x = frmStar.ScaleWidth
Poly(3).y = frmStar.ScaleHeight * 0.35 '0.45

Poly(4).x = frmStar.ScaleWidth * 0.82 '0.7
Poly(4).y = frmStar.ScaleHeight * 0.66  '.66

Poly(5).x = frmStar.ScaleWidth * 0.91 '0.8
Poly(5).y = frmStar.ScaleHeight

Poly(6).x = frmStar.ScaleWidth * 0.5
Poly(6).y = frmStar.ScaleHeight * 0.87 '0.77

Poly(7).x = frmStar.ScaleWidth * 0.09 '0.2
Poly(7).y = frmStar.ScaleHeight

Poly(8).x = frmStar.ScaleWidth * 0.18
Poly(8).y = frmStar.ScaleHeight * 0.66

Poly(9).x = frmStar.ScaleWidth * 0.001
Poly(9).y = frmStar.ScaleHeight * 0.35 '0.45

Poly(10).x = frmStar.ScaleWidth * 0.32 '0.37
Poly(10).y = frmStar.ScaleHeight * 0.26 '0.43

' Sets background color to green no red for contrast.
frmStar.BackColor = &H11FF77
' Polygon function creates unfilled polygon on screen.
' Remark FillRgn statement to see results.
'bool = Polygon(frmstar.hdc, Poly(1), NumCoords)
' Gets stock black brush.
'hbrush = GetStockObject(BLACKBRUSH)
' Creates region to fill with color.
'hRgn = CreatePolygonRgn(Poly(1), NumCoords, ALTERNATE)
' If the creation of the region was successful then color.
'If hRgn Then bool = FillRgn(frmstar.hdc, hRgn, hbrush)
' Print out some information.
'Print "FillRgn Return : "; bool
'Print "HRgn : "; hRgn
'Print "Hbrush : "; hbrush
'Trash = DeleteObject(hRgn)


'CenterIt fMainForm, Me
'Me.Move fMainForm.ScaleLeft + fMainForm.ScaleWidth * 0.7, _
    fMainForm.ScaleTop + fMainForm.ScaleHeight * 0.2, _
    fMainForm.ScaleWidth * 0.5, fMainForm.ScaleHeight * 0.5
'CenterIt Me, Picture1

'Me.Picture1.Move Me.ScaleLeft + Me.ScaleWidth / 2 - Me.Picture1.ScaleWidth / 2, Me.ScaleTop + Me.ScaleHeight / 2 - Me.Picture1.ScaleHeight / 2

Me.Picture1.Scale (-100, -100)-(100, 100)
Me.Picture1.Move Me.ScaleLeft + Me.ScaleWidth * 0.14, _
    Me.ScaleTop + Me.ScaleHeight * 0.17         '.14,.17
    
    ' Create the regions.
    wid = ScaleX(Width, vbTwips, vbPixels)
    hgt = ScaleY(Height, vbTwips, vbPixels)
    'outer_rgn = CreateEllipticRgn(wid * -0.03, hgt * -0.03, wid, hgt)
    outer_rgn = CreatePolygonRgn(Poly(1), 10, ALTERNATE)
    
    'border_width = (wid - ScaleWidth) / 2
    'title_height = hgt - border_width - ScaleHeight
    
    'wid = ScaleX(Picture1.Width, vbTwips, vbPixels)
    'hgt = ScaleY(Picture1.Height, vbTwips, vbPixels)
    
    'inner_rgn = CreateEllipticRgn( _
        wid * 0.04, hgt * 0.1, _
        wid * 0.66, hgt * 0.7)
    'inner_rgn = CreateEllipticRgn(wid * 0.03, hgt * 0.05, _
        wid * 0.72, hgt * 0.72)
    
    inner_rgn = CreateEllipticRgn(wid * 0.04, hgt * 0.04, _
        wid * 0.7, hgt * 0.7)
    
    'inner_rgn = CreateEllipticRgn( _
        0, 0, 0, 0)

    ' Subtract the inner region from the outer.
    'combined_rgn = CreateEllipticRgn(0, 0, 0, 0)
    'CombineRgn combined_rgn, outer_rgn, _
        inner_rgn, RGN_DIFF
    
    ' Restrict the window to the region.
    SetWindowRgn hwnd, outer_rgn, True
    SetWindowRgn Picture1.hwnd, inner_rgn, True
    DeleteObject combined_rgn
    DeleteObject inner_rgn
    DeleteObject outer_rgn

End Sub
Sub starPos()
frmStar.Move fMainForm.ScaleLeft + fMainForm.ScaleWidth * 1.7, _
        fMainForm.ScaleTop + fMainForm.ScaleHeight * 0.38

End Sub
