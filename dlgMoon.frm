VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form dlgMoon 
   BackColor       =   &H00600000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lunar Phase, Distance, Position and Zodiac for the selected date."
   ClientHeight    =   5835
   ClientLeft      =   3240
   ClientTop       =   3210
   ClientWidth     =   8010
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "dlgMoon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   7560
      Top             =   3720
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Click the Calendar for a date to see the Lunar details"
      Top             =   120
      Width           =   4695
      _Version        =   524288
      _ExtentX        =   8281
      _ExtentY        =   6588
      _StockProps     =   1
      BackColor       =   12648447
      Year            =   2022
      Month           =   10
      Day             =   17
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   16711680
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   0
      GridLinesColor  =   49152
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   255
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblDist 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1695
      Left            =   4000
      TabIndex        =   1
      Top             =   3960
      Width           =   4000
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPhase 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1695
      Left            =   150
      TabIndex        =   0
      Top             =   3960
      Width           =   4400
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   4920
      Picture         =   "dlgMoon.frx":8FBA
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "dlgMoon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Moon-phase calculation
'   Thanx to the Javascript code of Roger W. Sinnott, Sky & Telescope, June 16, 2006.
'   Which was in turn adapted from a BASIC program from the Astronomical Computing column of Sky & Telescope, April 1994
'   And now it is back to basic(s)

Option Explicit
'Dim n0 As Double
'Dim f0 As Double
Dim AG As Double
Dim DI As Double
Dim LA As Double
Dim LO As Double
Dim Phase As String
Dim Zodiac As String
Dim d As Date
Const earthRad = 6371.0088  'km
Dim JD As Double
Dim CalendarClicked As Boolean
Dim i As Integer


Function ValidateForm(Optional mDateTime As Date)
On Local Error GoTo validateErr

If ((mDateTime = "12:00:00 AM") Or (mDateTime = "00:00:00") Or IsNull(mDateTime) Or (Not (IsDate(mDateTime)))) Then
    mDateTime = FormatDateTime(Calendar1.Value, vbLongTime)
End If
If CalendarClicked Then
    mDateTime = FormatDateTime(Calendar1.Value, vbLongTime)
End If

JD = calcJD(mDateTime, Val(city.TimeZone))       'calcJD(Calendar1.Value, Val(city.TimeZone))
'JD = jday(year(mDateTime), month(mDateTime), day(mDateTime), hour(mDateTime), Minute(mDateTime), Second(mDateTime))

initialize
calculate mDateTime

Call moonElong(mDateTime)

lblDist.Caption = "Moons's Distance is " & DI & " Earth Radii" & vbCrLf & "Moons's Distance is " & DI * earthRad & " km" & vbCrLf
lblDist.Caption = lblDist.Caption & "Moon's Zodiac is " & Zodiac & vbCrLf & _
    "Moon's Ecliptic latitude is " & LA & vbCrLf & _
    "Moon's Ecliptic longitude is " & LO

lblPhase.Caption = lblPhase.Caption & vbCrLf & "Moon's Age is " & AG & " Days"
lblPhase.Caption = lblPhase.Caption & vbCrLf & Phase
dlgMoon.Caption = "Lunar Phase, Distance, Position and Zodiac for " & mDateTime

Exit Function
validateErr:
MsgBox Err.Description & " Error Number:" & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text
Resume Next

End Function

Function moonElong(Optional mDateTime As Date)
On Local Error GoTo meerr
Dim moonNumStr As String
   Dim meeDT As Double
   Dim meeT As Double
   Dim meeT2 As Double
   Dim meeT3 As Double
   Dim meeD As Double
   Dim meeM1 As Double
   Dim meeM As Double
   Dim elong As Double
   Dim moonNum As Double
   Dim moonImage As String
   Dim moonPhase As String
   Dim tmp As Double
   Dim tZone As String
   Dim dtZone As Double
   
   tZone = city.TimeZone
   dtZone = Val(tZone)
   If mDateTime = "12:00:00 AM" Then
    mDateTime = Now
   End If
   If CalendarClicked Then
    mDateTime = Me.Calendar1.Value
   End If
    JD = calcJD(mDateTime, dtZone)
    'JD = jday(year(mDateTime), month(mDateTime), day(mDateTime), hour(mDateTime), Minute(mDateTime), Second(mDateTime))

    tmp = JD - 2382148#
    tmp = tmp * tmp
    tmp = tmp / 41048480#
   meeDT = tmp / 86400#
   meeT = (JD + meeDT - 2451545#) / 36525#
   meeT2 = (meeT ^ 2)
   meeT3 = (meeT ^ 3)
   meeD = 297.85 + (445267.1115 * meeT) - (0.00163 * meeT2) + (meeT3 / 545868)
       meeD = D2R(proper_ang(meeD))
   meeM1 = 134.96 + (477198.8676 * meeT) + (0.008997 * meeT2) + (meeT3 / 69699)
       meeM1 = D2R(proper_ang(meeM1))
   meeM = 357.53 + (35999.0503 * meeT)
       meeM = D2R(proper_ang(meeM))
   elong = R2D(meeD) + 6.29 * Sin(meeM1)
       elong = elong - 2.1 * Sin(meeM)
       elong = elong + 1.27 * Sin(2 * meeD - meeM1)
       elong = elong + 0.66 * Sin(2 * meeD)
       elong = proper_ang(elong)
       elong = Round(elong)
   moonNum = ((elong + 6.43) / 360) * 28
       moonNum = Round(moonNum)
       moonNum = Abs(moonNum)
       If (moonNum = 28) Then
            moonNum = 0
       End If
       If (moonNum > 28) Then
            moonNum = 0
       End If
       
       moonNumStr = Trim$(str$(moonNum))
       If (moonNum < 10) Then
            moonNumStr = "0" & Abs(moonNum)
       End If
   moonImage = App.Path & "\moon\moon" & moonNumStr & ".gif"
'
'   moonPhase = " new Moon"
'     If ((moonNum > 3) And (moonNum < 11)) Then moonPhase = " first quarter"
'     If ((moonNum > 10) And (moonNum < 18)) Then moonPhase = " full Moon"
'     If ((moonNum > 17) And (moonNum < 25)) Then moonPhase = " last quarter"
'
'    If ((moonNum = 1) Or (moonNum = 8) Or (moonNum = 15) Or (moonNum = 22)) Then
'           moonPhase = " 1 day past" + moonPhase
'    End If
'     If ((moonNum = 2) Or (moonNum = 9) Or (moonNum = 16) Or (moonNum = 23)) Then
'           moonPhase = " 2 days past" + moonPhase
'     End If
'     If ((moonNum = 3) Or (moonNum = 10) Or (moonNum = 17) Or (moonNum = 24)) Then
'           moonPhase = " 3 days past" + moonPhase
'     End If
'     If ((moonNum = 4) Or (moonNum = 11) Or (moonNum = 18) Or (moonNum = 25)) Then
'           moonPhase = " 3 days before" + moonPhase
'     End If
'     If ((moonNum = 5) Or (moonNum = 12) Or (moonNum = 19) Or (moonNum = 26)) Then
'           moonPhase = " 2 days before" + moonPhase
'     End If
'     If ((moonNum = 6) Or (moonNum = 13) Or (moonNum = 20) Or (moonNum = 27)) Then
'           moonPhase = " 1 day before" + moonPhase
'     End If
   
'lblPhase.Caption = Trim(moonPhase) & vbCrLf & " Hijri Date " & GetHijriDate(Calendar1.Value)
lblPhase.Caption = "Hijri Date " & GetHijriDate(Calendar1.Value)
Image1.Picture = LoadPicture(moonImage)
Exit Function
meerr:
MsgBox Err.Description & " Error Number:" & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text
Resume Next

End Function

Private Sub Calendar1_Click()
Call ValidateForm(Calendar1.Value)
CalendarClicked = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Me.Hide
End If
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Calendar1.Value = Now
Randomize Timer

Timer1.Interval = 1000
Timer1.Enabled = True



Call ValidateForm(Now)

End Sub

Function proper_ang(big As Double) As Double
Dim tmp As Double
If (big > 0) Then
        
    tmp = big / 360#
    tmp = (tmp - Round(tmp)) * 360#
    
Else
    tmp = Round(Abs(big / 360#)) + 1
    tmp = big + tmp * 360#
End If

proper_ang = tmp

End Function


'zodiac, age, distance------------------------------------------------------------------

Sub initialize()

d = Now 'Calendar1.Value
AG = 0 'f0   ' Moon's age
DI = 0 'f0   ' Moon's distance in earth radii
LA = 0 'f0   ' Moon's ecliptic latitude
LO = 0 'f0   ' Moon's ecliptic longitude
Phase = " "
Zodiac = " "

End Sub

Sub calculate(Optional mDate As Date)

If ((mDate = "12:00:00 AM") Or (mDate = "00:00:00")) Then
    mDate = Calendar1.Value
End If
Call moon_posit(mDate)

AG = Round2(AG)
DI = Round2(DI)
LA = Round2(LA)
LO = Round2(LO)


End Sub
' compute moon position and phase
Function moon_posit(mDate As Date) As String

    Dim YY As Double
    YY = 0 'n0
    Dim Mm As Double
    Mm = 0 'n0
    Dim K1 As Double
    K1 = 0 'n0
    Dim K2 As Double
    K2 = 0 'n0
    Dim K3 As Double
    K3 = 0 'n0
    Dim JD1 As Double
    JD1 = 0 'n0
    Dim IP As Double
    IP = 0 'f0
    Dim DP As Double
    DP = 0 'f0
    Dim NP As Double
    NP = 0 'f0
    Dim RP As Double
    RP = 0 'f0

JD1 = calcJD(mDate, Val(city.TimeZone)) 'Calendar1.Value for mdate
'JD1 = jday(year(mDate), month(mDate), day(mDate), hour(mDate), Minute(mDate), Second(mDate))

    ' calculate moon's age in days
    IP = normalize((JD1 - 2451550.1) / 29.530588853)
    AG = IP * 29.53

    If (AG < 1.84566) Then
        Phase = "NEW"
    ElseIf (AG < 5.53699) Then
        Phase = "Evening crescent"
    ElseIf (AG < 9.22831) Then
        Phase = "First quarter"
    ElseIf (AG < 12.91963) Then
        Phase = "Waxing gibbous"
    ElseIf (AG < 16.61096) Then
        Phase = "FULL"
    ElseIf (AG < 20.30228) Then
        Phase = "Waning gibbous"
    ElseIf (AG < 23.99361) Then
        Phase = "Last quarter"
    ElseIf (AG < 27.68493) Then
        Phase = "Morning crescent"
    Else
        Phase = "NEW"
    End If
    IP = IP * 2 * PI                  ' Convert phase to radians

    ' calculate moon's distance
    DP = 2 * PI * normalize((JD1 - 2451562.2) / 27.55454988)
    DI = 60.4 - 3.3 * Cos(DP) - 0.6 * Cos(2 * IP - DP) - 0.5 * Cos(2 * IP)

    ' calculate moon's ecliptic latitude
    NP = 2 * PI * normalize((JD1 - 2451565.2) / 27.212220817)
    LA = 5.1 * Sin(NP)

    ' calculate moon's ecliptic longitude
    RP = normalize((JD1 - 2451555.8) / 27.321582241)
    LO = 360 * RP + 6.3 * Sin(DP) + 1.3 * Sin(2 * IP - DP) + 0.7 * Sin(2 * IP)

    If (LO < 33.18) Then
        Zodiac = "Pisces"
    ElseIf (LO < 51.16) Then
        Zodiac = "Aries"
    ElseIf (LO < 93.44) Then
        Zodiac = "Taurus"
    ElseIf (LO < 119.48) Then
        Zodiac = "Gemini"
    ElseIf (LO < 135.3) Then
        Zodiac = "Cancer"
    ElseIf (LO < 173.34) Then
        Zodiac = "Leo"
    ElseIf (LO < 224.17) Then
        Zodiac = "Virgo"
    ElseIf (LO < 242.57) Then
        Zodiac = "Libra"
    ElseIf (LO < 271.26) Then
        Zodiac = "Scorpio"
    ElseIf (LO < 302.49) Then
        Zodiac = "Sagittarius"
    ElseIf (LO < 311.72) Then
        Zodiac = "Capricorn"
    ElseIf (LO < 348.58) Then
        Zodiac = "Aquarius"
    Else
        Zodiac = "Pisces"
    End If
    ' so longitude is not greater than 360!
    If (LO > 360) Then LO = LO - 360
moon_posit = Phase
End Function

' normalize values to range 0...1

Function normalize(v As Double) As Double

    v = v - Round(v)
    If (v < 0) Then
        v = v + 1
    End If
   normalize = v

End Function


Function Round2(x As Double) As Double
    Round2 = Round(100 * x) / 100#
End Function

Private Sub Timer1_Timer()
On Local Error GoTo timererr

Randomize Timer
Static tick As Byte
tick = tick + 1
If tick > 27 Then
    tick = 0
End If

If Not CalendarClicked Then
    Call ValidateForm(Now)
End If

Exit Sub

timererr:
MsgBox " Timer " & Err.Description & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End Sub
