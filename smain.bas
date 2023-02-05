Attribute VB_Name = "smain"
Option Explicit
Public Const tpi As Double = 6.28318530717958
Public Const degs  As Double = 57.2957795130823
Public Const rads As Double = 1.74532925199433E-02

Public Const SW_SHOW = 5       ' Displays Window in its current size
                                      ' and position
Public Const SW_SHOWNORMAL = 1 ' Restores Window if Minimized or
                                      ' Maximized

Public Declare Function ShellExecute Lib "shell32.dll" Alias _
   "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
   String, ByVal lpFile As String, ByVal lpParameters As String, _
   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function FindExecutable Lib "shell32.dll" Alias _
   "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
   String, ByVal lpResult As String) As Long


Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) _
As Long
Public Declare Function GetCapture Lib "user32" () As Long

Public Declare Function GetWindowLong Lib "user32" Alias _
  "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias _
  "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
  ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" _
  (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, _
  ByVal dwFlags As Long) As Long

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long


Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT = &H20&
Public Const LWA_ALPHA = &H2&
'Public bTrans As Byte ' The level of transparency (0 - 255)

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type POINTAPI
        x As Long
        y As Long
End Type

Type CityType
    cityname As String
    Latitude As Double
    Longitude As Double
    Region As String
    country As String
    CityID As Long
    TimeZone As String
    MSL As Double
    isHanafi As Boolean
    isNorthernHemisphere As Boolean
End Type

Type salattype
    Fajr As String
    Zuhr As String
    AsrShafi As String
    AsrHanafi As String
    Maghrib As String
    Isha As String
    Sunrise As String
    
    Ishraq As String
    Duha As String
    
    DFajr As Date
    DZuhr As Date
    DAsrShafi As Date
    DAsrHanafi As Date
    DMaghrib As Date
    DIsha As Date
    DSunrise As Date

    DIshraq As Date
    DDuha As Date
End Type

Type ParamType
    EquationofTime As Double
    Latitude As Double
    Longitude As Double
    TimeZone As Double
    ReferenceLongitude As Double
    Declination As Double
    z As Double
    U As Double
    'V As Double
    VFajr As Double
    VIsha As Double
    W As Double
    x As Double
    TwilightAngle As Double
    TwilightAngleFajr As Double
    TwilightAngleIsha As Double
End Type

Type degreeformat
    Degrees As Currency
    Minutes As Currency
End Type

Type settingstype
    AdanEnabled As Boolean
    WarningEnabled As Boolean
    WarningMinutes As Integer
    Transparency As Byte
    TwilightAngle As Byte
    TwilightAngleFajr As Byte
    TwilightAngleIsha As Byte
    adan As String
    FAdan As String
    ClockFace As String
    HourHandColor As Long
    MinuteHandColor As Long
    SecondHandColor As Long
    WarningTxt As String
    AdjustmentDays As Integer
End Type

Public Settings1 As settingstype
Public city As CityType
Public Namaz As salattype
Public Param As ParamType
Public TwilightAngle   As Currency
'Public isHanafi As Boolean

Public Declare Function CreateEllipticRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateEllipticRgnIndirect Lib "GDI32" (lpRect As RECT) As Long
Public Declare Function CreatePolygonRgn Lib "GDI32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreatePolyPolygonRgn Lib "GDI32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long

Public Declare Function CreateRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "GDI32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Public Declare Function GetStockObject Lib "GDI32" (ByVal nIndex As Long) As Long
Public Declare Function FillRgn Lib "GDI32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hbrush As Long) As Long

Public Declare Function Polygon Lib "GDI32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Const ALTERNATE = 1 ' ALTERNATE and WINDING are
Public Const WINDING = 2   ' constants for FillMode.

Public Const PI = 3.14159265358979

Public Const SRCCOPY = &HCC0020

Public Declare Function SetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function StretchBlt& Lib "GDI32" (ByVal hDC&, ByVal x&, ByVal y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal XSrc&, ByVal YSrc&, ByVal nSrcWidth&, ByVal nSrcHeight&, ByVal dwRop&)
Public Poly() As POINTAPI

Public fMainForm As frmMain
Public Cnn As ADODB.Connection
Public DataPath As String
Public Waqt As salattype
Public StopTimer As Boolean
Public Declare Function sndPlaySound Lib "winmm" Alias _
    "sndPlaySoundA" (ByVal lpszSoundName As String, _
    ByVal uFlags As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias _
"PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, _
ByVal dwFlags As Long) As Long
    
' flag uitzetten
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Const SND_FILENAME = &H20000     ' Name is a file name.

Public AdanCalled As Boolean

Public Declare Function EnumCalendarInfo Lib "kernel32" Alias "EnumCalendarInfoA" (ByVal lpCalInfoEnumProc As Long, ByVal Locale As Long, ByVal Calendar As Long, ByVal CalType As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'user defined type required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
 cbSize As Long
 hwnd As Long
 uId As Long
 uFlags As Long
 uCallBackMessage As Long
 hIcon As Long
 szTip As String * 64
End Type
'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201      'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Right Button down
Public Const WM_RBUTTONUP = &H205       'Right Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Right Double-click

Public nid As NOTIFYICONDATA

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long


   Const SW_HIDE = 0
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, _
   ByVal wCmd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) _
   As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Const GW_HWNDNEXT = 2
Public Const GW_OWNER = 4

Public starttime As Date


'Registry functions
   Public Const REG_SZ As Long = 1
   Public Const REG_DWORD As Long = 4

   Public Const HKEY_CLASSES_ROOT = &H80000000
   Public Const HKEY_CURRENT_USER = &H80000001
   Public Const HKEY_LOCAL_MACHINE = &H80000002
   Public Const HKEY_USERS = &H80000003

   Public Const ERROR_NONE = 0
   Public Const ERROR_BADDB = 1
   Public Const ERROR_BADKEY = 2
   Public Const ERROR_CANTOPEN = 3
   Public Const ERROR_CANTREAD = 4
   Public Const ERROR_CANTWRITE = 5
   Public Const ERROR_OUTOFMEMORY = 6
   Public Const ERROR_INVALID_PARAMETER = 7
   Public Const ERROR_ACCESS_DENIED = 8
   Public Const ERROR_INVALID_PARAMETERS = 87
   Public Const ERROR_NO_MORE_ITEMS = 259

   Public Const KEY_ALL_ACCESS = &H3F
   Public Const REG_OPTION_NON_VOLATILE = 0

   Declare Function RegCloseKey Lib "advapi32.dll" _
   (ByVal hKey As Long) As Long
   
   Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
   "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
   As Long, ByVal samDesired As Long, ByVal lpjoysurfAttributes _
   As Long, phkResult As Long, lpdwDisposition As Long) As Long
   
   Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
   "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
   Long) As Long
   
   Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As String, lpcbData As Long) As Long
   
   Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, lpData As _
   Long, lpcbData As Long) As Long
   
   Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As Long, lpcbData As Long) As Long
   
   Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
   String, ByVal cbData As Long) As Long
   
   Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
   ByVal cbData As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" _
    Alias "RegDeleteKeyA" (ByVal hKey As Long, _
    ByVal lpSubKey As String) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" _
    Alias "RegDeleteValueA" (ByVal hKey As Long, _
    ByVal lpValueName As String) As Long
    
Public Declare Function RegEnumKey Lib "advapi32.dll" _
    Alias "RegEnumKeyA" (ByVal hKey As Long, _
    ByVal dwIndex As Long, ByVal lpName As String, _
    ByVal cbName As Long) As Long



Public Declare Function RegEnumValue Lib "advapi32.dll" _
Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex _
As Long, ByVal lpValueName As String, lpcbValueName As Long, _
ByVal lpReserved As Long, lpType As Long, lpData As Byte, _
lpcbData As Long) As Long

Public Declare Function RegFlushKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32" _
   Alias "RegQueryValueExA" (ByVal hKey As Long, _
   ByVal lpValueName As String, _
   ByVal lpReserved As Long, ByRef lpType As Long, _
   ByVal szData As String, _
   ByRef lpcbData As Long) As Long

Public Const ERROR_SUCCESS = 0
Public Const ERROR_ARENA_TRASHED = 7

Public CityEditMode As Boolean
Public Degree As degreeformat

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public f_Maximised As Boolean
Public HijriMonths(1 To 12) As String
Public Const QiblaLon = 39.823333
Public Const QiblaLat = 21.423333
'Public DSS As DirectSS
Public pathclockpic As String
Public DaysArray(1 To 12) As Long
Public AdjustmentDays As Integer



Sub Main()
On Local Error GoTo mainerr
    Dim strCnn As String
    If App.PrevInstance Then
        End
    End If
    frmSplash.Show
    frmSplash.Refresh
    starttime = Now
    DataPath = App.Path & "\" & App.exename & ".mdb"
    strCnn = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & DataPath & ";"
    Set Cnn = New ADODB.Connection
    Cnn.Open strCnn
    SetOnce App.exename
    Install
    If Not reg(App.exename, "01-Jan-2109", 999999999) Then
        MsgBox "Century Completed"
    End If
    Settings1 = GetSettings()
    Set fMainForm = New frmMain
    Load fMainForm
    fMainForm.Hide
    LoadCurrentCity
    TwilightAngle = 18
    LoadHijriMonths
'    isUpDateAvailable App.Major & "." & App.Minor & "." & App.Revision, "http://iqsoft.co.in/salat/version.txt"
    If InStr(Command$, "m") Then
        fMainForm.WindowState = vbMinimized
        f_Maximised = False
        frmStar.WindowState = vbMinimized
        fMainForm.mpopsysminimize.Enabled = False
        fMainForm.mpopsysrestore.Enabled = True
        HideMe frmMain
        HideMe frmStar
    Else
        fMainForm.Show
    End If
    RefreshWaqt
    Unload frmSplash
    Exit Sub
mainerr:
fMainForm.sbStatusBar.Panels(1).Text = Err.Description
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End Sub

Sub LoadResStrings(frm As Form)
    On Error Resume Next


    Dim ctl As Control
    Dim obj As Object
    Dim fnt As Object
    Dim sCtlType As String
    Dim nVal As Integer


    'set the form's caption
    frm.Caption = LoadResString(CInt(frm.Tag))
    

    'set the font
    Set fnt = frm.Font
    fnt.Name = LoadResString(20)
    fnt.Size = CInt(LoadResString(21))
    

    'set the controls' captions using the caption
    'property for menu items and the Tag property
    'for all other controls
    For Each ctl In frm.Controls
        Set ctl.Font = fnt
        sCtlType = TypeName(ctl)
        If sCtlType = "Label" Then
            ctl.Caption = LoadResString(CInt(ctl.Tag))
        ElseIf sCtlType = "Menu" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "TabStrip" Then
            For Each obj In ctl.Tabs
                obj.Caption = LoadResString(CInt(obj.Tag))
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "Toolbar" Then
            For Each obj In ctl.Buttons
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                obj.Text = LoadResString(CInt(obj.Tag))
            Next
        Else
            nVal = 0
            nVal = Val(ctl.Tag)
            If nVal > 0 Then ctl.Caption = LoadResString(nVal)
            nVal = 0
            nVal = Val(ctl.ToolTipText)
            If nVal > 0 Then ctl.ToolTipText = LoadResString(nVal)
        End If
    Next


End Sub

Function CalculateTimes(cty As CityType, ThisDay As Date) As salattype
On Local Error GoTo cterr
CalculateTimes.AsrHanafi = ""
CalculateTimes.AsrShafi = ""
CalculateTimes.Fajr = ""
CalculateTimes.Isha = ""
CalculateTimes.Maghrib = ""
CalculateTimes.Sunrise = ""
CalculateTimes.Zuhr = ""
CalculateTimes.Ishraq = ""
CalculateTimes.Duha = ""

Dim JulDay As Double
Dim tmp As Double

Param.TwilightAngle = TwilightAngle

Param.Declination = GetDeclination(ThisDay)

JulDay = calcJD(ThisDay, Val(cty.TimeZone))

Param.EquationofTime = calcET(JulDay)

Param.ReferenceLongitude = Val(cty.TimeZone) * 15 ' actually *60/4

Param.z = 12 + ((Param.ReferenceLongitude - cty.Longitude) / 15) '/ 60 made a big mistake and lost months by that division
Param.z = Param.z - Param.EquationofTime / 60 ' is it z-ET or+ ET? it worked for- ET

tmp = Sin(D2R(-0.8333 - (0.0347 * (cty.MSL) ^ 0.5)))
tmp = tmp - (Sin(D2R(Param.Declination)) * Sin(D2R(cty.Latitude)))

tmp = tmp / (Cos(D2R(Param.Declination)) * Cos(D2R(cty.Latitude)))

Param.U = (1 / 15) * R2D(ArcCos(tmp))
Param.TwilightAngleFajr = Settings1.TwilightAngleFajr
If Param.TwilightAngleFajr = 0 Then
    Param.TwilightAngleFajr = Param.TwilightAngle
End If
tmp = -Sin(D2R(Param.TwilightAngleFajr))
tmp = tmp - (Sin(D2R(Param.Declination)) * Sin(D2R(cty.Latitude)))
tmp = tmp / (Cos(D2R(Param.Declination)) * Cos(D2R(cty.Latitude)))

Param.VFajr = (1 / 15) * R2D(ArcCos(tmp))

Param.TwilightAngleIsha = Settings1.TwilightAngleIsha
If Param.TwilightAngleIsha = 0 Then
    Param.TwilightAngleIsha = Param.TwilightAngle
End If
tmp = -Sin(D2R(Param.TwilightAngleIsha))
tmp = tmp - (Sin(D2R(Param.Declination)) * Sin(D2R(cty.Latitude)))
tmp = tmp / (Cos(D2R(Param.Declination)) * Cos(D2R(cty.Latitude)))

Param.VIsha = (1 / 15) * R2D(ArcCos(tmp))

tmp = cty.Latitude - Param.Declination
tmp = Tan(D2R(tmp))
tmp = ArcCot(1 + Abs(tmp))
tmp = Sin(tmp)
tmp = tmp - (Sin(D2R(Param.Declination)) * Sin(D2R(cty.Latitude)))
tmp = tmp / (Cos(D2R(Param.Declination)) * Cos(D2R(cty.Latitude)))

Param.W = (1 / 15) * R2D(ArcCos(tmp))

'Param.W = (1 / 15) * R2D(ArcCos((Sin(ArcCot(1 + Tan((D2R(cty.Latitude) _
- D2R(Param.Declination))))) - (Sin(D2R(Param.Declination)) _
* Sin(D2R(cty.Latitude)))) / (Cos(D2R(Param.Declination)) _
* Cos(D2R(cty.Latitude)))))


Param.x = (1 / 15) * R2D(ArcCos((Sin(ArcCot(2 + Tan(Abs((D2R(cty.Latitude) _
- D2R(Param.Declination)))))) - (Sin(D2R(Param.Declination)) _
* Sin(D2R(cty.Latitude)))) / (Cos(D2R(Param.Declination)) _
* Cos(D2R(cty.Latitude)))))




CalculateTimes.Zuhr = ConvertParams(Param.z)

CalculateTimes.AsrHanafi = ConvertParams(Param.z + Param.x)

CalculateTimes.AsrShafi = ConvertParams(Param.z + Param.W)

CalculateTimes.Fajr = ConvertParams(Param.z - Param.VFajr)

CalculateTimes.Isha = ConvertParams(Param.z + Param.VIsha)

CalculateTimes.Maghrib = ConvertParams(Param.z + Param.U)

CalculateTimes.Sunrise = ConvertParams(Param.z - Param.U)


CalculateTimes.Ishraq = ConvertParams(Param.z - Param.U + 0.27) ' 16 minutes
'+ConvertParams(Param.ReferenceLongitude / 60*cty.TimeZone) ' actually sunrise +15 minutes

CalculateTimes.Duha = ConvertParams(Param.z - (Param.U / 2))
'45 degrees from east horizon


CalculateTimes.DZuhr = CDate(CalculateTimes.Zuhr)

CalculateTimes.DAsrHanafi = CDate(CalculateTimes.AsrHanafi)

CalculateTimes.DAsrShafi = CDate(CalculateTimes.AsrShafi)

CalculateTimes.DFajr = CDate(CalculateTimes.Fajr)

CalculateTimes.DIsha = CDate(CalculateTimes.Isha)

CalculateTimes.DMaghrib = CDate(CalculateTimes.Maghrib)

CalculateTimes.DSunrise = CDate(CalculateTimes.Sunrise)


CalculateTimes.DIshraq = CDate(CalculateTimes.Ishraq)
CalculateTimes.DDuha = CDate(CalculateTimes.Duha)


Exit Function
cterr:
fMainForm.sbStatusBar.Panels(1).Text = "in calculatetimes " & Err.Description & Err.Number

fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next

End Function

Function calcJD(ThisTime As Date, tZone As Double) As Double
On Local Error GoTo cjderr
Dim y As Variant
Dim m As Byte
Dim d As Byte
Dim t As Date
Dim a As Double
Dim b As Double
Dim jdg As Double
Dim jd0 As Double
Dim ThisDay As Date
Dim TimeinHours As Double

ThisDay = FormatDateTime(ThisTime, vbLongDate)
    y = year(ThisDay)
    m = Month(ThisDay)
    d = Day(ThisDay)
    t = FormatDateTime(ThisTime, vbLongTime)
    
If (m <= 2) Then
    y = y - 1
    m = m + 12
End If
     a = Math.Round(y / 100)
     b = 2 - a + Math.Round(a / 4)
    d = d + t / 24
    jdg = Math.Round(365.25 * (y + 4716)) + Math.Round(30.6001 * (m + 1)) + d + b - 1524.5
' Correct the time for UT from local standard time
'- this only sets up calc to give the right jd for 0h UT
    jdg = jdg - Val(city.TimeZone) / 24
    jd0 = Math.Round(jdg + 0.5) - 0.5 ' Julian day 0h UT
    'add localtime-timezone for local JD
    TimeinHours = Hour(t)
    TimeinHours = TimeinHours + Minute(t) / 60
    
    calcJD = jd0 + (TimeinHours - Val(city.TimeZone)) / 24
    If (TimeinHours - Val(city.TimeZone)) >= 24 Then
        calcJD = calcJD - 1
    End If
Exit Function
cjderr:
fMainForm.sbStatusBar.Panels(1).Text = "in calculate JD " & Err.Description
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next

End Function

Function calcET(jday) As Double
On Local Error GoTo ceterr
Dim radian As Double
Dim t As Double
Dim tau As Double
Dim obl As Double
Dim L As Double
Dim m As Double
Dim E As Double
Dim y As Double
Dim ET As Double


        radian = 180 / PI
        t = (jday - 2451545#) / 36525
        tau = t / 10
        obl = 23.439291 - 0.0013 * t
        obl = obl / radian
        L = 280.46644567 + 360007.6982779 * tau + 0.03032 * tau * tau _
            + (tau ^ 3) / 49931 - (tau ^ 4) / 15299
        L = L / radian
        m = 357.05291 + 35999.0503 * t _
            - 0.0001559 * t * t - 0.00000048 * t * t * t
        m = m / radian
        E = 0.01670617 - 0.000042037 * t - 0.0000001236 * t * t
        y = ((Tan(obl / 2)) ^ 2)
        ET = y * Sin(2 * L) - 2 * E * Sin(m) _
            + 4 * E * y * Sin(m) * Cos(2 * L) - y * y / 2 * Sin(4 * L)
    ET = ET * radian * 4
    calcET = ET
    Exit Function
ceterr:
    fMainForm.sbStatusBar.Panels(1).Text = "In calculate ET " & Err.Description & Err.Number
    fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text
    Resume Next
    
End Function

Function ArcCos(x As Double) As Double
On Local Error Resume Next
ArcCos = Atn(-x / (Sqr(Abs(-x * x + 1)))) + 2 * Atn(1)
End Function
Function ArcCot(x As Double) As Double
On Local Error Resume Next
ArcCot = Atn(x) + 2 * Atn(1)
End Function

Function GetDeclination(ThisDay As Date) As Double
On Local Error GoTo gderr
Dim tmp As Double
Dim n As Long
Dim Taud As Double 'day angle

'eqn from wikipedia
'eqn from http://holodeck.st.usm.edu/vrcomputing/vrc_t/tutorials/solar/declination.shtml
n = DateDiff("d", "01-Jan-" & year(ThisDay), ThisDay)
'tmp = 2 * PI / 365
'tmp = tmp * (N + 10)
'tmp = -23.45 * Cos(tmp)

'tmp = tmp * (284 + N + 1)
'tmp = 23.45 * Sin(tmp) '+ 23.45 here

Taud = 2 * PI * (n) / 365
tmp = 0.006918 - 0.39912 * Cos(Taud) + 0.070257 * Sin(Taud) - 0.006758 * Cos(2 * Taud) _
+ 0.000907 * Sin(2 * Taud) - 0.002697 * Cos(3 * Taud) + 0.00148 * Sin(3 * Taud)

'tmp = tmp * D2R(city.Latitude) 'added by me check!!

GetDeclination = R2D(tmp)
Exit Function
gderr:
fMainForm.sbStatusBar.Panels(1).Text = "In getdeclination " & Err.Description & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text
Resume Next

End Function

Function GetCity(cid As Long, LocationTable As String) As CityType
On Local Error GoTo gcerr
Dim rsCity As ADODB.Recordset
Set rsCity = New ADODB.Recordset

 GetCity.CityID = 0
 GetCity.cityname = ""
 GetCity.country = ""
 GetCity.Latitude = 0
 GetCity.Longitude = 0
 GetCity.MSL = 0
 GetCity.Region = ""
 GetCity.TimeZone = 0

rsCity.Open "select * from " & LocationTable & " where CityId=" & cid & ";", Cnn, adOpenKeyset, adLockOptimistic
If rsCity.RecordCount <= 0 Then
    MsgBox "City not found in database"
    Exit Function
End If
rsCity.MoveLast
 GetCity.CityID = rsCity!CityID
 GetCity.cityname = rsCity!city
 GetCity.country = rsCity!country
 GetCity.Latitude = rsCity!Latitude
 GetCity.Longitude = rsCity!Longitude
 'GetCity.isHanafi = Val(rsCity!isHanafi)
 If Not IsNull(rsCity!MSL) Then
    GetCity.MSL = rsCity!MSL
Else
    GetCity.MSL = 0
End If
 If Not IsNull(rsCity!isHanafi) Then
    If rsCity!isHanafi = 0 Then
        GetCity.isHanafi = False 'rsCity!isHanafi
    Else
        GetCity.isHanafi = True
    End If
 Else
    GetCity.isHanafi = False
 End If
 If Not IsNull(rsCity!isNorthernHemisphere) Then
    If rsCity!isNorthernHemisphere = 0 Then
        GetCity.isNorthernHemisphere = False 'rsCity!isNorthernHemisphere
    Else
        GetCity.isNorthernHemisphere = True
    End If
 Else
    GetCity.isNorthernHemisphere = True
 End If

 GetCity.Region = rsCity!Region
 GetCity.TimeZone = Replace(rsCity!TimeZone, ":", ".")

rsCity.Close
Set rsCity = Nothing
Exit Function
gcerr:
fMainForm.sbStatusBar.Panels(1).Text = "in getcity " & Err.Description & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End Function

Sub LoadCurrentCity()
On Local Error GoTo lccerr
Dim rsCurrentLoc As ADODB.Recordset
Set rsCurrentLoc = New ADODB.Recordset

rsCurrentLoc.Open "select * from CurrentLoc order by ID;", Cnn, adOpenKeyset, adLockOptimistic
If rsCurrentLoc.RecordCount <= 0 Then
    MsgBox "Select a City from database"
    'frmLocations.Show
    dlgSelect.Show
    rsCurrentLoc.Close
    Set rsCurrentLoc = Nothing
    Exit Sub
End If
rsCurrentLoc.MoveLast ' last city set current
 city.CityID = rsCurrentLoc!CityID
 city.cityname = rsCurrentLoc!city
 city.country = rsCurrentLoc!country
 city.Latitude = rsCurrentLoc!Latitude
 city.Longitude = rsCurrentLoc!Longitude
 If Not IsNull(rsCurrentLoc!MSL) Then
    city.MSL = rsCurrentLoc!MSL
 Else
    city.MSL = 0
 End If
 If Not IsNull(rsCurrentLoc!isHanafi) Then
    If rsCurrentLoc!isHanafi = 0 Then
        city.isHanafi = False 'rsCurrentLoc!isHanafi
    Else
        city.isHanafi = True
    End If
 Else
    city.isHanafi = False
 End If
 'isHanafi = city.isHanafi
 If Not IsNull(rsCurrentLoc!isNorthernHemisphere) Then
    If rsCurrentLoc!isNorthernHemisphere = 0 Then
        city.isNorthernHemisphere = False 'rsCurrentLoc!isNorthernHemisphere
    Else
        city.isNorthernHemisphere = True
    End If
 Else
    city.isNorthernHemisphere = True
 End If
 
 city.Region = rsCurrentLoc!Region
 city.TimeZone = Replace(rsCurrentLoc!TimeZone, ":", ".")

rsCurrentLoc.Close
Set rsCurrentLoc = Nothing
Exit Sub
lccerr:
fMainForm.sbStatusBar.Panels(1).Text = "In load current city " & Err.Description & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next

End Sub
Sub RefreshWaqt()
On Local Error GoTo rwerr
    Dim lret As Long
    Dim i As Byte
    'Dim lpCIEP As Long
    Waqt = CalculateTimes(city, FormatDateTime(Now, vbGeneralDate))
    'lRet = EnumCalendarInfo(lpCIEP, 2, 6, 32)
    With fMainForm
        For i = 0 To .lblWaqt.Count - 1
            .lblWaqt(i).FontSize = 12
            .lblWaqt(i).FontBold = True
        Next i
        .lblWaqt(0).Caption = "Fajr"
        .lblWaqt(1).Caption = "Sun Rise"
        .lblWaqt(2).Caption = "Zuhr"
        .lblWaqt(3).Caption = "Asr"
        .lblWaqt(4).Caption = "Maghrib"
        .lblWaqt(5).Caption = "Isha"
        .lblWaqt(6).Caption = Waqt.Fajr
        .lblWaqt(7).Caption = Waqt.Sunrise
        .lblWaqt(8).Caption = Waqt.Zuhr
        'isHanafi = city.isHanafi
        If city.isHanafi Then
            .lblWaqt(9).Caption = Waqt.AsrHanafi
        Else
            .lblWaqt(9).Caption = Waqt.AsrShafi
        End If
        .lblWaqt(10).Caption = Waqt.Maghrib
        .lblWaqt(11).Caption = Waqt.Isha
        
        .lblWaqt(12).Caption = "Ishraq"
        .lblWaqt(13).Caption = Waqt.Ishraq
        
        .lblWaqt(14).Caption = "Duha"
        .lblWaqt(15).Caption = Waqt.Duha
                
        .lblLoc = "Location: " & city.cityname & vbCr & _
        city.Region & ", " & city.country
        .lblHijri.Caption = "Hijri: " & GetHijriDate(FormatDateTime(Now, vbGeneralDate))
        
        .lblNextSalat = GetNextsalat()
    End With
Exit Sub
rwerr:
fMainForm.sbStatusBar.Panels(1).Text = "in refresh waqt " & Err.Description & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next

End Sub

Sub bmp_rotate(pic1 As PictureBox, pic2 As PictureBox, ByVal theta As Single)
' bmp_rotate(pic1, pic2, theta)
' Rotate the image in a picture box.
'   pic1 is the picture box with the bitmap to rotate
'   pic2 is the picture box to receive the rotated bitmap
'   theta is the angle of rotation
Dim c1x As Long, c1y As Long
Dim c2x As Long, c2y As Long
Dim a As Single
Dim p1x As Long, p1y As Long
Dim p2x As Long, p2y As Long
Dim n As Long, r   As Long
Dim pic1hDC As Long
Dim pic2hDC As Long
Dim c0 As Long
Dim c1 As Long
Dim c2 As Long
Dim c3 As Long
Dim xret As Long

c1x = pic1.ScaleWidth \ 2
c1y = pic1.ScaleHeight \ 2
c2x = pic2.ScaleWidth \ 2
c2y = pic2.ScaleHeight \ 2

If c2x < c2y Then n = c2y Else n = c2x
n = n - 1
pic1hDC = pic1.hDC
pic2hDC = pic2.hDC

For p2x = 0 To n
  For p2y = 0 To n
    If p2x = 0 Then a = PI / 2 Else a = Atn(p2y / p2x)
    r = Sqr(1 * p2x * p2x + 1 * p2y * p2y)
    p1x = r * Cos(a + theta)
    p1y = r * Sin(a + theta)
    c0 = GetPixel(pic1hDC, c1x + p1x, c1y + p1y)
    c1 = GetPixel(pic1hDC, c1x - p1x, c1y - p1y)
    c2 = GetPixel(pic1hDC, c1x + p1y, c1y - p1x)
    c3 = GetPixel(pic1hDC, c1x - p1y, c1y + p1x)
    If c0 <> -1 Then xret = SetPixel(pic2hDC, c2x + p2x, c2y + p2y, c0)
    If c1 <> -1 Then xret = SetPixel(pic2hDC, c2x - p2x, c2y - p2y, c1)
    If c2 <> -1 Then xret = SetPixel(pic2hDC, c2x + p2y, c2y - p2x, c2)
    If c3 <> -1 Then xret = SetPixel(pic2hDC, c2x - p2y, c2y + p2x, c3)
  Next
  DoEvents
Next
End Sub

Function D2R(Dg As Double) As Double
D2R = Dg * PI / 180
End Function
Function R2D(Rd As Double) As Double
R2D = Rd * 180 / PI
End Function

Function ConvertDateString( _
    ByRef StringIn As String, _
    ByRef OldCalendar As Integer, _
    ByVal NewCalendar As Integer, _
    ByRef NewFormat As String) As String
On Local Error GoTo calerr
    Dim SavedCal As Integer
    Dim d As Date
    Dim s As String
    
    '// Save VBA Calendar setting to restore when finished
    SavedCal = Calendar
    
    '// Convert date to new calendar and format
    Calendar = OldCalendar      ' Change to StringIn calendar
    d = CDate(StringIn)        ' Convert from String to Date
    Calendar = NewCalendar      ' Change to calendar of new string
    s = CStr(d)          ' Convert to short format String
    ConvertDateString = Format(s, NewFormat)      ' Reformat
    
    
    '// Restore VBA Calendar setting
    Calendar = SavedCal
    Exit Function
calerr:
    fMainForm.sbStatusBar.Panels(1).Text = "In convert datestring " & Err.Description & Err.Number
    fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End Function


Function GetHijriDate(GregorianDate As Date) As String '// Convert to Hijri date and return in Long Date format
On Local Error GoTo GHDErr
Dim mnthNo As Byte
Dim lefthalf As String
Dim righthalf As String
    Dim HijriDate As String
    Dim GD As String
    Dim DecrementHijriMonth As Boolean
    
    
    GD = CDate(GregorianDate)
    HijriDate = ConvertDateString(GD, 0, 1, "dd-mmm-yyyy")

lefthalf = Left$(HijriDate, 2) + Settings1.AdjustmentDays 'Hijri adjustment by user
If Val(lefthalf) <= 0 Then
    lefthalf = "30"
    DecrementHijriMonth = True
End If
'************************


DateTime.Calendar = vbCalHijri
HijriDate = CStr(Format(GregorianDate, "Long Date"))
HijriDate = LTrim(RTrim(HijriDate))
mnthNo = Month(CStr(Format(GregorianDate, "Short Date")))
If DecrementHijriMonth Then
    mnthNo = mnthNo - 1
    If mnthNo = 0 Then
        mnthNo = 12
    End If
End If
'lefthalf = Left$(HijriDate, InStr(HijriDate, " "))
righthalf = Mid$(HijriDate, InStr(Len(lefthalf) + 1, HijriDate, " "))
righthalf = Right$(righthalf, 8)

righthalf = Right$(righthalf, 4)
'HijriDate = lefthalf & vbCr & HijriMonths(mnthNo) & " " & righthalf
HijriDate = lefthalf & " " & HijriMonths(mnthNo) & " " & righthalf
DateTime.Calendar = vbCalGreg

'*************************************


GetHijriDate = HijriDate '+ AdjustmentDays : already done above

Exit Function
GHDErr:
fMainForm.sbStatusBar.Panels(1).Text = "In getHijriDate " & Err.Description & Err.Number

fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End Function

Function ConvertParams(Param As Double) As String
On Local Error GoTo cperr
Dim Hr As Integer
Dim Mt As Integer
Dim ampm As String
Dim cp As String
Dim Par As Double
Par = Abs(Param)
Hr = Int(Par)
Mt = (Par - Hr) * 60
If Mt = 60 Then
    Hr = Hr + 1
    Mt = 0
End If
If Hr > 12 Then
    ampm = "PM"
    Hr = Hr - 12
ElseIf Hr = 12 Then
    If Mt = 0 Then
        ampm = "NOON"
    Else
        ampm = "PM"
    End If
Else
    ampm = "AM"
End If
cp = Format(CStr(Hr), "0#") & ":" & Format(CStr(Mt), "0#") & " " & ampm

ConvertParams = cp
Exit Function
cperr:
fMainForm.sbStatusBar.Panels(1).Text = " In convert params " ^ Err.Description & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End Function

Function salatAdan() As Long
On Local Error GoTo saerr
Dim AdanTime As salattype

salatAdan = 0
AdanTime = CalculateTimes(city, Now)
If city.isHanafi Then
    salatAdan = DateDiff("s", AdanTime.AsrHanafi, Format(Now, "Medium Time"))
Else
    salatAdan = DateDiff("s", AdanTime.AsrShafi, Format(Now, "Medium Time"))
End If
If salatAdan <= 60 And salatAdan > 0 Then
    CallAdan (0)
    Exit Function
End If
salatAdan = DateDiff("s", AdanTime.Isha, Format(Now, "Medium Time"))
If salatAdan <= 60 And salatAdan > 0 Then
    CallAdan (0)
    Exit Function
End If
salatAdan = DateDiff("s", AdanTime.Maghrib, Format(Now, "Medium Time"))
If salatAdan <= 60 And salatAdan > 0 Then
    CallAdan (0)
    Exit Function
End If
salatAdan = DateDiff("s", AdanTime.Zuhr, Format(Now, "Medium Time"))
If salatAdan <= 60 And salatAdan > 0 Then
    CallAdan (0)
    Exit Function
End If
salatAdan = DateDiff("s", AdanTime.Fajr, Format(Now, "Medium Time"))
If salatAdan <= 60 And salatAdan >= 0 Then
    CallAdan (1)
    Exit Function
End If
salatAdan = DateDiff("s", AdanTime.Sunrise, Format(Now, "Medium Time"))
If salatAdan <= 60 And salatAdan > 0 Then
    CallWarning "Sunrise", , 0
End If
    
Exit Function

saerr:
fMainForm.sbStatusBar.Panels(1).Text = "in salatadan....." & Err.Description
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next

End Function

Sub CallAdan(AdanType As Byte, Optional adan As String)
On Local Error GoTo adanerr
'Dim ret As Long
Settings1 = GetSettings()
If Settings1.AdanEnabled = False Then
    MsgBox "It is time for Adan. Adan is currently disabled. To hear adan change the settings. "
    Exit Sub
End If
Static ResetTick As Integer

If adan = "" Then
    If AdanType = 0 Then
        adan = Settings1.adan 'App.Path & "\adan\" & "adan.wav"
    Else
        adan = Settings1.FAdan 'App.Path & "\adan\" & "fajradan.wav"
    End If
End If

If AdanCalled Then
    ResetTick = ResetTick + 1
    If ResetTick >= 300 Then 'adan length is 4:37
        AdanCalled = False
        ResetTick = 0
    End If
Else
    If InStr(adan, ":") = 0 Then
        adan = App.Path & "\adan\" & adan
    End If
    'ret = PlaySound(adan, &O0, SND_ASYNC Or SND_FILENAME)
    DoEvents
'    If ret <> 0 Then
'        ret = sndPlaySound(adan, SND_ASYNC Or SND_NODEFAULT)
'    End If
'    If ret <> 0 Then
        With fMainForm.MMControl1
           .Notify = False
           .Wait = True
           .Shareable = False
           .DeviceType = "WaveAudio"
           .FileName = adan

           ' Open the MCI WaveAudio device.
           .Command = "Open"
           
           .Command = "Play"
        End With
 '   End If
    AdanCalled = True
    
End If
Exit Sub
adanerr:
fMainForm.sbStatusBar.Panels(1).Text = " Adan " & Err.Description & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End Sub


Function ClckHnd_rotate(pic As PictureBox, theta As Currency, Optional HandLength As Byte, Optional HandWidth As Byte, Optional HandColor As Long) As Currency
On Local Error GoTo chrerr
Dim cx As Long, cy As Long
'Dim a As Single
Dim p1x As Long, p1y As Long
'Static p2x As Long, p2y As Long
Dim r  As Long
If IsNull(HandWidth) Then
    HandWidth = 2
End If
If IsNull(HandLength) Then
    HandLength = 80
End If
If IsNull(HandColor) Then
    HandColor = &HFF00FF
End If
pic.ForeColor = HandColor
pic.DrawWidth = HandWidth
cx = 0
cy = 0
'If p2x = 0 Then
 '   p2x = cx
  '  p2y = cy
'End If
'a = 0 'PI / 60 'Atn(p2y / p2x)
r = HandLength 'Sqr(1 * p2x * p2x + 1 * p2y * p2y)
p1x = r * Cos(theta)
p1y = r * Sin(theta)
'pic.DrawMode = 7
'pic.Line (cx, cy)-(p2x, p2y), &HFFFFFF
'pic.DrawMode = 7 'xor
pic.Line (cx, cy)-(p1x, p1y)
'p2x = p1x
'p2y = p1y
Exit Function
chrerr:
fMainForm.sbStatusBar.Panels(1).Text = " In clock hand rotate " & Err.Description & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End Function

Function GetNextsalat() As String
On Local Error GoTo gnserr
Dim Nextsalat As salattype
Dim TimeRemaining As Double
Dim tmp As Double
Dim salatname As String
Dim Hr As String
Dim Mt As String
Dim ThisTime As Date

ThisTime = Format(Now, "Medium Time")


Nextsalat = CalculateTimes(city, Now) 'CalculateTimes(city, Now)

If city.isHanafi Then
    tmp = DateDiff("n", ThisTime, Nextsalat.DAsrHanafi)
Else
    tmp = DateDiff("n", ThisTime, Nextsalat.DAsrShafi)
End If
salatname = "ASR"
If tmp > 0 Then
    TimeRemaining = tmp
Else
    TimeRemaining = 99999999#
End If
tmp = DateDiff("n", ThisTime, Nextsalat.DMaghrib)
If tmp > 0 And TimeRemaining > tmp Then
    salatname = "MAGRIB"
    TimeRemaining = tmp
End If
tmp = DateDiff("n", ThisTime, Nextsalat.DIsha)
If tmp > 0 And TimeRemaining > tmp Then
    salatname = "ISHA"
    TimeRemaining = tmp
End If
tmp = DateDiff("n", ThisTime, Nextsalat.DFajr)
If tmp < 0 Then ' required to check fajr
    tmp = 24 * 60 + tmp
End If
If tmp > 0 And TimeRemaining > tmp Then
    salatname = "FAJR"
    TimeRemaining = tmp
End If
tmp = DateDiff("n", ThisTime, Nextsalat.DSunrise)
If tmp > 0 And TimeRemaining > tmp Then
    salatname = "SUNRISE"
    TimeRemaining = tmp
End If
tmp = DateDiff("n", ThisTime, Nextsalat.DZuhr)
If tmp > 0 And TimeRemaining > tmp Then
    salatname = "ZUHR"
    TimeRemaining = tmp
End If

Hr = CStr(Int(TimeRemaining / 60))
Mt = TimeRemaining Mod 60
If Val(Hr) > 1 Then
    Hr = Hr & " Hours and "
ElseIf Val(Hr) = 0 Then
    Hr = ""
Else
    Hr = Hr & " Hour and "
End If
If Val(Mt) > 1 Then
    Mt = Mt & " Minutes for "
ElseIf Val(Mt) = 0 Then
    Mt = ""
Else
    Mt = Mt & " Minute for "
End If
If (Val(Mt) = 0 And Val(Hr) = 0) Then
    Hr = ""
    Mt = ""
End If
GetNextsalat = Hr & Mt & salatname
Exit Function
gnserr:
fMainForm.sbStatusBar.Panels(1).Text = "In get next salat " & Err.Description & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End Function

Sub Install()
On Local Error GoTo inserr
If QueryValue("Software\Microsoft\Windows\CurrentVersion\Run", "salat") = "" Then
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Run", "salat", App.Path & "\" & App.exename & ".exe", 1&
End If
'If QueryValue("Software\Microsoft\Windows\CurrentVersion\RunServices", "salat") = "" Then
 '  SetKeyValue "Software\Microsoft\Windows\CurrentVersion\RunServices", "salat", App.Path & "\" & App.EXEName & ".exe q", 1&
'End If
'If QueryValueUser(".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Run", "salat") = "" Then
 '  SetKeyValueUser ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Run", "salat", App.Path & "\" & App.EXEName & ".exe", 1&
'End If
Exit Sub
inserr:
fMainForm.sbStatusBar.Panels(1).Text = "In Install " & Err.Description & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End Sub

Public Sub HideMe(frm As Form)
ShowWindow frm.hwnd, SW_HIDE
HideFromTaskList frm
'ReleaseCapture
'showcursor (True)
End Sub

Public Sub ShowMe(frm As Form)
Dim hWndScreen As Long
ShowWindow frm.hwnd, SW_SHOW
'showcursor (False)
hWndScreen = GetDesktopWindow()
ShowWindow hWndScreen, SW_SHOW 'SW_HIDE
'SetCapture hWndScreen
End Sub

Public Function UpTime(Optional surfsec As Double) As String
On Local Error GoTo uterr
If surfsec = 0 Then
    Dim UTime As Date
    
    UTime = CDate(Now) - starttime
    UpTime = str$(Hour(UTime)) & " Hrs"
    UpTime = UpTime & ":" & str$(Minute(UTime)) & " Min"
    UpTime = UpTime & ":" & str$(Second(UTime)) & " Sec"
Else
    Dim Hr As Long, Mts As Long
    
    Hr = Int(surfsec / 3600)
    
    Mts = Int(surfsec / 60) Mod 60
    
    UpTime = str$(Hr) & " Hrs"
    UpTime = UpTime & ":" & str$(Mts) & " Min"
    UpTime = UpTime & ":" & str$(surfsec Mod 60) & " Secs"
End If
Exit Function
uterr:
fMainForm.sbStatusBar.Panels(1).Text = "In Uptime " & Err.Description & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End Function

Sub HideFromTaskList(frm As Form)
Dim OwnerhWnd As Long
Dim ret As Long

frm.Visible = False
OwnerhWnd = GetWindow(frm.hwnd, GW_OWNER)
ret = ShowWindow(OwnerhWnd, SW_HIDE)

End Sub

Sub UnInstall()
On Local Error GoTo unerr
If QueryValue("Software\Microsoft\Windows\CurrentVersion\Run", "salat") <> "" Then
    RegDeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\salat"
End If
If QueryValue("Software\Microsoft\Windows\CurrentVersion\RunServices", "salat") <> "" Then
   RegDeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices\salat"
End If
If QueryValueUser(".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Run", "salat") <> "" Then
   RegDeleteKey HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Run\salat"
End If
Exit Sub
unerr:
fMainForm.sbStatusBar.Panels(1).Text = "In uninstall " & Err.Description & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End Sub

Function Dec2Deg(dec As Currency) As degreeformat
On Local Error Resume Next
Dim sign As Integer
sign = Sgn(dec)
Dec2Deg.Degrees = Int(Abs(dec))
Dec2Deg.Degrees = sign * Dec2Deg.Degrees
Dec2Deg.Minutes = (Abs(dec) - Int(Abs(dec))) * 60

End Function

Function Deg2Dec(Deg As degreeformat) As Currency
On Local Error Resume Next
Dim sign As Integer
sign = Sgn(Deg.Degrees)
Deg2Dec = Abs(Deg.Degrees) + Abs(Deg.Minutes) / 60
Deg2Dec = Deg2Dec * sign

End Function

Function GetNextPic(PicPath As String) As String
On Local Error GoTo slideshowerr
Dim PicFile
Static visited As Byte
Dim PicName
Dim i As Byte
Dim ext As String
If Right$(PicPath, 1) <> "\" Then
    PicPath = PicPath & "\"
End If
PicName = Dir(PicPath, vbArchive Or vbNormal Or vbReadOnly Or vbHidden)    ' Retrieve the first entry.

'Do While PicName <> ""   ' Start the loop.
   ' Ignore the current directory and the encompassing directory.
   If PicName <> "." And PicName <> ".." Then
      ' Use bitwise comparison to make sure PicName is NOT a directory.
      If (GetAttr(PicPath & PicName) And vbDirectory) <> vbDirectory Then
         ext = LCase$(Right$(PicName, 4)) ' Display entry only if it
         If ext = ".bmp" Or ext = ".jpg" Or ext = ".gif" Or ext = ".wmf" Then
            For i = 1 To visited
                PicName = Dir
            Next
            GetNextPic = PicPath & PicName
'            If PicName <> oldpic Then Exit Do
            'If (Timer Mod 13) = 0 Then Exit Do
         End If
      End If   ' it represents a graphics file.
   End If
        'PicName = Dir     ' Get next entry.
'Loop
visited = visited + 1
If visited >= 16 Then
    visited = 1
End If
Exit Function
slideshowerr:
fMainForm.sbStatusBar.Panels(1).Text = "In get next pic " & Err.Description
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next

End Function

Function GetField(cmb As ComboBox, tbl As String, fild As String) As Long
On Local Error GoTo getfielderr
Dim i As Long
Dim present As Boolean
Dim fieldval As String
Dim rsFild As ADODB.Recordset
Set rsFild = New ADODB.Recordset
GetField = 0
Screen.MousePointer = vbHourglass
cmb.Clear
If InStr(LCase$(tbl), "select") = 0 Then
    rsFild.Open "select * from " & tbl & " order by '" & fild & "';", Cnn, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
Else
    rsFild.Open tbl, Cnn, adOpenKeyset, adLockOptimistic
End If
GetField = rsFild.RecordCount

If rsFild.RecordCount > 0 Then rsFild.MoveFirst
present = False
Do While Not rsFild.EOF
    fieldval = rsFild.Fields(fild).Value
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = fieldval Then
            present = True
            Exit For
        End If
    Next i
    If Not present Then
        cmb.AddItem fieldval
    End If
    present = False
    rsFild.MoveNext
    fMainForm.lblLoc.Caption = "Please wait..... Loading City ..." & fieldval
    DoEvents
Loop
fMainForm.lblLoc.Caption = "Loaded " & rsFild.RecordCount & " Cities of the world..."
DoEvents
rsFild.Close
DoEvents
Set rsFild = Nothing
cmb.Text = "Select " & fild
Screen.MousePointer = vbNormal
Exit Function
getfielderr:
fMainForm.sbStatusBar.Panels(1).Text = "In get field " & Err.Description & " " & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next

End Function

Sub AddCurrentCity()
On Local Error GoTo addcurcityerr
Dim rsCurrentLoc As ADODB.Recordset
Set rsCurrentLoc = New ADODB.Recordset

'add to currentloc
rsCurrentLoc.Open "select * from CurrentLoc", Cnn, adOpenStatic, adLockOptimistic
    rsCurrentLoc.AddNew
  
  rsCurrentLoc!CityID = city.CityID 'cityid is not unique in currentloc
  rsCurrentLoc!city = city.cityname
  rsCurrentLoc!country = city.country
  rsCurrentLoc!Latitude = city.Latitude
  rsCurrentLoc!Longitude = city.Longitude
  rsCurrentLoc!Region = city.Region
  rsCurrentLoc!TimeZone = city.TimeZone
  rsCurrentLoc!MSL = city.MSL
  rsCurrentLoc!isHanafi = city.isHanafi
  rsCurrentLoc!isNorthernHemisphere = city.isNorthernHemisphere
rsCurrentLoc.Update
rsCurrentLoc.Close


DoEvents

Set rsCurrentLoc = Nothing

Exit Sub
addcurcityerr:
fMainForm.sbStatusBar.Panels(1).Text = "In add current city " & Err.Description & " " & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next

End Sub
Sub LoadHijriMonths()
    HijriMonths(1) = "Muharram"
    HijriMonths(2) = "Safar"
    HijriMonths(3) = "Rabi Ul Awwal"
    HijriMonths(4) = "Rabi Ul Akhar"
    HijriMonths(5) = "Jamad Ul Awwal"
    HijriMonths(6) = "Jamad Ul Akhar"
    HijriMonths(7) = "Rajab"
    HijriMonths(8) = "Sha'aban"
    HijriMonths(9) = "Ramadan"
    HijriMonths(10) = "Shawwal"
    HijriMonths(11) = "Dhul Qa'ad"
    HijriMonths(12) = "Dhul Hijja"

End Sub

Sub CreateCrescent(frm As Form)
Const RGN_DIFF = 4

Dim outer_rgn As Long
Dim inner_rgn As Long
Dim combined_rgn As Long
Dim wid As Single
Dim hgt As Single
Dim border_width As Single
Dim title_height As Single

    If frm.WindowState = vbMinimized Then Exit Sub
    
    ' Create the regions.
    wid = frm.ScaleX(frm.Width, vbTwips, vbPixels)
    hgt = frm.ScaleY(frm.Height, vbTwips, vbPixels)
    outer_rgn = CreateEllipticRgn( _
    0, 0, wid, hgt)
    
    border_width = (wid - frm.ScaleWidth) / 2
    title_height = hgt - border_width - frm.ScaleHeight
    inner_rgn = CreateEllipticRgn( _
        wid * 0.55, hgt * 0.45, _
        wid, 0)

    ' Subtract the inner region from the outer.
    combined_rgn = CreateEllipticRgn(0, 0, 0, 0)
    CombineRgn combined_rgn, outer_rgn, _
        inner_rgn, RGN_DIFF
    
    ' Restrict the window to the region.
    SetWindowRgn frm.hwnd, combined_rgn, True

    DeleteObject combined_rgn
    DeleteObject inner_rgn
    DeleteObject outer_rgn


End Sub


Sub SetTransparency(frm As Form, Transparency As Byte)
    Dim lOldStyle As Long

    'bTrans = Transparency '200 '128
    lOldStyle = GetWindowLong(frm.hwnd, GWL_EXSTYLE)
    SetWindowLong frm.hwnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED 'Or WS_EX_TRANSPARENT 'last term for mouse events to go thro the form
    SetLayeredWindowAttributes frm.hwnd, 0, Transparency, LWA_ALPHA

End Sub

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
    As Long

    If Topmost = True Then 'Make the window topmost
       SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
          0, FLAGS)
    Else
       SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
          0, 0, FLAGS)
       SetTopMostWindow = False
    End If
 End Function


Sub StopAdan()
Dim ret As Long
ret = sndPlaySound("", SND_ASYNC Or SND_FILENAME)
If ret <> 0 Then

        With fMainForm.MMControl1
           .Notify = False
           .Wait = False
           .Shareable = False
           .DeviceType = "WaveAudio"
           '.FileName = ""
           ' Open the MCI WaveAudio device.
           .Command = "Close"
           '.Command = "Play"
        End With

End If
AdanCalled = True

End Sub

Sub ViewReport(CityID As Long, isHanafi As Boolean, Optional TimePeriod As Integer, Optional startDate As Date)
On Local Error GoTo vderr
Dim fl As String
Dim i As Integer
Dim ReportCity As CityType
Dim tmpDate As Date
Dim tmpsalat As salattype

If TimePeriod = 0 Then TimePeriod = 365
If Not (IsDate(startDate)) Then
    startDate = "01-Jan-" & year(Now)
End If
ReportCity = GetCity(CityID, "Locations")

fl = App.Path & "\Salat_Times_at_City_" & ReportCity.cityname & ".html"

Close #1
Open fl For Output As #1
Print #1, "<!--Created by Mohamed Iqbal Pallipurath-->"
Print #1, "<html>"
Print #1, "<head><Title>Salat Times for City " & ReportCity.cityname & " </Title></head>"
Print #1, "<body onload=print() >"
Print #1, "<center>"
Print #1, "<table border=" & Chr(34) & "1" & Chr(34) & " >"

Print #1, "<tr><center>Salat Times for City " & ReportCity.cityname & "&nbsp;Latitude: " & ReportCity.Latitude & "&nbsp;Longitude: " & ReportCity.Longitude & "</center></tr>"
Print #1, "<tr><td>Date</td>"
Print #1, " <td>Fajr</td>"
Print #1, " <td>Sunrise</td>"
Print #1, " <td>Ishraq</td>"
Print #1, " <td>Duha</td>"
Print #1, " <td>Zuhr</td>"
Print #1, " <td>Asr (Shafi)</td>"
Print #1, " <td>Asr (Hanafi)</td>"
Print #1, " <td>Magrib</td>"
Print #1, " <td>Isha</td></tr>"

Print #1, "<font face=" & Chr(34) & " Arial " & Chr(34) & "size=" & Chr(34) & 6 & Chr(34) & "style=" & Chr(34) & "font-family: Arial" & Chr(34) & ">"

For i = 0 To TimePeriod
    tmpDate = DateAdd("d", i, startDate)
    tmpsalat = CalculateTimes(ReportCity, tmpDate)
    Print #1, "<tr>"
    Print #1, "<td>" & tmpDate & "</td>"
    
    Print #1, "<td>" & tmpsalat.Fajr & "</td>"
    Print #1, "<td>" & tmpsalat.Sunrise & "</td>"
    Print #1, "<td>" & tmpsalat.Ishraq & "</td>"
    Print #1, "<td>" & tmpsalat.Duha & "</td>"
    Print #1, "<td>" & tmpsalat.Zuhr & "</td>"
    Print #1, "<td>" & tmpsalat.AsrShafi & "</td>"
    Print #1, "<td>" & tmpsalat.AsrHanafi & "</td>"
    Print #1, "<td>" & tmpsalat.Maghrib & "</td>"
    Print #1, "<td>" & tmpsalat.Isha & "</td>"
    Print #1, "</tr>"
Next i
Print #1, "</font>"
Print #1, "<hr>"


Print #1, "</table><br><br><font size=1>Generated by Salat &copy 2020-2025 <a href=" & Chr(34) & "mailto:mohamediqbalp@gmail.com" & Chr(34) & ">Mohamed Iqbal Pallipurath </a></font></body></html>"
Print #1, "</center>"
Print #1, "</body>"
Print #1, "</html>"
Close #1

DoEvents
OpenDefaultBrowser fl


Exit Sub
vderr:
fMainForm.sbStatusBar.Panels(1).Text = "In View Report " & Err.Description & " No " & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next

End Sub


Function salatWarning() As Currency
On Local Error GoTo swerr
Dim AdanTime As salattype

AdanTime = CalculateTimes(city, Now)
If city.isHanafi Then
    If DateDiff("n", AdanTime.AsrHanafi, Format(Now, "Medium Time")) = (-1) * Settings1.WarningMinutes Then
        CallWarning "Asr", , Settings1.WarningMinutes
    End If
Else
    If DateDiff("n", AdanTime.AsrShafi, Format(Now, "Medium Time")) = (-1) * Settings1.WarningMinutes Then
        CallWarning "Asr", , Settings1.WarningMinutes
    End If
End If
If DateDiff("n", AdanTime.Isha, Format(Now, "Medium Time")) = (-1) * Settings1.WarningMinutes Then
    CallWarning "Isha", , Settings1.WarningMinutes
End If
If DateDiff("n", AdanTime.Maghrib, Format(Now, "Medium Time")) = (-1) * Settings1.WarningMinutes Then
    CallWarning "Maghrib", , Settings1.WarningMinutes
End If
If DateDiff("n", AdanTime.Zuhr, Format(Now, "Medium Time")) = (-1) * Settings1.WarningMinutes Then
    CallWarning "Zuhr", , Settings1.WarningMinutes
End If
If DateDiff("n", AdanTime.Fajr, Format(Now, "Medium Time")) = (-1) * Settings1.WarningMinutes Then
    CallWarning "Fajr", , Settings1.WarningMinutes
End If
If DateDiff("n", AdanTime.Sunrise, Format(Now, "Medium Time")) = (-1) * Settings1.WarningMinutes Then
    CallWarning "Sunrise", , Settings1.WarningMinutes
End If
Exit Function
swerr:
fMainForm.sbStatusBar.Panels(1).Text = "in salatwarning...." & Err.Description
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End Function

Sub CallWarning(salatname As String, Optional Warning As String, Optional WMinutes As Integer)
On Local Error GoTo cwerr
Dim WarningWav As String

WarningWav = App.Path & "\adan\" & "BISMI.WAV"
If Settings1.WarningEnabled = False Then
   MsgBox "It is now " & Settings1.WarningMinutes & " minutes for adan of " & salatname & " . Warnings are disabled. To enable warnings, select options."
    Exit Sub
End If

Static ResetTick As Integer
Static WarningCalled As Boolean
If WMinutes = 0 Then
    WMinutes = 10
End If
If Warning = "" Then
    Warning = WMinutes & " minutes remain for "
Else
    Warning = WMinutes & Warning
End If

If WarningCalled Then
    ResetTick = ResetTick + 1
    If ResetTick >= 60 * WMinutes Then 'Warninglength is 10mts
        WarningCalled = False
        ResetTick = 0
    End If
Else
    'ReadIt Warning & salatname
    
        With fMainForm.MMControl1
           .Notify = False
           .Wait = True
           .Shareable = False
           .DeviceType = "WaveAudio"
           .FileName = WarningWav
           ' Open the MCI WaveAudio device.
           .Command = "Open"
           .Command = "Play"
        End With
    
    
    WarningCalled = True
End If
Exit Sub
cwerr:
fMainForm.sbStatusBar.Panels(1).Text = "in callwarning..." & Err.Description

fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next

End Sub

Sub ReadIt(wStr As String)
On Local Error GoTo readiterr
'Dim t2s As DirectSS
'Set t2s = New DirectSS

'fMainForm.DirectSS1.Initialized = True
'fMainForm.DirectSS1.Select (1)
'fMainForm.DirectSS1.Speed = 140
'fMainForm.DirectSS1.Gender (0)
'fMainForm.DirectSS1.Speaker (2)
'fMainForm.DirectSS1.Speak wStr
'fMainForm.DirectSS1.Speak wStr
Exit Sub
readiterr:
    fMainForm.sbStatusBar.Panels(1).Text = "In read it " & Err.Description & Err.Number
    fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text
Resume Next

End Sub

Function isUpDateAvailable(CurrentVersion As String, UpdatePath As String) As Boolean
'frmBrowser.brwWebBrowser.ExecWB OLECMDID_ALLOWUILESSSAVEAS, OLECMDEXECOPT_DONTPROMPTUSER, "http://iqsoft.co.in/salat/salat.html", App.Path & "\test.html"

End Function

Function CityExists(Cntry As String, Regn As String, cty As String) As Boolean
Dim rstmp As ADODB.Recordset
Set rstmp = New ADODB.Recordset
CityExists = False
rstmp.Open "select * from Locations where Country='" & Cntry & "' and region='" & _
    Regn & "' and city='" & cty & "';", Cnn, adOpenStatic, adLockOptimistic
If rstmp.RecordCount <= 0 Then
    CityExists = False
Else
    CityExists = True
End If
rstmp.Close
Set rstmp = Nothing

End Function

Sub AddLocation(cty As CityType)
Dim rsLoc As ADODB.Recordset
Set rsLoc = New ADODB.Recordset

'first add to locations
Set rsLoc = New ADODB.Recordset
rsLoc.Open "select * from Locations order by cityid", Cnn, adOpenStatic, adLockOptimistic
rsLoc.MoveLast
city.CityID = rsLoc!CityID + 1
    rsLoc.AddNew
  rsLoc!CityID = city.CityID
  rsLoc!city = cty.cityname
  rsLoc!country = cty.country
  rsLoc!Latitude = cty.Latitude
  rsLoc!Longitude = cty.Longitude
  rsLoc!Region = cty.Region
  rsLoc!TimeZone = cty.TimeZone
  rsLoc!MSL = cty.MSL
  rsLoc!isHanafi = cty.isHanafi
  rsLoc!isNorthernHemisphere = cty.isNorthernHemisphere
rsLoc.Update
DoEvents
city.CityID = rsLoc!CityID 'for adding to currentloc

rsLoc.Close
Set rsLoc = Nothing

End Sub

Function GetQibla(lat As Double, lon As Double) As Double
GetQibla = 0
Dim q As Double
q = Atn(Sin(D2R(QiblaLon - lon)) / (Cos(D2R(lat)) _
* Tan(D2R(QiblaLat)) - Sin(D2R(lat)) * Cos(D2R(QiblaLon - lon))))

'q = Atn(Sin(D2R(QiblaLon - lon)) / (Sin(D2R(lat)) _
* Cos(D2R(QiblaLon - lon)) - Cos(D2R(lat)) * Tan(D2R(QiblaLat))))


GetQibla = q

End Function

Function GetSettings() As settingstype

If QueryValue("Software\iqsoft\" & App.exename & "\settings", "adanenabled") = "" Then
    CreateNewKey "Software\iqsoft\" & App.exename & "\" & "settings", HKEY_LOCAL_MACHINE
    SetKeyValue "Software\iqsoft\" & App.exename & "\" & "settings", "adanenabled", "1", 1&
    SetKeyValue "Software\iqsoft\" & App.exename & "\" & "settings", "warningenabled", "1", 1&
    SetKeyValue "Software\iqsoft\" & App.exename & "\" & "settings", "transparency", "128", 1&
    SetKeyValue "Software\iqsoft\" & App.exename & "\" & "settings", "adan", App.Path & "\adan\adan.wav", 1&
    SetKeyValue "Software\iqsoft\" & App.exename & "\" & "settings", "fajradan", App.Path & "\adan\fajradan.wav", 1&
    SetKeyValue "Software\iqsoft\" & App.exename & "\" & "settings", "twilightangle", "18", 1&
    SetKeyValue "Software\iqsoft\" & App.exename & "\" & "settings", "twilightanglefajr", "18", 1&
    SetKeyValue "Software\iqsoft\" & App.exename & "\" & "settings", "twilightangleisha", "18", 1&
    SetKeyValue "Software\iqsoft\" & App.exename & "\" & "settings", "warningminutes", "10", 1&
    SetKeyValue "Software\iqsoft\" & App.exename & "\" & "settings", "clockface", App.Path & "\clock.bmp", 1&
    SetKeyValue "Software\iqsoft\" & App.exename & "\" & "settings", "secondhandcolor", &HFF0101, 1&
    SetKeyValue "Software\iqsoft\" & App.exename & "\" & "settings", "minutehandcolor", &HFFAA01, 1&
    SetKeyValue "Software\iqsoft\" & App.exename & "\" & "settings", "hourhandcolor", &H101FF, 1&
    SetKeyValue "Software\iqsoft\" & App.exename & "\" & "settings", "warningtxt", " Minutes to Adan ", 1&
    SetKeyValue "Software\iqsoft\" & App.exename & "\" & "settings", "adjustmentdays", 0, 1&

Else
    GetSettings.AdanEnabled = QueryValue("Software\iqsoft\" & App.exename & "\settings", "adanenabled")
    GetSettings.WarningEnabled = QueryValue("Software\iqsoft\" & App.exename & "\settings", "warningenabled")
    GetSettings.Transparency = QueryValue("Software\iqsoft\" & App.exename & "\settings", "transparency")
    GetSettings.TwilightAngle = QueryValue("Software\iqsoft\" & App.exename & "\settings", "twilightangle")
    GetSettings.TwilightAngleFajr = QueryValue("Software\iqsoft\" & App.exename & "\settings", "twilightanglefajr")
    GetSettings.TwilightAngleIsha = QueryValue("Software\iqsoft\" & App.exename & "\settings", "twilightangleisha")
    GetSettings.FAdan = QueryValue("Software\iqsoft\" & App.exename & "\settings", "fajradan")
    GetSettings.adan = QueryValue("Software\iqsoft\" & App.exename & "\settings", "adan")
    GetSettings.WarningMinutes = QueryValue("Software\iqsoft\" & App.exename & "\settings", "warningminutes")
    GetSettings.ClockFace = QueryValue("Software\iqsoft\" & App.exename & "\settings", "clockface")
    GetSettings.HourHandColor = QueryValue("Software\iqsoft\" & App.exename & "\settings", "hourhandcolor")
    GetSettings.MinuteHandColor = QueryValue("Software\iqsoft\" & App.exename & "\settings", "minutehandcolor")
    GetSettings.SecondHandColor = QueryValue("Software\iqsoft\" & App.exename & "\settings", "secondhandcolor")
    GetSettings.WarningTxt = QueryValue("Software\iqsoft\" & App.exename & "\settings", "warningtxt")
    GetSettings.AdjustmentDays = QueryValue("Software\iqsoft\" & App.exename & "\settings", "adjustmentdays")
End If

End Function

Function GetFiles(FilePath As String, FileExtension As String) As String
On Local Error GoTo slideshowerr
Dim PicFile
Static visited As Byte
Dim PicName
Dim i As Byte
Dim ext As String
Static firstfile As String
If Right$(FilePath, 1) <> "\" Then
    FilePath = FilePath & "\"
End If
PicName = Dir(FilePath, vbArchive Or vbNormal Or vbReadOnly Or vbHidden)    ' Retrieve the first entry.

'Do While PicName <> ""   ' Start the loop.
   ' Ignore the current directory and the encompassing directory.
   If PicName <> "." And PicName <> ".." Then
      ' Use bitwise comparison to make sure PicName is NOT a directory.
      If (GetAttr(FilePath & PicName) And vbDirectory) <> vbDirectory Then
         ext = LCase$(Right$(PicName, 4)) ' Display entry only if it
         If ext = FileExtension Then
            For i = 1 To visited
                PicName = Dir
            Next
            GetFiles = PicName
'            If PicName <> oldpic Then Exit Do
            'If (Timer Mod 13) = 0 Then Exit Do
         End If
      End If   ' it represents a graphics file.
   End If
        'PicName = Dir     ' Get next entry.
'Loop
If visited = 0 Then
    firstfile = GetFiles
End If
visited = visited + 1
If visited > 0 Then
   If GetFiles = "" Then
        visited = 0
    End If
End If
Exit Function
slideshowerr:
fMainForm.sbStatusBar.Panels(1).Text = Err.Description
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next

End Function

Sub SetSettings(ProgSettings As settingstype)
With ProgSettings ' set default values
    If .adan = "" Then
        .AdanEnabled = 1
        .WarningEnabled = True
        .Transparency = 128
        .adan = "adan.wav"
        .FAdan = "fajradan.wav"
        .TwilightAngle = 18
        .TwilightAngleFajr = 18
        .TwilightAngleIsha = 18
        .WarningMinutes = 10
        .ClockFace = App.Path & "\Clocks\clock.bmp"
        .SecondHandColor = &HFF0101
        .MinuteHandColor = &HFF0101
        .HourHandColor = &H101FF
        .AdjustmentDays = 0
    End If
End With

SetKeyValue "Software\iqsoft\" & App.exename & "\" & _
"settings", "twilightangle", ProgSettings.TwilightAngle, 1&
SetKeyValue "Software\iqsoft\" & App.exename & "\" & _
"settings", "twilightanglefajr", ProgSettings.TwilightAngleFajr, 1&
SetKeyValue "Software\iqsoft\" & App.exename & "\" & _
"settings", "twilightangleisha", ProgSettings.TwilightAngleIsha, 1&


SetKeyValue "Software\iqsoft\" & App.exename & "\" & _
"settings", "adan", ProgSettings.adan, 1&
SetKeyValue "Software\iqsoft\" & App.exename & "\" & _
"settings", "adanenabled", ProgSettings.AdanEnabled, 1&
SetKeyValue "Software\iqsoft\" & App.exename & "\" & _
"settings", "fajradan", ProgSettings.FAdan, 1&
SetKeyValue "Software\iqsoft\" & App.exename & "\" & _
"settings", "transparency", ProgSettings.Transparency, 1&
SetKeyValue "Software\iqsoft\" & App.exename & "\" & _
"settings", "warningenabled", ProgSettings.WarningEnabled, 1&
SetKeyValue "Software\iqsoft\" & App.exename & "\" & _
"settings", "warningminutes", ProgSettings.WarningMinutes, 1&
SetKeyValue "Software\iqsoft\" & App.exename & "\" & _
"settings", "clockface", ProgSettings.ClockFace, 1&
SetKeyValue "Software\iqsoft\" & App.exename & "\" & _
"settings", "hourhandcolor", ProgSettings.HourHandColor, 1&
SetKeyValue "Software\iqsoft\" & App.exename & "\" & _
"settings", "minutehandcolor", ProgSettings.MinuteHandColor, 1&
SetKeyValue "Software\iqsoft\" & App.exename & "\" & _
"settings", "secondhandcolor", ProgSettings.SecondHandColor, 1&
SetKeyValue "Software\iqsoft\" & App.exename & "\" & _
"settings", "adjustmentdays", ProgSettings.AdjustmentDays, 1&


End Sub

Public Sub OpenDefaultBrowser(weburl As String)
Dim r As Long
   r = ShellExecute(0, "open", weburl, 0, 0, 1)

End Sub
Public Sub SaveDefaultBrowser(fname As String)
Dim r As Long
   r = ShellExecute(0, "save", fname, 0, 0, 1)

End Sub

Public Sub PlaysndDx(SoundPath As String)
'    Dim BuffDesc As DSBufferDesc
'    ' get enumeration object
'    Set DSEnum = DX.GetDSEnum
'
'    ' select the first sound device, and create the Direct Sound object
'    Set DIS = DX.DirectSoundCreate(DSEnum.GetGuid(1))
'
'    ' Set the Cooperative Level to normal
'    DIS.setCooperativeLevel Me.hwnd, DSSCL_NORMAL
'
'    ' load the wave file, and create the buffer for it
'    Set DSSecBuffer = DIS.CreateSoundBufferFromFile(SoundPath, BuffDesc)
'
'    DSSecBuffer.play DSBPLAY_DEFAULT
End Sub

Public Function daysInFebruary(year As Long) As Long
  If (year > 1582) Then
     If (year Mod 4 = 0) And ((Not (year Mod 100 = 0)) Or (year Mod 400 = 0)) Then
        daysInFebruary = 29
     Else
        daysInFebruary = 28
     End If
  Else
    If (year Mod 4 = 0) Then
        daysInFebruary = 29
    Else
        daysInFebruary = 28
    End If
  End If
End Function

Public Function GetDaysArray(MonthNum As Byte) As Byte
Dim i As Integer
Dim DaysArray(12) As Byte
For i = 1 To MonthNum - 1
    DaysArray(i) = 31
    If (i = 4 Or i = 6 Or i = 9 Or i = 11) Then DaysArray(i) = 30
    If (i = 2) Then DaysArray(i) = 29
Next i
GetDaysArray = DaysArray(MonthNum)
End Function

Function GetNearestCardinalDir(qibla As Double) As String
Dim minDiff As Double
minDiff = 0
If minDiff < qibla Then
    GetNearestCardinalDir = R2D(qibla) & " degrees from north to east (Clockwise)"
End If
minDiff = Abs(PI / 2 - qibla)

End Function


Function synconv(tdate As String, toHijri As Boolean) As String
'assumes tdate is in format yyyy-mm-dd
'toHijri = true is to islamic conversion
'toHijri = false is from islamic conversion
'default conYear = 354.366
'default conMont = 29.5258
Dim conYear As Double
Dim conMonth As Double
Dim dd As Long
Dim JD As Long
Dim L, n, i, j, k, tMonth, tDay, tYear As Long
Dim OutGoing As String
Dim td As String
Dim HYear, hymos As Integer
Dim HMonth, Hday As Integer
Dim dmod As Double

conYear = 354.366
conMonth = 29.5285

If toHijri Then
dd = DateDiff("d", "1/1/1800", tdate)
JD = dd + 2378497

'the following section of code borrowed from:
'http://www.couprie.f2s.com/calmath/
'many thanks
L = JD - 1948440 + 10632
n = ((L - 1) \ 10631)
L = L - 10631 * n + 354
j = (((10985 - L) \ 5316)) * (((50 * L) \ 17719)) + ((L \ 5670)) * (((43 * L) \ 15238))
L = L - (((30 - j) \ 15)) * (((17719 * j) \ 50)) - ((j \ 16)) * (((15238 * j) \ 43)) + 29
tMonth = ((24 * L) \ 709)
tDay = L - ((709 * tMonth) \ 24)
tYear = 30 * n + j - 30

OutGoing = tYear & "-" & tMonth & "-" & tDay

Else
td = Split(tdate, "-")

HYear = CInt(Left(td, 1))
hymos = (HYear - 1) * 12
dmod = (((HYear - 1214) Mod 10) * 0.0000433) + conMonth
HMonth = Round((CInt(Mid(td, 1, 1)) + hymos) * dmod, 0)
Hday = CInt(Mid(td, 2, 1)) + HMonth

JD = Hday + 1948440

'the following section of code borrowed from:
'http://www.couprie.f2s.com/calmath/
'many thanks
If (JD > 2299160) Then
L = JD + 68569
n = ((4 * L) \ 146097)
L = L - ((146097 * n + 3) \ 4)
i = ((4000 * (L + 1)) \ 1461001)
L = L - ((1461 * i) \ 4) + 31
j = ((80 * L) \ 2447)
tDay = L - ((2447 * j) \ 80)
L = (j \ 11)
tMonth = j + 2 - 12 * L
tYear = 100 * (n - 49) + i + L
Else
j = JD + 1402
k = ((j - 1) \ 1461)
L = j - 1461 * k
n = ((L - 1) \ 365) - (L \ 1461)
i = L - 365 * n + 30
j = ((80 * i) \ 2447)
tDay = i - ((2447 * j) \ 80)
i = (j \ 11)
tMonth = j + 2 - 12 * i
tYear = 4 * k + n + i - 4716
End If

OutGoing = tYear & "-" & tMonth & "-" & tDay

End If

synconv = OutGoing

End Function

Function GetHijriDayMonthYear(GregorianDate As Date, HijriDay As String, HijriMonth As String, HijriYear As String) As String
On Local Error GoTo GHDMYErr
Dim mnthNo As Byte
Dim HijriDate As String
Dim GD As String
Dim DecrementHijriMonth As Boolean
GD = CDate(GregorianDate)
HijriDate = ConvertDateString(GD, 0, 1, "dd-mmm-yyyy")

HijriDay = Left$(HijriDate, 2) + Settings1.AdjustmentDays 'Hijri adjustment by user
If Val(HijriDay) <= 0 Then
    HijriDay = "30"
    DecrementHijriMonth = True
End If
'************************


DateTime.Calendar = vbCalHijri
HijriDate = CStr(Format(GregorianDate, "Long Date"))
HijriDate = LTrim(RTrim(HijriDate))
mnthNo = Month(CStr(Format(GregorianDate, "Short Date")))
If DecrementHijriMonth Then
    mnthNo = mnthNo - 1
    If mnthNo = 0 Then
        mnthNo = 12
    End If
End If
'HijriDay = Left$(HijriDate, InStr(HijriDate, " "))
HijriYear = Mid$(HijriDate, InStr(Len(HijriDay) + 1, HijriDate, " "))
HijriYear = Right$(HijriYear, 8)

HijriYear = Right$(HijriYear, 4)
HijriMonth = HijriMonths(mnthNo)
HijriDate = HijriDay & " " & HijriMonths(mnthNo) & " " & HijriYear
DateTime.Calendar = vbCalGreg

'*************************************


GetHijriDayMonthYear = HijriDate


Exit Function
GHDMYErr:
fMainForm.sbStatusBar.Panels(1).Text = "In getHijriDate " & Err.Description & Err.Number

fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next

End Function

Function getGregorianDate(HijriDate As String) As String
On Local Error GoTo GGDErr
Dim Gdate As String
Dim HD As String
If Len(HijriDate) > 10 Then
    DateTime.Calendar = vbCalHijri
    HD = CStr(getShortHijriDate(HijriDate))
Else
    HD = HijriDate
End If
Gdate = ConvertDateString(HD, vbCalHijri, vbCalGreg, "dd-mmm-yyyy")


DateTime.Calendar = vbCalGreg
Gdate = CStr(Format(Gdate, "Long Date"))
Gdate = LTrim(RTrim(Gdate))
DateTime.Calendar = vbCalHijri

'*************************************


getGregorianDate = Gdate


Exit Function
GGDErr:
fMainForm.sbStatusBar.Panels(1).Text = "In getHijriDate " & Err.Description & Err.Number

fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next


End Function

Function getShortHijriDate(longHijriDate As String) As String
Dim i As Integer
Dim HMonthNo As Integer
Dim HDayNo As Integer
Dim HYear As Integer
Dim pos As Integer
Dim HMonth As String

LoadHijriMonths
For i = 1 To 12
    pos = InStr(longHijriDate, HijriMonths(i))
    If pos > 0 Then
        HMonthNo = i
        Exit For
    End If
Next i
If HMonthNo < 10 Then
    HMonth = "0" & Trim(str(HMonthNo))
Else
    HMonth = str(HMonthNo)
End If
longHijriDate = Trim(longHijriDate)
HDayNo = Left$(longHijriDate, 2)
HYear = Mid$(longHijriDate, Len(longHijriDate) - 3)
getShortHijriDate = str(HDayNo) & "-" & HMonth & "-" & Trim(str(HYear))

End Function

'
'   The function below implements Paul Schlyter's simplification
'   of van Flandern and Pulkkinen's method for finding the geocentric
'   ecliptic positions of the Moon to an accuracy of about 1 to 4 arcmin.
'
'   I can probably reduce the number of variables, and there must
'   be a quicker way of declaring variables!
'
'   The VBA trig functions have been used throughout for speed,
'   note how the atan function returns values in domain -pi to pi
'
Public Function smoon(ByVal d As Double, Index As Integer) As Double
    Dim Nm As Double, im As Double, wm As Double, am As Double, ecm As Double, _
    Mm As Double, em As Double, Ms As Double, ws As Double, xv As Double, _
    yv As Double, vm As Double, rm As Double, x As Double, y As Double, _
    z As Double, lon As Double, lat As Double, ls As Double, lm As Double, _
    dm As Double, F As Double, dlon As Double, dlat As Double
    '   Paul's routine uses a slightly different definition of
    '   the day number - I adjust for it below. Remember that VBA
    '   defaults to 'pass by reference' so this change in d
    '   will be visible to other functions unless you set d to 'ByVal'
    '   to force it to be passed by value!
    d = d + 1.5
    '   moon elements
    Nm = range360(125.1228 - 0.0529538083 * d) * rads
    im = 5.1454 * rads
    wm = range360(318.0634 + 0.1643573223 * d) * rads
    am = 60.2666  '(Earth radii)
    ecm = 0.0549
    Mm = range360(115.3654 + 13.0649929509 * d) * rads
    '   position of Moon
    em = Mm + ecm * Sin(Mm) * (1# + ecm * Cos(Mm))
    xv = am * (Cos(em) - ecm)
    yv = am * (Sqr(1# - ecm * ecm) * Sin(em))
    vm = ATan2(xv, yv)
    '   If vm < 0 Then vm = tpi + vm
    rm = Sqr(xv * xv + yv * yv)
    x = rm * (Cos(Nm) * Cos(vm + wm) - Sin(Nm) * Sin(vm + wm) * Cos(im))
    y = rm * (Sin(Nm) * Cos(vm + wm) + Cos(Nm) * Sin(vm + wm) * Cos(im))
    z = rm * (Sin(vm + wm) * Sin(im))
    '   moons geocentric long and lat
    lon = ATan2(x, y)
    If lon < 0 Then lon = tpi + lon
    lat = Atn(z / Sqr(x * x + y * y))
    '   mean longitude of sun
    ws = range360(282.9404 + 0.0000470935 * d) * rads
    Ms = range360(356.047 + 0.9856002585 * d) * rads
    '   perturbations
    '   first calculate arguments below,
    'Ms, Mm             Mean Anomaly of the Sun and the Moon
    'Nm                 Longitude of the Moon's node
    'ws, wm             Argument of perihelion for the Sun and the Moon
    ls = Ms + ws       'Mean Longitude of the Sun  (Ns=0)
    lm = Mm + wm + Nm  'Mean longitude of the Moon
    dm = lm - ls        'Mean elongation of the Moon
    F = lm - Nm        'Argument of latitude for the Moon
    ' then add the following terms to the longitude
    ' note amplitudes are in degrees, convert at end
    Select Case Index
        Case 1  '   distance terms earth radii
            rm = rm - 0.58 * Cos(Mm - 2 * dm)
            rm = rm - 0.46 * Cos(2 * dm)
            smoon = rm
        Case 2  '   latitude terms
            dlat = -0.173 * Sin(F - 2 * dm)
            dlat = dlat - 0.055 * Sin(Mm - F - 2 * dm)
            dlat = dlat - 0.046 * Sin(Mm + F - 2 * dm)
            dlat = dlat + 0.033 * Sin(F + 2 * dm)
            dlat = dlat + 0.017 * Sin(2 * Mm + F)
            smoon = lat * degs + dlat
        Case 3  '   longitude terms
            dlon = -1.274 * Sin(Mm - 2 * dm)        '(the Evection)
            dlon = dlon + 0.658 * Sin(2 * dm)       '(the Variation)
            dlon = dlon - 0.186 * Sin(Ms)           '(the Yearly Equation)
            dlon = dlon - 0.059 * Sin(2 * Mm - 2 * dm)
            dlon = dlon - 0.057 * Sin(Mm - 2 * dm + Ms)
            dlon = dlon + 0.053 * Sin(Mm + 2 * dm)
            dlon = dlon + 0.046 * Sin(2 * dm - Ms)
            dlon = dlon + 0.041 * Sin(Mm - Ms)
            dlon = dlon - 0.035 * Sin(dm)           '(the Parallactic Equation)
            dlon = dlon - 0.031 * Sin(Mm + Ms)
            dlon = dlon - 0.015 * Sin(2 * F - 2 * dm)
            dlon = dlon + 0.011 * Sin(Mm - 4 * dm)
            smoon = lon * degs + dlon
    End Select
End Function
'


Private Function range360(x)
'
'   returns an angle x in the range 0 to 360
'   used to prevent the huge values of degrees
'   that you get from mean longitude formulas
'
'   this function is private to this module,
'   you won't find it in the Function Wizard,
'   and you can't use it on a spreadsheet.
'   If you want it on the spreadsheet, just remove
'   the 'private' keyword above.
'
range360 = x - 360 * Int(x / 360)
End Function


Function ATan2(y As Double, x As Double) As Double
ATan2 = Atn(y / x)

End Function
