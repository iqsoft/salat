Attribute VB_Name = "codelib"
Option Explicit
Option Compare Text
Public Rapps As Integer
Public User1 As usertype
Public NoOfTries As Integer
Public starttime As Date
Public LoginSucceeded As Boolean
Public AdminLogin As Boolean
Public week As Integer
Public UserPwd As String, AdminPwd As String
Public pathToPix As String


Public Declare Function RasGetConnectStatus Lib "rasapi32" _
Alias "RasGetConnectStatusA" _
(ByVal hrasconn As Long, ByVal connectstatus As Long) As Long

Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

   Const SW_SHOW = 5

   Const SW_HIDE = 0

'To kill Apps (time/date properties)

   Public Declare Function WaitForSingleObject Lib "kernel32" _
         (ByVal hHandle As Long, _
         ByVal dwMilliseconds As Long) As Long

      Public Declare Function FindWindow Lib "user32" _
         Alias "FindWindowA" _
         (ByVal lpClassName As String, _
         ByVal lpWindowName As String) As Long

      Public Declare Function PostMessage Lib "user32" _
         Alias "PostMessageA" _
         (ByVal hwnd As Long, _
         ByVal wMsg As Long, _
         ByVal wParam As Long, _
         ByVal lParam As Long) As Long

      Public Declare Function IsWindow Lib "user32" _
         (ByVal hwnd As Long) As Long

      'Constants used by the API functions
      Const WM_CLOSE = &H10
      Const INFINITE = &HFFFFFFFF


      Public Type LUID
         UsedPart As Long
         IgnoredForNowHigh32BitPart As Long
      End Type

      Public Type TOKEN_PRIVILEGES
         PrivilegeCount As Long
         TheLuid As LUID
         Attributes As Long
      End Type

      ' Beginning of Code
      Public Const EWX_SHUTDOWN As Long = 1
      Public Const EWX_FORCE As Long = 4
      Public Const EWX_REBOOT = 2

      Public Declare Function ExitWindowsEx Lib "user32" ( _
         ByVal dwOptions As Long, ByVal dwReserved As Long) As Long

      Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
      
      Public Declare Function OpenProcessToken Lib "advapi32" ( _
         ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, _
         TokenHandle As Long) As Long
      
      Public Declare Function LookupPrivilegeValue Lib "advapi32" _
         Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, _
         ByVal lpName As String, lpLuid As LUID) As Long
      
      Public Declare Function AdjustTokenPrivileges Lib "advapi32" ( _
         ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
         NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
         PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
      

      Declare Function TerminateProcess Lib "kernel32" _
         (ByVal hProcess As Long, _
         ByVal uExitCode As Long) As Long
 


Public Const SPI_SCREENSAVERRUNNING = 97&
      
Public Declare Function SystemParametersInfo Lib "user32" _
          Alias "SystemParametersInfoA" _
         (ByVal uAction As Long, _
          ByVal uParam As Long, _
         ByVal lpvparam As Boolean, _
          ByVal fuWinIni As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Type restrictedapptype
    title As String
    stime As Date
    etime As Date
    week As String
End Type
'user defined type required by Shell_NotifyIcon API call
'Public Type NOTIFYICONDATA
' cbSize As Long
' hwnd As Long
' uId As Long
' uFlags As Long
' uCallBackMessage As Long
' hIcon As Long
' szTip As String * 64
'End Type


'
''constants required by Shell_NotifyIcon API call:
'Public Const NIM_ADD = &H0
'Public Const NIM_MODIFY = &H1
'Public Const NIM_DELETE = &H2
'Public Const NIF_MESSAGE = &H1
'Public Const NIF_ICON = &H2
'Public Const NIF_TIP = &H4
'Public Const WM_MOUSEMOVE = &H200
'Public Const WM_LBUTTONDOWN = &H201      'Button down
'Public Const WM_LBUTTONUP = &H202       'Button up
'Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
'Public Const WM_RBUTTONDOWN = &H204     'Button down
'Public Const WM_RBUTTONUP = &H205       'Button up
'Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

'Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Public nid As NOTIFYICONDATA
'Public Declare Function showcursor Lib "user32" (ByVal bShow As Long) As Long

Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
   (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) _
   As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, _
   ByVal wCmd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) _
   As Long

Public Const GW_HWNDNEXT = 2
Public Const GW_OWNER = 4

Public Declare Function FormatMessage Lib "kernel32" _
    Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, _
    ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Public Const FORMAT_MESSAGE_FROM_STRING = &H400
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF


Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
'Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public LogFile As String
Public LicensetoKill As Boolean
Public Rapp(31) As restrictedapptype

Type regtype
    regcode As String * 10
End Type
Public fRegcode As regtype

Type companytype
    cname As String
    phone As String
    kgst As String
    cst As String
    address As String
    website As String
    webusername As String
    webpassword As String
End Type
Type usertype
    uname As String
    password As String
    starttime As Date
    EndTime As Date
    surftime As Date
    charge As Currency
    question As String
    answer As String
    supervisoraccess As Boolean
End Type


Public Function reg(exename As String, Optional expirydate As String, Optional numberofuses As Long) As Boolean
On Local Error GoTo regerr
'Dim fl As String
Dim dd As Date
Dim tmp As String
Dim pmi As String
Dim pwd As String
Dim Registered As Boolean
Dim regcode As String
Dim Used As Long

reg = False
'fl = App.Path & "\" & exename & ".reg"
'Open fl For Random As #1
'Get #1, 1, fRegcode

If expirydate = "" Then
    expirydate = "01-Jan-2006"
End If
If numberofuses = 0 Then
    numberofuses = 40
End If
regcode = QueryValue("Software\iqsoft\" & exename & "\regcode", "regcode")
If regcode = "" Then
    Registered = False
ElseIf regcode = EncryptIt("bismilla", "user") Then
    Registered = True
    reg = True
'    fRegcode.regcode = regcode
 '   Put #1, 1, fRegcode
    Exit Function
Else
    Registered = False
End If
If QueryValue("Software\iqsoft\" & exename & "\regcode", "used") = "" Then
    'If fRegcode.regcode = "" Then
        Used = 0
    'Else
'        reg = False
        
 '       Exit Function
    'End If
    CreateNewKey "Software\iqsoft\" & exename & "\regcode", HKEY_LOCAL_MACHINE
Else
    Used = Val(EncryptIt(QueryValue("Software\iqsoft\" & exename & "\regcode", "used"), "bismilla"))
End If
If QueryValueCurrentUser("Software\iqsoft\" & exename & "\regcode", "used") = "" Then
    'If fRegcode.regcode = "" Then
        Used = 0
    'Else
'        reg = False
        
 '       Exit Function
    'End If
    CreateNewKey "Software\iqsoft\" & exename & "\regcode", HKEY_CURRENT_USER
Else
    Used = Val(EncryptIt(QueryValueCurrentUser("Software\iqsoft\" & exename & "\regcode", "used"), "bismilla"))
End If

SetKeyValue "Software\iqsoft\" & exename & "\regcode", "used", EncryptIt(str(Used + 1), "bismilla"), 1&
SetKeyValueCurrentUser "Software\iqsoft\" & exename & "\regcode", "used", EncryptIt(str(Used + 1), "bismilla"), 1&


pmi = Weekday(Now) & day(Now) & month(Now) & year(Now) 'LoadResString(105)
dd = DateDiff("d", CDate(expirydate), Now)
If Not Registered Then
If (dd > 0) Or Used > numberofuses Then
    tmp = LoadResString(103)
    MsgBox tmp
    tmp = InputBox(LoadResString(104))
    If tmp = EncryptIt("bismilla", pmi) Then
        'ok
        reg = True
        CreateNewKey "Software\iqsoft\" & exename & "\regcode", HKEY_LOCAL_MACHINE
        SetKeyValue "Software\iqsoft\" & exename & "\regcode", "regcode", EncryptIt("bismilla", "user"), 1&
        
        CreateNewKey "Software\iqsoft\" & exename & "\regcode", HKEY_CURRENT_USER
        SetKeyValueCurrentUser "Software\iqsoft\" & exename & "\regcode", "regcode", EncryptIt("bismilla", "user"), 1&
        LogIt exename, "Registration code supplied"
    Else
        reg = False
        LogIt exename, "Expired and no registration code"
        KillApp exename
    End If
Else
        reg = True
     '   fRegcode.regcode = EncryptIt(Str(Used + 1), "bismilla")
End If
End If
'Put #1, 1, fRegcode
'Close #1
pmi = ""
tmp = ""
Exit Function
regerr:
MsgBox Err.Description & Err.Number
Resume Next
End Function

Public Sub SetOnce(exename)
On Local Error GoTo setonceerr 'Resume Next
Dim UserPwd As String, AdminPwd As String
DoEvents
If QueryValue("Software\iqsoft\" & exename & "\admin", "password") = "" Then
    CreateNewKey "Software\iqsoft\" & exename & "\admin", HKEY_LOCAL_MACHINE
    CreateNewKey "Software\iqsoft\" & exename & "\user", HKEY_LOCAL_MACHINE
    
    UserPwd = "2( >>#>%" 'LoadResString(101)
    AdminPwd = "*(0&$)" 'LoadResString(102)
    SetKeyValue "Software\iqsoft\" & exename & "\admin", "password", AdminPwd, 1&
    SetKeyValue "Software\iqsoft\" & exename & "\admin", "question", "admin", 1&
    SetKeyValue "Software\iqsoft\" & exename & "\admin", "answer", "yes", 1&
    SetKeyValue "Software\iqsoft\" & exename & "\admin", "supervisoraccess", 1, 1&
    SetKeyValue "Software\iqsoft\" & exename & "\admin", "pathtopix", App.Path & "\waha\", 1&

    SetKeyValue "Software\iqsoft\" & exename & "\user", "password", UserPwd, 1&
    SetKeyValue "Software\iqsoft\" & exename & "\user", "question", "admin", 1&
    SetKeyValue "Software\iqsoft\" & exename & "\user", "answer", "no", 1&
    SetKeyValue "Software\iqsoft\" & exename & "\user", "supervisoraccess", 0, 1&
    SetKeyValue "Software\iqsoft\" & exename & "\user", "pathtopix", App.Path & "\waha\", 1&
End If
If QueryValueCurrentUser("Software\iqsoft\" & exename & "\admin", "password") = "" Then
    CreateNewKey "Software\iqsoft\" & exename & "\admin", HKEY_CURRENT_USER
    CreateNewKey "Software\iqsoft\" & exename & "\user", HKEY_CURRENT_USER

    SetKeyValueCurrentUser "Software\iqsoft\" & exename & "\admin", "password", AdminPwd, 1&
    SetKeyValueCurrentUser "Software\iqsoft\" & exename & "\admin", "question", "admin", 1&
    SetKeyValueCurrentUser "Software\iqsoft\" & exename & "\admin", "answer", "yes", 1&
    SetKeyValueCurrentUser "Software\iqsoft\" & exename & "\admin", "supervisoraccess", 1, 1&
    SetKeyValueCurrentUser "Software\iqsoft\" & exename & "\admin", "pathtopix", App.Path & "\waha\", 1&

    SetKeyValueCurrentUser "Software\iqsoft\" & exename & "\user", "password", UserPwd, 1&
    SetKeyValueCurrentUser "Software\iqsoft\" & exename & "\user", "question", "admin", 1&
    SetKeyValueCurrentUser "Software\iqsoft\" & exename & "\user", "answer", "no", 1&
    SetKeyValueCurrentUser "Software\iqsoft\" & exename & "\user", "supervisoraccess", 0, 1&
    SetKeyValueCurrentUser "Software\iqsoft\" & exename & "\user", "pathtopix", App.Path & "\waha\", 1&
End If
pathToPix = QueryValue("software\iqsoft\" & exename & "\pix", "pixpath")
If pathToPix = "" Then
    CreateNewKey "software\iqsoft\" & exename & "\pix", HKEY_LOCAL_MACHINE
    SetKeyValue "software\iqsoft\" & exename & "\pix", "pixpath", "c:\windows\", 1&
End If
If QueryValue("software\iqsoft\" & exename, "smtp") = "" Then
    SetKeyValue "software\iqsoft\" & exename, "smtp", "gmail.com", 1&
End If

If QueryValue("software\iqsoft\" & exename, "fromaddress") = "" Then
    SetKeyValue "software\iqsoft\" & exename, "fromaddress", "mohamediqbalp@gmail.com", 1&
End If
Exit Sub
setonceerr:
MsgBox Err.Number & Err.Description
MsgBox ApiError(Err.Number)
Resume Next
End Sub


Sub CenterObj(mama As Object, kid As Object)
On Local Error Resume Next
DoEvents
kid.Left = (mama.Width - kid.Width) / 2
kid.Top = (mama.Height - kid.Height) / 2
End Sub

Public Function KillApp(WCaption As String) As String
Dim hWindow As Long
 Dim lngResult As Long
 Dim lngReturnValue As Long

 hWindow = FindWindow(vbNullString, WCaption)
 lngReturnValue = PostMessage(hWindow, WM_CLOSE, vbNull, vbNull)
 lngResult = WaitForSingleObject(hWindow, INFINITE)

 'Does the handle still exist?
 If IsWindow(hWindow) = 0 Then
    'The handle still exists. Use the TerminateProcess function
    'to close all related processes to this handle.
    KillApp = WCaption & " Handle still exists."
    Dim pid As Long
    Dim hProcess As Long
    'hProcess = GetCurrentProcess
    'Call TerminateProcess(hProcess, 0&)
        Call GetWindowThreadProcessId(hWindow, pid)
           
           ' Open the process with all access.
           
           hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
           
           ' Terminate the process.
           
           Call TerminateProcess(hProcess, 0&)

        
 Else
    'Handle does not exist.
  KillApp = WCaption & " Program closed."
 End If

End Function

Public Sub waitfor(sec As Double)
On Local Error Resume Next
Dim EndTime As Date
    EndTime = DateAdd("s", sec, Now)
    Do Until Now > EndTime
       DoEvents
    Loop
End Sub

Sub LogIt(logfilename As String, Optional Status As String, Optional usrname As String, Optional usrpwd As String)
On Local Error Resume Next
'Dim t As New AppMessages.Logging
Dim LogData As String
Dim pwd As String
LogFile = logfilename
If Right(LogFile, 4) <> ".log" Then
    LogFile = LogFile & ".log"
End If
If usrpwd <> "" And usrname <> "" Then
    pwd = EncryptIt(usrpwd, usrname)
End If
LogData = "" & vbCrLf
LogData = LogData & "Start ************************************" & vbCrLf
LogData = LogData & "User Name: " & usrname & vbCrLf
LogData = LogData & Format$(Now) & vbCrLf
LogData = LogData & pwd & vbCrLf
'LogData = LogData & "Loginsucceeded=" & LoginSucceeded & vbCrLf
'LogData = LogData & "Adminlogin=" & AdminLogin & vbCrLf
LogData = LogData & "Week=" & GetWeek(Now) & vbCrLf
LogData = LogData & Status & vbCrLf
LogData = LogData & "End **************************************" & vbCrLf
't.StartLog App.Path & "\" & LogFile, mLogToFile

't.AddLogEvent LogData, vbLogEventTypeInformation
'Set t = Nothing
App.StartLogging LogFile, vbLogToFile
App.LogEvent LogData, vbLogEventTypeInformation

End Sub


Public Function EncryptIt(plaintext As String, key As String) As String
Dim x As Long, L As Long
Dim Char As Integer
Dim ptext As String
' plaintext = the string you wish to encrypt or decrypt.
' key = the password with which to encrypt the string.
EncryptIt = ""
ptext = plaintext
If ptext = "" Or key = "" Then Exit Function
key = UCase$(key)
L = Len(key)
For x = 1 To Len(ptext)
   Char = Asc(Mid$(key, (x Mod L) - L * ((x Mod L) = 0), 1))
   Mid$(ptext, x, 1) = Chr$(Asc(Mid$(ptext, x, 1)) Xor Char)
Next x
EncryptIt = ptext
End Function


Public Function SetValueEx(ByVal hKey As Long, sValueName As String, _
lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    Select Case lType
        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, _
                                           lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, _
lType, lValue, 4)
        End Select
End Function

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As _
String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        ' For strings
        Case REG_SZ:
            sValue = String(cch, 0)
lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, _
sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch - 1)
            Else
                vValue = Empty
            End If
        ' For DWORDS
        Case REG_DWORD:
lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, _
lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:
       QueryValueEx = lrc
       Exit Function
QueryValueExError:
       Resume QueryValueExExit
   End Function




Public Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
       Dim hNewKey As Long         'handle to the new key
       Dim lRetVal As Long         'result of the RegCreateKeyEx function

       lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
                 vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                 0&, hNewKey, lRetVal)
       RegCloseKey (hNewKey)
End Sub

Public Sub SetKeyValue(sKeyName As String, sValueName As String, _
   vValueSetting As Variant, lValueType As Long)
       Dim lRetVal As Long         'result of the SetValueEx function
       Dim hKey As Long         'handle of open key

       'open the specified key
       lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, _
                              KEY_ALL_ACCESS, hKey)
       lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
       RegCloseKey (hKey)
   End Sub

   Public Function QueryValue(sKeyName As String, sValueName As String) As Variant
       Dim lRetVal As Long         'result of the API functions
       Dim hKey As Long         'handle of opened key
       Dim vValue As Variant      'setting of queried value
       Dim lpwD As Variant
       Dim lpC As String
       Dim dwO As Long
       Dim lpSA As Long
       
       lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, _
   KEY_ALL_ACCESS, hKey)
       lRetVal = RegCreateKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, _
    lpC, dwO, KEY_ALL_ACCESS, lpSA, hKey, lpwD)
       lRetVal = QueryValueEx(hKey, sValueName, vValue)
       QueryValue = vValue
       'MsgBox ApiError(lRetVal)
       RegCloseKey (hKey)
   End Function


Public Sub SetKeyValueUser(sKeyName As String, sValueName As String, _
   vValueSetting As Variant, lValueType As Long)
       Dim lRetVal As Long         'result of the SetValueEx function
       Dim hKey As Long         'handle of open key

       'open the specified key
       lRetVal = RegOpenKeyEx(HKEY_USERS, sKeyName, 0, _
                              KEY_ALL_ACCESS, hKey)
       lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
       RegCloseKey (hKey)
   End Sub

Public Sub SetKeyValueCurrentUser(sKeyName As String, sValueName As String, _
   vValueSetting As Variant, lValueType As Long)
       Dim lRetVal As Long         'result of the SetValueEx function
       Dim hKey As Long         'handle of open key
       Dim lpwD As Variant
       Dim lpC As String
       Dim dwO As Long
       Dim lpSA As Long
       

       'open the specified key
       lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, _
                              KEY_ALL_ACCESS, hKey)
       lRetVal = RegCreateKeyEx(HKEY_CURRENT_USER, sKeyName, 0, _
    lpC, dwO, KEY_ALL_ACCESS, lpSA, hKey, lpwD)
If lRetVal <> ERROR_SUCCESS Then
    ApiError (lRetVal)
End If
       lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
       RegCloseKey (hKey)
   End Sub

   Public Function QueryValueUser(sKeyName As String, sValueName As String) As Variant
       Dim lRetVal As Long         'result of the API functions
       Dim hKey As Long         'handle of opened key
       Dim vValue As Variant      'setting of queried value

       lRetVal = RegOpenKeyEx(HKEY_USERS, sKeyName, 0, _
   KEY_ALL_ACCESS, hKey)
       lRetVal = QueryValueEx(hKey, sValueName, vValue)
       QueryValueUser = vValue
       RegCloseKey (hKey)
   End Function

   Public Function QueryValueCurrentUser(sKeyName As String, sValueName As String) As Variant
       Dim lRetVal As Long         'result of the API functions
       Dim hKey As Long         'handle of opened key
       Dim vValue As Variant      'setting of queried value

       lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, _
   KEY_ALL_ACCESS, hKey)
       lRetVal = QueryValueEx(hKey, sValueName, vValue)
       QueryValueCurrentUser = vValue
       RegCloseKey (hKey)
   End Function



Public Function NumToText(dblVal As Double) As String
    Static Ones(0 To 9) As String
    Static Teens(0 To 9) As String
    Static Tens(0 To 9) As String
    Static Thousands(0 To 4) As String
    Static bInit As Boolean
    Dim i As Integer, bAllZeros As Boolean, bShowThousands As Boolean
    Dim strVal As String, strBuff As String, strTemp As String
    Dim nCol As Integer, nChar As Integer

    'Only handles positive values
    If dblVal < 0 Then
        dblVal = -1 * dblVal
    End If

    If bInit = False Then
        'Initialize array
        bInit = True
        Ones(0) = "zero"
        Ones(1) = "one"
        Ones(2) = "two"
        Ones(3) = "three"
        Ones(4) = "four"
        Ones(5) = "five"
        Ones(6) = "six"
        Ones(7) = "seven"
        Ones(8) = "eight"
        Ones(9) = "nine"
        Teens(0) = "ten"
        Teens(1) = "eleven"
        Teens(2) = "twelve"
        Teens(3) = "thirteen"
        Teens(4) = "fourteen"
        Teens(5) = "fifteen"
        Teens(6) = "sixteen"
        Teens(7) = "seventeen"
        Teens(8) = "eighteen"
        Teens(9) = "nineteen"
        Tens(0) = ""
        Tens(1) = "ten"
        Tens(2) = "twenty"
        Tens(3) = "thirty"
        Tens(4) = "forty"
        Tens(5) = "fifty"
        Tens(6) = "sixty"
        Tens(7) = "seventy"
        Tens(8) = "eighty"
        Tens(9) = "ninety"
        Thousands(0) = ""
        Thousands(1) = "thousand"   'US numbering
        Thousands(2) = "million"
        Thousands(3) = "billion"
        Thousands(4) = "trillion"
    End If
    'Trap errors
    On Error GoTo NumToTextError
    'Get fractional part
    Dim frac As Double
    frac = CDbl((dblVal - Int(dblVal)) * 100)
    If frac >= 0.01 Then
        strBuff = "and " & NumToText(frac) & " paise "          '"/100"
    End If
    'Convert rest to string and process each digit
    strVal = CStr(Int(dblVal))
    'Non-zero digit not yet encountered
    bAllZeros = True
    'Iterate through string
    For i = Len(strVal) To 1 Step -1
        'Get value of this digit
        nChar = Val(Mid$(strVal, i, 1))
        'Get column position
        nCol = (Len(strVal) - i) + 1
        'Action depends on 1's, 10's or 100's column
        Select Case (nCol Mod 3)
            Case 1  '1's position
                bShowThousands = True
                If i = 1 Then
                    'First digit in number (last in loop)
                    strTemp = Ones(nChar) & " "
                ElseIf Mid$(strVal, i - 1, 1) = "1" Then
                    'This digit is part of "teen" number
                    strTemp = Teens(nChar) & " "
                    i = i - 1   'Skip tens position
                ElseIf nChar > 0 Then
                    'Any non-zero digit
                    strTemp = Ones(nChar) & " "
                Else
                    'This digit is zero. If digit in tens and hundreds column
                    'are also zero, don't show "thousands"
                    bShowThousands = False
                    'Test for non-zero digit in this grouping
                    If Mid$(strVal, i - 1, 1) <> "0" Then
                        bShowThousands = True
                    ElseIf i > 2 Then
                        If Mid$(strVal, i - 2, 1) <> "0" Then
                            bShowThousands = True
                        End If
                    End If
                    strTemp = ""
                End If
                'Show "thousands" if non-zero in grouping
                If bShowThousands Then
                    If nCol > 1 Then
                        strTemp = strTemp & Thousands(nCol \ 3)
                        If bAllZeros Then
                            strTemp = strTemp & " "
                        Else
                            strTemp = strTemp & ", "
                        End If
                    End If
                    'Indicate non-zero digit encountered
                    bAllZeros = False
                End If
                strBuff = strTemp & strBuff
            Case 2  '10's position
                If nChar > 0 Then
                    If Mid$(strVal, i + 1, 1) <> "0" Then
                        strBuff = Tens(nChar) & "-" & strBuff
                    Else
                        strBuff = Tens(nChar) & " " & strBuff
                    End If
                End If
            Case 0  '100's position
                If nChar > 0 Then
                    strBuff = Ones(nChar) & " hundred " & strBuff
                End If
        End Select
    Next i
    'Convert first letter to upper case
    strBuff = UCase$(Left$(strBuff, 1)) & Mid$(strBuff, 2)
EndNumToText:
    'Return result
    NumToText = strBuff
    Exit Function
NumToTextError:
    strBuff = "#Error#"
    Resume EndNumToText
End Function

Public Function SpellNumber(Num As String) As String
SpellNumber = ""
Dim ch As String
Dim Degree As Long
Dim Place As String
Dim GT9 As Boolean
Dim GT19 As Boolean
Dim WasTeen As Boolean
Dim Nu As String
Dim ChOld As String, Ch2Dig As String
Dim DecPos As Integer
Dim StrLen As Integer

ChOld = ""
DecPos = InStr(1, Num, ".", 1)
 While Num <> ""
    If DecPos > 0 Then
        Degree = (InStr(1, Num, ".", 1) - 1)
        StrLen = Len(Num)
    Else
        Degree = Len(Num)
    End If
    ch = Left$(Num, 1)
    Ch2Dig = Left$(Num, 2)
    GT9 = False
    GT19 = False
    Place = ""
    
    Select Case Degree
        Case Is > 9
            SpellNumber = " Number beyond present capacity. "
        Case 9
            Place = "Crore "
            GT9 = True
        Case 8
            Place = "Crore "
            GT9 = False
        Case 7
            Place = "Lakh "
            GT9 = True
        Case 6
            Place = "Lakh "
            GT9 = False
        Case 5
            Place = "Thousand "
            GT9 = True
        Case 4
            Place = "Thousand "
            GT9 = False
        Case 3
            Place = "Hundred "
            GT9 = False
        Case 2
            GT9 = True
            Place = ""
        Case 1
            GT9 = False
            Place = ""
        Case Else
        
    End Select
    
    If GT9 Then
        Select Case (CStr(Ch2Dig))
            Case Is > 19
                GT19 = True
            Case Is < 19
                GT19 = False
            Case 19
                GT19 = False
                Nu = "Nineteen "
            Case Else
        End Select
    End If
    
    Select Case ch
        Case "9"
            If GT9 Then Nu = "Ninety " Else Nu = "Nine "
        Case "8"
            If GT9 Then Nu = "Eighty " Else Nu = "Eight "
        Case "7"
            If GT9 Then Nu = "Seventy " Else Nu = "Seven "
        Case "6"
            If GT9 Then Nu = "Sixty " Else Nu = "Six "
        Case "5"
            If GT9 Then Nu = "Fifty " Else Nu = "Five "
        Case "4"
            If GT9 Then Nu = "Fourty " Else Nu = "Four "
        Case "3"
            If GT9 Then Nu = "Thirty " Else Nu = "Three "
        Case "2"
            If GT9 And GT19 Then Nu = "Twenty "
            
            
            If Not GT9 And Not GT19 Then Nu = "Two "
        Case "1"
            If GT9 And Not GT19 Then
                Select Case (Ch2Dig)
                    Case "19"
                        Nu = "Nineteen "
                    Case "18"
                        Nu = "Eighteen "
                    Case "17"
                        Nu = "Seventeen "
                    Case "16"
                        Nu = "Sixteen "
                    Case "15"
                        Nu = "Fifteen "
                    Case "14"
                        Nu = "Fourteen "
                    Case "13"
                        Nu = "Thirteen "
                    Case "12"
                        Nu = "Twelve "
                    Case "11"
                        Nu = "Eleven "
                    Case "10"
                        Nu = "Ten "
                    Case Else
                End Select
            End If
            
            If Not GT9 And Not GT19 Then Nu = "One "
        Case "0"
            Nu = ""
            If Place = "Hundred " Then Place = ""
            If DecPos > 0 Then Nu = "Zero "
        Case "."
            Nu = "Decimal "
            
        Case Else
    
    End Select
       
    If Not GT9 And Not GT19 Then
        If ChOld = "1" Then
            If WasTeen Then
                SpellNumber = SpellNumber & Place
            Else
                SpellNumber = SpellNumber & Nu & Place
            End If
        Else
            If ch = "0" Then
                If Place = "Lakh " Or Place = "Crore " Then
                    SpellNumber = SpellNumber & Place
                Else
                    If DecPos > 0 Then
                        SpellNumber = SpellNumber & Nu
                    Else
                        SpellNumber = SpellNumber
                    End If
                End If
            Else
                SpellNumber = SpellNumber & Nu & Place
            End If
            
        End If
        WasTeen = False
    End If
    
    If GT9 And Not GT19 Then
        SpellNumber = SpellNumber & Nu
        WasTeen = True
    End If
    If GT9 And GT19 Then
        SpellNumber = SpellNumber & Nu
        WasTeen = False
    End If
    
    If DecPos > 0 Then
        Num = Right$(Num, StrLen - 1)
    Else
        Num = Right$(Num, Degree - 1)
    End If
    ChOld = ch

Wend

            
End Function


Function getAppWindow(appTitle As String, appHandle As Long, frmApps As Form) _
       As Boolean
    Dim dummyVariable As Long
    Dim lenTitle As Integer
    Dim winTitle As String * 64

        'initialize the function return as False
        getAppWindow = 0
        If appTitle = "" Then appTitle = "nothing"
        lenTitle = Len(appTitle)

        'Get the handle of the first child of the desktop window
        appHandle = GetTopWindow(0)

        'Loop through all top-level windows and search for the sub-string
        'in the Window title
        Do Until appHandle = 0
            dummyVariable = GetWindowText(appHandle, winTitle, 63)
            If Left(winTitle, lenTitle) = appTitle Then
                getAppWindow = -1
                Exit Function
            Else
                If InStr(1, winTitle, Chr(0)) >= 2 Then
                    frmApps.Combo1.AddItem winTitle
                End If
                appHandle = GetWindow(appHandle, GW_HWNDNEXT)
            End If
        Loop
If frmApps.Combo1.ListCount > 0 Then frmApps.Combo1.ListIndex = 0
End Function

Private Sub AdjustToken()

   Const TOKEN_ADJUST_PRIVILEGES = &H20
   Const TOKEN_QUERY = &H8
   Const SE_PRIVILEGE_ENABLED = &H2
   Dim hdlProcessHandle As Long
   Dim hdlTokenHandle As Long
   Dim tmpLuid As LUID
   Dim tkp As TOKEN_PRIVILEGES
   Dim tkpNewButIgnored As TOKEN_PRIVILEGES
   Dim lBufferNeeded As Long

   hdlProcessHandle = GetCurrentProcess()
   OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
      TOKEN_QUERY), hdlTokenHandle

   ' Get the LUID for shutdown privilege.
   LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid

   tkp.PrivilegeCount = 1    ' One privilege to set
   tkp.TheLuid = tmpLuid
   tkp.Attributes = SE_PRIVILEGE_ENABLED

   ' Enable the shutdown privilege in the access token of this
   ' process.
   AdjustTokenPrivileges hdlTokenHandle, False, tkp, _
      Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded

End Sub

Public Sub forceshutdown()
   AdjustToken
   ExitWindowsEx (EWX_SHUTDOWN Or EWX_FORCE Or EWX_REBOOT), &HFFFF
End Sub

Sub DisableCtrlAltDel()
Dim lngRet As Long
Dim blnOld As Boolean
lngRet = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, blnOld, 0&)
End Sub

Public Sub EnableCtrlAltDel()
Dim lngRet As Long
Dim blnOld As Boolean
lngRet = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, blnOld, 0&)
End Sub


Public Function UpTime() As String
Dim UTime As Date

UTime = CDate(Now) - starttime
UpTime = str$(hour(UTime)) & " Hrs"
UpTime = UpTime & ":" & str$(Minute(UTime)) & " Min"
UpTime = UpTime & ":" & str$(Second(UTime)) & " Sec"

End Function

Function getRunningApp(wTitle As String) As String
getRunningApp = ""
Dim appHandle As Long
Dim dummyVariable As Long
Dim winTitle As String * 256
winTitle = wTitle
'Get the handle of the first child of the desktop window
        appHandle = GetTopWindow(0)

        'Loop through all top-level windows and search for the sub-string
        'in the Window title
        Do Until appHandle = 0
            dummyVariable = GetWindowText(appHandle, winTitle, 255)
            getRunningApp = getRunningApp & winTitle & vbCrLf
            appHandle = GetWindow(appHandle, GW_HWNDNEXT)
        Loop
If getRunningApp = "" Then
    getRunningApp = "Currently running Apps:" & "NIL"
Else
    getRunningApp = "Currently running Apps:" & getRunningApp
End If

End Function
Sub DisableCAD()
Dim lpvparam As Boolean
    Dim x As Long

    x = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, lpvparam, 0&)

End Sub
Sub EnableCAD()
Dim lpvparam As Boolean
    Dim x As Long

    x = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, lpvparam, 0&)

End Sub


Sub CenterIt(mama As Object, kid As Object)
On Local Error Resume Next
Dim kLeft As Long, kTop As Long, mwidth As Long, kWidth As Long, mHeight As Long, kHeight As Long
'kLeft = kid.Left
'kTop = kid.Top
kWidth = kid.Width
kHeight = kid.Height
mwidth = mama.ScaleWidth
mHeight = mama.ScaleHeight

kLeft = mama.Left + (mwidth - kWidth) / 2
kTop = mama.Top + (mHeight - kHeight) / 2

kid.Move kLeft, kTop

End Sub
Function IsConnected() As Boolean
Dim hrc As Long, cns As Long, lret As Long
IsConnected = False
lret = RasGetConnectStatus(hrc, cns)
If lret = 0 Then
    IsConnected = True
Else
    IsConnected = False
End If
End Function

Function RunningProc(appTitle As String, appHandle As Long) As String
   Dim dummyVariable As Long
    Dim lenTitle As Integer
    Dim winTitle As String * 64

        'initialize the function return as empty
        RunningProc = ""
        'Get the handle of the first child of the desktop window
        appHandle = GetTopWindow(0)
        'Loop through all top-level windows and search for the sub-string
        'in the Window title
        Do Until appHandle = 0
            dummyVariable = GetWindowText(appHandle, winTitle, 63)
            'If Left(winTitle, lenTitle) = appTitle Then
              '  RunningProc = winTitle
             '   Exit Function
            'Else
                lenTitle = InStr(1, winTitle, Chr(0))
                If lenTitle > 1 Then
                    winTitle = RTrim(Left(winTitle, lenTitle - 1))
                    RunningProc = RunningProc & winTitle & vbCrLf '", "
                End If
                appHandle = GetWindow(appHandle, GW_HWNDNEXT)
            'End If
        Loop

End Function
Function GetFullTitle(appTitle As String, appHandle As Long) As Boolean
    Dim dummyVariable As Long
    Dim lenTitle As Integer
    Dim winTitle As String * 64

        'initialize the function return as False
        GetFullTitle = 0
        'Get the handle of the first child of the desktop window
        appHandle = GetTopWindow(0)
        'Loop through all top-level windows and search for the sub-string
        'in the Window title
        Do Until appHandle = 0
            dummyVariable = GetWindowText(appHandle, winTitle, 63)
            If InStr(1, winTitle, appTitle) > 0 Then
                GetFullTitle = -1
                appTitle = winTitle
                Exit Function
            Else
                appHandle = GetWindow(appHandle, GW_HWNDNEXT)
            End If
        Loop

End Function
Function CheckUse(elapsedseconds As Integer) As Boolean
Dim UTime As Date
CheckUse = False
UTime = CDate(Now) - starttime
If (Int(Second(UTime)) Mod elapsedseconds) = 0 Then
    CheckUse = True
End If

End Function

Function GetWeek(aday As Date) As String
Dim tmpweek As VbDayOfWeek
GetWeek = "Friday"
tmpweek = Weekday(aday)
Select Case tmpweek
    Case vbFriday
        GetWeek = "Friday"
    Case vbSaturday
        GetWeek = "Saturday"
    Case vbSunday
        GetWeek = "Sunday"
    Case vbMonday
        GetWeek = "Monday"
    Case vbTuesday
        GetWeek = "Tuesday"
    Case vbWednesday
        GetWeek = "Wednesday"
    Case vbThursday
        GetWeek = "Thursday"
End Select
End Function
Function GetWndFullName(appTitle As String, appHandle As Long) As String
Dim dummyVariable As Long
'Dim lenTitle As Integer
Dim winTitle As String * 128

        'initialize the function return as empty
        GetWndFullName = ""
        'Get the handle of the first child of the desktop window
        appHandle = GetTopWindow(0)
        'Loop through all top-level windows and search for the sub-string
        'in the Window title
        Do Until appHandle = 0
            dummyVariable = GetWindowText(appHandle, winTitle, 127&)
            If InStr(winTitle, appTitle) > 0 Then
                GetWndFullName = winTitle
                Exit Function
            Else
                appHandle = GetWindow(appHandle, GW_HWNDNEXT)
            End If
        Loop


End Function

' get the last occurrence of a string
Public Function InstrLast(ByVal Start As Long, _
   ByVal Source As String, ByVal Search As String, _
   Optional Cmp As VbCompareMethod) As Long
      
      Start = Start - 1
      Do
         Start = InStr(Start + 1, Source, Search, Cmp)
         If Start = 0 Then Exit Function
         InstrLast = Start
      Loop
End Function

Function GetMonthYear(adate As Date) As String
Dim lmonth As Integer
GetMonthYear = ""
lmonth = month(adate)
Select Case lmonth
    Case 1
        GetMonthYear = "January"
    Case 2
        GetMonthYear = "February"
    Case 3
        GetMonthYear = "March"
    Case 4
        GetMonthYear = "April"
    Case 5
        GetMonthYear = "May"
    Case 6
        GetMonthYear = "June"
    Case 7
        GetMonthYear = "July"
    Case 8
        GetMonthYear = "August"
    Case 9
        GetMonthYear = "September"
    Case 10
        GetMonthYear = "October"
    Case 11
        GetMonthYear = "November"
    Case 12
        GetMonthYear = "December"
    Case Else
    GetMonthYear = ""
End Select
GetMonthYear = GetMonthYear & " " & str$(year(adate))
End Function


Public Function LPad(ValIn As Variant, nDec As Integer, _
                      WidthOut As Integer) As String
'
' Formatting function left pads with spaces, using specified
' number of decimal digits.
'
   If IsNumeric(ValIn) Then
      If nDec > 0 Then
         LPad = Right$(Space$(WidthOut) & _
                Format$(ValIn, "0." & String$(nDec, "0")), _
                WidthOut)
      Else
         LPad = Right$(Space$(WidthOut) & Format$(ValIn, "0"), WidthOut)
      End If
   Else
      LPad = Right$(Space$(WidthOut) & ValIn, WidthOut)
   End If
End Function


Function GetCompany(appexename As String) As companytype
GetCompany.address = QueryValue("Software\iqsoft\" & appexename & "\company", "address")
GetCompany.cname = QueryValue("Software\iqsoft\" & appexename & "\company", "name")
GetCompany.cst = QueryValue("Software\iqsoft\" & appexename & "\company", "cst")
GetCompany.kgst = QueryValue("Software\iqsoft\" & appexename & "\company", "kgst")
GetCompany.phone = QueryValue("Software\iqsoft\" & appexename & "\company", "phone")

End Function

Sub ChangePwd(Username As String, appexename As String)
On Local Error Resume Next
Dim pwd As String, Pwd1 As String, Pwd2 As String
Dim uPwd As String, tPwd As String
Dim aUser As usertype
If Username = "" Then Exit Sub
Call GetUserDetails(Username, appexename, uPwd)
If aUser.uname <> "" Then
    pwd = InputBox("Type the old Password", "Change Password for " & aUser.uname, "")
    If pwd <> uPwd Then
        MsgBox "YOU IMPOSTER!!", vbCritical, "Caught a thief"
        LogIt LogFile, "Unauthorised attempt to change password of " & aUser.uname, aUser.password
        Exit Sub
    End If
    Pwd1 = InputBox("Type the new Password", "Change Password for " & aUser.uname, "")
    If Pwd1 = "" Then Exit Sub
    If Pwd1 = aUser.uname Then
        MsgBox "Cant use your ID as password", vbCritical, "Sorry!!"
        Exit Sub
    End If
    Pwd2 = InputBox("Type the new Password again to verify", "Change Password for " & aUser.uname, "")
    If Pwd1 = Pwd2 Then
        Pwd1 = EncryptIt(Pwd1, UCase$(aUser.uname))
        SetKeyValue "Software\iqsoft\" & appexename & "\" & aUser.uname, "password", Pwd1, 1&
    Else
        MsgBox "Passwords dont match", vbCritical, "Try again after learning to type!!"
        Exit Sub
    End If
Else
    MsgBox "No such user!", vbCritical, "I don't know this guy!"
    Exit Sub
End If
End Sub
Function GetUserDetails(Username As String, appexename As String, Optional password As String, Optional question As String, Optional answer As String, Optional supervisoraccess As Boolean) As Boolean
On Local Error GoTo getusererr
If QueryValue("Software\iqsoft\" & appexename & "\" & Username, "password") _
 <> "" Then
    Username = Username
    password = EncryptIt(QueryValue("Software\iqsoft\" & appexename & "\" & Username, "password"), Username)
    question = EncryptIt(QueryValue("Software\iqsoft\" & appexename & "\" & Username, "question"), Username)
    answer = EncryptIt(QueryValue("Software\iqsoft\" & appexename & "\" & Username, "answer"), Username)
    supervisoraccess = QueryValue("Software\iqsoft\" & appexename & "\" & Username, "supervisoraccess")
    GetUserDetails = True
   
ElseIf QueryValueCurrentUser("Software\iqsoft\" & appexename & "\" & Username, "password") _
 <> "" Then
    Username = Username
    password = EncryptIt(QueryValueCurrentUser("Software\iqsoft\" & appexename & "\" & Username, "password"), Username)
    question = EncryptIt(QueryValueCurrentUser("Software\iqsoft\" & appexename & "\" & Username, "question"), Username)
    answer = EncryptIt(QueryValueCurrentUser("Software\iqsoft\" & appexename & "\" & Username, "answer"), Username)
    supervisoraccess = QueryValueCurrentUser("Software\iqsoft\" & appexename & "\" & Username, "supervisoraccess")
    GetUserDetails = True
Else
    Username = ""
    password = ""
    question = ""
    answer = ""
    supervisoraccess = False
    GetUserDetails = False
End If
Exit Function
getusererr:
MsgBox Err.Description
Resume Next

End Function

Sub Forgot(appexename As String)
Dim ans As String, pquestion As String, pans As String, pwd As String
ans = InputBox("Forgot your password? No problem!! I'll tell you what it is." & vbCrLf & "what is your ID (User Name/ Initial)?", "Tell me your ID", User1.uname)
If ans = "" Then Exit Sub
'ans = UCase$(ans)
Call GetUserDetails(ans, appexename, pwd)
If User1.uname <> "" Then
    pquestion = User1.question
    pans = InputBox("What is the answer to this question you asked? " & vbCrLf & pquestion, "Finding your forgotten password")
    If pans = User1.answer Then
        MsgBox "Your Password is " & pwd
    Else
        MsgBox "Sorry! Cant tell you the password"
        Exit Sub
    End If
Else
    MsgBox "No Such User!!"
    Exit Sub
End If

End Sub
Function CreateRandNo(lower As Long, upper As Long) As Long
Dim lastID As Long
Randomize
lastID = ((upper - lower) * Rnd + lower)
lastID = Int(lastID)
CreateRandNo = lastID

End Function

' Validate a credit card numbers
' Returns True if valid, False if invalid
'
' Example:
'  If IsValidCreditCardNumber(Value:="1234-123456-12345", IsRequired:=True)

Function IsValidCreditCardNumber(Value As Variant, Optional ByVal IsRequired As _
    Boolean = True) As Boolean
    Dim strTemp As String
    Dim intCheckSum As Integer
    Dim blnDoubleFlag As Boolean
    Dim intDigit As Integer
    Dim i As Integer

    On Error GoTo ErrorHandler

    IsValidCreditCardNumber = True
    Value = Trim$(Value)

    If IsRequired And Len(Value) = 0 Then
        IsValidCreditCardNumber = False
    End If

    ' If after stripping out non-numerics, there is nothing left,
    '  they entered junk
    For i = 1 To Len(Value)
        If IsNumeric(Mid$(Value, i, 1)) Then strTemp = strTemp & Mid$(Value, i, _
            1)
    Next
    If IsRequired And Len(strTemp) = 0 Then
        IsValidCreditCardNumber = False
    End If

    'Handle different lengths for different credit card types
    Select Case Mid$(strTemp, 1, 1)
        Case "3"    'Amex
            If Len(strTemp) <> 15 Then
                IsValidCreditCardNumber = False
            Else
                Value = Mid$(strTemp, 1, 4) & "-" & Mid$(strTemp, 5, _
                    6) & "-" & Mid$(strTemp, 11, 5)
            End If
        Case "4"    'Visa
            If Len(strTemp) <> 16 Then
                IsValidCreditCardNumber = False
            Else
                Value = Mid$(strTemp, 1, 4) & "-" & Mid$(strTemp, 5, _
                    4) & "-" & Mid$(strTemp, 9, 4) & "-" & Mid$(strTemp, 13, 4)
            End If
        Case "5"    'Mastercard
            If Len(strTemp) <> 16 Then
                IsValidCreditCardNumber = False
            Else
                Value = Mid$(strTemp, 1, 4) & "-" & Mid$(strTemp, 5, _
                    4) & "-" & Mid$(strTemp, 9, 4) & "-" & Mid$(strTemp, 13, 4)
            End If
        Case Else      'Discover - Dont know rules yet
            If Len(strTemp) > 20 Then
                IsValidCreditCardNumber = False
            End If
    End Select

    'Now check for Check Sum (Mod 10)
    intCheckSum = 0
           ' Start with 0 intCheckSum
    blnDoubleFlag = 0
           ' Start with a non-doubling
    For i = Len(strTemp) To 1 Step -1                   ' Working backwards
        intDigit = Asc(Mid$(strTemp, i, 1))             ' Isolate character
        If intDigit > 47 Then                           ' Skip if not a intDigit
            If intDigit < 58 Then
                intDigit = intDigit - 48                ' Remove ASCII bias
                If blnDoubleFlag Then
                       ' If in the "double-add" phase
                    intDigit = intDigit + intDigit      '   then double first
                    If intDigit > 9 Then
                        intDigit = intDigit - 9         ' Cast nines
                    End If
                End If
                blnDoubleFlag = Not blnDoubleFlag       ' Flip doubling flag
                intCheckSum = intCheckSum + intDigit    ' Add to running sum
                If intCheckSum > 9 Then                 ' Cast tens
                    intCheckSum = intCheckSum - 10      ' (same as MOD 10 but
                                                        ' faster)
                End If
            End If
        End If
    Next

    If intCheckSum <> 0 Then                            '  Must sum to zero
        IsValidCreditCardNumber = False
    End If

ExitMe:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "IsValidCreditCardNumber", Err.Description

End Function


Function GetFileNamefromPath(names As String) As String
GetFileNamefromPath = names

Dim fname As String

Dim i As Long
For i = 0 To Len(names)
    fname = Right(names, i)
    If InStr(fname, "\") Then Exit For
Next i

GetFileNamefromPath = Right(fname, Len(fname) - 1)

End Function

Function GetPathFromFullFilename(names As String) As String
Dim fname As String
Dim fpath As String
Dim i As Long
Dim epn As Long
For i = 0 To Len(names)
    fname = Right(names, i)
    If InStr(fname, "\") Then
        Exit For
    End If
Next i
fpath = Left(names, Len(names) - Len(fname) + 1)
epn = InStr(Len(fpath) + 1, names, Chr(0), vbTextCompare)
If epn > 0 Then
    fpath = Left(names, epn - 1)
End If
GetPathFromFullFilename = fpath

End Function

Function ApiError(ByVal E As Long) As String
Dim s As String
Dim c As Long
Dim pNull As Variant

s = String(256, 0)
c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, pNull, E, 0&, s, Len(s), ByVal pNull)
    If c Then
        ApiError = Left$(s, c)
    End If
End Function

Function stripChar(str As String, ch As String) As String
    Dim i As Integer

    stripChar = str
    For i = 1 To Len(ch)
        stripChar = Replace(stripChar, Mid$(ch, i, 1), "")
    Next i

End Function


