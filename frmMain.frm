VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "salat"
   ClientHeight    =   5805
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2040
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "one"
            Object.ToolTipText     =   "Settings"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "two"
            Object.ToolTipText     =   "Website"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "three"
            Object.ToolTipText     =   "Select your location"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "four"
            Object.ToolTipText     =   "Create a report of salat timings for any location"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "moon"
            Object.ToolTipText     =   "Lunar Date, Position, Phase"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   5430
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   661
      SimpleText      =   "Salat by Iqsoft"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   3731
            Text            =   "Iqsoft Software Consultants "
            TextSave        =   "SCRL"
            Object.ToolTipText     =   "Iqsoft Software Consultants "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3731
            TextSave        =   "21:32"
            Object.ToolTipText     =   "Salat by Iqsoft"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   3731
            Text            =   "https://iqsoft.co.in"
            TextSave        =   "SCRL"
            Object.ToolTipText     =   "https://iqsoft.co.in"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   4920
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8FBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AD2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CA06
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E648
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10322
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12094
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MCI.MMControl MMControl1 
      Height          =   855
      Left            =   0
      TabIndex        =   18
      Top             =   4200
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   1508
      _Version        =   393216
      BorderStyle     =   0
      RecordMode      =   0
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PlayVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   "adan\adan.wav"
   End
   Begin VB.Timer Timer1 
      Left            =   5520
      Top             =   960
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   4080
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblHijri 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   300
      TabIndex        =   22
      Top             =   1320
      Width           =   3500
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   600
      TabIndex        =   21
      Top             =   600
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWaqt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waqt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   15
      Left            =   2685
      TabIndex        =   17
      Top             =   2160
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWaqt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waqt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   330
      Index           =   14
      Left            =   2685
      TabIndex        =   16
      Top             =   1800
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWaqt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waqt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   13
      Left            =   1605
      TabIndex        =   15
      Top             =   2160
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWaqt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waqt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   330
      Index           =   12
      Left            =   1605
      TabIndex        =   14
      Top             =   1800
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1485
      Left            =   2467
      Picture         =   "frmMain.frx":13CD6
      Top             =   3960
      Width           =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      BorderWidth     =   5
      X1              =   3240
      X2              =   3720
      Y1              =   4560
      Y2              =   3960
   End
   Begin VB.Label lblQibla 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   6135
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNextSalat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   3240
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWaqt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waqt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   11
      Left            =   4852
      TabIndex        =   11
      Top             =   2880
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWaqt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waqt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   10
      Left            =   3772
      TabIndex        =   10
      Top             =   2880
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWaqt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waqt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   9
      Left            =   2692
      TabIndex        =   9
      Top             =   2880
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWaqt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waqt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   8
      Left            =   1612
      TabIndex        =   8
      Top             =   2880
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWaqt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waqt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   7
      Left            =   525
      TabIndex        =   7
      Top             =   2160
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWaqt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waqt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   6
      Left            =   532
      TabIndex        =   6
      Top             =   2880
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWaqt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waqt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   5
      Left            =   4852
      TabIndex        =   5
      Top             =   2520
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWaqt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waqt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   4
      Left            =   3772
      TabIndex        =   4
      Top             =   2520
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWaqt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waqt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   3
      Left            =   2692
      TabIndex        =   3
      Top             =   2520
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWaqt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waqt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   2
      Left            =   1612
      TabIndex        =   2
      Top             =   2520
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWaqt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waqt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   330
      Index           =   1
      Left            =   525
      TabIndex        =   1
      Top             =   1800
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWaqt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waqt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   0
      Left            =   532
      TabIndex        =   0
      Top             =   2520
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mpopsys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mpopsysoptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mpopsysstopsound 
         Caption         =   "&Stop Sound"
      End
      Begin VB.Menu mpopsysstartminimized 
         Caption         =   "S&tart Minimized"
      End
      Begin VB.Menu msep4 
         Caption         =   "-"
      End
      Begin VB.Menu msyslunar 
         Caption         =   "Lunar Phase"
      End
      Begin VB.Menu mpopsyssep1 
         Caption         =   "-"
      End
      Begin VB.Menu mpopsysrestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mpopsysminimize 
         Caption         =   "&Minimize"
      End
      Begin VB.Menu mpopsep2 
         Caption         =   "-"
      End
      Begin VB.Menu mpopabout 
         Caption         =   "About"
      End
      Begin VB.Menu mpopsyssep0 
         Caption         =   "-"
      End
      Begin VB.Menu mpopsysexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnudummysalat 
      Caption         =   "salat                       "
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuLocations 
      Caption         =   "Location"
      Begin VB.Menu mnuLocationsNew 
         Caption         =   "New (Your Latitude and  Longitude)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuLocationsOpen 
         Caption         =   "Open from Database"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLocationsselect 
         Caption         =   "Select Country, Region, City"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLocationsEdit 
         Caption         =   "Edit Existing Location"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "1012"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "1019"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "1020"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "1021"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "1022"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "1023"
      End
      Begin VB.Menu mnureports 
         Caption         =   "Reports"
      End
      Begin VB.Menu mnuViewHijriCalendar 
         Caption         =   "Hijri Calendar"
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "1024"
      End
      Begin VB.Menu mnuviewbar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuviewminimise 
         Caption         =   "Minimise to System Tray"
      End
      Begin VB.Menu mnuviewtop 
         Caption         =   "Always on top"
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "1025"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "1026"
      End
      Begin VB.Menu mnuuninstall 
         Caption         =   "Uninstall"
      End
      Begin VB.Menu mnusounds 
         Caption         =   "Sounds"
      End
      Begin VB.Menu mnutoolssep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvertHijriDate 
         Caption         =   "Convert Hijri Date"
      End
      Begin VB.Menu mnutoolsmoonphase 
         Caption         =   "Moon Phase"
      End
      Begin VB.Menu mnuMoonSet 
         Caption         =   "Moon Set"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "1027"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "1028"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "1029"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuchkversion 
         Caption         =   "Check for New Version"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "1030"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub Form_DblClick()
    Call mnuviewminimise_Click
End Sub

Private Sub Form_Load()
On Local Error GoTo frmloaderr
Randomize Timer
LoadResStrings Me

Timer1.Interval = 100
Timer1.Enabled = True
Me.MouseIcon = LoadResPicture(101, vbResIcon)
Me.Icon = LoadResPicture(101 + Int(Rnd * 5), vbResIcon)
       
With nid
 .cbSize = Len(nid)
 .hwnd = Me.hwnd
 .uId = vbNull
 .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
 .uCallBackMessage = WM_MOUSEMOVE
 .hIcon = Me.Icon
 .szTip = "Starting up..."
End With
Shell_NotifyIcon NIM_ADD, nid

CreateCrescent Me
Settings1 = GetSettings()
TwilightAngle = Settings1.TwilightAngle

SetTransparency Me, Settings1.Transparency

lblLoc.FontBold = True
lblLoc.FontSize = 10



Exit Sub
frmloaderr:
MsgBox " fmainform load " & Err.Description & Err.Number

fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 Then
    Result = SetForegroundWindow(Me.hwnd)
    Me.PopupMenu Me.mpopsys
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error Resume Next
'this procedure receives the callbacks from the System Tray icon.
Dim Result As Long
Dim msg As Long
 'the value of X will vary depending upon the scalemode setting
 If Me.ScaleMode = vbPixels Then
  msg = x
 Else
  msg = x \ Screen.TwipsPerPixelX
 End If
 Select Case msg
    Case WM_LBUTTONUP   'WM_LBUTTONDOWN,   482, 483 ',  '514 restore form window
        Restore
    Case WM_RBUTTONUP  '484, 485, 486
        Result = SetForegroundWindow(Me.hwnd)
        Me.PopupMenu Me.mpopsys
    Case WM_RBUTTONDOWN
    
    Case WM_LBUTTONDOWN
        'ReadIt "Bismilla"
        CallWarning "startup"
    Case WM_LBUTTONDBLCLK
        Restore
    Case WM_RBUTTONDBLCLK
        HideMe Me
        HideMe frmStar
    Case Else
      'showcursor (False)

End Select

End Sub

Private Sub Form_Resize()
CreateCrescent Me
If Not Me.WindowState = vbMinimized Then
    Me.Move (Screen.Width - Me.ScaleWidth) / 2, (Screen.Height - Me.ScaleHeight) / 2
    ShowMe frmStar
    starMov
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
    Dim i As Integer
    MMControl1.Command = "Close"
    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
Shell_NotifyIcon NIM_DELETE, nid
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.title, "Settings", "MainLeft", Me.Left
        SaveSetting App.title, "Settings", "MainTop", Me.Top
        SaveSetting App.title, "Settings", "MainWidth", Me.Width
        SaveSetting App.title, "Settings", "MainHeight", Me.Height
    End If


End Sub

Private Sub lblHijri_Click()
ShowMe frmHijri

End Sub

Private Sub lblLoc_Click()

MsgBox "Location: " & city.cityname & vbCr & "Latitude: " & _
        city.Latitude & vbCr & "Longitude: " & city.Longitude & vbCr, vbOKOnly, " Location : " & city.cityname
        
End Sub

Private Sub mnuchkversion_Click()
CheckVersion
End Sub

Private Sub mnuConvertHijriDate_Click()
frmHijri.Show
End Sub

Private Sub mnuexit_Click()
Unload Me
End
End Sub

Private Sub mnuLocationsEdit_Click()
'CityEditMode = True
'CenterIt Me, frmLocations
'frmLocations.Show

 '   CityEditMode = False
    CenterIt fMainForm, dlgEditCity
    dlgEditCity.Show

End Sub

Private Sub mnuLocationsNew_Click()
CenterIt Me, frmNewLoc
frmNewLoc.Show
Timer1.Enabled = False
End Sub

Private Sub mnuLocationsOpen_Click()
CityEditMode = False
frmLocations.Show
CenterIt Me, frmLocations

Timer1.Enabled = False
End Sub

Private Sub mnuLocationsselect_Click()
Me.lblLoc.Caption = "Please wait... Loading Cities of the world..."
CenterIt Screen, dlgSelect
dlgSelect.Show
Timer1.Enabled = False

End Sub


Private Sub mnuMoonSet_Click()
'Moon set time
ShowMe dlgMoonSet
CenterIt Me, dlgMoonSet
End Sub

Private Sub mnureports_Click()
CenterIt Screen, dlgReport
dlgReport.Show
End Sub

Private Sub mnusounds_Click()
CallAdan 0
CenterIt Me, frmSettings
frmSettings.ZOrder 0
frmSettings.Show

End Sub

Private Sub mnutoolsmoonphase_Click()
CenterIt Screen, dlgMoon
dlgMoon.Show
dlgMoon.ZOrder 0

End Sub

Private Sub mnuuninstall_Click()
UnInstall
MsgBox "Salat is now Uninstalled!!(Will not start with Windows)"
End Sub

Private Sub mnuViewHijriCalendar_Click()
frmHijri.Show
End Sub

Private Sub mnuviewminimise_Click()
Me.WindowState = vbMinimized
f_Maximised = False
frmStar.WindowState = vbMinimized
HideMe Me
HideMe frmStar

End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show 'vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
'    Dim nRet As Integer
'
'
'    'if there is no helpfile for this project display a message to the user
'    'you can set the HelpFile for your application in the
'    'Project Properties dialog
'    If Len(App.HelpFile) = 0 Then
'        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
'    Else
'        On Error Resume Next
'        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
'        If Err Then
'            MsgBox Err.Description
'        End If
'    End If
OpenDefaultBrowser "salathelp.html"

End Sub

Private Sub mnuHelpContents_Click()
'    Dim nRet As Integer
'    'if there is no helpfile for this project display a message to the user
'    'you can set the HelpFile for your application in the
'    'Project Properties dialog
'    If Len(App.HelpFile) = 0 Then
'        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
'    Else
'        On Error Resume Next
'        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
'        If Err Then
'            MsgBox Err.Description
'        End If
'    End If
OpenDefaultBrowser ("salathelp.html")

End Sub


Private Sub mnuToolsOptions_Click()
    frmSettings.Show vbModal, Me
End Sub

Private Sub mnuviewtop_Click()
If Me.mnuviewtop.Checked Then
    SetTopMostWindow Me.hwnd, False
    SetTopMostWindow frmStar.hwnd, False
    Me.mnuviewtop.Checked = False
Else
    SetTopMostWindow Me.hwnd, True
    SetTopMostWindow frmStar.hwnd, True
    Me.mnuviewtop.Checked = True
End If

End Sub

Private Sub mnuViewWebBrowser_Click()
OpenDefaultBrowser "http://iqsoft.co.in"
End Sub

Private Sub mnuViewOptions_Click()
    frmSettings.Show vbModal, Me
End Sub

Private Sub mnuViewRefresh_Click()
 Me.Refresh
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    Toolbar1.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me
End
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String


    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    'ToDo: add code to process the opened file

End Sub

Private Sub mpopabout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mpopsysexit_Click()
End
End Sub

Private Sub mpopsysminimize_Click()
Me.WindowState = vbMinimized
f_Maximised = False
frmStar.WindowState = vbMinimized
Me.mpopsysminimize.Enabled = False
Me.mpopsysrestore.Enabled = True
HideMe Me
HideMe frmStar

End Sub

Private Sub mpopsysoptions_Click()
CenterIt Screen, frmSettings
frmSettings.Show
End Sub

Private Sub mpopsysrestore_Click()

Restore
Me.mpopsysrestore.Enabled = False
Me.mpopsysminimize.Enabled = True

End Sub

Private Sub mpopsysstartminimized_Click()
If Me.mpopsysstartminimized.Checked Then
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Run", "salat", App.Path & "\" & App.exename & ".exe -m", 1&
Else
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Run", "salat", App.Path & "\" & App.exename & ".exe", 1&
End If
Me.mpopsysstartminimized.Checked = Not (Me.mpopsysstartminimized.Checked)

End Sub

Private Sub mpopsysstopsound_Click()
StopAdan
End Sub

Private Sub msyslunar_Click()

CenterIt Screen, dlgMoon
dlgMoon.Show
dlgMoon.ZOrder 0

End Sub

Private Sub sbStatusBar_PanelClick(ByVal Panel As MSComctlLib.Panel)
If Panel.Index = 1 Then
    MsgBox Panel.Text, vbInformation, "IQSoft Salat Information"
    Panel.Text = "IQSoft Software Consultants"
ElseIf Panel.Index = 2 Then
    MsgBox "Next Salat: " & GetNextsalat()
Else
    MsgBox "Date " & Date$ & " Time " & Time$, vbInformation, "IQSoft Salat Information"
End If
End Sub

Private Sub Timer1_Timer()
On Local Error GoTo timererr
Randomize Timer
RefreshWaqt
salatAdan
salatWarning
Static tick As Byte
tick = tick + 1
If tick > 27 Then
    tick = 0
End If
Dim resindex As Integer
resindex = 101 + tick
Me.Icon = LoadResPicture(resindex, vbResIcon)
With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = " " & vbCr & GetNextsalat & vbCr & vbCr & "  "
End With
Shell_NotifyIcon NIM_MODIFY, nid

DrawQibla

Exit Sub
timererr:
MsgBox " Timer " & Err.Description & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End Sub

Sub Restore()
Dim Result As Long
ShowMe Me
Result = SetForegroundWindow(Me.hwnd)
Me.WindowState = vbNormal 'Focus
ShowMe frmStar
frmStar.WindowState = vbNormal
f_Maximised = True
starMov
fMainForm.mpopsysminimize.Enabled = True
fMainForm.mpopsysrestore.Enabled = False

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
        
Select Case Button.key
    Case "one"
        Call mnuToolsOptions_Click
    Case "two"
        Call mnuViewWebBrowser_Click
    Case "three"
        Call mnuLocationsselect_Click
    Case "four"
        Call mnureports_Click
    Case "moon"
        dlgMoon.Show vbModal
        dlgMoon.ZOrder 0
    Case "nine"
        frmAbout.Show vbModal
    Case Else
        frmAbout.Show vbModal
End Select

End Sub

Sub DrawQibla()
Dim q As Double
Dim cenq As POINTAPI
'cenq.x = Me.image1.Left + Me.image1.Width / 2 'Me.ScaleWidth / 3
'cenq.Y = Me.image1.Top + Me.image1.Height / 2 'Me.ScaleHeight / 2
cenq.x = Me.Image1.Left + Me.Image1.Width / 2 'Me.ScaleWidth / 3
cenq.y = Me.Image1.Top + Me.Image1.Height / 2 'Me.ScaleHeight / 2

q = GetQibla(city.Latitude, city.Longitude)
If q < 0 Then
    q = (PI / 2) + q
End If
Me.Line1.BorderWidth = 2
'Me.Line2.BorderWidth = 2

Me.Line1.X1 = cenq.x
Me.Line1.Y1 = cenq.y
Me.Line1.X2 = cenq.x - Cos(q) * (Me.Image1.Width / 2)
Me.Line1.Y2 = cenq.y - Sin(q) * (Me.Image1.Height / 2)

'Me.Line2.X1 = cenq.x
'Me.Line2.Y1 = cenq.Y
'Me.Line2.X2 = cenq.x
'Me.Line2.Y2 = cenq.Y - Me.Image1.Height / 2
GetNearestCardinalDir (q)
lblQibla.Caption = "Qibla Direction= " & Format(R2D((PI / 2) - q), "##0.#0") & " Degrees from North Anticlockwise"
Me.Line1.ZOrder 0
'Me.Line2.ZOrder 0


End Sub

Sub CheckVersion()
Dim updateURL As String
Dim ver As String
Dim verNo As Single
Dim res As VbMsgBoxResult

updateURL = "https://pallipurath.in/salat/"

ver = ""
ver = Inet1.OpenURL("https://pallipurath.in/salat/version.txt")
While ver = ""
    DoEvents
    Me.sbStatusBar.Panels(1).Text = "Please wait ...checking for new version.."
Wend

ver = stripChar(ver, ".")
verNo = Val(ver)
If verNo > Val(App.Major & App.Minor & App.Revision) Then
   res = MsgBox("New Version " & ver & " Do you want to download the new version?", vbYesNoCancel, "New Version found")
   If res = vbYes Then
    OpenDefaultBrowser "https://pallipurath.in/salat"
   End If
Else
 MsgBox "Congratulations, You have the latest version of salat. "
End If
End Sub

Sub starMov()
ShowMe frmStar
frmStar.WindowState = vbNormal
frmStar.Move fMainForm.ScaleLeft + fMainForm.ScaleWidth * 1.65, _
    fMainForm.ScaleTop + fMainForm.ScaleHeight * 0.28

End Sub
