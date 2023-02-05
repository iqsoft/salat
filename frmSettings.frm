VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tb 
      Height          =   4200
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7408
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Location"
      TabPicture(0)   =   "frmSettings.frx":8FBA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraCon(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Sounds"
      TabPicture(1)   =   "frmSettings.frx":8FD6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Calculation"
      TabPicture(2)   =   "frmSettings.frx":8FF2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Interface"
      TabPicture(3)   =   "frmSettings.frx":900E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Clock"
      TabPicture(4)   =   "frmSettings.frx":902A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Startup"
      TabPicture(5)   =   "frmSettings.frx":9046
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame6 
         Caption         =   "Startup options"
         Height          =   3615
         Left            =   -74880
         TabIndex        =   53
         Top             =   120
         Width           =   6255
         Begin VB.CheckBox chkStartup 
            Caption         =   "Start Minimised?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   960
            TabIndex        =   81
            Top             =   1200
            Width           =   2715
         End
         Begin VB.CheckBox chkStartup 
            Caption         =   "Start With Windows?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   80
            Top             =   600
            Width           =   2715
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   4800
            TabIndex        =   56
            Tag             =   "1060"
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   4800
            TabIndex        =   55
            Tag             =   "1059"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "&Apply"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   4800
            TabIndex        =   54
            Tag             =   "1058"
            Top             =   1800
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Clock Settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   -74880
         TabIndex        =   49
         Top             =   120
         Width           =   6375
         Begin VB.ComboBox cmbClock 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2160
            TabIndex        =   70
            Top             =   1920
            Width           =   2415
         End
         Begin VB.CommandButton cmdColor 
            Height          =   345
            Index           =   2
            Left            =   4080
            Picture         =   "frmSettings.frx":9062
            Style           =   1  'Graphical
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   1200
            Width           =   360
         End
         Begin VB.CommandButton cmdColor 
            Height          =   345
            Index           =   1
            Left            =   4080
            Picture         =   "frmSettings.frx":992C
            Style           =   1  'Graphical
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   720
            Width           =   360
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   5280
            Top             =   2760
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton cmdColor 
            Height          =   345
            Index           =   0
            Left            =   4080
            Picture         =   "frmSettings.frx":A1F6
            Style           =   1  'Graphical
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   240
            Width           =   360
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   5040
            TabIndex        =   52
            Tag             =   "1060"
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   5040
            TabIndex        =   51
            Tag             =   "1059"
            Top             =   960
            Width           =   1095
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "&Apply"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   5040
            TabIndex        =   50
            Tag             =   "1058"
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Clock Face "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   480
            TabIndex        =   71
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Second Hand Color"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   480
            TabIndex        =   67
            Top             =   1320
            Width           =   3000
         End
         Begin VB.Label Label12 
            Caption         =   "Minute Hand Color"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   480
            TabIndex        =   66
            Top             =   840
            Width           =   3000
         End
         Begin VB.Label Label12 
            Caption         =   "Hour Hand Color"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   65
            Top             =   360
            Width           =   3000
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Interface Settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   -74880
         TabIndex        =   45
         Top             =   120
         Width           =   6375
         Begin MSComctlLib.Slider sliTransparency 
            Height          =   375
            Left            =   120
            TabIndex        =   83
            Top             =   1560
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   661
            _Version        =   393216
            LargeChange     =   20
            SmallChange     =   5
            Max             =   255
            TickFrequency   =   20
         End
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   45
            Left            =   120
            Picture         =   "frmSettings.frx":AAC0
            ScaleHeight     =   45
            ScaleWidth      =   4455
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   1440
            Width           =   4455
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   4800
            TabIndex        =   48
            Tag             =   "1060"
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   4800
            TabIndex        =   47
            Tag             =   "1059"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "&Apply"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   4800
            TabIndex        =   46
            Tag             =   "1058"
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Transparency Level"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   840
            Width           =   4095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Calculation"
         Height          =   3495
         Left            =   -74880
         TabIndex        =   41
         Top             =   240
         Width           =   6255
         Begin MSComctlLib.Slider slrAdjDays 
            DragIcon        =   "frmSettings.frx":B3F6
            Height          =   255
            Left            =   2400
            TabIndex        =   82
            ToolTipText     =   "Adjust left to decrease, right to increase by a day"
            Top             =   2760
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            MousePointer    =   4
            MouseIcon       =   "frmSettings.frx":12408
            LargeChange     =   1
            Min             =   -5
            Max             =   5
            TickStyle       =   2
            TextPosition    =   1
         End
         Begin VB.TextBox txtTAIsha 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   75
            Text            =   "18"
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox txtTAFajr 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   74
            Text            =   "18"
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   58
            Text            =   "18"
            Top             =   600
            Width           =   615
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   4800
            TabIndex        =   44
            Tag             =   "1060"
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   4800
            TabIndex        =   43
            Tag             =   "1059"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "&Apply"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   4800
            TabIndex        =   42
            Tag             =   "1058"
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "Days to adjust in lunar month"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   3
            Left            =   255
            TabIndex        =   60
            Top             =   2520
            Width           =   1965
         End
         Begin VB.Label Label9 
            Caption         =   "Twilight Angle for Isha"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   240
            TabIndex        =   73
            Top             =   1920
            Width           =   2895
         End
         Begin VB.Label Label9 
            Caption         =   "Twilight Angle for Fajr"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   240
            TabIndex        =   72
            Top             =   1200
            Width           =   2775
         End
         Begin VB.Label Label9 
            Caption         =   "Default Twilight Angle"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   57
            Top             =   600
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Set Sounds"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   6255
         Begin VB.TextBox txtWarningtxt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            TabIndex        =   78
            Text            =   " Minutes to Adan"
            Top             =   2760
            Width           =   2775
         End
         Begin VB.CommandButton cmdPlayFAdan 
            Height          =   375
            Left            =   3960
            Picture         =   "frmSettings.frx":1942A
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   1200
            Width           =   375
         End
         Begin VB.CommandButton cmdPlayAdan 
            Height          =   375
            Left            =   3960
            Picture         =   "frmSettings.frx":19974
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   720
            Width           =   375
         End
         Begin VB.ComboBox cmbFAdan 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1920
            TabIndex        =   62
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   4920
            TabIndex        =   40
            Tag             =   "1060"
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   4920
            TabIndex        =   39
            Tag             =   "1059"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "&Apply"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   4920
            TabIndex        =   38
            Tag             =   "1058"
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            TabIndex        =   37
            Text            =   "10"
            Top             =   2280
            Width           =   615
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Enable Warning"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   35
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Enable Adan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.ComboBox cmbAdan 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1920
            TabIndex        =   33
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label13 
            Caption         =   "Warning Text"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   79
            Top             =   2880
            Width           =   2895
         End
         Begin VB.Label Label11 
            Caption         =   "Select Fajr Adan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   63
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label Label8 
            Caption         =   "Warning Minutes before Adan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   36
            Top             =   2400
            Width           =   2895
         End
         Begin VB.Label Label7 
            Caption         =   "Select Adan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   32
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.Frame fraCon 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3585
         Index           =   0
         Left            =   -74880
         TabIndex        =   1
         Tag             =   "1054"
         Top             =   120
         Width           =   6375
         Begin VB.CommandButton cmdApply 
            Caption         =   "&Apply"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   5160
            TabIndex        =   30
            Tag             =   "1058"
            Top             =   1440
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   5160
            TabIndex        =   29
            Tag             =   "1059"
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   5160
            TabIndex        =   28
            Tag             =   "1060"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtTimezone 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3720
            TabIndex        =   21
            ToolTipText     =   "Negative Values for West "
            Top             =   720
            Width           =   1300
         End
         Begin VB.TextBox txtCountry 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1320
            TabIndex        =   20
            Top             =   720
            Width           =   1300
         End
         Begin VB.TextBox txtRegion 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3720
            TabIndex        =   19
            Top             =   240
            Width           =   1300
         End
         Begin VB.TextBox txtTwilight 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   3720
            TabIndex        =   18
            Top             =   2880
            Width           =   852
         End
         Begin VB.TextBox txtMSL 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1800
            TabIndex        =   17
            Top             =   2880
            Width           =   852
         End
         Begin VB.CheckBox chkHanafi 
            Caption         =   "Hanafi Asr"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   4920
            TabIndex        =   16
            Top             =   2880
            Width           =   1095
         End
         Begin VB.TextBox txtcityname 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1320
            TabIndex        =   15
            Top             =   240
            Width           =   1300
         End
         Begin VB.Frame Frame1 
            Caption         =   "Longitude"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   732
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   4935
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Index           =   5
               Left            =   1080
               TabIndex        =   12
               Top             =   240
               Width           =   492
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Index           =   4
               Left            =   2520
               TabIndex        =   11
               Top             =   240
               Width           =   735
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Index           =   3
               Left            =   3360
               TabIndex        =   10
               Top             =   240
               Width           =   972
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Degrees"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Minutes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   1680
               TabIndex        =   13
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Latitude"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   1920
            Width           =   6135
            Begin VB.CheckBox chkHemisphere 
               Caption         =   "Northern Hemisphere"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   4560
               TabIndex        =   6
               Top             =   240
               Value           =   1  'Checked
               Width           =   1455
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Index           =   2
               Left            =   3360
               TabIndex        =   5
               Top             =   240
               Width           =   972
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Index           =   1
               Left            =   2520
               TabIndex        =   4
               Top             =   240
               Width           =   735
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Index           =   0
               Left            =   1080
               TabIndex        =   3
               Top             =   240
               Width           =   492
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Minutes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   8
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Degrees"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   7
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Label Label6 
            Caption         =   "Time Zone"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   27
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Country"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   200
            TabIndex        =   26
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Region"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Twilight Angle"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2760
            TabIndex        =   24
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Height in metres above sea level"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   23
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label lblCityName 
            Caption         =   "City Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   195
            TabIndex        =   22
            Top             =   240
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IOK As Boolean
Dim city1 As CityType

Private Sub Check1_Click()
If Check1.Value = vbChecked Then
    Settings1.AdanEnabled = True
Else
    Settings1.AdanEnabled = False
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = vbChecked Then
    Settings1.WarningEnabled = True
Else
    Settings1.WarningEnabled = False
End If
End Sub

Private Sub chkHanafi_Click()
city.isHanafi = Not city.isHanafi

End Sub

Private Sub chkHemisphere_Click()
city.Latitude = -1 * city.Latitude

End Sub

Private Sub cmbClock_Change()
pathclockpic = App.Path & "\clocks\" & cmbClock.Text
frmStar.Refresh
End Sub

Private Sub cmbClock_Click()
pathclockpic = App.Path & "\clocks\" & cmbClock.Text
frmStar.Refresh

End Sub

Private Sub cmdApply_Click(Index As Integer)
Select Case Index
    Case 0  'basic location ,namaz settings
        ApplyVals
    Case 1
        Settings1.WarningMinutes = Val(Text2.Text)
        If Check2.Value = vbChecked Then
            Settings1.WarningEnabled = True
        Else
            Settings1.WarningEnabled = False
        End If
        If Check1.Value = vbChecked Then
            Settings1.AdanEnabled = True
        Else
            Settings1.AdanEnabled = False
        End If
        Settings1.adan = cmbAdan.Text
        Settings1.FAdan = cmbFAdan.Text
        Settings1.WarningTxt = txtWarningtxt.Text
    Case 2
        Settings1.TwilightAngle = Val(Text3.Text)
        Settings1.TwilightAngleFajr = Val(txtTAFajr)
        Settings1.TwilightAngleIsha = Val(txtTAIsha)
        Settings1.AdjustmentDays = slrAdjDays.Value
    Case 3
        ApplyInterface
    Case 4      'clock and hands
        Settings1.ClockFace = cmbClock.Text
        
    Case 5  'startup
        If chkStartup(0).Value = vbChecked Then
        
            If chkStartup(1).Value = vbChecked Then
                SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Run", "salat", App.Path & "\" & App.exename & ".exe -m", 1&
            Else
                SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Run", "salat", App.Path & "\" & App.exename & ".exe", 1&
            End If
        
        Else
            If QueryValue("Software\Microsoft\Windows\CurrentVersion\Run", "salat") <> "" Then
                RegDeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\salat"
            End If
            If QueryValue("Software\Microsoft\Windows\CurrentVersion\RunServices", "salat") <> "" Then
               RegDeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices\salat"
            End If
            If QueryValueUser(".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Run", "salat") <> "" Then
               RegDeleteKey HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Run\salat"
            End If
        
        
        
        End If
        
        
    Case Else
End Select
IOK = True
SetSettings Settings1
End Sub

Private Sub cmdCancel_Click(Index As Integer)
Unload Me
End Sub


Private Sub cmdColor_Click(Index As Integer)
'CommonDialog1.CancelError = True
Select Case Index
    Case 0
        CommonDialog1.ShowColor
        Settings1.HourHandColor = CommonDialog1.Color
        
    Case 1
        CommonDialog1.ShowColor
        Settings1.MinuteHandColor = CommonDialog1.Color
    
    Case 2
        CommonDialog1.ShowColor
        Settings1.SecondHandColor = CommonDialog1.Color
    
    Case Else
End Select
End Sub

Private Sub cmdOK_Click(Index As Integer)
Call cmdApply_Click(Index)
Unload Me


End Sub

Private Sub cmdPlayAdan_Click()
'Call sndPlaySound(App.Path & "\adan\" & cmbAdan.Text & Chr$(0), SND_FILENAME Or SND_ASYNC)
AdanCalled = False
CallAdan 0, App.Path & "\adan\" & cmbAdan.Text & Chr$(0)
AdanCalled = False

End Sub

Private Sub cmdPlayFAdan_Click()
'Call sndPlaySound(App.Path & "\adan\" & cmbFAdan.Text & Chr$(0), SND_ASYNC Or SND_FILENAME)
AdanCalled = False
CallAdan 1, App.Path & "\adan\" & cmbFAdan.Text & Chr$(0)
AdanCalled = False

End Sub

Private Sub Form_Load()
On Local Error GoTo loaderr
Dim lng As degreeformat
Dim lat As degreeformat
Dim tmpadan As String
Dim tmpclock As String

city1 = GetCity(city.CityID, "CurrentLoc")
If city1.CityID = 0 Then
    city1 = GetCity(city.CityID, "Locations")
End If
txtcityname = city1.cityname
txtCountry = city1.country
txtMSL = city1.MSL
txtRegion = city1.Region
txtTimezone = city1.TimeZone
txtTwilight = TwilightAngle
lng = Dec2Deg(CCur(city1.Longitude))
lat = Dec2Deg(CCur(city1.Latitude))
If city1.isHanafi Then
    chkHanafi.Value = vbChecked
Else
    chkHanafi.Value = vbUnchecked
End If
If city1.Latitude < 0 Then
    chkHemisphere.Value = vbUnchecked
Else
    chkHemisphere.Value = vbChecked
End If

Text1(0) = lat.Degrees
Text1(1) = lat.Minutes
Text1(2) = city1.Latitude
Text1(3) = city1.Longitude
Text1(4) = lng.Minutes
Text1(5) = lng.Degrees


Settings1 = GetSettings()
sliTransparency.Value = Settings1.Transparency

Do
    tmpadan = GetFiles(App.Path & "\adan", ".wav")
    cmbAdan.AddItem tmpadan
    cmbFAdan.AddItem tmpadan
Loop Until tmpadan = ""
If InStr(Settings1.adan, ":") <> 0 Then
    cmbAdan.Text = GetFileNamefromPath(Settings1.adan)
    cmbFAdan.Text = GetFileNamefromPath(Settings1.FAdan)
Else
    cmbAdan.Text = Settings1.adan
    cmbFAdan.Text = Settings1.FAdan
End If
Do
    tmpclock = GetFiles(App.Path & "\Clocks", ".bmp")
    If LCase$(Right$(tmpclock, 4)) = ".bmp" Then
        cmbClock.AddItem tmpclock
    End If
Loop Until tmpclock = ""

If Settings1.ClockFace <> "" Then
    cmbClock.Text = GetFileNamefromPath(Settings1.ClockFace)
End If
txtTAFajr = Settings1.TwilightAngleFajr
txtTAIsha = Settings1.TwilightAngleIsha
Text3.Text = Settings1.TwilightAngle
txtWarningtxt.Text = Settings1.WarningTxt

slrAdjDays.Value = Settings1.AdjustmentDays  ' slider set

Exit Sub
loaderr:
MsgBox Err.Description
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'clear memory
If IOK = False Then
    If MsgBox("This will cancel all setting changes!", vbOKCancel) = vbCancel Then
        Cancel = True
        Exit Sub
    End If
End If

Set frmSettings = Nothing
End Sub

Private Sub sliTransparency_Change()
Settings1.Transparency = sliTransparency.Value '* 2.5

End Sub

Private Sub sliTransparency_Click()
Settings1.Transparency = sliTransparency.Value '* 2.5
End Sub


Private Sub slrAdjDays_Change()
AdjustmentDays = CInt(slrAdjDays.Value)
End Sub

Private Sub slrAdjDays_Click()
AdjustmentDays = CInt(slrAdjDays.Value)
End Sub

Private Sub Text1_Change(Index As Integer)
On Local Error GoTo stackerr
Select Case Index
    Case 0
        Degree.Degrees = Val(Text1(0))
        Degree.Minutes = Val(Text1(1))
        Text1(2) = CStr(Deg2Dec(Degree))
    Case 1
        Degree.Degrees = Val(Text1(0))
        Degree.Minutes = Val(Text1(1))
        Text1(2) = CStr(Deg2Dec(Degree))
    Case 2
        Text1(0) = CStr(Dec2Deg(CCur(Val(Text1(Index)))).Degrees)
        Text1(1) = CStr(Dec2Deg(CCur(Val(Text1(Index)))).Minutes)
        city1.Latitude = Val(Text1(2).Text)
    Case 3
        Text1(4) = CStr(Dec2Deg(CCur(Val(Text1(Index)))).Minutes) 'minutes
       Text1(5) = CStr(Dec2Deg(CCur(Val(Text1(Index)))).Degrees) 'Degrees
        city1.Longitude = Val(Text1(3).Text)
    Case 4
        Degree.Degrees = Val(Text1(5))
        Degree.Minutes = Val(Text1(4))
        Text1(3) = CStr(Deg2Dec(Degree))
    Case 5
        Degree.Degrees = Val(Text1(5))
        Degree.Minutes = Val(Text1(4))
        Text1(3) = CStr(Deg2Dec(Degree))
        
End Select
Exit Sub
stackerr:
If Err.Number = 28 Then
    Err.Clear
    fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text
    Resume Next
End If


End Sub

Private Sub Text2_Change()
Settings1.WarningMinutes = Val(Text2.Text)
End Sub

Private Sub txtAdjDays_Change()

End Sub

Private Sub txtMSL_Change()
city.MSL = Val(txtMSL.Text)
End Sub

Private Sub txtTAFajr_Change()
Settings1.TwilightAngleFajr = Val(txtTAFajr.Text)
Param.TwilightAngleFajr = Val(txtTAFajr.Text)
End Sub

Private Sub txtTAIsha_Change()
Param.TwilightAngleIsha = Val(txtTAIsha.Text)
Settings1.TwilightAngleIsha = Val(txtTAIsha.Text)
End Sub

Private Sub txtTwilight_Change()
Param.TwilightAngle = Val(txtTwilight.Text)
Settings1.TwilightAngle = Val(txtTwilight.Text)
End Sub

Sub ApplyVals()
Dim rsSave As ADODB.Recordset
Set rsSave = New ADODB.Recordset
rsSave.Open "select * from Locations where CityId=" & city1.CityID & ";", Cnn, adOpenKeyset, adLockOptimistic
If rsSave.RecordCount <= 0 Then
    rsSave.Close
    Set rsSave = Nothing
    Exit Sub
End If
rsSave!CityID = city1.CityID
rsSave!city = city1.cityname
rsSave!country = city1.country
rsSave!Latitude = city1.Latitude
rsSave!Longitude = city1.Longitude
rsSave!MSL = city1.MSL
rsSave!Region = city1.Region
rsSave!TimeZone = city1.TimeZone
If chkHanafi.Value = vbChecked Then
    rsSave!isHanafi = True
    city1.isHanafi = True
Else
    rsSave!isHanafi = False
    city1.isHanafi = False
End If
If chkHemisphere.Value = vbChecked Then
    rsSave!isNorthernHemisphere = True
    city1.isNorthernHemisphere = True
Else
    rsSave!isNorthernHemisphere = False
    city1.isNorthernHemisphere = False
End If

rsSave!TwilightAngle = TwilightAngle

rsSave.Update
rsSave.Close


Set rsSave = New ADODB.Recordset
rsSave.Open "select * from CurrentLoc where CityId=" & city1.CityID & ";", Cnn, adOpenKeyset, adLockOptimistic
If rsSave.RecordCount <= 0 Then
    rsSave.Close
    Set rsSave = Nothing
    Exit Sub
End If
rsSave!CityID = city1.CityID
rsSave!city = city1.cityname
rsSave!country = city1.country
rsSave!Latitude = city1.Latitude
rsSave!Longitude = city1.Longitude
rsSave!MSL = city1.MSL
rsSave!Region = city1.Region
rsSave!TimeZone = city1.TimeZone
If chkHanafi.Value = vbChecked Then
    rsSave!isHanafi = True
    city1.isHanafi = True
Else
    rsSave!isHanafi = False
    city1.isHanafi = False
End If
If chkHemisphere.Value = vbChecked Then
    rsSave!isNorthernHemisphere = True
    city1.isNorthernHemisphere = True
Else
    rsSave!isNorthernHemisphere = False
    city1.isNorthernHemisphere = False
End If

rsSave!TwilightAngle = TwilightAngle

rsSave.Update
rsSave.Close


Set rsSave = Nothing
city = GetCity(city1.CityID, "CurrentLoc")


End Sub

Sub ApplyInterface()
SetTransparency fMainForm, sliTransparency.Value
SetTransparency frmStar, sliTransparency.Value

SetKeyValue "Software\iqsoft\" & App.exename & "\" & "settings", "transparency", sliTransparency.Value, 1&
End Sub

Private Sub txtWarningtxt_Change()
Settings1.WarningTxt = txtWarningtxt.Text
End Sub
