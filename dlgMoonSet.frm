VERSION 5.00
Begin VB.Form dlgMoonSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Moon Set Times"
   ClientHeight    =   5160
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6330
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "dlgMoonSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   2640
      Top             =   360
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moon's Properties Now..."
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   6135
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   2
         Left            =   2500
         TabIndex        =   8
         Top             =   1800
         Width           =   3300
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   1
         Left            =   2500
         TabIndex        =   7
         Top             =   1200
         Width           =   3300
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Index           =   0
         Left            =   2500
         TabIndex        =   6
         Top             =   600
         Width           =   3300
      End
      Begin VB.Label Label1 
         Caption         =   "Moon's Distance"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Moon's Longitude"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Moon's Latitude"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
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
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "dlgMoonSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mLat As Double
Dim mLon As Double
Dim mR As Double
Dim JD As Double
Const earthRad = 6371.0088  'km


Private Sub CancelButton_Click()
Unload Me

End Sub

Private Sub Form_Load()
On Local Error GoTo frmloaderr
Randomize Timer
Timer1.Interval = 100
Timer1.Enabled = True

Exit Sub
frmloaderr:
MsgBox " Moon set " & Err.Description & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next

End Sub

Private Sub OKButton_Click()
Unload Me

End Sub

Private Sub Timer1_Timer()
On Local Error GoTo timer1err
Dim i As Byte
For i = 0 To 2
    Label2(i).Caption = ""
Next i
'JD = jday(year(Now), month(Now), day(Now), hour(Now), Minute(Now), Second(Now))
JD = calcJD(Now, Val(city.TimeZone))
mR = smoon(JD, 1)  'Radius
mLat = smoon(JD, 2) 'latitude
mLon = smoon(JD, 3) 'longitude
Label2(0).Caption = mLat & " Degrees"
Label2(1).Caption = mLon & " Degrees"
Label2(2).Caption = mR * earthRad & " km"


Exit Sub

timer1err:
MsgBox " Timer " & Err.Description & Err.Number
fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next

End Sub
