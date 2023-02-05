VERSION 5.00
Begin VB.Form dlgReport 
   Caption         =   "Create Salat Times for a year"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5145
   FillColor       =   &H80000014&
   Icon            =   "dlgReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Create report for a period of:"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   4935
      Begin VB.OptionButton optAny 
         Caption         =   "Any Period"
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "One Week"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "One Month"
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optYear 
         Caption         =   "One Year"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Hanafi School"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Select"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selct your Country, Region or State and City"
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3495
      Begin VB.ComboBox cmbCountry 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cmbRegion 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox cmbCity 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Country"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Region/State"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "City"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
   End
End
Attribute VB_Name = "dlgReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQL As String
Dim Loaded As Boolean
Dim repCity As CityType
Dim isHanafi As Boolean
Dim OptionTimePeriod As Integer

Private Sub CancelButton_Click()
Unload Me

End Sub

Private Sub cmbCountry_Change()
'LoadRegion

End Sub

Private Sub cmbCountry_Click()
LoadRegion
End Sub

Private Sub cmbRegion_Change()
LoadCity
End Sub

Private Sub cmbRegion_Click()

LoadCity
End Sub

Private Sub Form_Load()
On Local Error Resume Next
fMainForm.lblLoc.Caption = "Please wait... Loading Cities of the world..."
GetField Me.cmbCountry, "Locations", "Country"
Me.Refresh
Loaded = True
OptionTimePeriod = "Year"
End Sub
Sub LoadCity()
If Not Loaded Then
    MsgBox "Select a country and region first"
    Exit Sub
End If
cmbCity.Clear
strSQL = "Select * from Locations where Region='" & cmbRegion.Text & "';"
Call GetField(Me.cmbCity, strSQL, "City")

End Sub

Sub LoadRegion()
If Not Loaded Then
    MsgBox "Select a country first"
    Exit Sub
End If
cmbRegion.Clear
cmbCity.Clear
strSQL = "Select * from Locations where Country='" & cmbCountry.Text & "';"
GetField Me.cmbRegion, strSQL, "Region"

End Sub

Private Sub Form_Unload(Cancel As Integer)
fMainForm.Timer1.Enabled = True
End Sub

Private Sub OKButton_Click()

Dim startDate As Date
Dim rsCity As ADODB.Recordset
Set rsCity = New ADODB.Recordset
rsCity.Open "select * from Locations where Country ='" & cmbCountry & "' and region ='" & cmbRegion & "' and city ='" & cmbCity & "';", Cnn, adOpenKeyset, adLockOptimistic
If rsCity.RecordCount <= 0 Then
    MsgBox "sorry, can't find your city... Add it to the database."
    rsCity.Close
    Set rsCity = Nothing
    frmNewLoc.Show
    Exit Sub
End If
repCity.CityID = CLng(rsCity!CityID)
repCity = GetCity(repCity.CityID, "Locations")
rsCity.Close
Set rsCity = Nothing
If Check1.Value = vbChecked Then
    isHanafi = True
Else
    isHanafi = False
End If
If OptionTimePeriod = 0 Then
    OptionTimePeriod = 365
End If
startDate = InputBox("Starting from which Date?", "Start Date", "01-Jan-" & year(Now))
ViewReport repCity.CityID, isHanafi, OptionTimePeriod, startDate

Unload Me
End Sub

Private Sub optAny_Click()
OptionTimePeriod = InputBox("Show Salat Timings for how many days?", "Custom period", 365)


End Sub

Private Sub optMonth_Click()
OptionTimePeriod = 30
End Sub

Private Sub optWeek_Click()
OptionTimePeriod = 7
End Sub

Private Sub optYear_Click()
OptionTimePeriod = 365
End Sub
