VERSION 5.00
Begin VB.Form dlgEditCity 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit City"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "dlgEditCity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTimezone 
      Height          =   288
      Left            =   3240
      TabIndex        =   26
      ToolTipText     =   "Negative Values for West "
      Top             =   600
      Width           =   1092
   End
   Begin VB.TextBox txtCountry 
      Height          =   288
      Left            =   960
      TabIndex        =   24
      Top             =   600
      Width           =   1092
   End
   Begin VB.TextBox txtRegion 
      Height          =   288
      Left            =   3240
      TabIndex        =   22
      Top             =   120
      Width           =   1092
   End
   Begin VB.TextBox txtTwilight 
      Height          =   288
      Left            =   5040
      TabIndex        =   20
      Top             =   2760
      Width           =   852
   End
   Begin VB.TextBox txtMSL 
      Height          =   288
      Left            =   2640
      TabIndex        =   18
      Top             =   2760
      Width           =   852
   End
   Begin VB.CheckBox chkHanafi 
      Caption         =   "Hanafi Asr"
      Height          =   372
      Left            =   4680
      TabIndex        =   17
      Top             =   1200
      Width           =   972
   End
   Begin VB.TextBox txtcityname 
      Height          =   288
      Left            =   960
      TabIndex        =   16
      Top             =   120
      Width           =   1092
   End
   Begin VB.Frame Frame1 
      Caption         =   "Longitude"
      Height          =   732
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   4212
      Begin VB.TextBox Text1 
         Height          =   288
         Index           =   5
         Left            =   840
         TabIndex        =   11
         Top             =   240
         Width           =   492
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Index           =   4
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   492
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Index           =   3
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   972
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Degrees"
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   732
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes"
         Height          =   252
         Index           =   2
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   612
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Latitude"
      Height          =   732
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   5772
      Begin VB.CheckBox chkHemisphere 
         Caption         =   "Northern Hemisphere"
         Height          =   372
         Left            =   4200
         TabIndex        =   27
         Top             =   240
         Value           =   1  'Checked
         Width           =   1332
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Index           =   2
         Left            =   3000
         TabIndex        =   7
         Top             =   240
         Width           =   972
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Index           =   1
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   492
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   492
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes"
         Height          =   252
         Index           =   1
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   612
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Degrees"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   732
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Time Zone"
      Height          =   252
      Left            =   2280
      TabIndex        =   25
      Top             =   600
      Width           =   852
   End
   Begin VB.Label Label5 
      Caption         =   "Country"
      Height          =   252
      Left            =   120
      TabIndex        =   23
      Top             =   600
      Width           =   732
   End
   Begin VB.Label Label4 
      Caption         =   "Region"
      Height          =   252
      Left            =   2280
      TabIndex        =   21
      Top             =   120
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "Twilight Angle"
      Height          =   252
      Left            =   3840
      TabIndex        =   19
      Top             =   2760
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "Height in metres above sea level"
      Height          =   252
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   2412
   End
   Begin VB.Label lblCityName 
      Caption         =   "City Name"
      Height          =   252
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   852
   End
End
Attribute VB_Name = "dlgEditCity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim city1 As CityType


Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub chkHanafi_Click()
city1.isHanafi = chkHanafi.Value
End Sub

Private Sub chkHemisphere_Click()
city1.isNorthernHemisphere = chkHemisphere.Value
End Sub

Private Sub Form_Load()
Dim lng As degreeformat
Dim lat As degreeformat

city1 = GetCity(city.CityID, "CurrentLoc")
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



End Sub

Private Sub Form_Unload(Cancel As Integer)
CityEditMode = False

End Sub

Private Sub OKButton_Click()
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
rsSave.Open "select * from CurrentLoc where CityId=" & city1.CityID & " order by ID;", Cnn, adOpenKeyset, adLockOptimistic
If rsSave.RecordCount <= 0 Then
    rsSave.Close
    Set rsSave = Nothing
    Exit Sub
End If
rsSave.MoveLast
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

Unload Me
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
  '      Text1(0) = CStr(Dec2Deg(CCur(Val(Text1(Index)))).Degrees)
   '     Text1(1) = CStr(Dec2Deg(CCur(Val(Text1(Index)))).Minutes)
        city1.Latitude = Val(Text1(2).Text)
    Case 3
     '   Text1(4) = CStr(Dec2Deg(CCur(Val(Text1(Index)))).Minutes)
      '  Text1(5) = CStr(Dec2Deg(CCur(Val(Text1(Index)))).Degrees)
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
    fMainForm.sbStatusBar.Panels(1).ToolTipText = fMainForm.sbStatusBar.Panels(1).Text: Resume Next
End If


End Sub

Private Sub txtcityname_Change()
city1.cityname = txtcityname.Text
End Sub

Private Sub txtCountry_Change()
city1.country = txtCountry.Text
End Sub

Private Sub txtMSL_Change()
city1.MSL = Val(txtMSL.Text)
End Sub

Private Sub txtRegion_Change()
city1.Region = txtRegion.Text
End Sub

Private Sub txtTimezone_Change()
city1.TimeZone = txtTimezone.Text
End Sub

Private Sub txtTwilight_Change()
TwilightAngle = Val(txtTwilight.Text)
End Sub
