VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6840
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1053"
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1658
      TabIndex        =   1
      Tag             =   "1060"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2888
      TabIndex        =   3
      Tag             =   "1059"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4088
      TabIndex        =   5
      Tag             =   "1058"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3540
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3597.097
      ScaleMode       =   0  'User
      ScaleWidth      =   6412.68
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   6345
      Begin VB.Frame fraCon 
         Caption         =   "Interface"
         Height          =   3346
         Index           =   3
         Left            =   -19789
         TabIndex        =   11
         Tag             =   "1057"
         Top             =   118
         Width           =   5442
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3540
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3597.097
      ScaleMode       =   0  'User
      ScaleWidth      =   6412.68
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   6345
      Begin VB.Frame fraCon 
         Caption         =   "Reports"
         Height          =   3346
         Index           =   2
         Left            =   -19789
         TabIndex        =   10
         Tag             =   "1056"
         Top             =   118
         Width           =   5442
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   120
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   6412.68
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   6345
      Begin VB.Frame fraCon 
         Caption         =   "Location"
         Height          =   3346
         Index           =   0
         Left            =   208
         TabIndex        =   4
         Tag             =   "1054"
         Top             =   120
         Width           =   6026
         Begin VB.Frame Frame1 
            Caption         =   "Latitude"
            Height          =   732
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   2040
            Width           =   5772
            Begin VB.TextBox Text1 
               Height          =   288
               Index           =   0
               Left            =   840
               TabIndex        =   29
               Top             =   240
               Width           =   492
            End
            Begin VB.TextBox Text1 
               Height          =   288
               Index           =   1
               Left            =   2040
               TabIndex        =   28
               Top             =   240
               Width           =   492
            End
            Begin VB.TextBox Text1 
               Height          =   288
               Index           =   2
               Left            =   3000
               TabIndex        =   27
               Top             =   240
               Width           =   972
            End
            Begin VB.CheckBox chkHemisphere 
               Caption         =   "Northern Hemisphere"
               Height          =   372
               Left            =   4200
               TabIndex        =   26
               Top             =   240
               Value           =   1  'Checked
               Width           =   1332
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Degrees"
               Height          =   252
               Index           =   0
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   732
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Minutes"
               Height          =   252
               Index           =   1
               Left            =   1440
               TabIndex        =   30
               Top             =   240
               Width           =   612
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Longitude"
            Height          =   732
            Index           =   1
            Left            =   200
            TabIndex        =   19
            Top             =   1080
            Width           =   4212
            Begin VB.TextBox Text1 
               Height          =   288
               Index           =   3
               Left            =   3000
               TabIndex        =   22
               Top             =   240
               Width           =   972
            End
            Begin VB.TextBox Text1 
               Height          =   288
               Index           =   4
               Left            =   2040
               TabIndex        =   21
               Top             =   240
               Width           =   492
            End
            Begin VB.TextBox Text1 
               Height          =   288
               Index           =   5
               Left            =   840
               TabIndex        =   20
               Top             =   240
               Width           =   492
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Minutes"
               Height          =   252
               Index           =   2
               Left            =   1440
               TabIndex        =   24
               Top             =   240
               Width           =   612
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Degrees"
               Height          =   252
               Index           =   3
               Left            =   120
               TabIndex        =   23
               Top             =   240
               Width           =   732
            End
         End
         Begin VB.TextBox txtcityname 
            Height          =   288
            Left            =   1080
            TabIndex        =   18
            Top             =   240
            Width           =   1092
         End
         Begin VB.CheckBox chkHanafi 
            Caption         =   "Hanafi Asr"
            Height          =   372
            Left            =   4680
            TabIndex        =   17
            Top             =   1320
            Width           =   972
         End
         Begin VB.TextBox txtMSL 
            Height          =   288
            Left            =   2760
            TabIndex        =   16
            Top             =   2880
            Width           =   852
         End
         Begin VB.TextBox txtTwilight 
            Height          =   288
            Left            =   4800
            TabIndex        =   15
            Top             =   2880
            Width           =   852
         End
         Begin VB.TextBox txtRegion 
            Height          =   288
            Left            =   3360
            TabIndex        =   14
            Top             =   240
            Width           =   1092
         End
         Begin VB.TextBox txtCountry 
            Height          =   288
            Left            =   1080
            TabIndex        =   13
            Top             =   720
            Width           =   1092
         End
         Begin VB.TextBox txtTimezone 
            Height          =   288
            Left            =   3360
            TabIndex        =   12
            ToolTipText     =   "Negative Values for West "
            Top             =   720
            Width           =   1092
         End
         Begin VB.Label lblCityName 
            Caption         =   "City Name"
            Height          =   255
            Left            =   200
            TabIndex        =   37
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Height in metres above sea level"
            Height          =   255
            Left            =   200
            TabIndex        =   36
            Top             =   2880
            Width           =   2415
         End
         Begin VB.Label Label3 
            Caption         =   "Twilight Angle"
            Height          =   255
            Left            =   3720
            TabIndex        =   35
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Region"
            Height          =   255
            Left            =   2400
            TabIndex        =   34
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Country"
            Height          =   255
            Left            =   200
            TabIndex        =   33
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Time Zone"
            Height          =   255
            Left            =   2400
            TabIndex        =   32
            Top             =   720
            Width           =   855
         End
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7488
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Locations"
            Object.ToolTipText     =   "Options for location"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sounds"
            Object.ToolTipText     =   "Options for sounds"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports"
            Object.ToolTipText     =   "Options for reports"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Interface"
            Object.ToolTipText     =   "Options for Interface"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   6412.68
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   6345
      Begin VB.Frame fraCon 
         Caption         =   "Sounds"
         Height          =   3346
         Index           =   1
         Left            =   -19789
         TabIndex        =   9
         Tag             =   "1055"
         Top             =   118
         Width           =   5442
         Begin VB.CheckBox chkSnd 
            Caption         =   "Sound ON"
            Height          =   615
            Left            =   1920
            TabIndex        =   38
            Top             =   720
            Value           =   1  'Checked
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim city1 As CityType

Private Sub chkHanafi_Click()
city1.isHanafi = Not city1.isHanafi

End Sub

Private Sub chkHemisphere_Click()
city1.Latitude = -1 * city1.Latitude

End Sub

Private Sub Form_Load()
Dim lng As degreeformat
Dim lat As degreeformat

city1 = GetCity(city.CityID)
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

  For i = 0 To fraCon.Count - 1
   With fraCon(i)
      .Move tbsOptions.ClientLeft, _
      tbsOptions.ClientTop, _
      tbsOptions.ClientWidth, _
      tbsOptions.ClientHeight
   End With
  Next i
fraCon(0).ZOrder 0


End Sub

Private Sub cmdApply_Click()
ApplyVals
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
ApplyVals
Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    i = tbsOptions.SelectedItem.Index
    'handle ctrl+tab to move to the next tab
    If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
        If i = 1 Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.Tabs.Count)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i - 1)
        End If
    End If


End Sub



Private Sub tbsOptions_Click()
fraCon(tbsOptions.SelectedItem.Index - 1).ZOrder 0

    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
            
            fraCon(i).Left = 210
            fraCon(i).Enabled = True
        fraCon(i).ZOrder 0
        fraCon(i).Left = tbsOptions.Left + tbsOptions.ClientLeft
    fraCon(i).Top = tbsOptions.Top + tbsOptions.ClientTop
    fraCon(i).Width = tbsOptions.ClientWidth
    fraCon(i).Height = tbsOptions.ClientHeight


        picOptions(i).Left = tbsOptions.Left + tbsOptions.ClientLeft
    picOptions(i).Top = tbsOptions.Top + tbsOptions.ClientTop
    picOptions(i).Width = tbsOptions.ClientWidth
    picOptions(i).Height = tbsOptions.ClientHeight
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        
            fraCon(i).Left = -20000
            fraCon(i).Enabled = False
        
        End If
    Next
    

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
    Resume Next
End If


End Sub

Private Sub txtMSL_Change()
city1.MSL = Val(txtMSL.Text)
End Sub

Private Sub txtTwilight_Change()
Param.TwilightAngle = Val(txtTwilight.Text)
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
city = GetCity(city1.CityID)


End Sub
