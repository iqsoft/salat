VERSION 5.00
Begin VB.Form dlgSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select your city"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "dlgSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Selct your Country, Region or State and City"
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox cmbCity 
         Height          =   315
         Left            =   2280
         TabIndex        =   5
         Top             =   2040
         Width           =   2055
      End
      Begin VB.ComboBox cmbRegion 
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Top             =   1200
         Width           =   2055
      End
      Begin VB.ComboBox cmbCountry 
         Height          =   315
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "City"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Region/State"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Country"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   1095
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
      Caption         =   "Select"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "dlgSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim strSQL As String
Dim Loaded As Boolean

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
Dim ans As VbMsgBoxResult
Dim rsCity As ADODB.Recordset
Set rsCity = New ADODB.Recordset
rsCity.Open "select * from Locations where Country ='" & cmbCountry & "' and region ='" & cmbRegion & "' and city ='" & cmbCity & "';", Cnn, adOpenKeyset, adLockOptimistic
If rsCity.RecordCount <= 0 Then
    ans = MsgBox("sorry, can't find your city... Do you want to add it to the database?", vbYesNo)
    If ans = vbNo Then
        Exit Sub
    End If
    frmNewLoc.Show
    Unload Me
    Exit Sub
End If
city.CityID = CLng(rsCity!CityID)
city = GetCity(city.CityID, "Locations")
rsCity.Close
Set rsCity = Nothing

AddCurrentCity


Unload Me
End Sub
