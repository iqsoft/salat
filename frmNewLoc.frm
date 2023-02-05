VERSION 5.00
Begin VB.Form frmNewLoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recently Used Locations"
   ClientHeight    =   3690
   ClientLeft      =   1095
   ClientTop       =   360
   ClientWidth     =   4905
   Icon            =   "frmNewLoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4905
   Begin VB.CheckBox Check1 
      Caption         =   "Hanafi School"
      DataField       =   "isHanafi"
      Height          =   375
      Left            =   2040
      TabIndex        =   28
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MSL"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   12
      Top             =   2010
      Width           =   2800
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   4905
      TabIndex        =   24
      Top             =   3084
      Width           =   4908
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   300
         Left            =   1920
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   360
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Select"
         Height          =   300
         Left            =   3600
         TabIndex        =   16
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   1920
         TabIndex        =   26
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add New"
         Height          =   300
         Left            =   360
         TabIndex        =   25
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   4905
      TabIndex        =   18
      Top             =   3384
      Width           =   4908
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Picture         =   "frmNewLoc.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Picture         =   "frmNewLoc.frx":0D44
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Picture         =   "frmNewLoc.frx":1086
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Picture         =   "frmNewLoc.frx":13C8
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   23
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CityId"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   13
      Top             =   2340
      Width           =   2800
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TimeZone"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   1685
      Width           =   2800
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Longitude"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1360
      Width           =   2800
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Latitude"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1035
      Width           =   2800
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Country"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   710
      Width           =   2800
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Region"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   385
      Width           =   2800
   End
   Begin VB.TextBox txtFields 
      DataField       =   "City"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   2800
   End
   Begin VB.Label lblLabels 
      Caption         =   "Metres Above Sea Level"
      Height          =   252
      Index           =   7
      Left            =   120
      TabIndex        =   27
      Top             =   2016
      Width           =   1812
   End
   Begin VB.Label lblLabels 
      Caption         =   "CityId:"
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   2340
      Width           =   1812
   End
   Begin VB.Label lblLabels 
      Caption         =   "TimeZone:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1690
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Longitude:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1364
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Latitude:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1038
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Country:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   712
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Region:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   386
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "City:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmNewLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents PrimaryCLS As clsLocations
Attribute PrimaryCLS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Form_Load()
  Set PrimaryCLS = New clsLocations

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.DataMember = "Primary"
    Set oText.DataSource = PrimaryCLS
  Next
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      'cmdClose_Click
      Unload Me
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
    Case vbKeyReturn
        cmdClose_Click
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
  fMainForm.Timer1.Enabled = True
End Sub

Private Sub PrimaryCLS_MoveComplete()
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(PrimaryCLS.AbsolutePosition)
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  PrimaryCLS.AddNew
  lblStatus.Caption = "Add record"
  mbAddNewFlag = True
  SetButtons False

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  PrimaryCLS.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  PrimaryCLS.Cancel
  SetButtons True
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  PrimaryCLS.Update
  SetButtons True
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
Dim rsCity As ADODB.Recordset
Set rsCity = New ADODB.Recordset

    city.cityname = txtFields(0).Text
    city.Region = txtFields(1).Text
    city.country = txtFields(2).Text
    city.Latitude = txtFields(3).Text
    city.Longitude = txtFields(4).Text
    city.TimeZone = txtFields(5).Text
    city.MSL = txtFields(7).Text
    If Check1.Value = vbChecked Then
        city.isHanafi = True
    Else
        city.isHanafi = False
    End If


rsCity.Open "select * from CurrentLoc where Country ='" & city.country & "' and region ='" & _
    city.Region & "' and city ='" & city.cityname & "';", Cnn, adOpenStatic, adLockOptimistic
If rsCity.RecordCount <= 0 Then
    If Not CityExists(txtFields(2), txtFields(1), txtFields(0)) Then
        AddLocation city
        city = GetCity(city.CityID, "Locations")
    End If
Else
    AddCurrentCity
End If


rsCity.Close
Set rsCity = Nothing
 
Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  PrimaryCLS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  PrimaryCLS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  PrimaryCLS.MoveNext
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  PrimaryCLS.MovePrevious
  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

