VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLocations 
   Caption         =   "Locations"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
   Icon            =   "frmLocations.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5915
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Select City"
      Height          =   495
      Left            =   4030
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmLocations.frx":08CA
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   2
      WrapCellPointer =   -1  'True
      RowDividerStyle =   3
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Select A City"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GridClicked As Boolean


Private Sub Command1_Click()

If Not GridClicked Then
    Unload Me
End If
Dim salattimes As salattype
Dim oldcity As CityType

If DataGrid1.Col <> 0 Then
    DataGrid1.Col = 0
End If
oldcity = city
city.CityID = DataGrid1.Text
city = GetCity(city.CityID, "Locations")
If CityEditMode Then
    CityEditMode = False
    CenterIt fMainForm, dlgEditCity
    dlgEditCity.Show
    Unload Me
    Exit Sub
End If
If MsgBox("Set the current Location as " & city.cityname & "?", vbQuestion + vbYesNo, "Current Location Changed!!") = vbNo Then
    city = oldcity

    Unload Me
    Exit Sub
End If

AddCurrentCity

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DataGrid1_Click()
GridClicked = True
End Sub

Private Sub Form_Activate()
If CityEditMode Then
    Me.Caption = "Select Location to Edit"
Else
    Me.Caption = "Select Location"
End If

End Sub

Private Sub Form_Load()
On Local Error Resume Next
StopTimer = True
Dim rsLoc As ADODB.Recordset
Dim strSQL As String
Set rsLoc = New ADODB.Recordset

If CityEditMode Then
    strSQL = "Select * from CurrentLoc order by City"
Else
    strSQL = "Select * from Locations order by City"
End If
rsLoc.Open strSQL, Cnn, adOpenStatic, adLockOptimistic

Set DataGrid1.DataSource = rsLoc

DataGrid1.Columns(0).DataField = rsLoc.Fields(0)
DataGrid1.Columns(1).DataField = rsLoc.Fields(1)
DataGrid1.Columns(2).DataField = rsLoc.Fields(2)
DataGrid1.Columns(3).DataField = rsLoc.Fields(3)
DataGrid1.Columns(4).DataField = rsLoc.Fields(4)
DataGrid1.Columns(5).DataField = rsLoc.Fields(5)
DataGrid1.Columns(6).DataField = rsLoc.Fields(6)
If Not IsNull(rsLoc.Fields(7)) Then
    If DataGrid1.Columns(7).DataField = "" Then
        DataGrid1.Columns(7).DataField = 0
    Else
        DataGrid1.Columns(7).DataField = rsLoc.Fields(7)
    End If
Else
    DataGrid1.Columns(7).DataField = "0"
End If
If Not IsNull(rsLoc.Fields(8)) Then
    DataGrid1.Columns(8).DataField = rsLoc.Fields(8)
Else
    DataGrid1.Columns(8).DataField = False
End If
If Not IsNull(rsLoc.Fields(9)) Then
    DataGrid1.Columns(9).DataField = rsLoc.Fields(9)
Else
    DataGrid1.Columns(9).DataField = True
End If
DataGrid1.Columns(0).Width = 600
DataGrid1.Columns(1).Width = 1600
DataGrid1.Columns(2).Width = 1500
DataGrid1.Columns(3).Width = 1500
DataGrid1.Columns(4).Width = 800
DataGrid1.Columns(5).Width = 800
DataGrid1.Columns(6).Width = 800
DataGrid1.Columns(7).Width = 800
DataGrid1.Columns(8).Width = 800
DataGrid1.Columns(9).Width = 1800



End Sub

Sub CreateLocations() 'to be used only once
Dim dbs As Database
Dim strSQL As String
Set dbs = OpenDatabase(DataPath)

strSQL = "Select all Cities.CityId, Cities.City, Regions.Region, Countries.Country," _
& "Cities.Latitude, Cities.Longitude, Cities.TimeZone into Locations from Cities inner join " _
& "(Regions inner join Countries on  Regions.CountryID = Countries.CountryID)" _
& "on Regions.RegionID = Cities.RegionID order by City"
dbs.Execute strSQL & ";"
DoEvents
dbs.Close

End Sub

Private Sub Form_Resize()
'CenterIt Screen, Me
'Me.Height = Screen.Height - 100
'DataGrid1.Height = Me.Height - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
StopTimer = False
fMainForm.Timer1.Enabled = True
End Sub
