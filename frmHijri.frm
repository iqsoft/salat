VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHijri 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hijri Date"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8310
   Icon            =   "frmHijri.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   8310
   Begin VB.Frame Frame1 
      Caption         =   "Hijri Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   3480
      TabIndex        =   1
      Top             =   200
      Width           =   4575
      Begin VB.CommandButton cmdToGreg 
         Caption         =   "Convert to Gregorian Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   6
         Top             =   1200
         Width           =   3615
      End
      Begin VB.ComboBox cmbHyear 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   5
         Text            =   "Hijri year"
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cmbHmonth 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         TabIndex        =   4
         Text            =   "Hijri Month"
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox cmbHDay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   3
         Text            =   "Day"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblHijri 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   495
         TabIndex        =   2
         Top             =   3000
         Width           =   90
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Gregorian Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   200
      Width           =   2925
      Begin MSComCtl2.DTPicker dtpGreg 
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   127664129
         CurrentDate     =   44586
      End
      Begin VB.Label lblGreg 
         Caption         =   "Click on a date above to see the Hijri Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Width           =   2700
      End
   End
End
Attribute VB_Name = "frmHijri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Hdate As String

Private Sub cmdToGreg_Click()
Dim Gdate As String
Dim Hdate As String

Hdate = Trim(cmbHDay.Text) & " " & Trim(cmbHmonth.Text) & " " & Trim(cmbHyear.Text)
Gdate = getGregorianDate(Hdate)
lblGreg.Caption = Gdate
'Gdate = Replace(Gdate, " ", "-")
Gdate = FormatDateTime(Gdate, vbLongDate)
dtpGreg.Value = DateValue(Gdate)
lblHijri.Caption = Hdate
End Sub


Private Sub dtpGreg_Change()
setHijri
lblGreg.Caption = dtpGreg.Value
End Sub

Private Sub dtpGreg_Click()
setHijri
lblGreg.Caption = dtpGreg.Value
End Sub

Private Sub dtpGreg_DateClick(ByVal DateClicked As Date)
setHijri

End Sub

Private Sub dtpGreg_DateDblClick(ByVal DateDblClicked As Date)
setHijri

End Sub

Private Sub dtpGreg_DblClick()
setHijri

End Sub

Private Sub dtpGreg_LostFocus()
setHijri

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then 'escape key
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim i As Byte
CenterObj Screen, Me
dtpGreg.Value = Now
setHijri
For i = 1 To 30
    cmbHDay.AddItem str(i)
Next i
For i = 1 To 254
    cmbHyear.AddItem str(1300 + i)
Next i

For i = 1 To 12
    cmbHmonth.AddItem HijriMonths(i)
Next i

setHijriCombos (Date)


End Sub

Sub setHijri()
Hdate = GetHijriDate(dtpGreg.Value)
lblHijri = Hdate
setHijriCombos (dtpGreg.Value)
End Sub

Sub setHijriCombos(GregDate As Date)
Dim Hday As String
Dim HMonth As String
Dim HYear As String

Hdate = GetHijriDayMonthYear(GregDate, Hday, HMonth, HYear)
cmbHDay.Text = Hday
cmbHmonth.Text = HMonth
cmbHyear.Text = HYear


End Sub

