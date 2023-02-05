VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   7450
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   7380
      Begin VB.PictureBox picLogo 
         Height          =   2376
         Left            =   960
         Picture         =   "frmSplash.frx":886A
         ScaleHeight     =   2310
         ScaleWidth      =   2955
         TabIndex        =   2
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LicenseTo"
         ForeColor       =   &H00000080&
         Height          =   975
         Left            =   2280
         TabIndex        =   1
         Tag             =   "1052"
         Top             =   240
         Width           =   1935
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   720
         Left            =   1800
         TabIndex        =   9
         Tag             =   "1051"
         Top             =   4440
         Width           =   2190
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CompanyProduct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   1800
         TabIndex        =   8
         Tag             =   "1050"
         Top             =   3960
         Width           =   3000
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Platform"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   3120
         TabIndex        =   7
         Tag             =   "1049"
         Top             =   5040
         Width           =   3420
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   3840
         TabIndex        =   6
         Tag             =   "1048"
         Top             =   5520
         Width           =   2610
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Warning"
         ForeColor       =   &H00000080&
         Height          =   795
         Left            =   1800
         TabIndex        =   3
         Tag             =   "1047"
         Top             =   6480
         Width           =   3735
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Tag             =   "1046"
         Top             =   6120
         Width           =   3855
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Tag             =   "1045"
         Top             =   5880
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'LoadResStrings Me
    lblProductName.FontSize = 14
    lblProductName.FontBold = True
    lblProductName.Caption = UCase$(App.title)
    
    lblVersion.FontSize = 12
    lblVersion.FontBold = False
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
 
 
    lblCompany.FontSize = 10
    lblCompany.FontBold = True
    lblCompany = "IQSoft Software Consultants http://iqsoft.co.in"
    
    lblCompanyProduct.FontSize = 14
    lblCompanyProduct.FontBold = True
    lblCompanyProduct = "Prayer Timings for the Universe"
    
    lblCopyright.FontSize = 10
    lblCopyright.FontBold = True
    lblCopyright = "(c) 2017-2021 Mohamed Iqbal Pallipurath"
    
    lblLicenseTo = "Unregistered Version. For your Registered copy send email to mohamediqbalp@gmail.com"
    lblPlatform.FontSize = 12
    lblPlatform = "Windows 7/8/10"
    lblWarning = "This software is freeware so long as no commercial use is made of it."
    
CreateCrescent Me

SetTransparency Me, Settings1.Transparency

End Sub

