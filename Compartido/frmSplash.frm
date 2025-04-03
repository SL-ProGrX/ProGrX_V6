VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3555
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   9780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   9013
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   120
      Top             =   1920
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   4  'Align Right
      Height          =   3555
      Left            =   9615
      TabIndex        =   3
      Top             =   0
      Width           =   165
      _ExtentX        =   291
      _ExtentY        =   6271
      _Version        =   393216
      Appearance      =   0
      Max             =   120
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Label lblDerechos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":000C
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   492
      Left            =   -120
      TabIndex        =   2
      Top             =   3240
      Width           =   9612
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "lbl"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Image imgSplash 
      Appearance      =   0  'Flat
      Height          =   1320
      Left            =   1920
      Picture         =   "frmSplash.frx":0098
      Top             =   1080
      Width           =   5025
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 lblTitulo.Caption = App.ProductName
 lblVersion.Caption = "v" & App.Major & "." & App.Minor & "." & Format(App.Revision, "00") & ".r" & GLOBALES.SysVersion
 
 
 lblTitulo.ForeColor = RGB(70, 111, 178)
 lblDerechos.ForeColor = RGB(70, 111, 178)
 
 
' imgSplash.Width = Me.Width
' imgSplash.Height = Me.Height
End Sub



Private Sub Timer1_Timer()
If prgBar.Value = prgBar.Max Then
  Timer1.Interval = 0
  Unload Me
Else
  prgBar.Value = prgBar.Value + 1
End If
End Sub
