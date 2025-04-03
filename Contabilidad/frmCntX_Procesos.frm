VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmCntX_Procesos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "..."
   ClientHeight    =   996
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   996
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Left            =   240
      Top             =   720
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   170
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   5535
      _ExtentX        =   9758
      _ExtentY        =   296
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Image Image1 
      Height          =   576
      Left            =   120
      Picture         =   "frmCntX_Procesos.frx":0000
      Top             =   120
      Width           =   576
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "frmCntX_Procesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer


Private Sub Form_Load()
i = 1
End Sub

Private Sub TimerX_Timer()
prgBar.Max = 1000

If i = 1000 Then i = 1

prgBar.Value = i

i = 1

End Sub
