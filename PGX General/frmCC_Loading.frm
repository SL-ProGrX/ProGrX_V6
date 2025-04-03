VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmCC_Loading 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.Label lblEstado 
      Height          =   1215
      Left            =   5760
      TabIndex        =   0
      Top             =   720
      Width           =   3615
      _Version        =   1245185
      _ExtentX        =   6376
      _ExtentY        =   2143
      _StockProps     =   79
      Caption         =   "Iniciando..."
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgX 
      Height          =   3300
      Left            =   0
      Picture         =   "frmCC_Loading.frx":0000
      Stretch         =   -1  'True
      Tag             =   "1200"
      Top             =   -360
      Width           =   9555
   End
End
Attribute VB_Name = "frmCC_Loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Me.BackColor = RGB(70, 111, 178)

End Sub
