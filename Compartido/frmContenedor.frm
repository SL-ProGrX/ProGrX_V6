VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#19.1#0"; "Codejock.SkinFramework.v19.1.0.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmContenedor 
   Caption         =   "Contenedor"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11532
   Icon            =   "frmContenedor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   11532
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgLogon 
      Left            =   5280
      Top             =   0
      _ExtentX        =   974
      _ExtentY        =   974
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContenedor.frx":071A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContenedor.frx":108E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContenedor.frx":186D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolbarIcons 
      Left            =   4440
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContenedor.frx":2225
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContenedor.frx":29DE
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContenedor.frx":311B
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContenedor.frx":37A5
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContenedor.frx":3E9C
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContenedor.frx":4829
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContenedor.frx":5247
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContenedor.frx":5A03
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContenedor.frx":644E
            Key             =   "Open"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport Crt 
      Left            =   0
      Top             =   0
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin XtremeSuiteControls.CommonDialog CD 
      Left            =   6480
      Top             =   120
      _Version        =   1245185
      _ExtentX        =   339
      _ExtentY        =   339
      _StockProps     =   4
   End
   Begin XtremeSuiteControls.CommonDialog dlg 
      Left            =   6240
      Top             =   120
      _Version        =   1245185
      _ExtentX        =   339
      _ExtentY        =   339
      _StockProps     =   4
   End
   Begin XtremeSuiteControls.TrayIcon TrayIcon 
      Left            =   3720
      Top             =   120
      _Version        =   1245185
      _ExtentX        =   339
      _ExtentY        =   339
      _StockProps     =   16
      Text            =   "TrayIcon1"
      Picture         =   "frmContenedor.frx":6BF1
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   2640
      Top             =   120
      _Version        =   1245185
      _ExtentX        =   508
      _ExtentY        =   508
      _StockProps     =   0
   End
   Begin VB.Image imgBanner_Reportes 
      Height          =   960
      Left            =   0
      Picture         =   "frmContenedor.frx":7043
      Top             =   6240
      Width           =   24480
   End
   Begin VB.Image imgBanner_Tramites 
      Height          =   960
      Left            =   0
      Picture         =   "frmContenedor.frx":804D
      Top             =   4080
      Width           =   24480
   End
   Begin VB.Image imgBanner_Consultas 
      Height          =   960
      Left            =   0
      Picture         =   "frmContenedor.frx":90A5
      Top             =   5160
      Width           =   24480
   End
   Begin VB.Image imgBanner_Procesar 
      Height          =   960
      Left            =   0
      Picture         =   "frmContenedor.frx":9FEA
      Top             =   3000
      Width           =   24480
   End
   Begin VB.Image imgBanner_Mantenimiento 
      Height          =   960
      Left            =   0
      Picture         =   "frmContenedor.frx":B1D4
      Top             =   1920
      Width           =   24480
   End
   Begin VB.Image imgBanner_01 
      Height          =   960
      Left            =   0
      Picture         =   "frmContenedor.frx":C2AE
      Top             =   840
      Width           =   23040
   End
End
Attribute VB_Name = "frmContenedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


