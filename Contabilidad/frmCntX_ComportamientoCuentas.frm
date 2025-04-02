VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmCntX_ComportamientoCuentas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Histograma de Cuentas Contables"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6588
   HelpContextID   =   16
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   6588
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lsw 
      Height          =   1575
      Left            =   0
      TabIndex        =   5
      Top             =   1020
      Width           =   2895
      _ExtentX        =   5101
      _ExtentY        =   2773
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSChart20Lib.MSChart MsChart 
      Height          =   1452
      Left            =   120
      OleObjectBlob   =   "frmCntX_ComportamientoCuentas.frx":0000
      TabIndex        =   4
      Top             =   3360
      Width           =   1452
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyyy"
      Format          =   191365123
      CurrentDate     =   36878
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyyy"
      Format          =   191365123
      CurrentDate     =   36878
   End
   Begin XtremeSuiteControls.PushButton cmdReporte 
      Height          =   312
      Left            =   4920
      TabIndex        =   6
      Top             =   120
      Width           =   1452
      _Version        =   1245187
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Reporte"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
   End
   Begin XtremeShortcutBar.ShortcutCaption lbl 
      Height          =   372
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   3972
      _Version        =   1245187
      _ExtentX        =   7006
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Catálogo de Cuentas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   4560
      X2              =   4560
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmCntX_ComportamientoCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Set Me.MouseIcon = frmContenedor.MouseIcon

End Sub

Private Sub Form_Resize()
On Error Resume Next


lbl.Width = Me.Width - 150
lsw.Width = lbl.Width

lsw.Height = Me.Height / 2

MsChart.Move lsw.Left, lsw.Height + lsw.Top, lsw.Width / 2, lsw.Height / 2
End Sub
