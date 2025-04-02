VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmCR_PolizasControl 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Pólizas"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10950
   Icon            =   "frmCR_PolizasControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2412
      Left            =   120
      TabIndex        =   28
      Top             =   1320
      Width           =   10692
      _Version        =   1441793
      _ExtentX        =   18860
      _ExtentY        =   4254
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      GridLines       =   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbModificaciones 
      Height          =   2532
      Left            =   2880
      TabIndex        =   8
      Top             =   4320
      Width           =   2412
      _Version        =   1441793
      _ExtentX        =   4254
      _ExtentY        =   4466
      _StockProps     =   79
      Caption         =   "Modificaciones"
      ForeColor       =   4210752
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnEliminar 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Eliminar Cierre"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_PolizasControl.frx":6852
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnActualizar 
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Actualizar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_PolizasControl.frx":6DF6
         ImageAlignment  =   4
      End
   End
   Begin XtremeSuiteControls.GroupBox gbCierres 
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   2535
      _Version        =   1441793
      _ExtentX        =   4466
      _ExtentY        =   4254
      _StockProps     =   79
      Caption         =   "Cierres"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   330
         Left            =   840
         TabIndex        =   16
         Top             =   480
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCierreCorte 
         Height          =   330
         Left            =   840
         TabIndex        =   15
         Top             =   960
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.PushButton btnNuevo 
         Height          =   492
         Left            =   840
         TabIndex        =   17
         Top             =   1440
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Nuevo"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_PolizasControl.frx":74F6
         ImageAlignment  =   4
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.CheckBox chkCierrePreliminar 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Preliminar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkCierreDefinitivo 
      BackColor       =   &H0000C000&
      Caption         =   "&Definitivo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   3
      Top             =   7125
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9960
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PolizasControl.frx":7C0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PolizasControl.frx":1E5D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PolizasControl.frx":34F93
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PolizasControl.frx":4A105
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PolizasControl.frx":5F277
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PolizasControl.frx":5F3BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PolizasControl.frx":65C21
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PolizasControl.frx":6C483
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.GroupBox gbInformes 
      Height          =   2532
      Left            =   5520
      TabIndex        =   9
      Top             =   4320
      Width           =   2532
      _Version        =   1441793
      _ExtentX        =   4466
      _ExtentY        =   4466
      _StockProps     =   79
      Caption         =   "Informes de Pólizas"
      ForeColor       =   4210752
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.RadioButton rbReportes 
         Height          =   252
         Index           =   0
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Inclusiones"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnInformes 
         Height          =   492
         Left            =   360
         TabIndex        =   18
         Top             =   1920
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Reporte"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_PolizasControl.frx":72CE5
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.RadioButton rbReportes 
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   21
         Top             =   720
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Exclusiones"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton rbReportes 
         Height          =   252
         Index           =   2
         Left            =   360
         TabIndex        =   22
         Top             =   1080
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cambios"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton rbReportes 
         Height          =   252
         Index           =   3
         Left            =   360
         TabIndex        =   23
         Top             =   1440
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "General"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   2532
      Left            =   8280
      TabIndex        =   10
      Top             =   4320
      Width           =   2532
      _Version        =   1441793
      _ExtentX        =   4466
      _ExtentY        =   4466
      _StockProps     =   79
      Caption         =   "Informes Contables"
      ForeColor       =   4210752
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.RadioButton rbContables 
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Informe Resumen"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnInformeContable 
         Height          =   492
         Left            =   480
         TabIndex        =   19
         Top             =   1920
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Reporte"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_PolizasControl.frx":733EC
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.RadioButton rbContables 
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Informe Detallado"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton rbContables 
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Pólizas con déficit (Balanza)"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton rbContables 
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Pólizas con Saldos a Favor"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtPoliza 
      Height          =   495
      Left            =   2160
      TabIndex        =   29
      Top             =   120
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeShortcutBar.ShortcutCaption lblCierre 
      Height          =   330
      Left            =   1440
      TabIndex        =   31
      Top             =   3840
      Width           =   9375
      _Version        =   1441793
      _ExtentX        =   16536
      _ExtentY        =   582
      _StockProps     =   14
      Caption         =   "Info. Cierre Actual"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   330
      Left            =   120
      TabIndex        =   30
      Top             =   3840
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   582
      _StockProps     =   14
      Caption         =   "Cierre Actual: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   ">>> Seleccione el Cierre que desea visualizar <<<"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   5655
   End
   Begin VB.Label lblPoliza 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   6615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Póliza"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Image imgBanner 
      Height          =   732
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12372
   End
End
Attribute VB_Name = "frmCR_PolizasControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean, vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


Private Sub btnActualizar_Click()

On Error GoTo vError

If lblCierre.Tag = "" Then
 MsgBox "Seleccione un Corte", vbExclamation
 Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "exec spCrdPolizasActualizacion"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
MsgBox "Polizas Actualizadas Satisfactoriamente...", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnEliminar_Click()

If lblCierre.Tag = "" Then
 MsgBox "Seleccione un Corte", vbExclamation
 Exit Sub
End If

'Solo Puede Eliminar Cierres Preliminares
End Sub

Private Sub btnInformes_Click()
Dim vPoliza As String, vPolizaGeneral As Integer
Dim vTipo As String, vCorte As Integer


If lblCierre.Tag = "" Then
 MsgBox "Seleccione un Corte", vbExclamation
 Exit Sub
End If

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized

    .Connect = glogon.ConectRPT

    .WindowTitle = "Reportes - Control Pólizas"

    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "Fecha='Fecha:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "usuario='Usuario: " & glogon.Usuario & "'"

    vPoliza = txtPoliza.Text
    vPolizaGeneral = IIf(fxPolizaGeneral(vPoliza), 1, 0)
    vCorte = Mid(lblCierre.Tag, 2, 100)
    vTipo = Mid(lblCierre.Tag, 1, 1)

    If vPolizaGeneral = 0 Then
          .ReportFileName = SIFGlobal.fxPathReportes("CrdControlPolizas.rpt")
          .SelectionFormula = "{CRD_POLIZAS_CONTROL.TIPO} = '" & vTipo & "' and {CRD_POLIZAS_CONTROL.COD_POLIZA} = '" _
                            & vPoliza & "' and {CRD_POLIZAS_CONTROL.COD_CORTE} = " & vCorte
          Select Case True
            Case rbReportes(0).Value 'Inclusiones
              .Formulas(3) = "subtitulo='Poliza: " & vPoliza & "   Cierre " & IIf(vTipo = "P", "Preliminar", "Definitivo") & "  Corte : " & lblCierre.Caption & " Inclusiones'"
              .SelectionFormula = .SelectionFormula & " AND {CRD_POLIZAS_CONTROL.MONTO_ANTERIOR}=0 AND {CRD_POLIZAS_CONTROL.MONTO_ACTUAL}>0"
            Case rbReportes(1).Value 'Exclusiones
              .Formulas(3) = "subtitulo='Poliza: " & vPoliza & "   Cierre " & IIf(vTipo = "P", "Preliminar", "Definitivo") & "  Corte : " & lblCierre.Caption & " Exclusiones'"
              .SelectionFormula = .SelectionFormula & " AND {CRD_POLIZAS_CONTROL.MONTO_ANTERIOR}>0 AND {CRD_POLIZAS_CONTROL.MONTO_ACTUAL}=0"
            Case rbReportes(2).Value 'Modificaciones
              .Formulas(3) = "subtitulo='Poliza: " & vPoliza & "   Cierre " & IIf(vTipo = "P", "Preliminar", "Definitivo") & "  Corte : " & lblCierre.Caption & " Modificaciones'"
              .SelectionFormula = .SelectionFormula & " AND {CRD_POLIZAS_CONTROL.MONTO_ANTERIOR} <> {CRD_POLIZAS_CONTROL.MONTO_ACTUAL} AND " _
                                & " {CRD_POLIZAS_CONTROL.MONTO_ANTERIOR} > 0 AND {CRD_POLIZAS_CONTROL.MONTO_ACTUAL} > 0"
            Case rbReportes(3).Value 'General
              .Formulas(3) = "subtitulo='Poliza: " & vPoliza & "   Cierre " & IIf(vTipo = "P", "Preliminar", "Definitivo") & "  Corte : " & lblCierre.Caption & " General'"
              .SelectionFormula = .SelectionFormula & " AND {CRD_POLIZAS_CONTROL.MONTO_ACTUAL} > 0"
          End Select
    Else
      'Poliza General

          .ReportFileName = SIFGlobal.fxPathReportes("CrdControlPolizasGen.rpt")
          .SelectionFormula = "{CRD_POLIZAS_GENERAL.TIPO} = '" & vTipo & "' and {CRD_POLIZAS_GENERAL.COD_POLIZA} = '" _
                            & vPoliza & "' and {CRD_POLIZAS_GENERAL.COD_CORTE} = vCorte"

          Select Case True
            Case rbReportes(0).Value 'Inclusiones
              .Formulas(3) = "subtitulo='Poliza: " & vPoliza & "   Cierre " & IIf(vTipo = "P", "Preliminar", "Definitivo") & "  Corte : " & lblCierre.Caption & " Inclusiones'"
              .SelectionFormula = .SelectionFormula & " AND {CRD_POLIZAS_GENERAL.MONTO_ANTERIOR}=0 AND {CRD_POLIZAS_GENERAL.MONTO_ACTUAL}>0"
            Case rbReportes(1).Value 'Exclusiones
              .Formulas(3) = "subtitulo='Poliza: " & vPoliza & "   Cierre " & IIf(vTipo = "P", "Preliminar", "Definitivo") & "  Corte : " & lblCierre.Caption & " Exclusiones'"
              .SelectionFormula = .SelectionFormula & " AND {CRD_POLIZAS_GENERAL.MONTO_ANTERIOR}>0 AND {CRD_POLIZAS_GENERAL.MONTO_ACTUAL}=0"
            Case rbReportes(2).Value 'Modificaciones
              .Formulas(3) = "subtitulo='Poliza: " & vPoliza & "   Cierre " & IIf(vTipo = "P", "Preliminar", "Definitivo") & "  Corte : " & lblCierre.Caption & " Modificaciones'"
              .SelectionFormula = .SelectionFormula & " AND {CRD_POLIZAS_GENERAL.MONTO_ANTERIOR} <> {CRD_POLIZAS_GENERAL.MONTO_ACTUAL} AND " _
                                & " {CRD_POLIZAS_GENERAL.MONTO_ANTERIOR} > 0 AND {CRD_POLIZAS_GENERAL.MONTO_ACTUAL} > 0"
            Case rbReportes(3).Value 'General
              .Formulas(3) = "subtitulo='Poliza: " & vPoliza & "   Cierre " & IIf(vTipo = "P", "Preliminar", "Definitivo") & "  Corte : " & lblCierre.Caption & " General'"
              .SelectionFormula = .SelectionFormula & " AND {CRD_POLIZAS_GENERAL.MONTO_ACTUAL} > 0"
          End Select

    End If
    .PrintReport
End With

Me.MousePointer = vbDefault


End Sub

Private Sub btnNuevo_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass
                            
strSQL = "exec spCrdPolizasCierre '" & Mid(cboTipo.Text, 1, 1) & "','" & Format(dtpCierreCorte.Value, "yyyy/mm/dd") _
       & " 23:59:59','" & Format(dtpCierreCorte.Value, "yyyy/mm/dd") _
       & " 23:59:59','" & glogon.Usuario & "','" & txtPoliza.Text & "',0"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
MsgBox "Cierre : " & cboTipo.Text & " Realizado Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub chkCierreDefinitivo_Click()
chkCierrePreliminar.Value = IIf((chkCierreDefinitivo.Value = vbChecked), vbUnchecked, vbChecked)
Call sbCierreLsw
End Sub

Private Sub chkCierrePreliminar_Click()
chkCierreDefinitivo.Value = IIf((chkCierrePreliminar.Value = vbChecked), vbUnchecked, vbChecked)
Call sbCierreLsw
End Sub



Private Function fxPolizaGeneral(pPoliza As String) As Boolean
Dim vResultado As Boolean

strSQL = "select Poliza_General from CRD_CATALOGO_POLIZAS" _
       & " where COD_POLIZA = '" & pPoliza & "'"
Call OpenRecordSet(rs, strSQL)
 If rs!Poliza_general = 0 Then
    vResultado = False
 Else
    vResultado = True
 End If
rs.Close

fxPolizaGeneral = vResultado

End Function



Private Sub FlatScrollBar_Change()

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_poliza,descripcion from crd_catalogo_polizas"

    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_poliza > '" & txtPoliza.Text & "' order by cod_poliza asc"
    Else
       strSQL = strSQL & " where cod_poliza < '" & txtPoliza.Text & "' order by cod_poliza desc"
    End If

    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtPoliza.Text = rs!cod_poliza
      lblPoliza.Caption = rs!Descripcion
      Call sbCierreLsw
    End If
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  MsgBox "Consulte a Su Administrador de Base de Datos, sobre Transacciones con TOP y Record Count", vbInformation

End Sub


Private Sub Form_Activate()
vModulo = 11
End Sub

Private Sub Form_Load()

vModulo = 11

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True

cboTipo.AddItem "Preliminar"
cboTipo.AddItem "Definitivo"

cboTipo.Text = "Preliminar"

lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Id Cierre", 1200, vbCenter
lsw.ColumnHeaders.Add , , "Corte", 1600, vbCenter
lsw.ColumnHeaders.Add , , "Tipo", 1600, vbCenter
lsw.ColumnHeaders.Add , , "Usuario", 2400, vbLeftJustify
lsw.ColumnHeaders.Add , , "Fecha", 2400, vbLeftJustify

lsw.ListItems.Clear

dtpCierreCorte.Value = fxFechaServidor

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbLimpia()

End Sub



Private Sub sbCierreLsw()

Me.MousePointer = vbHourglass

lblCierre.Caption = ">>> Seleccione un Cierre <<<"
lblCierre.Tag = ""

strSQL = "select cod_corte,Registro_Usuario,Fecha_corte,Registro_Fecha,Tipo" _
       & " from crd_polizas_cortes" _
       & " where cod_poliza = '" & txtPoliza.Text & "' and  Tipo in("

If chkCierreDefinitivo.Value = vbChecked Then
   strSQL = strSQL & "'D',"
End If

If chkCierrePreliminar.Value = vbChecked Then
   strSQL = strSQL & "'P',"
End If

strSQL = strSQL & "'') group by  cod_corte,Registro_Usuario,Fecha_corte,Registro_Fecha,Tipo"
Call OpenRecordSet(rs, strSQL)

vPaso = True
lsw.ListItems.Clear

Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!cod_corte)
      itmX.SubItems(1) = Format(rs!fecha_corte, "dd/mm/yyyy")
      itmX.SubItems(2) = IIf((rs!Tipo = "P"), "Preliminar", "Definitivo")
      itmX.SubItems(3) = rs!Registro_Usuario
      itmX.SubItems(4) = rs!Registro_Fecha

  rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault

End Sub




Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

lblCierre.Caption = "Id: " & Item.Text & Space(15) & "Corte: " & Item.SubItems(1) & Space(15) & "Tipo: " & Item.SubItems(2)
lblCierre.Tag = Mid(Item.SubItems(2), 1, 1) & Item.Text
End Sub
