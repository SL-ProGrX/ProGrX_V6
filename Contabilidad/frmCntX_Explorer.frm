VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCntX_Explorer 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Explorar: Contabilidad"
   ClientHeight    =   4545
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16875
   HelpContextID   =   2008
   Icon            =   "frmCntX_Explorer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   16875
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lswExplorer 
      Height          =   2175
      Left            =   2760
      TabIndex        =   31
      Top             =   840
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   3831
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
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox fraConsultaAvanzada 
      Height          =   1815
      Left            =   2685
      TabIndex        =   4
      Top             =   810
      Width           =   13935
      _Version        =   1441793
      _ExtentX        =   24580
      _ExtentY        =   3201
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkFechasTodas 
         Height          =   252
         Left            =   11760
         TabIndex        =   30
         Top             =   120
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   612
         Left            =   11760
         TabIndex        =   5
         Top             =   600
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCntX_Explorer.frx":6852
      End
      Begin XtremeSuiteControls.FlatEdit txtCentroCosto 
         Height          =   312
         Left            =   720
         TabIndex        =   17
         Top             =   960
         Width           =   852
         _Version        =   1441793
         _ExtentX        =   1503
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
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
      Begin XtremeSuiteControls.FlatEdit txtCAsiento 
         Height          =   312
         Left            =   720
         TabIndex        =   18
         Top             =   120
         Width           =   852
         _Version        =   1441793
         _ExtentX        =   1503
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
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
      Begin XtremeSuiteControls.FlatEdit txtUnidad 
         Height          =   312
         Left            =   720
         TabIndex        =   19
         Top             =   600
         Width           =   852
         _Version        =   1441793
         _ExtentX        =   1503
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
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
      Begin XtremeSuiteControls.FlatEdit txtDAsiento 
         Height          =   312
         Left            =   1560
         TabIndex        =   21
         Top             =   120
         Width           =   2772
         _Version        =   1441793
         _ExtentX        =   4890
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUnidadDesc 
         Height          =   312
         Left            =   1560
         TabIndex        =   22
         Top             =   600
         Width           =   2772
         _Version        =   1441793
         _ExtentX        =   4890
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCentroCostoDesc 
         Height          =   312
         Left            =   1560
         TabIndex        =   23
         Top             =   960
         Width           =   2772
         _Version        =   1441793
         _ExtentX        =   4890
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNAsiento 
         Height          =   312
         Left            =   5640
         TabIndex        =   20
         Top             =   120
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4254
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtReferencia 
         Height          =   312
         Left            =   5640
         TabIndex        =   24
         Top             =   1320
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4254
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   312
         Left            =   5640
         TabIndex        =   26
         Top             =   960
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4254
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   312
         Left            =   5640
         TabIndex        =   28
         Top             =   600
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4254
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtLineas 
         Height          =   312
         Left            =   12360
         TabIndex        =   29
         Top             =   1320
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "1000"
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaInicio 
         Height          =   315
         Left            =   9240
         TabIndex        =   32
         Top             =   120
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   550
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaCorte 
         Height          =   312
         Left            =   10440
         TabIndex        =   33
         Top             =   120
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   550
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   612
         Left            =   13080
         TabIndex        =   36
         ToolTipText     =   "Exportar a Excel"
         Top             =   600
         Width           =   612
         _Version        =   1441793
         _ExtentX        =   1080
         _ExtentY        =   1080
         _StockProps     =   79
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCntX_Explorer.frx":7270
      End
      Begin XtremeSuiteControls.FlatEdit txtDivisa 
         Height          =   312
         Left            =   9240
         TabIndex        =   27
         Top             =   600
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4254
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
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
      Begin XtremeSuiteControls.FlatEdit txtCuenta 
         Height          =   312
         Left            =   9240
         TabIndex        =   25
         Top             =   960
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4254
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
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
      Begin XtremeSuiteControls.FlatEdit txtCuentaCorte 
         Height          =   312
         Left            =   9240
         TabIndex        =   37
         Top             =   1320
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4254
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
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
      Begin XtremeSuiteControls.ProgressBar ProgressBarX 
         Height          =   132
         Left            =   11760
         TabIndex        =   39
         Top             =   480
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   233
         _StockProps     =   93
         BackColor       =   -2147483633
         Scrolling       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtMovimiento 
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   56
         Top             =   1320
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMovimiento 
         Height          =   315
         Index           =   1
         Left            =   2880
         TabIndex        =   57
         Top             =   1320
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "999999999999.99"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboMov 
         Height          =   315
         Left            =   720
         TabIndex        =   58
         Top             =   1320
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Mov"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   55
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Corte"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   6
         Left            =   8220
         TabIndex        =   38
         Top             =   1320
         Width           =   1188
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Líneas:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   11760
         TabIndex        =   16
         Top             =   1356
         Width           =   648
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Asiento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   0
         Left            =   4536
         TabIndex        =   15
         Top             =   120
         Width           =   948
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   708
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   8220
         TabIndex        =   13
         Top             =   120
         Width           =   588
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Documento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   1
         Left            =   4536
         TabIndex        =   12
         Top             =   600
         Width           =   1188
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   2
         Left            =   4536
         TabIndex        =   11
         Top             =   960
         Width           =   828
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   708
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "C.C."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   708
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Divisa"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   3
         Left            =   8256
         TabIndex        =   8
         Top             =   600
         Width           =   1188
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Inicio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   4
         Left            =   8220
         TabIndex        =   7
         Top             =   960
         Width           =   1188
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   5
         Left            =   4560
         TabIndex        =   6
         Top             =   1320
         Width           =   1188
      End
   End
   Begin XtremeSuiteControls.GroupBox gbBarra 
      Height          =   480
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   16455
      _Version        =   1441793
      _ExtentX        =   29025
      _ExtentY        =   847
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnSemaforo 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   41
         ToolTipText     =   "Refrescar"
         Top             =   0
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Checked         =   -1  'True
         Picture         =   "frmCntX_Explorer.frx":7A75
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnSemaforo 
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   42
         ToolTipText     =   "Refrescar"
         Top             =   0
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCntX_Explorer.frx":8099
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnSemaforo 
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   43
         ToolTipText     =   "Refrescar"
         Top             =   0
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCntX_Explorer.frx":86B5
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnSemaforo 
         Height          =   375
         Index           =   3
         Left            =   1080
         TabIndex        =   44
         ToolTipText     =   "Refrescar"
         Top             =   0
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCntX_Explorer.frx":8CD3
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   45
         Top             =   0
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Editar"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCntX_Explorer.frx":93BA
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   46
         Top             =   0
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Borrar"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCntX_Explorer.frx":99B5
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   47
         Top             =   0
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Refrescar"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCntX_Explorer.frx":9F59
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   3
         Left            =   5640
         TabIndex        =   48
         Top             =   0
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCntX_Explorer.frx":A659
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   4
         Left            =   6840
         TabIndex        =   49
         Top             =   0
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Mayorizar"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCntX_Explorer.frx":AD60
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   5
         Left            =   8400
         TabIndex        =   50
         Top             =   0
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Detalle"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Checked         =   -1  'True
         Picture         =   "frmCntX_Explorer.frx":B479
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   6
         Left            =   9720
         TabIndex        =   51
         Top             =   0
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Movimiento"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCntX_Explorer.frx":BB81
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnVisualizar 
         Height          =   375
         Index           =   0
         Left            =   11760
         TabIndex        =   52
         Top             =   0
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Notas"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCntX_Explorer.frx":C289
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnVisualizar 
         Height          =   375
         Index           =   1
         Left            =   12960
         TabIndex        =   53
         Top             =   0
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Avanzada"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Checked         =   -1  'True
         Picture         =   "frmCntX_Explorer.frx":C9A2
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnConcilia 
         Height          =   375
         Left            =   14760
         TabIndex        =   54
         Top             =   0
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Conciliador"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCntX_Explorer.frx":D0A2
         ImageAlignment  =   0
      End
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   6480
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":D7C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":D8D7
            Key             =   "imgFolder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":D9F3
            Key             =   "imgOpcion"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":DB12
            Key             =   "imgUsuario"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":DC19
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":DD4E
            Key             =   "imgAsientos"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":DE52
            Key             =   "imgPlantillas"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":DF59
            Key             =   "imgCuentas"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":E067
            Key             =   "imgAreas"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":E15B
            Key             =   "imgDiferidos"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":E279
            Key             =   "imgPresupuesto"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":E392
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":E4B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":E5C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":E6E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":E806
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraNotas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   2760
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox txtNotas 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   2160
      Left            =   5400
      ScaleHeight     =   940.557
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   1
      Top             =   825
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   6960
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":E930
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":15192
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":1B9F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":22256
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":28AB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":2F31A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":35B7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":3C3DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView ArbolExp 
      Height          =   2160
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   3810
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgLista"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":42C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":494A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":4FD04
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":56566
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":5CDC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":6362A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":69E8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":706EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":76F50
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":7796E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":780E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":78852
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Explorer.frx":78FC4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption lblTitle 
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   34
      Top             =   480
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4678
      _ExtentY        =   582
      _StockProps     =   14
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
   Begin XtremeShortcutBar.ShortcutCaption lblTitle 
      Height          =   330
      Index           =   1
      Left            =   2760
      TabIndex        =   35
      Top             =   480
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4678
      _ExtentY        =   582
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   6
   End
   Begin VB.Image imgSplitter 
      Height          =   2145
      Left            =   2565
      MousePointer    =   9  'Size W E
      Top             =   945
      Width           =   150
   End
End
Attribute VB_Name = "frmCntX_Explorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vPaso As Boolean
Dim vNode As Node
Dim mbMoving As Boolean
Dim x As Boolean
Const sglSplitLimit = 500

Private Sub ArbolExp_DblClick()
'Call tlbPrincipal_ButtonClick(tlbPrincipal.Buttons.Item(1))
End Sub

Private Function fxDescribeAsientos(vTipo As String, vDescripcion) As String
Dim rsX As New ADODB.Recordset, strSQL As String

If vTipo = "C" Then 'Codigo
  strSQL = "select Tipo_asiento as Resultado from CntX_Tipos_Asientos where cod_contabilidad = " _
         & gCntX_Parametros.CodigoConta & " and descripcion = '" & vDescripcion & "'"
Else 'Descripcion
  strSQL = "select Descripcion as Resultado from CntX_Tipos_Asientos where cod_contabilidad = " _
         & gCntX_Parametros.CodigoConta & " and cod_tipo = '" & vDescripcion & "'"
End If
Call OpenRecordSet(rsX, strSQL, 0)
If rsX.EOF And rsX.BOF Then
  fxDescribeAsientos = ""
Else
  fxDescribeAsientos = rsX!Resultado
End If
rsX.Close

End Function


Private Function fxIndiceCodigo(xkey As String) As String
xkey = Mid(xkey, 4, Len(xkey))
xkey = Mid(xkey, 1, Len(xkey) - 1)
fxIndiceCodigo = xkey
End Function

Private Function fxIndiceAsiento(xkey As String, vTipo As String) As String
Dim i As Long, strResultado As String, blnPaso As Boolean

xkey = fxIndiceCodigo(xkey)

blnPaso = True

If vTipo = "T" Then ' Tipo
  i = 1
  strResultado = ""
  Do While blnPaso
    If Mid(xkey, i, 1) <> "-" Then
     strResultado = strResultado & Mid(xkey, i, 1)
    Else
     blnPaso = False
    End If
    i = i + 1
  Loop
  
Else 'Numero

  i = 1
  strResultado = ""
  Do While blnPaso
    If Mid(xkey, i, 1) = "-" Then blnPaso = False
    i = i + 1
  Loop
  strResultado = Mid(xkey, i, 50) '50 es un default ningun asiento es tan largo

End If

fxIndiceAsiento = strResultado

End Function


Private Sub ArbolExp_Expand(ByVal Node As MSComctlLib.Node)
Dim rs As New ADODB.Recordset, strSQL As String
Dim frmX As Form


On Error GoTo vNodoError

Set vNode = Node

If Node.Tag = 1 Then Exit Sub

If Node.Index > 1 Then ArbolExp.Nodes.Remove Node.Child.Index

Node.Tag = 1

Select Case Node.Text
  
  Case "Catálogo de Cuentas"
      strSQL = "select tipo_cuenta,Descripcion from CntX_Tipos_Cuentas" _
             & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
             & " order by Prioridad"
      Call OpenRecordSet(rs, strSQL, 0)
      Do While Not rs.EOF
        Call sbCreaNodos(Node.Key, rs!Descripcion, "imgCuentas", True, "0x0" & rs!tipo_cuenta & "T")
        rs.MoveNext
      Loop
      rs.Close
 
  Case "Asientos"
             
        strSQL = "exec spCntx_Consulta_Asientos_Rsm " & gCntX_Parametros.CodigoConta _
               & "," & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes
      Call OpenRecordSet(rs, strSQL, 0)
      
      If rs.BOF And rs.EOF Then
            strSQL = "select tipo_asiento,descripcion from CntX_Tipos_Asientos" _
                   & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta & " and Activo = 1" _
                   & " order by Descripcion"
            Call OpenRecordSet(rs, strSQL, 0)
      End If
      
      Do While Not rs.EOF
        Call sbCreaNodos(Node.Key, rs!Descripcion, "imgAsientos", True, "0x0" & rs!Tipo_Asiento & "X")
        rs.MoveNext
      Loop
      rs.Close
 
  Case "Mantenimiento"
      Call sbCreaNodos(Node.Key, "Periodos", "imgFolder", False, "0x001M")
      Call sbCreaNodos(Node.Key, "Cierres", "imgFolder", False, "0x001M")
      Call sbCreaNodos(Node.Key, "Contabilidades", "imgFolder", False, "0x001M")
      Call sbCreaNodos(Node.Key, "Tipos de Cuentas", "imgFolder", False, "0x001M")
      Call sbCreaNodos(Node.Key, "Tipos de Asientos", "imgFolder", False, "0x001M")
      Call sbCreaNodos(Node.Key, "Divisas", "imgFolder", False, "0x001M")

'      Call sbCreaNodos(Node.Key, "Empresa", "imgFolder", False, "0x001M")
'     Call sbCreaNodos(Node.Key, "Usuarios", "imgFolder", False)
'     Call sbCreaNodos(Node.Key, "Parámetros", "imgFolder", False)
  
  Case "ProGrX: Contabilidad", "ProGrX: Contabilidad"
      'Nada
  Case "Areas de Trabajo"
      strSQL = "select cod_area,Descripcion from CntX_Area_Definicion where cod_contabilidad = " & gCntX_Parametros.CodigoConta
      Call OpenRecordSet(rs, strSQL, 0)
      Do While Not rs.EOF
        Call sbCreaNodos(Node.Key, rs!Descripcion, "imgFolder", False, "0x0" & rs!cod_area & "Y")
        rs.MoveNext
      Loop
      rs.Close
  
  Case "Administracion  de Diferidos"
      strSQL = "select cod_diferido,Descripcion from CntX_Diferidos where cod_contabilidad = " & gCntX_Parametros.CodigoConta
      Call OpenRecordSet(rs, strSQL, 0)
      Do While Not rs.EOF
        Call sbCreaNodos(Node.Key, rs!Descripcion, "imgCntX_Cuentas", True, "0x0" & rs!cod_diferido & "D")
        rs.MoveNext
      Loop
      rs.Close
  
  Case "Plantillas Asientos Porcentuales"
      strSQL = "select cod_plantilla,Descripcion from CntX_Plantilla_Rate where cod_contabilidad = " & gCntX_Parametros.CodigoConta
      Call OpenRecordSet(rs, strSQL, 0)
      Do While Not rs.EOF
        Call sbCreaNodos(Node.Key, rs!Descripcion, "imgFolder", False, "0x0" & rs!cod_plantilla & "R")
        rs.MoveNext
      Loop
      rs.Close
      
  Case "Plantillas Asientos Fijos"
      strSQL = "select cod_plantilla,Descripcion from CntX_Plantilla_Asientos where cod_contabilidad = " & gCntX_Parametros.CodigoConta
      Call OpenRecordSet(rs, strSQL, 0)
      Do While Not rs.EOF
        Call sbCreaNodos(Node.Key, rs!Descripcion, "imgFolder", False, "0x0" & rs!cod_plantilla & "F")
        rs.MoveNext
      Loop
      rs.Close
  
  Case "Presupuesto"
  
  
  Case Else 'SubCntX_Cuentas y Cntx_Asientos
      
    Select Case Right(Node.Key, 1)
        Case "X" 'Tipos de Asientos
            Call sbFormsCall("frmCntX_Procesos")
            Call sbFormActivo("frmCntX_Procesos", frmX)
            
            frmX.Caption = "Explorer"
            frmX.lbl.Caption = "Cargando..."
            frmX.lbl.Refresh
            strSQL = "select tipo_asiento,num_asiento,descripcion from Cntx_Asientos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                   & " and tipo_asiento = '" & fxIndiceCodigo(Node.Key) & "'" _
                   & " and anio = " & gCntX_Parametros.PeriodoAnio & " and mes = " & gCntX_Parametros.PeriodoMes _
                   & " order by num_asiento"
            Call OpenRecordSet(rs, strSQL, 0)
                        
            frmX.PrgBar.Max = rs.RecordCount + 1
            frmX.PrgBar.Value = 1
            
            Do While Not rs.EOF
              Call sbCreaNodos(Node.Key, Trim(rs!Num_Asiento) & "   [" & Trim(UCase(rs!Descripcion)) & "]", "imgFolder", False, "0x0" & rs!Tipo_Asiento & "-" & rs!Num_Asiento & "A")
              rs.MoveNext
              If frmX.PrgBar.Value < frmX.PrgBar.Max Then frmX.PrgBar.Value = frmX.PrgBar.Value + 1
            Loop
            rs.Close
            
            Unload frmX
            
        Case "T" 'Tipos de Cuentas
        
            strSQL = "select cod_cuenta,descripcion,acepta_movimientos from CntX_Cuentas where cuenta_madre = ''" _
                   & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                   & " and tipo_cuenta = '" & fxIndiceCodigo(Node.Key) & "' order by cod_cuenta"
            Call OpenRecordSet(rs, strSQL, 0)
            Do While Not rs.EOF
              
              If rs!Acepta_movimientos = 0 Then
                  Call sbCreaNodos(Node.Key, fxCntX_CuentaFormato(True, rs!cod_cuenta) & " - " & rs!Descripcion, "imgCntX_Cuentas", True, "0x0" & fxCntX_CuentaFormato(False, rs!cod_cuenta) & "C")
              Else
                  Call sbCreaNodos(Node.Key, fxCntX_CuentaFormato(True, rs!cod_cuenta) & " - " & rs!Descripcion, "imgFolder", False, "0x0" & fxCntX_CuentaFormato(False, rs!cod_cuenta) & "C")
              End If
              rs.MoveNext
            Loop
            rs.Close
        
        Case "D" 'Plantilla de Diferidos
        
            strSQL = "select cod_difPlantilla,descripcion,cod_diferido from CntX_diferido_plantilla" _
                   & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                   & " and cod_diferido = " & fxIndiceCodigo(Node.Key)
            Call OpenRecordSet(rs, strSQL, 0)
            Do While Not rs.EOF
              Call sbCreaNodos(Node.Key, rs!Descripcion, "imgFolder", False, "0x0" & rs!cod_diferido & "-" & rs!cod_difPlantilla & "E")
              rs.MoveNext
            Loop
            rs.Close
        
        
        Case Else 'Sub Cuentas
        
            strSQL = "select cod_cuenta,descripcion,acepta_movimientos from CntX_Cuentas where cuenta_madre = '" & fxCntX_CuentaFormato(False, fxIndiceCodigo(Node.Key)) _
                   & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta & " order by cod_cuenta"
            Call OpenRecordSet(rs, strSQL, 0)
            Do While Not rs.EOF
             ' Call sbCreaNodos(Node.Key, fxCntX_CuentaFormato(True, rs!cod_cuenta) & " - " & rs!descripcion, "imgFolder", IIf((rs!acepta_movimientos = "N"), True, False), "0x0" & fxCntX_CuentaFormato(False, rs!cod_cuenta) & "C")
              
              If rs!Acepta_movimientos = 0 Then
                  Call sbCreaNodos(Node.Key, fxCntX_CuentaFormato(True, rs!cod_cuenta) & " - " & rs!Descripcion, "imgCntX_Cuentas", True, "0x0" & fxCntX_CuentaFormato(False, rs!cod_cuenta) & "C")
              Else
                  Call sbCreaNodos(Node.Key, fxCntX_CuentaFormato(True, rs!cod_cuenta) & " - " & rs!Descripcion, "imgFolder", False, "0x0" & fxCntX_CuentaFormato(False, rs!cod_cuenta) & "C")
              End If
              
              rs.MoveNext
            Loop
            rs.Close
    End Select

End Select

Exit Sub

vNodoError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbMuestraMemos(strMemo As String)
Dim itmX As ListItem, lng As Long

lng = 1

Do While Len(strMemo) + 42 > lng
  Set itmX = lswExplorer.ListItems.Add(lswExplorer.ListItems.Count + 1, , "")
      itmX.SubItems(1) = Mid(strMemo, lng, 40)
  lng = lng + 40
Loop

End Sub

Private Sub sbMuestraParametros()
Dim itmX As ListViewItem, strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, i2 As Integer

lswExplorer.ListItems.Clear
lswExplorer.ColumnHeaders.Clear
With lswExplorer
  .ListItems.Clear
  .ColumnHeaders.Clear
  Select Case vNode.Text
    Case "Tipos de Asientos"
        .ColumnHeaders.Add , , "Código", 850
        .ColumnHeaders.Add , , "Descripción", 3850
        .ColumnHeaders.Add , , "Consecutivo", 1850
        
        strSQL = "select tipo_asiento,descripcion,consecutivo from CntX_Tipos_Asientos where cod_contabilidad = " _
              & gCntX_Parametros.CodigoConta & " order by tipo_asiento"
    Case "Tipos de Cuentas"
        .ColumnHeaders.Add , , "Código", 850
        .ColumnHeaders.Add , , "Descripción", 3850
        strSQL = "select tipo_cuenta,descripcion from CntX_Tipos_Cuentas where cod_contabilidad = " _
               & gCntX_Parametros.CodigoConta & " order by tipo_cuenta"
    Case "Contabilidades"
        .ColumnHeaders.Add , , "Código", 850
        .ColumnHeaders.Add , , "Descripción", 3850
        .ColumnHeaders.Add , , "Tel.Central", 1250, 2
        .ColumnHeaders.Add , , "Tel.Fax", 950
        .ColumnHeaders.Add , , "Contacto", 2950
        
        strSQL = "select cod_contabilidad,nombre,tel_central,tel_fax,contacto" _
                & " from CntX_Contabilidades order by nombre"
                
    Case "Periodos"
        .ColumnHeaders.Add , , "Año", 850
        .ColumnHeaders.Add , , "Mes", 850
        .ColumnHeaders.Add , , "Fecha Corte", 1850, vbCenter
        .ColumnHeaders.Add , , "Cerrado", 1200, vbCenter
        .ColumnHeaders.Add , , "Cierre: Usuario", 2200, vbCenter
        .ColumnHeaders.Add , , "Cierre: Fecha", 2200, vbCenter
        
        strSQL = "select Anio,Mes,PERIODO_CORTE, case estado when 'C' then 'CERRADO' when 'P' then 'PENDIENTE' end As Estado,cierre_usuario,Cierre_fecha from CntX_Periodos where cod_contabilidad = " _
               & gCntX_Parametros.CodigoConta & " order by Anio,Mes"
    
    Case "Empresa"
        .ColumnHeaders.Add , , "Nombre", 4850
        .ColumnHeaders.Add , , "Teléfono", 1200
        .ColumnHeaders.Add , , "Fax", 1200
        .ColumnHeaders.Add , , "Apto.Postal", 1820
        strSQL = "select Nombre,Telefono,fax,Apto_postal from CntX_Empresa_Registro"
    
    
    Case "Divisas"
        .ColumnHeaders.Add , , "Divisas", 850
        .ColumnHeaders.Add , , "Descripcion", 4050
        .ColumnHeaders.Add , , "TC.Venta", 1200, vbRightJustify
        .ColumnHeaders.Add , , "TC.Compra", 1200, vbRightJustify
        .ColumnHeaders.Add , , "Local", 1820, vbCenter
        strSQL = "select cod_divisa,descripcion,tc_venta,tc_compra,Divisa_local from CntX_Divisas" _
               & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta
    
    Case "Cierres"
        .ColumnHeaders.Add , , "In.Año", 850
        .ColumnHeaders.Add , , "In.Mes", 850
        .ColumnHeaders.Add , , "Co.Año", 850
        .ColumnHeaders.Add , , "Co.Mes", 850
        .ColumnHeaders.Add , , "Descripción", 2500
        .ColumnHeaders.Add , , "Gan.Per", 1300
        .ColumnHeaders.Add , , "Exc.Uti", 1300
        .ColumnHeaders.Add , , "Renta.Cta", 1300
        .ColumnHeaders.Add , , "Renta.", 1100
        .ColumnHeaders.Add , , "Vigente", 1300

        
        strSQL = "select inicio_anio,inicio_mes,corte_anio,corte_mes,descripcion" _
               & ",cuenta_ganper,cuenta_utilidad,cuenta_impRenta,impuesto_renta,case activo when 1 then 'Si' else 'NO' end as Vigente from cntx_cierres" _
               & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta
   
  End Select
  
  Call OpenRecordSet(rs, strSQL, 0)
  Do While Not rs.EOF
    If rs.Fields(1).Value <> "" Then
        Set itmX = .ListItems.Add(.ListItems.Count + 1, , rs.Fields(0).Value, 5)
        For i = 1 To (rs.Fields.Count - 1)
          itmX.SubItems(i) = IIf(IsNull(rs.Fields(i).Value), "", rs.Fields(i).Value)
        Next i
    End If
   rs.MoveNext
  Loop
  rs.Close
End With
End Sub

Private Function fxCedula(vCadena As String) As String
Dim strResultado As String, i As Integer

strResultado = ""

For i = 1 To Len(vCadena)
 If Mid(vCadena, i, 1) = "-" Then
   fxCedula = strResultado
   Exit Function
 Else
   strResultado = strResultado + Mid(vCadena, i, 1)
 End If
Next i

End Function

Private Function fxNotas(vTipo As String, vNumero As String) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select notas from Cntx_Asientos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and num_asiento = '" & vNumero & "' and tipo_asiento = '" & vTipo & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  fxNotas = rs!Notas & ""
Else
  fxNotas = ""
End If
rs.Close

End Function


Private Sub sbMuestraDetalleSubNodos()
Dim itmX As ListViewItem, strSQL As String, rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset, curDebitos As Currency, curCreditos As Currency
Dim curInicial As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case Right(vNode.Key, 1)
  Case "F" 'Asiento de Plantilla Asientos Fijos
        With lswExplorer
           strSQL = "select P.*,C.descripcion,C.cod_Cuenta_Mask " _
                  & " from CntX_Plantilla_detalle P inner join CntX_Cuentas C on P.cod_cuenta = C.cod_cuenta" _
                  & " and P.cod_cuenta = C.cod_cuenta" _
                  & " where P.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                  & " and P.cod_plantilla = " & fxIndiceCodigo(vNode.Key) _
                  & " order by P.num_linea"
           Call OpenRecordSet(rs, strSQL, 0)
           .ColumnHeaders.Add , , "Cuenta", 2800
           .ColumnHeaders.Add , , "Descripción", 3200
           .ColumnHeaders.Add , , "Débitos", 1400, vbRightJustify
           .ColumnHeaders.Add , , "Créditos", 1400, vbRightJustify
           .ColumnHeaders.Add , , "Detalle", 2200
           curDebitos = 0
           curCreditos = 0
           Do While Not rs.EOF
             Set itmX = .ListItems.Add(, , rs!Cod_Cuenta_Mask)
                 itmX.SubItems(1) = rs!Descripcion
                 itmX.SubItems(2) = Format(rs!Debitos, "Standard")
                 itmX.SubItems(3) = Format(rs!Creditos, "Standard")
                 itmX.SubItems(4) = ""
                 curDebitos = curDebitos + rs!Debitos
                 curCreditos = curCreditos + rs!Creditos
             rs.MoveNext
           Loop
           rs.Close
              
           Set itmX = .ListItems.Add(, , "")
               itmX.SubItems(2) = "___________"
               itmX.SubItems(3) = "___________"
           
           Set itmX = .ListItems.Add(, , "")
               itmX.SubItems(2) = Format(curDebitos, "Standard")
               itmX.SubItems(3) = Format(curCreditos, "Standard")
               itmX.TextBackColor = RGB(214, 234, 248)
        End With
  
  
  Case "R" 'Plantilla de Asientos Porcentuales
  
        With lswExplorer
           strSQL = "select P.*,C.descripcion,C.cod_Cuenta_Mask " _
                  & " from CntX_Plantilla_Rate_Detalle P inner join CntX_Cuentas C on P.cod_cuenta = C.cod_cuenta" _
                  & " and P.cod_cuenta = C.cod_cuenta" _
                  & " where P.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                  & " and P.cod_plantilla = " & fxIndiceCodigo(vNode.Key) _
                  & " order by P.num_linea"
           Call OpenRecordSet(rs, strSQL, 0)
           .ColumnHeaders.Add , , "Cuenta", 2800
           .ColumnHeaders.Add , , "Descripción", 3200
           .ColumnHeaders.Add , , "Débitos", 1800, vbRightJustify
           .ColumnHeaders.Add , , "Créditos", 1800, vbRightJustify
           .ColumnHeaders.Add , , "Detalle", 2200
           curDebitos = 0
           curCreditos = 0
           Do While Not rs.EOF
             Set itmX = .ListItems.Add(, , rs!Cod_Cuenta_Mask, 13)
                 itmX.SubItems(1) = rs!Descripcion
                 itmX.SubItems(2) = Format(rs!Debitos, "Standard")
                 itmX.SubItems(3) = Format(rs!Creditos, "Standard")
                 itmX.SubItems(4) = rs!Detalle
                 curDebitos = curDebitos + rs!Debitos
                 curCreditos = curCreditos + rs!Creditos
             rs.MoveNext
           Loop
           rs.Close
              
           Set itmX = .ListItems.Add(, , "")
               itmX.SubItems(2) = "___________"
               itmX.SubItems(3) = "___________"
           
           Set itmX = .ListItems.Add(, , "")
               itmX.SubItems(2) = Format(curDebitos, "Standard")
               itmX.SubItems(3) = Format(curCreditos, "Standard")
               itmX.TextBackColor = RGB(214, 234, 248)
        
        End With
  
  
  Case "E" 'Historial de Diferidos
      With lswExplorer
           strSQL = "select * from CntX_Diferido_Historico" _
                  & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                  & " AND COD_DIFPlantilla = " & fxIndiceAsiento(vNode.Key, "N") _
                  & " and cod_diferido = " & fxIndiceAsiento(vNode.Key, "T")
           Call OpenRecordSet(rs, strSQL, 0)
           
           .ColumnHeaders.Add , , "Asiento", 3200
           .ColumnHeaders.Add , , "Tipo", 1200
           .ColumnHeaders.Add , , "Fecha", 1200
           .ColumnHeaders.Add , , "Año", 800
           .ColumnHeaders.Add , , "Mes", 800
           .ColumnHeaders.Add , , "Usuario", 1200
           
           Do While Not rs.EOF
             Set itmX = .ListItems.Add(, , rs!Num_Asiento, 13)
                 itmX.SubItems(1) = rs!Tipo_Asiento
                 itmX.SubItems(2) = Format(rs!fecha, "yyyy/mm/dd")
                 itmX.SubItems(3) = rs!Anio
                 itmX.SubItems(4) = rs!Mes
                 itmX.SubItems(5) = rs!Usuario
                 
                 itmX.Tag = "Asiento"
                 
             rs.MoveNext
           Loop
           rs.Close
        End With
  
  Case "D" 'Listado de Plantillas de Diferidos para Ejecucion creadas
      With lswExplorer
           strSQL = "select * from CntX_diferido_plantilla" _
                  & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                  & " and cod_diferido = " & fxIndiceCodigo(vNode.Key)
           Call OpenRecordSet(rs, strSQL, 0)
           
           .ColumnHeaders.Add , , "Plantilla", 1200
           .ColumnHeaders.Add , , "Descripción", 3200
           .ColumnHeaders.Add , , "Monto", 1200, vbRightJustify
           .ColumnHeaders.Add , , "Acumulado", 1200, vbRightJustify
           .ColumnHeaders.Add , , "Pendiente", 1200, vbRightJustify
           .ColumnHeaders.Add , , "Plazo", 800, vbRightJustify
           .ColumnHeaders.Add , , "Inicio", 1200
           .ColumnHeaders.Add , , "Usuario", 1200
           .ColumnHeaders.Add , , "Documento", 1200
           
           Do While Not rs.EOF
             Set itmX = .ListItems.Add(, , rs!cod_difPlantilla, 13)
                 itmX.SubItems(1) = rs!Descripcion
                 itmX.SubItems(2) = Format(rs!monto_diferir, "Standard")
                 itmX.SubItems(3) = Format(rs!acumulado, "Standard")
                 itmX.SubItems(4) = Format(rs!monto_diferir - rs!acumulado, "Standard")
                 itmX.SubItems(5) = rs!Plazo
                 itmX.SubItems(6) = rs!Anio & "-" & Format(rs!Mes, "00")
                 itmX.SubItems(7) = rs!user_crea
                 itmX.SubItems(8) = rs!Documento
             rs.MoveNext
           Loop
           rs.Close
        End With
  
  
  Case "Y" 'Catalogo de Cuentas de las Areas de Trabajo
   If btnAccion(6).Checked = False Then
      
          With lswExplorer
           strSQL = "select A.cod_cuenta,C.descripcion,C.acepta_movimientos, C.cod_Cuenta_Mask" _
                  & " from CntX_Cuentas C inner join CntX_Area_Cuentas A" _
                  & " On C.cod_cuenta = A.cod_cuenta and C.cod_contabilidad = A.cod_contabilidad" _
                  & " where A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                  & " and A.cod_area = " & fxIndiceCodigo(vNode.Key) _
                  & " order by A.cod_cuenta"
           Call OpenRecordSet(rs, strSQL, 0)
           .ColumnHeaders.Add , , "Cuenta", 2000
           .ColumnHeaders.Add , , "Descripción", 3200
           .ColumnHeaders.Add , , "Acep.Mov.", 1200, 2
            
            Do While Not rs.EOF
                Set itmX = .ListItems.Add(, , rs!Cod_Cuenta_Mask, 13)
                    itmX.SubItems(1) = rs!Descripcion
                    itmX.SubItems(2) = IIf((rs!Acepta_movimientos = 1), "Sí", "No")
                If rs!Acepta_movimientos = 0 Then itmX.ForeColor = vbBlue
              rs.MoveNext
            Loop
            rs.Close
        End With
  
   Else 'Movimientos
    
         curDebitos = 0
         curCreditos = 0
         curInicial = 0
      
         With lswExplorer
            strSQL = "select B.*,A.descripcion,A.acepta_movimientos,A.cod_Cuenta_Mask" _
                   & " from CntX_Cuentas A inner join vCntX_Mov_Cuentas_General B" _
                   & " On A.cod_cuenta = B.cod_cuenta and A.cod_contabilidad = B.cod_contabilidad" _
                   & " inner join CntX_Area_Cuentas X on A.cod_cuenta = X.cod_cuenta" _
                   & " and A.cod_contabilidad = X.cod_contabilidad" _
                   & " where X.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                   & " and X.cod_area = " & fxIndiceCodigo(vNode.Key) _
                   & " and anio = " & gCntX_Parametros.PeriodoAnio & " and mes = " & gCntX_Parametros.PeriodoMes _
                   & " order by A.cod_cuenta"
            Call OpenRecordSet(rs, strSQL, 0)
            .ColumnHeaders.Add , , "Cuenta", 2800
            .ColumnHeaders.Add , , "Descripción", 3200
            .ColumnHeaders.Add , , "Inicial", 1800, 1
            .ColumnHeaders.Add , , "Débitos", 1800, 1
            .ColumnHeaders.Add , , "Créditos", 1800, 1
            .ColumnHeaders.Add , , "Mes", 1800, 1
            .ColumnHeaders.Add , , "Actual", 1800, 1
    
            Do While Not rs.EOF
              Set itmX = .ListItems.Add(, , rs!Cod_Cuenta_Mask, 13)
                  itmX.SubItems(1) = rs!Descripcion
                  itmX.SubItems(2) = Format(rs!saldo_inicial, "Standard")
                  itmX.SubItems(3) = Format(rs!total_debitos, "Standard")
                  itmX.SubItems(4) = Format(rs!total_creditos, "Standard")
                  itmX.SubItems(5) = Format(rs!total_debitos + rs!total_creditos, "Standard")
                  itmX.SubItems(6) = Format(rs!saldo_inicial + (rs!total_debitos + rs!total_creditos), "Standard")
                  
                  If rs!Acepta_movimientos = 1 Then
                    curDebitos = curDebitos + rs!total_debitos
                    curCreditos = curCreditos + rs!total_creditos
                    curInicial = curInicial + rs!saldo_inicial
                  Else
                    itmX.ForeColor = vbBlue
                  End If
              rs.MoveNext
            Loop
            rs.Close
            'Totales
                   
              Set itmX = .ListItems.Add(, , "")
                  itmX.SubItems(2) = "___________"
                  itmX.SubItems(3) = "___________"
                  itmX.SubItems(4) = "___________"
                  itmX.SubItems(5) = "___________"
                  itmX.SubItems(6) = "___________"
              
              Set itmX = .ListItems.Add(, , "TOTALES")
                  itmX.SubItems(2) = Format(curInicial, "Standard")
                  itmX.SubItems(3) = Format(curDebitos, "Standard")
                  itmX.SubItems(4) = Format(curCreditos, "Standard")
                  itmX.SubItems(5) = Format(curDebitos + curCreditos, "Standard")
                  itmX.SubItems(6) = Format(curInicial + (curDebitos + curCreditos), "Standard")
                  itmX.ForeColor = vbBlue
                  itmX.TextBackColor = RGB(214, 234, 248)
         End With
    
   End If 'Movimientos o Descripcion
  
  Case "T" 'Tipos de Cuentas
    If btnAccion(6).Checked = False Then
         With lswExplorer
            
            strSQL = "select cod_cuenta,A.descripcion,B.Descripcion as Res,A.acepta_movimientos,A.cod_Cuenta_Mask" _
                   & " from CntX_Cuentas A inner join CntX_Tipos_Cuentas B" _
                   & " On A.tipo_cuenta = B.tipo_cuenta and A.cod_contabilidad = B.cod_contabilidad" _
                   & " where A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                   & " and A.tipo_cuenta = '" & fxIndiceCodigo(vNode.Key) & "'" _
                   & " order by cod_cuenta"
                   
            Call OpenRecordSet(rs, strSQL, 0)
            .ColumnHeaders.Add , , "Cuenta", 2600
            .ColumnHeaders.Add , , "Descripción", 3200
            .ColumnHeaders.Add , , "Tipo", 2200
            .ColumnHeaders.Add , , "Acep.Mov.", 1200, 2
            
            Do While Not rs.EOF
              Set itmX = .ListItems.Add(, , rs!Cod_Cuenta_Mask, 13)
                  itmX.SubItems(1) = rs!Descripcion & ""
                  itmX.SubItems(2) = rs!res & ""
                  itmX.SubItems(3) = IIf((rs!Acepta_movimientos = 1), "Sí", "No")
              If rs!Acepta_movimientos = 0 Then itmX.ForeColor = vbBlue
              rs.MoveNext
            Loop
            rs.Close
         End With
      
      Else 'Movimientos
         curDebitos = 0
         curCreditos = 0
         curInicial = 0
      
         With lswExplorer
            strSQL = "select B.*,A.descripcion,A.acepta_movimientos,A.cod_Cuenta_Mask" _
                   & " from CntX_Cuentas A inner join vCntX_Mov_Cuentas_General B" _
                   & " On A.cod_cuenta = B.cod_cuenta and A.cod_contabilidad = B.cod_contabilidad" _
                   & " where A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                   & " and A.tipo_cuenta = '" & fxIndiceCodigo(vNode.Key) & "'" _
                   & " and anio = " & gCntX_Parametros.PeriodoAnio & " and mes = " & gCntX_Parametros.PeriodoMes _
                   & " and A.cuenta_madre = '' order by A.cod_cuenta"
            Call OpenRecordSet(rs, strSQL, 0)
            .ColumnHeaders.Add , , "Cuenta", 2600
            .ColumnHeaders.Add , , "Descripción", 3200
            .ColumnHeaders.Add , , "Inicial", 1800, 1
            .ColumnHeaders.Add , , "Débitos", 1800, 1
            .ColumnHeaders.Add , , "Créditos", 1800, 1
            .ColumnHeaders.Add , , "Mes", 1800, 1
            .ColumnHeaders.Add , , "Actual", 1800, 1
    
            Do While Not rs.EOF
              Set itmX = .ListItems.Add(, , rs!Cod_Cuenta_Mask, 13)
                  itmX.SubItems(1) = rs!Descripcion
                  itmX.SubItems(2) = Format(rs!saldo_inicial, "Standard")
                  itmX.SubItems(3) = Format(rs!total_debitos, "Standard")
                  itmX.SubItems(4) = Format(rs!total_creditos, "Standard")
                  itmX.SubItems(5) = Format(rs!total_debitos + rs!total_creditos, "Standard")
                  itmX.SubItems(6) = Format(rs!saldo_inicial + (rs!total_debitos + rs!total_creditos), "Standard")
                  
              If rs!Acepta_movimientos = 0 Then
                  itmX.ForeColor = vbBlue
              Else
                  curDebitos = curDebitos + rs!total_debitos
                  curCreditos = curCreditos + rs!total_creditos
                  curInicial = curInicial + rs!saldo_inicial
              End If
              rs.MoveNext
            Loop
            rs.Close
            'Totales
                   
              Set itmX = .ListItems.Add(, , "")
                  itmX.SubItems(2) = "___________"
                  itmX.SubItems(3) = "___________"
                  itmX.SubItems(4) = "___________"
                  itmX.SubItems(5) = "___________"
                  itmX.SubItems(6) = "___________"
       
       
              Set itmX = .ListItems.Add(, , "TOTALES")
                  itmX.SubItems(2) = Format(curInicial, "Standard")
                  itmX.SubItems(3) = Format(curDebitos, "Standard")
                  itmX.SubItems(4) = Format(curCreditos, "Standard")
                  itmX.SubItems(5) = Format(curDebitos + curCreditos, "Standard")
                  itmX.SubItems(6) = Format(curInicial + (curDebitos + curCreditos), "Standard")
                  
                  itmX.TextBackColor = RGB(214, 234, 248)
         End With
      
      End If
  
  Case "C" 'Codigos de Cuenta
    If btnAccion(6).Checked = False Then
        With lswExplorer
           strSQL = "select cod_cuenta,A.descripcion,B.Descripcion as Res,A.acepta_movimientos,A.cod_Cuenta_Mask" _
                  & " from CntX_Cuentas A inner join CntX_Tipos_Cuentas B" _
                  & " On A.tipo_cuenta = B.tipo_cuenta and A.cod_contabilidad = B.cod_contabilidad" _
                  & " where A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                  & " and A.cuenta_madre = '" & fxIndiceCodigo(vNode.Key) & "'" _
                  & " order by cod_cuenta"
           Call OpenRecordSet(rs, strSQL, 0)
           .ColumnHeaders.Add , , "Cuenta", 2600
           .ColumnHeaders.Add , , "Descripción", 3200
           .ColumnHeaders.Add , , "Tipo", 2200
           .ColumnHeaders.Add , , "Acep.Mov.", 1200, 2
            
            Do While Not rs.EOF
                Set itmX = .ListItems.Add(, , rs!Cod_Cuenta_Mask, 13)
                    itmX.SubItems(1) = rs!Descripcion
                    itmX.SubItems(2) = rs!res
                    itmX.SubItems(3) = IIf((rs!Acepta_movimientos = 1), "Sí", "No")
                If rs!Acepta_movimientos = 0 Then itmX.ForeColor = vbBlue
            rs.MoveNext
            Loop
            rs.Close
        End With
      
      Else 'Movimientos
         curDebitos = 0
         curCreditos = 0
         curInicial = 0
         
         'Si No Recibe Movimientos Mostrar las Sub CntX_Cuentas
         'Si Recibe Mov, Mostrar Asientos
         strSQL = "select acepta_movimientos from CntX_Cuentas where cod_cuenta = '" _
                & fxIndiceCodigo(vNode.Key) & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta
         Call OpenRecordSet(rs, strSQL, 0)
         If rs!Acepta_movimientos = 1 Then
             rs.Close
             With lswExplorer
                
                strSQL = "select * " _
                       & ", (MONTO_CREDITO + MONTO_DEBITO) / dbo.fxSys_Tipo_Cambio_Apl(TIPO_CAMBIO) as 'IMPORTE'" _
                       & " from Cntx_Asientos_detalle where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                       & " and fecha_asiento between '" & gCntX_Parametros.PeriodoAnio & "/" & gCntX_Parametros.PeriodoMes & "/01" _
                       & " 00:00:00' and '" & gCntX_Parametros.PeriodoAnio & "/" & gCntX_Parametros.PeriodoMes & "/" & fxCntX_UltimoDiaMes(gCntX_Parametros.PeriodoMes, gCntX_Parametros.PeriodoAnio) _
                       & " 23:59:59' and fecha_asiento is not null and cod_cuenta = '" & fxIndiceCodigo(vNode.Key) & "'"
                       
                Call OpenRecordSet(rs, strSQL, 0)
                .ColumnHeaders.Add , , "N°Asiento", 1400
                .ColumnHeaders.Add , , "Tipo", 1000
                .ColumnHeaders.Add , , "Detalle", 3200
                .ColumnHeaders.Add , , "Débitos", 1800, 1
                .ColumnHeaders.Add , , "Créditos", 1800, 1
                .ColumnHeaders.Add , , "Unidad", 1400, vbCenter
                .ColumnHeaders.Add , , "C.C.", 1400, vbCenter
                .ColumnHeaders.Add , , "Divisa", 1400, vbCenter
                .ColumnHeaders.Add , , "TipoCambio", 1400, 1
                .ColumnHeaders.Add , , "Importe", 1800, 1
                  
                    Do While Not rs.EOF
                        Set itmX = .ListItems.Add(, , rs!Num_Asiento, 5)
                            itmX.SubItems(1) = rs!Tipo_Asiento
                            itmX.SubItems(2) = IIf(IsNull(rs!Detalle), "", rs!Detalle)
                            itmX.SubItems(3) = Format(rs!monto_debito, "Standard")
                            itmX.SubItems(4) = Format(rs!monto_credito, "Standard")
                            
                            itmX.SubItems(5) = rs!Cod_Unidad & ""
                            itmX.SubItems(6) = rs!Cod_Centro_Costo & ""
                            itmX.SubItems(7) = rs!COD_DIVISA & ""
                            itmX.SubItems(8) = Format(rs!TIPO_CAMBIO, "Standard")
                            itmX.SubItems(9) = Format(rs!Importe, "Standard")
                            
                            itmX.Tag = "Asiento"
                            
                            curDebitos = curDebitos + rs!monto_debito
                            curCreditos = curCreditos + rs!monto_credito
                      rs.MoveNext
                    Loop

                    rs.Close
    
                'Totales
                       
                  Set itmX = .ListItems.Add(, , "")
                      itmX.SubItems(3) = "___________"
                      itmX.SubItems(4) = "___________"
           
                  Set itmX = .ListItems.Add(, , "TOTALES")
                      itmX.SubItems(3) = Format(curDebitos, "Standard")
                      itmX.SubItems(4) = Format(curCreditos, "Standard")
             
                      itmX.TextBackColor = RGB(214, 234, 248)
             End With
         
         
         
         
         Else 'Acepta Movimientos
         
             rs.Close
             With lswExplorer
                strSQL = "select A.cod_Cuenta_Mask,A.cod_cuenta,isnull(B.saldo_inicial,0) as Saldo_Inicial" _
                       & ",isnull(B.total_debitos,0) as Total_Debitos,A.descripcion" _
                       & ",isnull(B.total_creditos,0) as Total_creditos" _
                       & " from CntX_Cuentas A LEFT join  vCntX_Mov_Cuentas_General B" _
                       & " on A.cod_contabilidad = B.cod_contabilidad and A.cod_cuenta = B.cod_cuenta" _
                       & " and B.anio = " & gCntX_Parametros.PeriodoAnio & " and B.mes = " & gCntX_Parametros.PeriodoMes _
                       & " where A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                       & " and A.cuenta_madre = '" & fxIndiceCodigo(vNode.Key) & "'" _
                       & " order by A.cod_cuenta"
                Call OpenRecordSet(rs, strSQL, 0)
                .ColumnHeaders.Add , , "Cuenta", 2600
                .ColumnHeaders.Add , , "Descripción", 3200
                .ColumnHeaders.Add , , "Inicial", 1800, 1
                .ColumnHeaders.Add , , "Débitos", 1800, 1
                .ColumnHeaders.Add , , "Créditos", 1800, 1
                .ColumnHeaders.Add , , "Mes", 1800, 1
                .ColumnHeaders.Add , , "Actual", 1800, 1
                  
                    
'                    call sbFormsCall("frmCntX_Procesos.prgBar.Max = rs.RecordCount + 1
'                    call sbFormsCall("frmCntX_Procesos.prgBar.Value = 1
                    
                    Do While Not rs.EOF
                      If rs!cod_cuenta <> "" Then
                        Set itmX = .ListItems.Add(, , rs!Cod_Cuenta_Mask, 13)
                            itmX.SubItems(1) = rs!Descripcion
                            itmX.SubItems(2) = Format(rs!saldo_inicial, "Standard")
                            itmX.SubItems(3) = Format(rs!total_debitos, "Standard")
                            itmX.SubItems(4) = Format(rs!total_creditos, "Standard")
                            itmX.SubItems(5) = Format(rs!total_debitos + rs!total_creditos, "Standard")
                            itmX.SubItems(6) = Format(rs!saldo_inicial + (rs!total_debitos + rs!total_creditos), "Standard")
                            curDebitos = curDebitos + rs!total_debitos
                            curCreditos = curCreditos + rs!total_creditos
                            curInicial = curInicial + rs!saldo_inicial
                      End If
'                      If call sbFormsCall("frmCntX_Procesos.prgBar.Value < frmX.prgBar.Max Then frmX.prgBar.Value = frmX.prgBar.Value + 1
                      rs.MoveNext
                    Loop
                    rs.Close
    
                'Totales
                       
                  Set itmX = .ListItems.Add(, , "")
                      itmX.SubItems(2) = "___________"
                      itmX.SubItems(3) = "___________"
                      itmX.SubItems(4) = "___________"
                      itmX.SubItems(5) = "___________"
                      itmX.SubItems(6) = "___________"
           
           
                  Set itmX = .ListItems.Add(, , "TOTALES")
                      itmX.SubItems(2) = Format(curInicial, "Standard")
                      itmX.SubItems(3) = Format(curDebitos, "Standard")
                      itmX.SubItems(4) = Format(curCreditos, "Standard")
                      itmX.SubItems(5) = Format(curDebitos + curCreditos, "Standard")
                      itmX.SubItems(6) = Format(curInicial + (curDebitos + curCreditos), "Standard")
                      itmX.TextBackColor = RGB(214, 234, 248)
             
             End With
         
         End If 'Acepta Movimientos
       
       
       
       End If 'Pressed
  
  
  
   Case "X" 'Tipos de Asientos
     With lswExplorer
       .ColumnHeaders.Add , , "N°Asiento", 1850
       .ColumnHeaders.Add , , "Tipo", 1250, 2
       .ColumnHeaders.Add , , "Fecha", 1250, 2
       .ColumnHeaders.Add , , "Descripción", 3450
       .ColumnHeaders.Add , , "Debitos", 1850, 1
       .ColumnHeaders.Add , , "Creditos", 1850, 1
       .ColumnHeaders.Add , , "Mayorizado", 1450, 2
       .ColumnHeaders.Add , , "Referencia", 4450
       
        strSQL = "select A.Tipo_asiento,A.Num_asiento,A.descripcion,A.fecha_asiento,A.fecha_aplicado,A.Referencia, isnull(sum(D.monto_debito),0) as Debito, isnull(sum(D.monto_credito),0) as credito" _
               & " from Cntx_Asientos A left join Cntx_Asientos_Detalle D on A.cod_contabilidad = D.cod_contabilidad" _
               & " and A.num_asiento = D.num_Asiento and A.tipo_asiento = D.tipo_asiento" _
               & " where A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
               & " and A.anio = " & gCntX_Parametros.PeriodoAnio & " and A.Mes = " & gCntX_Parametros.PeriodoMes _
               & " and A.tipo_asiento = '" & fxIndiceCodigo(vNode.Key) & "'"
               
               
                 
        If btnSemaforo(1).Checked Then
           strSQL = strSQL & " and A.Fecha_Aplicado is not null"
        End If
        
        If btnSemaforo(2).Checked Then
           strSQL = strSQL & " and A.Fecha_Aplicado is null"
        End If
        
        'Si está marcado: Muestra Des-Balanceados, Si está desmarcado muestra TODOS
        If btnSemaforo(3).Checked Then
           strSQL = strSQL & " and A.Balanceado = 'N'"
        End If
        
       strSQL = strSQL & " group by A.Tipo_asiento,A.Num_asiento,A.descripcion,A.fecha_asiento,A.fecha_aplicado,A.Referencia" _
               & " order by A.Tipo_asiento,A.Num_asiento"
               
               
       Call OpenRecordSet(rs, strSQL, 0)
        curDebitos = 0
        curCreditos = 0
        
'        frmX.prgBar.Max = rs.RecordCount + 1
'        frmX.prgBar.Value = 1
        
        Do While Not rs.EOF
          curDebitos = curDebitos + rs!Debito
          curCreditos = curCreditos + rs!Credito
          
          Set itmX = .ListItems.Add(, , rs!Num_Asiento)
              itmX.SubItems(1) = rs!Tipo_Asiento
              itmX.SubItems(2) = rs!fecha_asiento
              itmX.SubItems(3) = rs!Descripcion
              itmX.SubItems(4) = Format(rs!Debito, "Standard")
              itmX.SubItems(5) = Format(rs!Credito, "Standard")
              itmX.SubItems(6) = IIf(IsNull(rs!Fecha_Aplicado), "NO", "SI")
              itmX.SubItems(7) = rs!Referencia & ""
          
'           If IsNull(rs!Fecha_Aplicado) Then
'              itmX.Icon = 15
'           Else
'              itmX.Icon = 14
'           End If
           
           If IsNull(rs!Fecha_Aplicado) Then
              itmX.Bold = True
              itmX.TextBackColor = RGB(252, 243, 207) 'Amarillo
            End If
           
           If rs!Debito <> rs!Credito Then
'              itmX.Bold = True
'              itmX.ForeColor = vbRed
'              itmX.Icon = 16
              itmX.ForeColor = vbRed
              itmX.Bold = True
              itmX.TextBackColor = RGB(250, 219, 216) 'Rojo
           
           End If
           
           itmX.Tag = "Asiento"
          
'          If frmX.prgBar.Value < frmX.prgBar.Max Then frmX.prgBar.Value = frmX.prgBar.Value + 1
          rs.MoveNext
        Loop
        rs.Close
        
          Set itmX = .ListItems.Add(, , "")
              itmX.SubItems(4) = "___________"
              itmX.SubItems(5) = "___________"
        
          Set itmX = .ListItems.Add(, , "")
              itmX.SubItems(3) = "Transac. No.: " & Format(.ListItems.Count - 2, "###,###,##0")
              itmX.SubItems(4) = Format(curDebitos, "Standard")
              itmX.SubItems(5) = Format(curCreditos, "Standard")
              itmX.TextBackColor = RGB(214, 234, 248)
        
     End With
     
   Case "A" 'Numero de Asiento
     
     With lswExplorer
       .ColumnHeaders.Add , , "Cuenta", 2600
       .ColumnHeaders.Add , , "Descripción", 3250
       .ColumnHeaders.Add , , "Debitos", 1850, 1
       .ColumnHeaders.Add , , "Creditos", 1850, 1
       .ColumnHeaders.Add , , "Documento", 2850
       .ColumnHeaders.Add , , "Detalle", 3850
       .ColumnHeaders.Add , , "Unidad", 900
       .ColumnHeaders.Add , , "C.C.", 900
       .ColumnHeaders.Add , , "Divisa", 900
       .ColumnHeaders.Add , , "T.C.", 900, 1
       .ColumnHeaders.Add , , "Importe", 1850, 1
       
       txtNotas = fxNotas(fxIndiceAsiento(vNode.Key, "T"), fxIndiceAsiento(vNode.Key, "N"))
       
       strSQL = "select A.*,B.descripcion as Res,B.cod_Cuenta_Mask" _
              & ", (A.MONTO_CREDITO+A.MONTO_DEBITO) / dbo.fxSys_Tipo_Cambio_Apl(A.TIPO_CAMBIO) as 'IMPORTE'" _
              & " from Cntx_Asientos_detalle A inner join CntX_Cuentas B" _
              & " On A.cod_cuenta = B.cod_cuenta and A.cod_contabilidad = B.cod_contabilidad" _
              & " where A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
              & " and tipo_asiento = '" & fxIndiceAsiento(vNode.Key, "T") & "'" _
              & " and num_asiento = '" & fxIndiceAsiento(vNode.Key, "N") & "'"
        Call OpenRecordSet(rs, strSQL, 0)
        curDebitos = 0
        curCreditos = 0
        
'        frmX.prgBar.Max = rs.RecordCount + 1
'        frmX.prgBar.Value = 1
        
        Do While Not rs.EOF
          
          curDebitos = curDebitos + rs!monto_debito
          curCreditos = curCreditos + rs!monto_credito
          
          Set itmX = .ListItems.Add(, , rs!Cod_Cuenta_Mask, 13)
              itmX.SubItems(1) = rs!res
              itmX.SubItems(2) = Format(rs!monto_debito, "Standard")
              itmX.SubItems(3) = Format(rs!monto_credito, "Standard")
              itmX.SubItems(4) = rs!Documento
              itmX.SubItems(5) = rs!Detalle
              itmX.SubItems(6) = rs!Cod_Unidad & ""
              itmX.SubItems(7) = rs!Cod_Centro_Costo & ""
              itmX.SubItems(8) = rs!COD_DIVISA & ""
              itmX.SubItems(9) = rs!TIPO_CAMBIO & ""
              itmX.SubItems(10) = Format(rs!Importe, "Standard")
        
'          If frmX.prgBar.Value < frmX.prgBar.Max Then frmX.prgBar.Value = frmX.prgBar.Value + 1
          rs.MoveNext
        Loop
        rs.Close
        
          Set itmX = .ListItems.Add(, , "")
              itmX.SubItems(2) = "___________"
              itmX.SubItems(3) = "___________"
        
          Set itmX = .ListItems.Add(, , "")
              itmX.SubItems(2) = Format(curDebitos, "Standard")
              itmX.SubItems(3) = Format(curCreditos, "Standard")
          If Abs(curDebitos - curCreditos) > 0 Then
             itmX.Text = "DESBALANCE >>"
             itmX.SubItems(1) = Format(curDebitos - curCreditos, "Standard")
             itmX.ForeColor = vbRed
          Else
               itmX.TextBackColor = RGB(214, 234, 248)
          End If
        
     End With
     
     
   Case Else
       'Mantenimiento
       Call sbMuestraParametros

End Select

'Unload frmX
Me.MousePointer = vbDefault

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
 
End Sub


Private Sub ArbolExp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim frmX As Form

For Each frmX In Forms
   If Mid(frmX.Name, 1, 3) = "MDI" Then
        Exit For
   End If
Next

If Button = 2 Then
   Call PopupMenu(frmX.mnuExplorerContable, , x, y)
   ' Call sbMenuPopUp_Show(0)
End If
End Sub

Private Sub ArbolExp_NodeClick(ByVal Node As MSComctlLib.Node)
Dim itmX As ListViewItem, strSQL As String, rs As New ADODB.Recordset
Dim curDebito As Currency, curCredito As Currency, lngMovimientos As Long

On Error Resume Next

Set vNode = Node

curDebito = 0
curCredito = 0
lngMovimientos = 0

lswExplorer.Enabled = True

lblTitle(0).Caption = UCase(vNode.Parent)   'Mid(vNode.FullPath, 1, (Len(vNode.FullPath) - Len(vNode.Text)) - 1)
lblTitle(1).Caption = UCase(vNode.Text) 'UCase (vNode.Parent) & " : " & UCase(vNode.Text)

lswExplorer.ListItems.Clear
lswExplorer.ColumnHeaders.Clear

'Habilita o No, Los botones de Borrar y Mayorizar
'Depende del nodo activo, solo activa en la secciones de Cntx_Asientos
Select Case Right(vNode.Key, 1)
  Case "A"
     
    btnAccion(1).Enabled = IIf((Right(vNode.Key, 1) = "A"), True, False)
    btnAccion(4).Enabled = IIf((Right(vNode.Key, 1) = "A"), True, False)
  Case "X"
    btnAccion(4).Enabled = IIf((Right(vNode.Key, 1) = "X"), True, False)
End Select

Select Case vNode.Text
  Case "ProGrX: Contabilidad"
     With lswExplorer
      .ColumnHeaders.Add , , "Empresa", 2450
      .ColumnHeaders.Add , , "Usuario", 2450
      .ColumnHeaders.Add , , "Fecha", 1450
       Set itmX = .ListItems.Add(.ListItems.Count + 1, , gCntX_Parametros.EmpresaLocal, 13)
           itmX.Tag = itmX.Index
           itmX.SubItems(1) = UCase(glogon.Usuario)
           itmX.SubItems(2) = Date
     End With
  
  
  Case "Asientos"
     With lswExplorer
        .ColumnHeaders.Add , , "Código", 1050
        .ColumnHeaders.Add , , "Descripción", 3850
        .ColumnHeaders.Add , , "Movimientos", 1200, vbCenter
        .ColumnHeaders.Add , , "Total Débitos", 2850, vbRightJustify
        .ColumnHeaders.Add , , "Total Créditos", 2850, vbRightJustify
        .ColumnHeaders.Add , , "AS: Total", 1600, vbCenter
        .ColumnHeaders.Add , , "AS: Aplicados", 1600, vbCenter
        .ColumnHeaders.Add , , "AS: Pendientes", 1600, vbCenter
        .ColumnHeaders.Add , , "AS: Desbalanceados", 1600, vbCenter
        
        strSQL = "exec spCntx_Consulta_Asientos_Rsm " & gCntX_Parametros.CodigoConta _
               & "," & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
           Set itmX = .ListItems.Add(, , rs!Tipo_Asiento)
               itmX.SubItems(1) = rs!Descripcion
               itmX.SubItems(2) = Format(rs!Movimientos, "###,###,###,##0")
               itmX.SubItems(3) = Format(rs!Debitos, "Standard")
               itmX.SubItems(4) = Format(rs!Creditos, "Standard")
               itmX.SubItems(5) = Format(rs!Asientos_Total, "###,###,###,##0")
               itmX.SubItems(6) = Format(rs!Asientos_Aplicados, "###,###,###,##0")
               itmX.SubItems(7) = Format(rs!Asientos_Pendientes, "###,###,###,##0")
               itmX.SubItems(8) = Format(rs!Asientos_Desbalanceados, "###,###,###,##0")
          
          curDebito = curDebito + rs!Debitos
          curCredito = curCredito + rs!Creditos
          lngMovimientos = lngMovimientos + rs!Movimientos
          
          rs.MoveNext
        Loop
        rs.Close
             
           Set itmX = .ListItems.Add(, , "")
               itmX.SubItems(1) = "TOTAL:"
               itmX.SubItems(2) = Format(lngMovimientos, "###,###,###,##0")
               itmX.SubItems(3) = Format(curDebito, "Standard")
               itmX.SubItems(4) = Format(curCredito, "Standard")
     
               itmX.Bold = True
               itmX.ForeColor = vbWhite
               itmX.TextBackColor = RGB(214, 234, 248)
     End With
  
  
  Case "Catálogo de Cuentas"
     With lswExplorer
        .ColumnHeaders.Add , , "Código", 850
        .ColumnHeaders.Add , , "Descripción", 3850
        .ColumnHeaders.Add , , "Clasificacion", 2850
        .ColumnHeaders.Add , , "Movimientos", 1200, vbCenter
        .ColumnHeaders.Add , , "Total Débitos", 2850, vbRightJustify
        .ColumnHeaders.Add , , "Total Créditos", 2850, vbRightJustify
        .ColumnHeaders.Add , , "Diferencia", 1850, vbRightJustify
        
        strSQL = "exec spCntx_Consulta_Tipo_Cuenta_Rsm " & gCntX_Parametros.CodigoConta _
               & "," & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
           Set itmX = .ListItems.Add(.ListItems.Count + 1, , rs!tipo_cuenta)
               itmX.SubItems(1) = rs!Descripcion
               itmX.SubItems(2) = fxCntX_TiposCuentas(rs!Clasificacion)
               itmX.SubItems(3) = Format(rs!Movimientos, "###,###,###,##0")
               itmX.SubItems(4) = Format(rs!Debitos, "Standard")
               itmX.SubItems(5) = Format(rs!Creditos, "Standard")
               itmX.SubItems(6) = Format(rs!Diferencia, "Standard")
          
          curDebito = curDebito + rs!Debitos
          curCredito = curCredito + rs!Creditos
          lngMovimientos = lngMovimientos + rs!Movimientos
          
          rs.MoveNext
        Loop
        rs.Close
  
           Set itmX = .ListItems.Add(, , "")
               itmX.SubItems(2) = "TOTAL:"
               itmX.SubItems(3) = Format(lngMovimientos, "###,###,###,##0")
               itmX.SubItems(4) = Format(curDebito, "Standard")
               itmX.SubItems(5) = Format(curCredito, "Standard")
               itmX.SubItems(6) = Format(curDebito - curCredito, "Standard")
               
               itmX.TextBackColor = RGB(214, 234, 248)
               itmX.Bold = True
  
     End With
  
  
  Case "Areas de Trabajo"
     With lswExplorer
        .ColumnHeaders.Add , , "Código", 850
        .ColumnHeaders.Add , , "Descripción", 3850
        strSQL = "select cod_area,descripcion from CntX_Area_Definicion where cod_contabilidad = " _
              & gCntX_Parametros.CodigoConta & " order by cod_area"
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
           Set itmX = .ListItems.Add(.ListItems.Count + 1, , rs!cod_area, 13)
               itmX.SubItems(1) = rs!Descripcion
          rs.MoveNext
        Loop
        rs.Close
     End With
  
  Case "Administracion  de Diferidos"
     With lswExplorer
        .ColumnHeaders.Add , , "Código", 850
        .ColumnHeaders.Add , , "Descripción", 3850
        strSQL = "select cod_diferido,descripcion from CntX_Diferidos where cod_contabilidad = " _
              & gCntX_Parametros.CodigoConta & " order by cod_diferido"
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
           Set itmX = .ListItems.Add(.ListItems.Count + 1, , rs!cod_diferido, 5)
               itmX.SubItems(1) = rs!Descripcion
          rs.MoveNext
        Loop
        rs.Close
     End With
  
  
  Case "Plantillas Asientos Porcentuales"
     With lswExplorer
        .ColumnHeaders.Add , , "Código", 850
        .ColumnHeaders.Add , , "Descripción", 3850
        strSQL = "select cod_plantilla,descripcion from CntX_Plantilla_Rate where cod_contabilidad = " _
              & gCntX_Parametros.CodigoConta & " order by cod_plantilla"
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
           Set itmX = .ListItems.Add(.ListItems.Count + 1, , rs!cod_plantilla, 13)
               itmX.SubItems(1) = rs!Descripcion
          rs.MoveNext
        Loop
        rs.Close
     End With
  
  Case "Plantillas Asientos Fijos"
  
     With lswExplorer
        .ColumnHeaders.Add , , "Código", 850
        .ColumnHeaders.Add , , "Descripción", 3850
        strSQL = "select cod_plantilla,descripcion from CntX_Plantilla_Asientos where cod_contabilidad = " _
              & gCntX_Parametros.CodigoConta & " order by cod_plantilla"
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
           Set itmX = .ListItems.Add(.ListItems.Count + 1, , rs!cod_plantilla, 13)
               itmX.SubItems(1) = rs!Descripcion
          rs.MoveNext
        Loop
        rs.Close
     End With
  
  
  
  Case "Presupuesto"
   '********************************* OJO PENDIENTE DE DESARROLLO
     With lswExplorer
        .ColumnHeaders.Add , , "Cuenta", 1850
        .ColumnHeaders.Add , , "Descripción", 3850
        .ColumnHeaders.Add , , "Pres_Presupuesto", 1250, vbRightJustify
        .ColumnHeaders.Add , , "Real", 1250, vbRightJustify
        .ColumnHeaders.Add , , "Diferencia", 1250, vbRightJustify
        .ColumnHeaders.Add , , "Pre.Original", 1250, vbRightJustify
        .ColumnHeaders.Add , , "(+) Ajuste", 1250, vbRightJustify
        .ColumnHeaders.Add , , "(-) Ajuste", 1250, vbRightJustify
        
        strSQL = "select P.*,C.descripcion,(isnull(M.saldo_inicial,0) + isnull(M.total_Debitos,0)" _
               & " + isnull(M.total_creditos,0)) as Real" _
               & " from Pres_Presupuesto P inner join CntX_Cuentas C on P.cod_cuenta = C.cod_cuenta and P.cod_contabilidad = C.cod_contabilidad" _
               & " left join vCntX_Mov_Cuentas_General M on P.cod_cuenta = M.cod_cuenta" _
               & " and P.anio = M.anio and P.mes = M.mes and P.cod_contabilidad = M.cod_contabilidad" _
               & " where P.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
               & " and P.anio = " & gCntX_Parametros.PeriodoAnio & " and P.Mes = " _
               & gCntX_Parametros.PeriodoMes
        
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
           Set itmX = .ListItems.Add(, , fxCntX_CuentaFormato(True, rs!cod_cuenta, 0), 13)
               itmX.SubItems(1) = rs!Descripcion
               itmX.SubItems(2) = Format(rs!presu_actual, "Standard")
               itmX.SubItems(3) = Format(rs!Real, "Standard")
               itmX.SubItems(4) = Format(rs!presu_actual - rs!Real, "Standard")
               itmX.SubItems(5) = Format(rs!presu_original, "Standard")
               itmX.SubItems(6) = Format(rs!ajuste_positivo, "Standard")
               itmX.SubItems(7) = Format(rs!ajuste_negativo, "Standard")
               If rs!presu_actual < rs!Real Then itmX.ForeColor = vbRed
          rs.MoveNext
        Loop
        rs.Close
     End With
  
  
  Case "Mantenimiento"
  
     With lswExplorer
      .ColumnHeaders.Add , , "", 2450
      .ColumnHeaders.Add , , "", 4450
       Set itmX = .ListItems.Add(.ListItems.Count + 1, , "Mantenimiento de:", 13)
           itmX.SubItems(1) = "[NECESITA CLAVE DE ADMINISTRADOR]"
           itmX.ForeColor = vbBlue
       Set itmX = .ListItems.Add(.ListItems.Count + 1, , "")
           itmX.SubItems(1) = "Tipos de Cuentas"
       Set itmX = .ListItems.Add(.ListItems.Count + 1, , "")
           itmX.SubItems(1) = "Tipos de Asientos"
       Set itmX = .ListItems.Add(.ListItems.Count + 1, , "")
           itmX.SubItems(1) = "Periodos"
       Set itmX = .ListItems.Add(.ListItems.Count + 1, , "")
           itmX.SubItems(1) = "Contabilidades"
       Set itmX = .ListItems.Add(.ListItems.Count + 1, , "")
           itmX.SubItems(1) = "Empresa"
       Set itmX = .ListItems.Add(.ListItems.Count + 1, , "")
           itmX.SubItems(1) = "Parámetros del Sistema"
           
     End With
  
  
  Case Else
 
    Call sbMuestraDetalleSubNodos

End Select

End Sub


Public Sub sbButtonPopUp(i As Integer)

On Error GoTo vError

Select Case i
 Case 1 'Editar
   Call btnAccion_Click(0)
 
 Case 2 'Borrar
   Call btnAccion_Click(1)
   
 Case 3 'Refrescar
   Call btnAccion_Click(2)
 
 Case 4 'Imprimir
   Call btnAccion_Click(3)
 
 Case 5 'Mayorizar
   Call btnAccion_Click(4)

End Select

vError:

End Sub

Private Sub btnAccion_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim frmX As Form, vEncuentra As Boolean


If Index = 5 Then
    btnAccion.Item(5).Checked = True
    btnAccion.Item(6).Checked = False
End If

If Index = 6 Then
    btnAccion.Item(5).Checked = False
    btnAccion.Item(6).Checked = True
End If


Select Case Index
  Case 0  'editar
     If vNode.Index > 1 Then
         Select Case Right(vNode.Key, 1)
             Case "A"
                 gCntX_Arbol.ArbolActivo = True
                 gCntX_Arbol.AsientoTipo = fxIndiceAsiento(vNode.Key, "T")
                 gCntX_Arbol.AsientoNumr = fxIndiceAsiento(vNode.Key, "N")
                 
                vEncuentra = False
                For Each frmX In Forms
                   If Trim(frmX.Name) = "frmCntX_Asientos" Then
                        frmX.Show
                        frmX.sbFormReLoad
                        vEncuentra = True
                   End If
                Next
                
                If Not vEncuentra Then
                    Call sbFormsCall("frmCntX_Asientos")
                End If
                 
                 gCntX_Arbol.ArbolActivo = False
             Case "T"
                 Call sbClassCall("Contabilidad", 0, "frmCntX_TiposCuentas")
             Case "X"
                 gCntX_Arbol.ArbolActivo = False
                 Call sbFormsCall("frmCntX_Asientos")
             
             Case "C"
                 Call sbClassCall("Contabilidad", 0, "frmCntX_CatalogoCuentas")
             
             Case "F" 'Plantilla Cntx_Asientos  Fijos
                 Call sbClassCall("Contabilidad", 0, "frmCntX_PlantillaAsientos")
             
             Case "R" 'Plantilla Cntx_Asientos RAte
                 Call sbClassCall("Contabilidad", 0, "frmCntX_PlantillaRate")
             
             Case "D" 'Tipos de Asientos
                 Call sbClassCall("Contabilidad", 0, "frmCntX_TiposAsientos")
             
             Case "E" 'Plantilla de Diferidos para Ejecucion
                 Call sbClassCall("Contabilidad", 0, "frmCntX_DiferidosCreacion")
             
             Case "Y" 'Areas de Trabajo
                 Call sbClassCall("Contabilidad", 0, "frmCntX_AreaDefinicion")
                 
             Case Else 'Mantenimiento"
               Select Case vNode.Text
                 Case "Tipos de Asientos"
                      Call sbClassCall("Contabilidad", 0, "frmCntX_TiposAsientos")
                 Case "Tipos de Cuentas"
                      Call sbClassCall("Contabilidad", 0, "frmCntX_TiposCuentas")
                 Case "Contabilidades"
                      Call sbFormsCall("frmCntX_Contabilidades", 1)
                 Case "Usuarios"
                      MsgBox "Opcion no asignada...!", vbInformation
                 Case "Periodos"
                      Call sbClassCall("Contabilidad", 0, "frmCntX_PeriodosDefinicion")
                 Case "Cierres"
                      Call sbClassCall("Contabilidad", 0, "frmCntX_Cierres")
                 Case "Parámetros"
                      MsgBox "Opcion no asignada...!", vbInformation
                 Case "Empresa"
                      Call sbClassCall("Contabilidad", 0, "frmCntX_Empresa")
                 Case "Divisas"
                      Call sbClassCall("Contabilidad", 0, "frmCntX_Divisas")
               End Select
               Call ArbolExp_NodeClick(vNode)
     End Select
       End If
  
  
  Case 1 'Borrar
  
       If Right(vNode.Key, 1) = "A" Then
       
          Call sbCntX_AsientoBorra(fxIndiceAsiento(vNode.Key, "T"), fxIndiceAsiento(vNode.Key, "N"), gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes)
          On Error Resume Next
          ArbolExp.Nodes.Remove vNode.Index
       End If
          
          
          
  Case 2 'refrescar
    Call sbRefrescaArbol
    
    
  Case 5, 6  ' Detalle, Movimientos
    On Error Resume Next
    Call ArbolExp_NodeClick(vNode)
    
  Case 3 'Imprimir
     If vNode.Index > 1 Then
         Select Case Right(vNode.Key, 1)
             Case "A" 'Asientos
                strSQL = "{Cntx_Asientos.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
                       & " AND {Cntx_Asientos.TIPO_ASIENTO} = '" & fxIndiceAsiento(vNode.Key, "T") & "' AND " _
                       & " {Cntx_Asientos.NUM_ASIENTO} = '" & fxIndiceAsiento(vNode.Key, "N") & "'"
                Call sbCntX_Reportes("ASIENTO", strSQL)
             Case "T" 'Tipos de Cuentas
               Call sbCntX_Reportes_Catalogos("Tipos_Cuentas")
             
             Case "X" 'Asientos por Tipos
                strSQL = "{Cntx_Asientos.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
                       & " AND {Cntx_Asientos.TIPO_ASIENTO} = '" & fxIndiceCodigo(vNode.Key) & "'" _
                       & " AND {Cntx_Asientos.MES} = " & gCntX_Parametros.PeriodoMes & " AND {Cntx_Asientos.ANIO} = " & gCntX_Parametros.PeriodoAnio
               Call sbCntX_Reportes("ASIENTO", strSQL)
             Case "C" 'Cuentas
                Call sbCntX_Reportes("CATALOGO", "{CntX_Tipos_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta, "Detalle")
             
             Case "F" 'Plantilla de Asientos Fijos
             Case "R" 'Plantilla de Asientos Rate
             Case "D" 'Plantilla Diferidos
             Case "E" 'Diferidos Creacion
             Case "Y" 'Areas de Trabajo
                strSQL = "{CntX_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
                       & " AND {CntX_Area_Cuentas.COD_AREA} = " & fxIndiceCodigo(vNode.Key)
                Call sbCntX_Reportes("CATALOGOAREAS", strSQL, vNode.Text)

             Case Else 'Mantenimiento"
               Select Case vNode.Text
                 Case "Tipos de Asientos"
                       Call sbCntX_Reportes_Catalogos("Tipos_Asientos")
                 Case "Tipos de CntX_Cuentas"
                       Call sbCntX_Reportes_Catalogos("Tipos_Cuentas")
                 Case "Contabilidades"
                 Case "Usuarios"
                       Call sbCntX_Reportes_Catalogos("usuarios")
                 Case "Periodos"
                       Call sbCntX_Reportes_Catalogos("Periodos")
                 Case "Cierres"
                       Call sbCntX_Reportes_Catalogos("cierres")
                 Case "Parámetros"
                       Call sbCntX_Reportes_Catalogos("parametros")
                 Case "Empresa"
               End Select
       End Select
     End If
    

  Case 4 'Aplicar, Mayorizar
     If vNode.Index > 1 Then
         Select Case Right(vNode.Key, 1)
             Case "A"
                Call sbFormsCall("frmCntX_Procesos")
                Call sbFormActivo("frmCntX_Procesos", frmX)
                
                frmX.Caption = "Mayorizando Asientos..."
                
                frmX.lbl.Caption = "Cargando Información..."
                frmX.lbl.Refresh
                
                strSQL = "select Asi.Tipo_Asiento,Asi.Num_Asiento,Asi.Fecha_Asiento,Asi.TS" _
                       & " from Cntx_Asientos Asi left join Cntx_Asientos_Detalle Ad on Asi.Cod_Contabilidad = Ad.cod_Contabilidad" _
                       & " and Asi.Tipo_Asiento = Ad.Tipo_Asiento and Asi.Num_Asiento = Ad.Num_Asiento" _
                       & " where Asi.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                       & "  AND Asi.TIPO_ASIENTO = '" & fxIndiceAsiento(vNode.Key, "T") _
                       & "' AND Asi.NUM_ASIENTO = '" & fxIndiceAsiento(vNode.Key, "N") & "'" _
                       & "  AND Asi.fecha_aplicado is null and Asi.balanceado = 'S'" _
                       & " group by Asi.Tipo_Asiento,Asi.Num_Asiento,Asi.Fecha_Asiento,Asi.TS" _
                       & " Having sum(Ad.Monto_Debito) - sum(Ad.Monto_Credito) = 0"
                Call OpenRecordSet(rs, strSQL, 0)
                frmX.PrgBar.Max = rs.RecordCount + 1
                frmX.PrgBar.Value = 1
         
                frmX.lbl.Caption = "Procesando Información..."
                frmX.lbl.Refresh
                
                Do While Not rs.EOF
                  If fxCntX_AsientoConcurrencia(rs!Tipo_Asiento, rs!Num_Asiento) = fxTsToHex(rs!TS) Then
                    Call sbCntX_Asiento_Mayorizar(rs!Tipo_Asiento, rs!Num_Asiento, rs!fecha_asiento)
                  End If
                  If frmX.PrgBar.Value < frmX.PrgBar.Max Then frmX.PrgBar.Value = frmX.PrgBar.Value + 1
                  rs.MoveNext
                Loop
                rs.Close
                
                Unload frmX
                
             Case "X"
                Call sbFormsCall("frmCntX_Procesos")
                Call sbFormActivo("frmCntX_Procesos", frmX)
                
                frmX.Caption = "Mayorizando Asientos..."
                
                frmX.lbl.Caption = "Aplicado Asientos: " & vNode.Text
                frmX.lbl.Refresh
                
                strSQL = "exec spCntX_AsientosAplicacionLote_TipoAsiento " & gCntX_Parametros.CodigoConta _
                        & "," & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes _
                        & ",'M','" & glogon.Usuario & "','" & fxIndiceCodigo(vNode.Key) & "'"
                Call ConectionExecute(strSQL, 0)

                Unload frmX
            Case Else
             If Right(vNode.Key, 1) = "D" And Mid(vNode.Key, 1, 8) = "Asientos" Then
                Call sbFormsCall("frmCntX_MayorizacionFull", 1)
             End If
         End Select
     
     End If
    
End Select

End Sub

Private Sub btnBuscar_Click()
Call sbConsulta_Analitico
End Sub


Public Sub sbMenuPopUp_Acciones(Index As Integer)

On Error GoTo vError

Select Case Index
  Case 1 'Editar
        Call sbButtonPopUp(1)
  Case 2 'Borrar
        Call sbButtonPopUp(2)
  Case 3 'Reporte
        Call sbButtonPopUp(4)
  Case 4 'Mayorizar
        Call sbButtonPopUp(5)
  Case 5 'Refrescar
        Call sbButtonPopUp(3)
  Case 6 'Cerrar
End Select


Exit Sub
        
vError:
        Me.MousePointer = vbDefault
        MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

'Private Sub sbMenuPopUp_Show(pItems As Integer)
'Dim oBar As XtremeCommandBars.CommandBar
'Dim oControl As XtremeCommandBars.CommandBarControl
'Dim oPopup As XtremeCommandBars.CommandBarPopup
'
'
'Set oBar = cbMenuPopUp.Add("Menú", xtpBarPopup)
'    oBar.EnableAnimation = True
'
'With oBar.Controls
'
'.Add xtpControlButton, 1, "Editar"
'.Add xtpControlButton, 2, "Borrar"
'.Add(xtpControlButton, 3, "Reporte").BeginGroup = True
'.Add(xtpControlButton, 4, "Mayorizar").BeginGroup = True
'.Add xtpControlButton, 5, "Refrescar"
'.Add(xtpControlButton, 6, "Cerrar").BeginGroup = True
'
'End With
'
''show it
'oBar.ShowPopup
'
'End Sub


Private Sub cbMenuPopUp_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call sbMenuPopUp_Acciones(Control.Id)
End Sub



Private Sub btnConcilia_Click()

Call sbClassCall("Contabilidad", 0, "frmCntX_Conciliacion_Mov")

End Sub

Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lswExplorer, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnSemaforo_Click(Index As Integer)
Dim i As Integer

For i = 0 To btnSemaforo.Count - 1
    btnSemaforo.Item(i).Checked = False
Next i

btnSemaforo.Item(Index).Checked = True

End Sub

Private Sub btnVisualizar_Click(Index As Integer)
Dim i As Integer

For i = 0 To btnVisualizar.Count - 1
    btnVisualizar.Item(i).Checked = False
Next i

btnVisualizar.Item(Index).Checked = True

Select Case Index
    Case 0 'Notas
        fraNotas.Width = lswExplorer.Width
        txtNotas.Width = lswExplorer.Width - 200
        fraNotas.Left = lswExplorer.Left
        
        'Set Height
        If fraNotas.Visible = True Then
          fraNotas.Visible = False
          lswExplorer.Height = ArbolExp.Height
        Else
          fraNotas.Visible = True
          lswExplorer.Height = ArbolExp.Height - (fraNotas.Height + 100)
          fraNotas.top = lswExplorer.top + lswExplorer.Height + 100
        End If
    
    
    Case 1 'Avanzada
       If btnVisualizar.Item(Index).Checked Then
          fraConsultaAvanzada.Visible = True
       Else
          fraConsultaAvanzada.Visible = False
       End If
       
       Call Form_Resize
 
 End Select


End Sub

Private Sub cboMov_Click()
If cboMov.Text = "NA" Then
    txtMovimiento(0).Enabled = False
Else
    txtMovimiento(0).Enabled = True
End If

txtMovimiento(1).Enabled = txtMovimiento(0).Enabled

End Sub

Private Sub Form_Activate()
On Error GoTo vError
  dtpFechaInicio.Value = CDate(gCntX_Parametros.PeriodoAnio & "/" & gCntX_Parametros.PeriodoMes & "/01")
  dtpFechaCorte.Value = DateAdd("d", -1, DateAdd("m", 1, dtpFechaInicio.Value))
vError:


End Sub

Private Sub Form_Load()
 Dim i As Integer

 
 vModulo = 20

cboMov.Clear
cboMov.AddItem "DB"
cboMov.AddItem "CR"
cboMov.AddItem "TD"
cboMov.AddItem "NA"
cboMov.Text = "NA"


Call cboMov_Click

With imgLista.ListImages
    For i = 1 To .Count
      lswExplorer.Icons.AddIcon .Item(i).ExtractIcon.Handle, i, xtpImageNormal
    Next i
End With

Me.BackColor = RGB(214, 234, 248)

 Call sbRefrescaArbol

  vPaso = False
  
  dtpFechaInicio.Value = CDate(gCntX_Parametros.PeriodoAnio & "/" & gCntX_Parametros.PeriodoMes & "/01")
  dtpFechaCorte.Value = DateAdd("d", -1, DateAdd("m", 1, dtpFechaInicio.Value))
  
  chkFechasTodas.Value = xtpUnchecked
  Call chkFechasTodas_Click

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'   If UnloadMode = 0 Then
'      Cancel = True
'      Me.WindowState = 1
'   End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left

    gbBarra.Width = Me.Width

End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With imgSplitter
        picSplitter.Move .Left, .top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub


'Private Sub TreeView1_DragDrop(Source As Control, x As Single, Y As Single)
'    If Source = imgSplitter Then
'        SizeControls x
'    End If
'End Sub


Private Sub SizeControls(x As Single)
    On Error Resume Next
'    tlbPrincipal.Width = 9205
'    CoolBar.Bands.Item(1).Width = 9495
'    CoolBar.Bands.Item(2).Width = 975
    'set the width
    If x < 3665 Then x = 3665
    If x > (Me.Width - 3665) Then x = Me.Width - 3665
    ArbolExp.Width = x
    imgSplitter.Left = x
    lswExplorer.Left = x + 40
    lswExplorer.Width = Me.Width - (ArbolExp.Width + 160)
    
    lblTitle(0).Width = ArbolExp.Width
    lblTitle(1).Left = lblTitle(0).Left + lblTitle(0).Width
    lblTitle(1).Width = Me.Width

    fraConsultaAvanzada.Width = lswExplorer.Width
    
    
    'set the top
    lswExplorer.top = ArbolExp.top
    imgSplitter.top = ArbolExp.top
    ArbolExp.Height = Me.Height - 1280
    
    fraNotas.Width = lswExplorer.Width - 60
    txtNotas.Width = fraNotas.Width - 240
    fraNotas.Left = lswExplorer.Left
    
    imgSplitter.Height = ArbolExp.Height
    
    fraConsultaAvanzada.top = lswExplorer.top
    fraConsultaAvanzada.Left = lswExplorer.Left
    
    'Ajusta la Lista
    If fraConsultaAvanzada.Visible Then
       lswExplorer.top = fraConsultaAvanzada.top + fraConsultaAvanzada.Height + 110
       lswExplorer.Height = ArbolExp.Height - lswExplorer.top + 700
    Else
       lswExplorer.Height = ArbolExp.Height
    End If
    
    'Set Height
    If fraNotas.Visible Then
      lswExplorer.Height = lswExplorer.Height - (fraNotas.Height + 100)
      fraNotas.top = lswExplorer.top + lswExplorer.Height + 100
    End If
    

End Sub


Public Sub sbRefrescaArbol()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vNode As Node, strOpciones  As String

strSQL = "select * from CntX_Contabilidades" _
       & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_contabilidad in(SELECT COD_CONTABILIDAD " _
       & "  From CNTX_CONTA_USUARIOS WHERE USUARIO = '" & glogon.Usuario & "')"

Call OpenRecordSet(rs, strSQL, 0)

With ArbolExp
  .Nodes.Clear
  'Crear Root
  Set vNode = .Nodes.Add(, , "ProGrX: Contabilidad", "ProGrX: Contabilidad", "imgRoot")
  'Crear Arbol Inicial
  If Not rs.EOF And Not rs.BOF Then
      If rs!ExpAsientos = 1 Then Call sbCreaNodos("ProGrX: Contabilidad", "Asientos", "imgAsientos2", True)
      If rs!ExpCuentas = 1 Then Call sbCreaNodos("ProGrX: Contabilidad", "Catálogo de Cuentas", "imgCuentas2", True)
      If rs!ExpAreas = 1 Then Call sbCreaNodos("ProGrX: Contabilidad", "Areas de Trabajo", "imgAreas", True)
    '  If rs!ExpPresupuesto = 1 Then Call sbCreaNodos("ProGrX: Contabilidad", "Presupuesto", "imgPresupuesto", True)
      If rs!ExpPlanFijo = 1 Then Call sbCreaNodos("ProGrX: Contabilidad", "Plantillas Asientos Fijos", "imgPlantillas", True)
      If rs!ExpPlanRate = 1 Then Call sbCreaNodos("ProGrX: Contabilidad", "Plantillas Asientos Porcentuales", "imgPlantillas", True)
      If rs!ExpDiferidos = 1 Then Call sbCreaNodos("ProGrX: Contabilidad", "Administracion  de Diferidos", "imgDiferidos", True)
      If rs!ExpMantenimiento = 1 Then Call sbCreaNodos("ProGrX: Contabilidad", "Mantenimiento", "imgOpcion", True)
      
      .Nodes(1).Expanded = True
      lswExplorer.ListItems.Clear
      lswExplorer.ColumnHeaders.Clear
      lswExplorer.ColumnHeaders.Add , , "", 120
  Else
     frmCntX_Seleccionar.Show vbModal
  End If
End With

rs.Close

btnAccion(1).Enabled = False
btnAccion(4).Enabled = False

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean, Optional xkey As String = "N")
Dim nodX As Node, vKey As String
On Error Resume Next

Set nodX = ArbolExp.Nodes.Add(vPadre, tvwChild)
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
    nodX.Image = vImagen
    If xkey = "N" Then
        nodX.Key = vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
    Else
        nodX.Key = xkey
    End If
    
vKey = nodX.Key

If vExpand Then
    Set nodX = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If
    
End Sub

Private Sub lswExplorer_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswExplorer.SortKey = ColumnHeader.Index - 1
  If lswExplorer.SortOrder = 0 Then lswExplorer.SortOrder = 1 Else lswExplorer.SortOrder = 0
  lswExplorer.Sorted = True
End Sub

'Private Sub lswExplorer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'On Error GoTo vError
'
'    lswExplorer.SortKey = ColumnHeader.Index - 1
'
'    If (lswExplorer.SortOrder = lvwAscending) Then
'        lswExplorer.SortOrder = lvwDescending
'    Else
'        lswExplorer.SortOrder = lvwAscending
'    End If
'
'    lswExplorer.Sorted = True
'    Exit Sub
'
'vError:
'   MsgBox "Ocurrió un error al ordenar los datos de la columna seleccionada.", vbCritical
'
'End Sub

Private Sub lswExplorer_DblClick()
Dim frmX As Form, vEncuentra As Boolean

If lswExplorer.ListItems.Count <= 0 Then Exit Sub


If lswExplorer.SelectedItem Is Nothing Then Exit Sub

With lswExplorer.SelectedItem


Select Case .Tag
  Case "Asiento"
    gCntX_Arbol.ArbolActivo = True
    gCntX_Arbol.AsientoTipo = .SubItems(1)
    gCntX_Arbol.AsientoNumr = .Text
    
    vEncuentra = False
    For Each frmX In Forms
       If Trim(frmX.Name) = "frmCntX_Asientos" Then
            frmX.Show
            frmX.sbFormReLoad
            vEncuentra = True
       End If
    Next
    
    If Not vEncuentra Then
        Call sbFormsCall("frmCntX_Asientos")
    End If
    
    gCntX_Arbol.ArbolActivo = False
     
  Case Else
End Select

End With




End Sub

Private Sub lswExplorer_KeyDown(KeyCode As Integer, Shift As Integer)
Dim frmX As Form

For Each frmX In Forms
   If Mid(frmX.Name, 1, 3) = "MDI" Then
        Exit For
   End If
Next

If KeyCode = vbKeyF10 Then
  If UCase(InputBox("Digite la clave de Activación del Módulo Profesional de ProGrX: Contabilidad", "Indique la Clave de Activación")) = "PBN" Then
    frmX.mnuProfesional.Visible = True
  Else
    frmX.mnuProfesional.Visible = False
  End If
End If
End Sub

Private Function fxExtraeCedula(strCedula As String) As String
Dim bolPaso As Boolean, strResultado As String, i As Integer

bolPaso = True
i = 1
strResultado = ""

Do While bolPaso
 If Mid(strCedula, i, 1) <> "-" Then
   strResultado = strResultado & Mid(strCedula, i, 1)
 Else
  bolPaso = False
 End If
 i = i + 1
Loop

fxExtraeCedula = strResultado

End Function

Private Sub lswExplorerX_BeforeLabelEdit(Cancel As Integer)

End Sub



'------------------------------------------------------------------------------------------------------------------------------------
Private Sub chkFechasTodas_Click()
If chkFechasTodas.Value = vbChecked Then
   dtpFechaInicio.Enabled = False
Else
   dtpFechaInicio.Enabled = True
End If

dtpFechaCorte.Enabled = dtpFechaInicio.Enabled
End Sub

Private Sub sbBuscar(Optional vBusca As Integer = 1)

Select Case vBusca
  Case 1 'Tipo de ASiento
     gBusquedas.Columna = "Tipo_Asiento"
     gBusquedas.Orden = "Tipo_Asiento"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     gBusquedas.Consulta = "select Tipo_Asiento,descripcion from CntX_Tipos_Asientos"
     frmBusquedas.Show vbModal
     txtCAsiento = gBusquedas.Resultado
  Case 2 'Descripcion del Tipo de Asiento
     gBusquedas.Columna = "Descripcion"
     gBusquedas.Orden = "Descripcion"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     gBusquedas.Consulta = "select Tipo_Asiento,descripcion from CntX_Tipos_Asientos"
     frmBusquedas.Show vbModal
     txtCAsiento = gBusquedas.Resultado
  Case 3 'Numero de Asiento
     gBusquedas.Columna = "Num_Asiento"
     gBusquedas.Orden = "Num_Asiento"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                       & " and tipo_asiento = '" & txtCAsiento & "'"
                       
                       
  Case 4 'Codigo Unidad
     gBusquedas.Columna = "cod_unidad"
     gBusquedas.Orden = "cod_unidad"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     gBusquedas.Consulta = "select cod_unidad as 'Unidad',Descripcion from CntX_Unidades"
     frmBusquedas.Show vbModal
     txtUnidad.Text = gBusquedas.Resultado
     txtUnidadDesc.Text = gBusquedas.Resultado2
  Case 5 'Descripcion de la Unidad
     gBusquedas.Columna = "Descripcion"
     gBusquedas.Orden = "Descripcion"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     gBusquedas.Consulta = "select cod_unidad as 'Unidad',Descripcion from CntX_Unidades"
     frmBusquedas.Show vbModal
     txtUnidad.Text = gBusquedas.Resultado
     txtUnidadDesc.Text = gBusquedas.Resultado2
                       
                       
  Case 6 'Codigo Centro de Costo
     gBusquedas.Columna = "cod_Centro_Costo"
     gBusquedas.Orden = "cod_Centro_Costo"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     gBusquedas.Consulta = "select cod_Centro_Costo as 'Centro',Descripcion from CntX_Centro_Costos"
     frmBusquedas.Show vbModal
     txtCentroCosto.Text = gBusquedas.Resultado
     txtCentroCostoDesc.Text = gBusquedas.Resultado2
  Case 7 'Descripcion de la Centro de Costo
     gBusquedas.Columna = "Descripcion"
     gBusquedas.Orden = "Descripcion"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     gBusquedas.Consulta = "select cod_Centro_Costo as 'Centro',Descripcion from CntX_Centro_Costos"
     frmBusquedas.Show vbModal
     txtCentroCosto.Text = gBusquedas.Resultado
     txtCentroCostoDesc.Text = gBusquedas.Resultado2
                       
  Case 8 'Divisa
     gBusquedas.Columna = "cod_divisa"
     gBusquedas.Orden = "cod_divisa"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     gBusquedas.Consulta = "select cod_divisa as 'Divisa',Descripcion from CntX_Divisas"
     frmBusquedas.Show vbModal
     txtDivisa.Text = gBusquedas.Resultado
                       
  Case 9, 10 'Cuenta Contable
     frmCntX_ConsultaCuentas.Show vbModal
     
     If vBusca = 9 Then
         txtCuenta.Text = fxCntX_CuentaFormato(True, gCuenta, 0)
     Else
         txtCuentaCorte.Text = fxCntX_CuentaFormato(True, gCuenta, 0)
     End If
                       
End Select
End Sub


Private Sub sbConsulta_Analitico()
Dim strSQL As String, curDebitos As Currency, curCreditos As Currency, lng As Long
Dim rs As New ADODB.Recordset, itmX As ListViewItem

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = "select Top " & txtLineas.Text & " Asi.TIPO_ASIENTO,Asi.NUM_ASIENTO,Asi.FECHA_ASIENTO,Asi.Descripcion,Asi.Fecha_Aplicado,Asi.Referencia" _
       & ",Cta.COD_CUENTA_MASK,Cta.DESCRIPCION as 'CTA_DESC', Det.COD_UNIDAD, Det.Cod_Centro_Costo" _
       & ",Det.COD_DIVISA,Det.TIPO_CAMBIO,Det.DOCUMENTO,Det.DETALLE, Asi.USER_CREA, Asi.USER_APLICA, Asi.USER_AUTORIZA , Asi.USER_MODIFICA" _
       & ",Det.MONTO_CREDITO,Det.MONTO_DEBITO, (Det.MONTO_CREDITO+Det.MONTO_DEBITO) / dbo.fxSys_Tipo_Cambio_Apl(Det.TIPO_CAMBIO) as 'IMPORTE'" _
       & " from CNTX_CUENTAS Cta inner join CNTX_ASIENTOS_DETALLE Det" _
       & " on Cta.COD_CONTABILIDAD = Det.COD_CONTABILIDAD and Cta.COD_CUENTA = Det.COD_CUENTA" _
       & " inner join CNTX_ASIENTOS Asi on Det.COD_CONTABILIDAD = Asi.COD_CONTABILIDAD" _
       & " and Det.TIPO_ASIENTO = Asi.TIPO_ASIENTO and Det.NUM_ASIENTO = Asi.NUM_ASIENTO" _
       & " Where Cta.COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta
       
       
If txtReferencia.Text <> "" Then
   strSQL = strSQL & " and Asi.Referencia like '%" & txtReferencia.Text & "%'"
End If
       
If txtCAsiento.Text <> "" Then
   strSQL = strSQL & " and Det.Tipo_Asiento = '" & txtCAsiento.Text & "'"
End If
       
If txtNAsiento.Text <> "" Then
   strSQL = strSQL & " and Det.Num_Asiento like '%" & txtNAsiento.Text & "%'"
End If
       
If txtUnidad.Text <> "" Then
   strSQL = strSQL & " and Det.Cod_Unidad = '" & txtUnidad.Text & "'"
End If
       
If txtCentroCosto.Text <> "" Then
   strSQL = strSQL & " and Det.Cod_Centro_Costo = '" & txtCentroCosto.Text & "'"
End If
       
If txtDivisa.Text <> "" Then
   strSQL = strSQL & " and Det.cod_Divisa = '" & txtDivisa.Text & "'"
End If
       
If txtDocumento.Text <> "" Then
   strSQL = strSQL & " and Det.documento like '%" & txtDocumento.Text & "%'"
End If
       
If txtDetalle.Text <> "" Then
   strSQL = strSQL & " and Det.detalle like '%" & txtDetalle.Text & "%'"
End If
       
If chkFechasTodas.Value = vbUnchecked Then
   strSQL = strSQL & " and Asi.Fecha_Asiento between  '" & Format(dtpFechaInicio.Value, "yyyy/mm/dd") _
          & " 00:00:00' and '" & Format(dtpFechaCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End If
       
       
'Filtro de Cuentas
If txtCuenta.Text <> "" And txtCuentaCorte.Text = "" Then
   strSQL = strSQL & " and Det.cod_cuenta = '" & fxCntX_CuentaFormato(False, txtCuenta.Text, 0) & "'"
End If
       
If txtCuentaCorte.Text <> "" And txtCuenta.Text = "" Then
   strSQL = strSQL & " and Det.cod_cuenta = '" & fxCntX_CuentaFormato(False, txtCuentaCorte.Text, 0) & "'"
End If
       
If txtCuenta.Text <> "" And txtCuentaCorte.Text <> "" Then
   strSQL = strSQL & " and Det.cod_cuenta between '" & fxCntX_CuentaFormato(False, txtCuenta.Text, 0) _
          & "' and '" & fxCntX_CuentaFormato(False, txtCuentaCorte.Text, 0) & "'"
End If
       
       
Select Case cboMov.Text
    Case "DB"
        strSQL = strSQL & " and Det.Monto_Debito between " & CCur(txtMovimiento(0).Text) _
               & " and " & CCur(txtMovimiento(1).Text)
    
    Case "CR"
        strSQL = strSQL & " and Det.Monto_Credito between " & CCur(txtMovimiento(0).Text) _
               & " and " & CCur(txtMovimiento(1).Text)
    Case "TD"
        strSQL = strSQL & " and (Det.Monto_Debito + Det.Monto_Credito) between " & CCur(txtMovimiento(0).Text) _
               & " and " & CCur(txtMovimiento(1).Text)
End Select
       
       
 With lswExplorer
 .ListItems.Clear
 .ColumnHeaders.Clear
 .ColumnHeaders.Add , , "N°Asiento", 1850
 .ColumnHeaders.Add , , "Tipo", 1250, 2
 .ColumnHeaders.Add , , "Fecha", 1250, 2
 .ColumnHeaders.Add , , "Descripción", 3450
 .ColumnHeaders.Add , , "Debitos", 1850, 1
 .ColumnHeaders.Add , , "Creditos", 1950, 1
 .ColumnHeaders.Add , , "Mayorizado", 1450, 2
 
 
 .ColumnHeaders.Add , , "Cuenta", 2150, vbCenter
 .ColumnHeaders.Add , , "Cuenta Describe", 4150, vbCenter
 .ColumnHeaders.Add , , "Unidad", 1150, vbCenter
 .ColumnHeaders.Add , , "Centro", 1150, vbCenter
 .ColumnHeaders.Add , , "Divisa", 1150, vbCenter
 .ColumnHeaders.Add , , "T.C.", 1450, vbRightJustify
 .ColumnHeaders.Add , , "Importe", 1850, vbRightJustify
 
 .ColumnHeaders.Add , , "Documento", 3150
 .ColumnHeaders.Add , , "Detalle", 4150
 .ColumnHeaders.Add , , "Referencia", 4150
 
 .ColumnHeaders.Add , , "Creado por:", 2500, vbCenter
 .ColumnHeaders.Add , , "Modificado por:", 2500, vbCenter
 .ColumnHeaders.Add , , "Autorizado por:", 2500, vbCenter
 .ColumnHeaders.Add , , "Aplicado por:", 2500, vbCenter
 
 
 Call OpenRecordSet(rs, strSQL)
 
 curDebitos = 0
 curCreditos = 0
  
 ProgressBarX.Visible = True
 ProgressBarX.Max = rs.RecordCount + 5
 
  
 Do While Not rs.EOF
          curDebitos = curDebitos + rs!monto_debito
          curCreditos = curCreditos + rs!monto_credito
          
          ProgressBarX.Value = .ListItems.Count + 1
          
          Set itmX = .ListItems.Add(, , rs!Num_Asiento)
              itmX.SubItems(1) = rs!Tipo_Asiento
              itmX.SubItems(2) = rs!fecha_asiento
              itmX.SubItems(3) = rs!Descripcion
              itmX.SubItems(4) = Format(rs!monto_debito, "Standard")
              itmX.SubItems(5) = Format(rs!monto_credito, "Standard")
              itmX.SubItems(6) = IIf(IsNull(rs!Fecha_Aplicado), "NO", "SI")
          
              itmX.SubItems(7) = rs!Cod_Cuenta_Mask & ""
              itmX.SubItems(8) = rs!Cta_Desc & ""
              itmX.SubItems(9) = rs!Cod_Unidad
              itmX.SubItems(10) = rs!Cod_Centro_Costo
              itmX.SubItems(11) = rs!COD_DIVISA
              itmX.SubItems(12) = rs!TIPO_CAMBIO
              itmX.SubItems(13) = Format(rs!Importe, "Standard")
              
              itmX.SubItems(14) = rs!Documento
              itmX.SubItems(15) = rs!Detalle & ""
              itmX.SubItems(16) = rs!Referencia & ""
           
           
              itmX.SubItems(17) = rs!user_crea & ""
              itmX.SubItems(18) = rs!user_modifica & ""
              itmX.SubItems(19) = rs!user_autoriza & ""
              itmX.SubItems(20) = rs!user_aplica & ""
           
           
           If IsNull(rs!Fecha_Aplicado) Then
              itmX.Bold = True
              itmX.TextBackColor = RGB(252, 243, 207)
            End If
           
           itmX.Tag = "Asiento"

   
   rs.MoveNext
 Loop
 rs.Close
 
  Set itmX = .ListItems.Add(, , "")
      itmX.SubItems(4) = "___________"
      itmX.SubItems(5) = "___________"

  Set itmX = .ListItems.Add(, , "")
      itmX.SubItems(4) = Format(curDebitos, "Standard")
      itmX.SubItems(5) = Format(curCreditos, "Standard")
      itmX.SubItems(6) = "Lineas: " & Format(.ListItems.Count - 1, "###,###,##0")
 
End With
 
ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub txtCAsiento_Change()
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select descripcion from CntX_Tipos_Asientos where cod_contabilidad = " _
       & gCntX_Parametros.CodigoConta & " and tipo_asiento = '" _
       & txtCAsiento.Text & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  txtDAsiento = rs!Descripcion
End If
rs.Close
End Sub


Private Sub txtCAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDAsiento.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(1)
End Sub

Private Sub txtCAsiento_LostFocus()
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select descripcion from CntX_Tipos_Asientos where cod_contabilidad = " _
       & gCntX_Parametros.CodigoConta & " and tipo_asiento = '" _
       & txtCAsiento.Text & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  txtDAsiento = rs!Descripcion
End If
rs.Close

End Sub

Private Sub txtCentroCosto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCentroCostoDesc.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(6)
End Sub

Private Sub txtCentroCostoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDivisa.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(7)
End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaCorte.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(9)
End Sub

Private Sub txtCuenta_LostFocus()
If Len(txtCuenta.Text) > 0 Then
    txtCuenta.Text = fxCntX_CuentaFormato(True, txtCuenta.Text, 0)
End If
End Sub


Private Sub txtCuentaCorte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCAsiento.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(10)
End Sub

Private Sub txtCuentaCorte_LostFocus()
If Len(txtCuentaCorte.Text) > 0 Then
    txtCuentaCorte.Text = fxCntX_CuentaFormato(True, txtCuentaCorte.Text, 0)
End If
End Sub


Private Sub txtDAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNAsiento.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(2)
End Sub

Private Sub txtDivisa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuenta.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(8)
End Sub


Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDetalle.SetFocus
End Sub

Private Sub txtMovimiento_GotFocus(Index As Integer)
On Error GoTo vError

txtMovimiento(Index).Text = CCur(txtMovimiento(Index).Text)

vError:

End Sub

Private Sub txtMovimiento_LostFocus(Index As Integer)
On Error GoTo vError

txtMovimiento(Index).Text = Format(CCur(txtMovimiento(Index).Text), "Standard")

vError:
End Sub

Private Sub txtNAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUnidad.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(3)
End Sub

Private Sub txtUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUnidadDesc.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(4)
End Sub

Private Sub txtUnidadDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(5)
End Sub




