VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmTES_Conciliacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conciliación Bancaria"
   ClientHeight    =   8235
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   14025
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   480
      Top             =   360
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7092
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   13812
      _Version        =   1441793
      _ExtentX        =   24363
      _ExtentY        =   12509
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   5
      Item(0).Caption =   "Historial"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "gHistorial"
      Item(1).Caption =   "Resumen"
      Item(1).ControlCount=   5
      Item(1).Control(0)=   "gbAcciones"
      Item(1).Control(1)=   "gbArchivo"
      Item(1).Control(2)=   "gbResumen(0)"
      Item(1).Control(3)=   "gbResumen(1)"
      Item(1).Control(4)=   "gbResumen(2)"
      Item(2).Caption =   "Resultados"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "gResultados"
      Item(2).Control(1)=   "gbResultados"
      Item(2).Control(2)=   "gbAutoRegistro"
      Item(3).Caption =   "Conciliación"
      Item(3).ControlCount=   10
      Item(3).Control(0)=   "gbConciliaFuente"
      Item(3).Control(1)=   "feMov_Conciliado"
      Item(3).Control(2)=   "Label1(1)"
      Item(3).Control(3)=   "feMov_Pendiente"
      Item(3).Control(4)=   "Label1(0)"
      Item(3).Control(5)=   "btnConcilia_Aplicar"
      Item(3).Control(6)=   "tcConcilia"
      Item(3).Control(7)=   "feMov_SelMonto"
      Item(3).Control(8)=   "Label1(24)"
      Item(3).Control(9)=   "feMov_SelCasos"
      Item(4).Caption =   "Informes"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "gbInformes"
      Begin XtremeSuiteControls.GroupBox gbInformes 
         Height          =   5412
         Left            =   -67240
         TabIndex        =   84
         Top             =   720
         Visible         =   0   'False
         Width           =   7812
         _Version        =   1441793
         _ExtentX        =   13779
         _ExtentY        =   9546
         _StockProps     =   79
         Caption         =   "Informes de la Conciliación "
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.RadioButton rbInformes 
            Height          =   372
            Index           =   0
            Left            =   720
            TabIndex        =   86
            Top             =   600
            Width           =   6012
            _Version        =   1441793
            _ExtentX        =   10604
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Informe de Conciliación"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnPeriodoInforme 
            Height          =   645
            Left            =   6000
            TabIndex        =   85
            Top             =   4200
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   1138
            _StockProps     =   79
            Caption         =   "Informes "
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
            Picture         =   "frmTES_Conciliacion.frx":0000
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.RadioButton rbInformes 
            Height          =   372
            Index           =   1
            Left            =   720
            TabIndex        =   87
            Top             =   960
            Width           =   6012
            _Version        =   1441793
            _ExtentX        =   10604
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Bancos: Pendientes de Conciliación"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.RadioButton rbInformes 
            Height          =   372
            Index           =   2
            Left            =   720
            TabIndex        =   88
            Top             =   1320
            Width           =   6012
            _Version        =   1441793
            _ExtentX        =   10604
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Bancos: Movimientos del Periodo"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.RadioButton rbInformes 
            Height          =   372
            Index           =   3
            Left            =   720
            TabIndex        =   99
            Top             =   1680
            Width           =   6012
            _Version        =   1441793
            _ExtentX        =   10604
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Libros: Pendientes de Conciliación"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.RadioButton rbInformes 
            Height          =   372
            Index           =   4
            Left            =   720
            TabIndex        =   100
            Top             =   2040
            Width           =   6012
            _Version        =   1441793
            _ExtentX        =   10604
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Libros: Movimientos del Periodo"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.RadioButton rbInformes 
            Height          =   372
            Index           =   5
            Left            =   720
            TabIndex        =   101
            Top             =   2400
            Width           =   6012
            _Version        =   1441793
            _ExtentX        =   10604
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Libros: Transferencias y Lotes"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.DateTimePicker dtpInforme 
            Height          =   375
            Index           =   0
            Left            =   2520
            TabIndex        =   115
            Top             =   4320
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   661
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
         Begin XtremeSuiteControls.DateTimePicker dtpInforme 
            Height          =   375
            Index           =   1
            Left            =   3960
            TabIndex        =   116
            Top             =   4320
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   661
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   375
            Left            =   720
            TabIndex        =   114
            Top             =   4320
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Fechas de Movimientos"
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
            Alignment       =   1
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   375
            Left            =   4920
            TabIndex        =   113
            Top             =   2040
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Label2"
         End
      End
      Begin XtremeSuiteControls.GroupBox gbAutoRegistro 
         Height          =   972
         Left            =   -70000
         TabIndex        =   74
         Top             =   6120
         Visible         =   0   'False
         Width           =   13812
         _Version        =   1441793
         _ExtentX        =   24363
         _ExtentY        =   1714
         _StockProps     =   79
         Caption         =   "Registro automático de casos seleccionados en el auxiliar de bancos y contabilidad"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnAutoRegistro 
            Height          =   330
            Left            =   10920
            TabIndex        =   75
            Top             =   480
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Auto Registro"
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
         End
         Begin XtremeSuiteControls.FlatEdit feAR_Monto 
            Height          =   312
            Left            =   9120
            TabIndex        =   76
            Top             =   480
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2984
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feAR_Casos 
            Height          =   312
            Left            =   8160
            TabIndex        =   77
            Top             =   480
            Width           =   972
            _Version        =   1441793
            _ExtentX        =   1714
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feAR_Cuenta 
            Height          =   312
            Left            =   720
            TabIndex        =   80
            Top             =   480
            Width           =   2052
            _Version        =   1441793
            _ExtentX        =   3619
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feAR_CuentaDesc 
            Height          =   312
            Left            =   2760
            TabIndex        =   79
            Top             =   480
            Width           =   4572
            _Version        =   1441793
            _ExtentX        =   8064
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkAutoReg_Contabilidad 
            Height          =   252
            Left            =   4200
            TabIndex        =   102
            Top             =   240
            Width           =   3132
            _Version        =   1441793
            _ExtentX        =   5524
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Registro en Contabilidad?  "
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            Appearance      =   16
            Value           =   1
            Alignment       =   1
         End
         Begin XtremeSuiteControls.PushButton btnPendiente 
            Height          =   330
            Left            =   12360
            TabIndex        =   103
            ToolTipText     =   "Pone la Transacción como Pendiente de Conciliación"
            Top             =   480
            Width           =   1332
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Pendiente"
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
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   22
            Left            =   -120
            TabIndex        =   81
            Top             =   480
            Width           =   612
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   21
            Left            =   6960
            TabIndex        =   78
            Top             =   480
            Width           =   972
         End
      End
      Begin XtremeSuiteControls.TabControl tcConcilia 
         Height          =   5172
         Left            =   -70000
         TabIndex        =   71
         Top             =   1440
         Visible         =   0   'False
         Width           =   13812
         _Version        =   1441793
         _ExtentX        =   24363
         _ExtentY        =   9123
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
         Appearance      =   10
         Color           =   32
         ItemCount       =   2
         Item(0).Caption =   "Asignar"
         Item(0).ControlCount=   7
         Item(0).Control(0)=   "gConcilia"
         Item(0).Control(1)=   "chkConciliaPendientes"
         Item(0).Control(2)=   "chkConciliaFiltroMontos"
         Item(0).Control(3)=   "chkConciliaFiltroFechas"
         Item(0).Control(4)=   "dtpConciliaInicio"
         Item(0).Control(5)=   "dtpConciliaCorte"
         Item(0).Control(6)=   "chkConciliaTodos"
         Item(1).Caption =   "Conciliados"
         Item(1).ControlCount=   4
         Item(1).Control(0)=   "lswConcilia"
         Item(1).Control(1)=   "lswConcilia_Lote"
         Item(1).Control(2)=   "Label1(26)"
         Item(1).Control(3)=   "btnConciliaReversa"
         Begin XtremeSuiteControls.ListView lswConcilia 
            Height          =   2892
            Left            =   -69880
            TabIndex        =   73
            Top             =   360
            Visible         =   0   'False
            Width           =   13572
            _Version        =   1441793
            _ExtentX        =   23939
            _ExtentY        =   5101
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
         End
         Begin XtremeSuiteControls.ListView lswConcilia_Lote 
            Height          =   1092
            Left            =   -69880
            TabIndex        =   94
            Top             =   3960
            Visible         =   0   'False
            Width           =   13572
            _Version        =   1441793
            _ExtentX        =   23939
            _ExtentY        =   1926
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
         End
         Begin XtremeSuiteControls.CheckBox chkConciliaPendientes 
            Height          =   252
            Left            =   11520
            TabIndex        =   105
            Top             =   360
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Mostrar Pendientes"
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
            Appearance      =   14
         End
         Begin XtremeSuiteControls.CheckBox chkConciliaFiltroMontos 
            Height          =   252
            Left            =   8760
            TabIndex        =   106
            Top             =   360
            Width           =   2772
            _Version        =   1441793
            _ExtentX        =   4890
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Filtrar Montos Aproximados"
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
            Appearance      =   14
         End
         Begin FPSpreadADO.fpSpread gConcilia 
            Height          =   4332
            Left            =   120
            TabIndex        =   72
            Top             =   720
            Width           =   13572
            _Version        =   524288
            _ExtentX        =   23939
            _ExtentY        =   7641
            _StockProps     =   64
            BackColorStyle  =   1
            BorderStyle     =   0
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   9
            ScrollBars      =   2
            SpreadDesigner  =   "frmTES_Conciliacion.frx":0707
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.PushButton btnConciliaReversa 
            Height          =   330
            Left            =   -58120
            TabIndex        =   104
            Top             =   3360
            Visible         =   0   'False
            Width           =   1812
            _Version        =   1441793
            _ExtentX        =   3196
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Des - Conciliar"
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
         End
         Begin XtremeSuiteControls.CheckBox chkConciliaFiltroFechas 
            Height          =   252
            Left            =   6840
            TabIndex        =   107
            Top             =   360
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Filtar Fechas"
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
            Appearance      =   14
         End
         Begin XtremeSuiteControls.DateTimePicker dtpConciliaInicio 
            Height          =   300
            Left            =   3960
            TabIndex        =   108
            Top             =   360
            Width           =   1332
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   529
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
         Begin XtremeSuiteControls.DateTimePicker dtpConciliaCorte 
            Height          =   300
            Left            =   5280
            TabIndex        =   109
            Top             =   360
            Width           =   1332
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   529
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
         Begin XtremeSuiteControls.CheckBox chkConciliaTodos 
            Height          =   255
            Left            =   960
            TabIndex        =   112
            Top             =   360
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Todos"
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
            Appearance      =   14
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Lote:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   26
            Left            =   -69880
            TabIndex        =   95
            Top             =   3720
            Visible         =   0   'False
            Width           =   1332
         End
      End
      Begin XtremeSuiteControls.GroupBox gbConciliaFuente 
         Height          =   972
         Left            =   -69880
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   13572
         _Version        =   1441793
         _ExtentX        =   23939
         _ExtentY        =   1714
         _StockProps     =   79
         Caption         =   "Movimiento a Conciliar"
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit feMov_Descripcion 
            Height          =   330
            Left            =   6600
            TabIndex        =   48
            Top             =   600
            Width           =   3852
            _Version        =   1441793
            _ExtentX        =   6794
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feMov_Importe 
            Height          =   330
            Left            =   5040
            TabIndex        =   49
            Top             =   600
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feMov_Documento 
            Height          =   330
            Left            =   3480
            TabIndex        =   50
            Top             =   600
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feMov_Fecha 
            Height          =   330
            Left            =   1320
            TabIndex        =   58
            Top             =   600
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feMov_Tipo 
            Height          =   330
            Left            =   2520
            TabIndex        =   59
            Top             =   600
            Width           =   972
            _Version        =   1441793
            _ExtentX        =   1714
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feMov_Estado 
            Height          =   330
            Left            =   10440
            TabIndex        =   60
            Top             =   600
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnConcilia_Buscar 
            Height          =   312
            Left            =   13200
            TabIndex        =   66
            Top             =   600
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "..."
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
         End
         Begin XtremeSuiteControls.FlatEdit feMov_Id 
            Height          =   330
            Left            =   120
            TabIndex        =   82
            Top             =   600
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboMovFiltro 
            Height          =   330
            Left            =   11640
            TabIndex        =   98
            Top             =   600
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
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
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Buscar:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   27
            Left            =   11640
            TabIndex        =   97
            Top             =   360
            Width           =   1212
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Estado:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   17
            Left            =   10440
            TabIndex        =   37
            Top             =   360
            Width           =   1212
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   16
            Left            =   6600
            TabIndex        =   36
            Top             =   360
            Width           =   1572
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Importe:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   15
            Left            =   5040
            TabIndex        =   35
            Top             =   360
            Width           =   1572
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Documento:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   14
            Left            =   3480
            TabIndex        =   34
            Top             =   360
            Width           =   1572
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   13
            Left            =   2640
            TabIndex        =   33
            Top             =   360
            Width           =   1572
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   12
            Left            =   1320
            TabIndex        =   32
            Top             =   360
            Width           =   1572
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Id Referencia:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   23
            Left            =   120
            TabIndex        =   83
            Top             =   360
            Width           =   1092
         End
      End
      Begin XtremeSuiteControls.GroupBox gbResultados 
         Height          =   732
         Left            =   -69760
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   13212
         _Version        =   1441793
         _ExtentX        =   23304
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Movimientos a Visualizar"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cboUbicacion 
            Height          =   312
            Left            =   1200
            TabIndex        =   30
            Top             =   360
            Width           =   2652
            _Version        =   1441793
            _ExtentX        =   4683
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
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
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboTipoDoc 
            Height          =   312
            Left            =   3840
            TabIndex        =   55
            Top             =   360
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
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
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnCasos_Buscar 
            Height          =   312
            Index           =   0
            Left            =   7800
            TabIndex        =   56
            Top             =   360
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "..."
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
         End
         Begin XtremeSuiteControls.ComboBox cboCasos_Estados 
            Height          =   312
            Left            =   5760
            TabIndex        =   57
            Top             =   360
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
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
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnCasos_Buscar 
            Height          =   312
            Index           =   1
            Left            =   8160
            TabIndex        =   67
            Top             =   360
            Width           =   972
            _Version        =   1441793
            _ExtentX        =   1714
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Exportar"
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
         End
         Begin XtremeSuiteControls.FlatEdit feResTotalMonto 
            Height          =   312
            Left            =   11520
            TabIndex        =   68
            Top             =   360
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2984
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feResTotalCasos 
            Height          =   312
            Left            =   10560
            TabIndex        =   70
            Top             =   360
            Width           =   972
            _Version        =   1441793
            _ExtentX        =   1714
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   20
            Left            =   9360
            TabIndex        =   69
            Top             =   360
            Width           =   972
         End
      End
      Begin XtremeSuiteControls.GroupBox gbAcciones 
         Height          =   975
         Left            =   -68440
         TabIndex        =   25
         Top             =   5880
         Visible         =   0   'False
         Width           =   9855
         _Version        =   1441793
         _ExtentX        =   17383
         _ExtentY        =   1720
         _StockProps     =   79
         Caption         =   "Acciones para el periodo en conciliación"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnPeriodoCerrar 
            Height          =   528
            Left            =   1200
            TabIndex        =   26
            Top             =   360
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   931
            _StockProps     =   79
            Caption         =   "Cerrar Proceso Conciliatorio"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmTES_Conciliacion.frx":0F05
         End
         Begin XtremeSuiteControls.PushButton btnPeriodoInicializa 
            Height          =   528
            Left            =   8400
            TabIndex        =   27
            Top             =   360
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   931
            _StockProps     =   79
            Caption         =   "Inicializa"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmTES_Conciliacion.frx":18F1
         End
         Begin XtremeSuiteControls.PushButton btnPeriodoActualiza 
            Height          =   528
            Left            =   6960
            TabIndex        =   28
            Top             =   360
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   931
            _StockProps     =   79
            Caption         =   "Actualizar"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmTES_Conciliacion.frx":227E
         End
         Begin XtremeSuiteControls.PushButton btnConcilia1a1 
            Height          =   528
            Index           =   0
            Left            =   3960
            TabIndex        =   110
            Top             =   360
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   931
            _StockProps     =   79
            Caption         =   "Concilia 1:1 Bancos"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmTES_Conciliacion.frx":2A40
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.PushButton btnConcilia1a1 
            Height          =   528
            Index           =   1
            Left            =   5520
            TabIndex        =   111
            Top             =   360
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   931
            _StockProps     =   79
            Caption         =   "Concilia 1:1 Libros"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmTES_Conciliacion.frx":3167
            ImageAlignment  =   4
         End
      End
      Begin XtremeSuiteControls.GroupBox gbArchivo 
         Height          =   1335
         Left            =   -68440
         TabIndex        =   21
         Top             =   4440
         Visible         =   0   'False
         Width           =   9855
         _Version        =   1441793
         _ExtentX        =   17383
         _ExtentY        =   2355
         _StockProps     =   79
         Caption         =   "Carga Archivo de Movimientos del Banco"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnArchivo 
            Height          =   432
            Index           =   0
            Left            =   8400
            TabIndex        =   24
            ToolTipText     =   "Busca Archivo de Carga"
            Top             =   480
            Width           =   492
            _Version        =   1441793
            _ExtentX        =   868
            _ExtentY        =   762
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
            Picture         =   "frmTES_Conciliacion.frx":388E
         End
         Begin XtremeSuiteControls.PushButton btnArchivo 
            Height          =   432
            Index           =   1
            Left            =   8880
            TabIndex        =   38
            ToolTipText     =   "Carga Archivo"
            Top             =   480
            Width           =   492
            _Version        =   1441793
            _ExtentX        =   868
            _ExtentY        =   762
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
            Picture         =   "frmTES_Conciliacion.frx":3F8E
         End
         Begin XtremeSuiteControls.PushButton btnArchivo 
            Height          =   432
            Index           =   2
            Left            =   9360
            TabIndex        =   39
            ToolTipText     =   "Información del Archivo a Cargar"
            Top             =   480
            Width           =   492
            _Version        =   1441793
            _ExtentX        =   868
            _ExtentY        =   762
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
            Picture         =   "frmTES_Conciliacion.frx":46A7
         End
         Begin XtremeSuiteControls.FlatEdit feArchivo 
            Height          =   432
            Left            =   1920
            TabIndex        =   23
            Top             =   480
            Width           =   6372
            _Version        =   1441793
            _ExtentX        =   11239
            _ExtentY        =   762
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeShortcutBar.ShortcutCaption scStatus 
            Height          =   255
            Left            =   1920
            TabIndex        =   117
            Top             =   960
            Visible         =   0   'False
            Width           =   6375
            _Version        =   1441793
            _ExtentX        =   11245
            _ExtentY        =   450
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
            Alignment       =   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Seleccione el Archivo:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   492
            Index           =   10
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   1572
         End
      End
      Begin XtremeSuiteControls.GroupBox gbResumen 
         Height          =   2292
         Index           =   0
         Left            =   -68440
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   4692
         _Version        =   1441793
         _ExtentX        =   8276
         _ExtentY        =   4043
         _StockProps     =   79
         Caption         =   "Libros"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit feLibrosSaldoActual 
            Height          =   312
            Left            =   2400
            TabIndex        =   9
            Top             =   480
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feBancosNC 
            Height          =   312
            Left            =   2400
            TabIndex        =   10
            Top             =   1320
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feBancosND 
            Height          =   312
            Left            =   2400
            TabIndex        =   11
            Top             =   960
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feLibrosSaldoConcilia 
            Height          =   312
            Left            =   2400
            TabIndex        =   12
            Top             =   1800
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Saldo a Conciliar"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   5
            Left            =   240
            TabIndex        =   8
            Top             =   1800
            Width           =   1932
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "(-) Notas de Débito"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   4
            Left            =   240
            TabIndex        =   7
            Top             =   960
            Width           =   1932
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "(+) Notas de Créditos"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   6
            Top             =   1320
            Width           =   1932
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Saldo Actual"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   5
            Top             =   480
            Width           =   1932
         End
      End
      Begin FPSpreadADO.fpSpread gHistorial 
         Height          =   6252
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   12852
         _Version        =   524288
         _ExtentX        =   22669
         _ExtentY        =   11028
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         ScrollBars      =   2
         SpreadDesigner  =   "frmTES_Conciliacion.frx":4DC0
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread gResultados 
         Height          =   4572
         Left            =   -70000
         TabIndex        =   2
         Top             =   1440
         Visible         =   0   'False
         Width           =   13692
         _Version        =   524288
         _ExtentX        =   24151
         _ExtentY        =   8064
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   11
         ScrollBars      =   2
         SpreadDesigner  =   "frmTES_Conciliacion.frx":560E
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.GroupBox gbResumen 
         Height          =   2292
         Index           =   1
         Left            =   -63160
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   4692
         _Version        =   1441793
         _ExtentX        =   8276
         _ExtentY        =   4043
         _StockProps     =   79
         Caption         =   "Estado de Cuenta"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit feBancosSaldoActual 
            Height          =   312
            Left            =   2280
            TabIndex        =   17
            Top             =   480
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feDepositosTransito 
            Height          =   312
            Left            =   2280
            TabIndex        =   18
            Top             =   960
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feChequesNoCobrados 
            Height          =   312
            Left            =   2280
            TabIndex        =   19
            Top             =   1320
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feBancosSaldoConcilia 
            Height          =   312
            Left            =   2280
            TabIndex        =   20
            Top             =   1800
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnBancos_Saldo_Update 
            Height          =   312
            Left            =   4200
            TabIndex        =   96
            ToolTipText     =   "Reemplaza el Saldo en Bancos"
            Top             =   480
            Width           =   492
            _Version        =   1441793
            _ExtentX        =   868
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Upd"
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Saldo a Conciliar"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   9
            Left            =   120
            TabIndex        =   16
            Top             =   1800
            Width           =   1932
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "(-) Cheques Girados, no cobrados"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   8
            Left            =   0
            TabIndex        =   15
            Top             =   1320
            Width           =   2052
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "(+) Depósitos en Tránsito"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   7
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   1932
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Saldo Actual"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   6
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   1932
         End
      End
      Begin XtremeSuiteControls.GroupBox gbResumen 
         Height          =   1332
         Index           =   2
         Left            =   -68440
         TabIndex        =   51
         Top             =   2880
         Visible         =   0   'False
         Width           =   9852
         _Version        =   1441793
         _ExtentX        =   17378
         _ExtentY        =   2350
         _StockProps     =   79
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnGuardar 
            Height          =   408
            Left            =   8520
            TabIndex        =   52
            Top             =   240
            Width           =   1332
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   720
            _StockProps     =   79
            Caption         =   "Guardar"
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
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmTES_Conciliacion.frx":5F08
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.FlatEdit feSaldoDiferencia 
            Height          =   312
            Left            =   2400
            TabIndex        =   53
            Top             =   336
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feNotas 
            Height          =   552
            Left            =   2400
            TabIndex        =   93
            Top             =   720
            Width           =   7452
            _Version        =   1441793
            _ExtentX        =   13144
            _ExtentY        =   974
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label1 
            Caption         =   "Notas:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   25
            Left            =   1080
            TabIndex        =   92
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Diferencia por Conciliar:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Index           =   11
            Left            =   1080
            TabIndex        =   54
            Top             =   240
            Width           =   1335
         End
      End
      Begin XtremeSuiteControls.PushButton btnConcilia_Aplicar 
         Height          =   330
         Left            =   -58120
         TabIndex        =   61
         Top             =   6720
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Conciliar"
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
      End
      Begin XtremeSuiteControls.FlatEdit feMov_Conciliado 
         Height          =   312
         Left            =   -68200
         TabIndex        =   62
         Top             =   6720
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit feMov_Pendiente 
         Height          =   312
         Left            =   -64840
         TabIndex        =   64
         Top             =   6720
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit feMov_SelMonto 
         Height          =   312
         Left            =   -60880
         TabIndex        =   89
         Top             =   6720
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit feMov_SelCasos 
         Height          =   312
         Left            =   -61480
         TabIndex        =   91
         Top             =   6720
         Visible         =   0   'False
         Width           =   612
         _Version        =   1441793
         _ExtentX        =   1080
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Seleccionados:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   24
         Left            =   -63160
         TabIndex        =   90
         Top             =   6720
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Pendiente:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   0
         Left            =   -66400
         TabIndex        =   65
         Top             =   6720
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Conciliado:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   1
         Left            =   -70360
         TabIndex        =   63
         Top             =   6720
         Visible         =   0   'False
         Width           =   1932
      End
   End
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   312
      Left            =   2760
      TabIndex        =   40
      Top             =   240
      Width           =   7692
      _Version        =   1441793
      _ExtentX        =   13573
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit feAnio 
      Height          =   312
      Left            =   2760
      TabIndex        =   43
      Top             =   600
      Width           =   732
      _Version        =   1441793
      _ExtentX        =   1291
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
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
   Begin XtremeSuiteControls.FlatEdit fePeriodo 
      Height          =   312
      Left            =   4080
      TabIndex        =   44
      Top             =   600
      Width           =   4692
      _Version        =   1441793
      _ExtentX        =   8276
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   16711680
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
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit feMes 
      Height          =   312
      Left            =   3480
      TabIndex        =   45
      Top             =   600
      Width           =   612
      _Version        =   1441793
      _ExtentX        =   1080
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
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
   Begin XtremeSuiteControls.FlatEdit fePeriodoEstado 
      Height          =   312
      Left            =   8760
      TabIndex        =   46
      Top             =   600
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnPeriodoNuevo 
      Height          =   312
      Left            =   10560
      TabIndex        =   47
      Top             =   600
      Width           =   372
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "..."
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
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   19
      Left            =   1560
      TabIndex        =   42
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
      DataField       =   "Banco"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   18
      Left            =   1560
      TabIndex        =   41
      Top             =   240
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmTES_Conciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub sbRefrescaInformacion()
Dim strResultado As String

On Error GoTo vError

vPaso = True

feAnio.Text = Val(feAnio.Text)
  
  Select Case Val(feMes.Text)
    Case 1
        strResultado = "ENERO DEL " & feAnio.Text
    Case 2
        strResultado = "FEBRERO DEL " & feAnio.Text
    Case 3
        strResultado = "MARZO DEL " & feAnio.Text
    Case 4
        strResultado = "ABRIL DEL " & feAnio.Text
    Case 5
        strResultado = "MAYO DEL " & feAnio.Text
    Case 6
        strResultado = "JUNIO DEL " & feAnio.Text
    Case 7
        strResultado = "JULIO DEL " & feAnio.Text
    Case 8
        strResultado = "AGOSTO DEL " & feAnio.Text
    Case 9
        strResultado = "SETIEMBRE DEL " & feAnio.Text
    Case 10
        strResultado = "OCTUBRE DEL " & feAnio.Text
    Case 11
        strResultado = "NOVIEMBRE DEL " & feAnio.Text
    Case 12
        strResultado = "DICIEMBRE DEL " & feAnio.Text
  End Select

  fePeriodo.Text = strResultado

vPaso = False
Exit Sub

vError:
    vPaso = False

End Sub

Private Sub sbArchivoBusca()


With frmContenedor.CD
    .InitDir = "C:\"
    .DialogTitle = "Localice Archivo de Depósitos del Banco [Microsoft EXCEL]"
    .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
    .ShowOpen

    If .FileName = "" Then
        MsgBox "Archivo no válido...", vbExclamation
        Exit Sub
    End If

    If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
        'Ok
    Else
        MsgBox "La Extensión del Archivo no es válido...", vbExclamation
        Exit Sub
    End If

    
    feArchivo.Text = .FileName
End With

End Sub


Private Sub sbArchivoCarga()
Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset

Dim i As Integer, iCampos As Integer, vExiste As Integer
Dim vFecha As Date, vDocumento As String, vImporte As Currency, vDescripcion As String
Dim vCedula As String, vNombre As String, vTipo As String

Dim curMonto As Currency, lCasos As Long

On Error GoTo vError

If feArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

'If feBancoDesc.Text = "" Then
'    MsgBox "No existe ninguna cuenta Bancaria, no se puede procesar el archivo...", vbCritical
'    Exit Sub
'End If

Me.MousePointer = vbHourglass


curMonto = 0
lCasos = 0 'Total

scStatus.Visible = True
scStatus.Caption = "Cargdo archivo, espere!"

Set rsExcel = Excel_Load(feArchivo.Text, "Import")

'Verifica Estructura del Archivo

iCampos = 0
For i = 0 To rsExcel.Fields.Count - 1
   Select Case UCase(rsExcel.Fields(i).Name)
      Case "DOCUMENTO", "FECHA", "IMPORTE", "DESCRIPCION", "TIPO", "SALDO"
        iCampos = iCampos + 1
      Case Else
      
   End Select
Next i

If iCampos < 6 Then
   Me.MousePointer = vbDefault
   MsgBox "1. No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "2. Los campos son Fecha, Tipo, Documento, Importe, Descripcion y Saldo", vbExclamation
   Exit Sub
End If



Dim vCount As Long

vCount = 0
strSQL = ""

Do While Not rsExcel.EOF
 vDocumento = Trim(rsExcel!Documento & "")
 vFecha = rsExcel!fecha
 vImporte = rsExcel!Importe
 vDescripcion = rsExcel!Descripcion & ""
 vTipo = rsExcel!Tipo
 vCount = vCount + 1

 If vImporte < 0 Then
    vTipo = "D"
 End If
 
 scStatus.Caption = "Cargando archivo...Registro No." & vCount
 DoEvents
 
 strSQL = strSQL & Space(10) & "exec spTes_Concilia_Banco_Mov " & cboBanco.ItemData(cboBanco.ListIndex) & ",'" & Format(vFecha, "yyyy/mm/dd") _
        & "','" & Mid(vDocumento, 1, 30) & "','" & Mid(vTipo, 1, 1) & "'," & Abs(vImporte) & ",'" & Mid(vDescripcion, 1, 150) _
        & "',0,'" & glogon.Usuario & "'"
 'Inserta Valores
 If Len(strSQL) > 25000 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
 End If
 
rsExcel.MoveNext
Loop
rsExcel.Close
    
 If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
 End If
    

scStatus.Caption = "Realizando la conciliación automática, espere!"
DoEvents
 
'Concilia y Actualiza
strSQL = "exec spTes_Concilia_Automatica " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)
                

scStatus.Caption = "Actualizando los resultados de la conciliación, espere!"
 DoEvents
 
strSQL = "exec spTes_Concilia_Periodo_Actualiza " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)
                
       
scStatus.Visible = False

'Totales
Me.MousePointer = vbDefault
MsgBox "Información Cargada Satisfactoriamente", vbInformation

Call sbPeriodo_Consulta(cboBanco.ItemData(cboBanco.ListIndex), feAnio.Text, feMes.Text)


Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub btnArchivo_Click(Index As Integer)
Dim vMensaje As String
  


Select Case Index
  
  Case 0 'buscar
        feArchivo.Text = ""
        
        If fePeriodoEstado.Text <> "Abierto" Then
          MsgBox "El Periodo no ha sido declarado o se encuentra Cerrado, verifique!", vbExclamation
          Exit Sub
        End If
        Call sbArchivoBusca

  Case 1 'Carga
        If fePeriodoEstado.Text <> "Abierto" Then
          MsgBox "El Periodo no ha sido declarado o se encuentra Cerrado, verifique!", vbExclamation
          Exit Sub
        End If
       
       Call sbArchivoCarga
       
'       Call btnPeriodoActualiza

  Case 2 'Info
     vMensaje = "-> FORMATO DEL ARCHIVO DE CARGA <-" & vbCrLf & vbCrLf _
              & " 1. Microsoft Excel" & vbCrLf _
              & " 2. Nombre de la Hoja.: Import" & vbCrLf _
              & " 3. Columnas.: FECHA, TIPO, DOCUMENTO, IMPORTE, DESCRIPCION, SALDO"
     
     MsgBox vMensaje, vbInformation
     
     
End Select
End Sub

Private Sub sbResultados_Consulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curTotal As Currency, i As Long

On Error GoTo vError

Me.MousePointer = vbHourglass
curTotal = 0

vPaso = True

strSQL = "exec spTes_Concilia_Periodo_Resultados " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'" & cboUbicacion.ItemData(cboUbicacion.ListIndex) _
       & "','" & cboTipoDoc.ItemData(cboTipoDoc.ListIndex) & "','" & Mid(cboCasos_Estados.Text, 1, 1) & "'"
Call OpenRecordSet(rs, strSQL, 0)

With gResultados

    .MaxRows = 0
    
    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      
      .col = 3
      .Text = CStr(rs!Id)
      .col = 4
      .Text = Format(rs!fecha, "dd/mm/yyyy")
      .col = 5
      .Text = rs!Tipo_Desc & ""
      .col = 6
      .Text = rs!Documento
      .col = 7
      .Text = Format(rs!Importe, "Standard")
      .col = 8
      .Text = rs!Descripcion & ""
      .col = 9
      .Text = rs!Estado & ""
      .col = 10
      .Text = CStr(rs!CONCILIA_ID_REF & "")
      
      curTotal = curTotal + rs!Importe
      rs.MoveNext
    Loop
    rs.Close

feResTotalCasos.Text = Format(.MaxRows, "###,###,##0")
feResTotalMonto.Text = Format(curTotal, "Standard")

feAR_Casos.Text = Format(0, "###,###,##0")
feAR_Monto.Text = Format(0, "Standard")

End With

vPaso = False


'--- Estado de los Botones de Auto Registro y Pendientes
btnPendiente.Visible = False
btnAutoRegistro.Visible = False
    
If Mid(cboCasos_Estados.Text, 1, 1) = "P" Then
   btnPendiente.Visible = True
End If

If Mid(cboCasos_Estados.Text, 1, 1) = "P" And Mid(cboTipoDoc.Text, 1, 1) <> "T" _
        And cboUbicacion.ItemData(cboUbicacion.ListIndex) = "B" Then
    btnAutoRegistro.Visible = True
End If


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub




Private Sub btnAutoRegistro_Click()
Dim strSQL As String, i As Long
Dim pCuenta As String

If Mid(fePeriodoEstado.Text, 1, 1) = "C" Then Exit Sub

If CCur(feAR_Monto.Text) = 0 Then
    MsgBox "No se ha indicado ningún movimiento para auto-registro!", vbExclamation
    Exit Sub
End If

If Not fxgCntCuentaValida(feAR_Cuenta.Text) Then
    MsgBox "La cuenta contable indicada para el auto-registro no es válida!", vbExclamation
    Exit Sub
End If

'Inicio
On Error GoTo vError

Me.MousePointer = vbHourglass

pCuenta = fxgCntCuentaFormato(False, feAR_Cuenta.Text, 0)

With gResultados
    strSQL = ""
    
    For i = 1 To .MaxRows
        .Row = i
        .col = 2
        If .Value = vbChecked Then
            .col = 3
            strSQL = strSQL & Space(10) & "exec spTes_Concilia_Auto_Registro " & cboBanco.ItemData(cboBanco.ListIndex) _
                   & "," & feAnio.Text & "," & feMes.Text & "," & .Text & ",'" & pCuenta & "','" & glogon.Usuario _
                   & "'," & chkAutoReg_Contabilidad.Value
        End If
        
        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If
        
    Next i

End With

'Revisa Ultimo Lote
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If


'Actualiza Resumen
strSQL = "exec spTes_Concilia_Periodo_Actualiza " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

'Refresca
Call sbResultados_Consulta

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnBancos_Saldo_Update_Click()
Dim strSQL As String

On Error GoTo vError

If Mid(fePeriodoEstado.Text, 1, 1) = "C" Then Exit Sub

Me.MousePointer = vbHourglass

'Concilia y Actualiza
strSQL = "exec spTes_Concilia_Periodo_Actualiza_Saldo_Cta " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & "," & CCur(feBancosSaldoActual.Text) & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)
                
strSQL = "exec spTes_Concilia_Periodo_Actualiza " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


'Consulta
Call sbPeriodo_Consulta(cboBanco.ItemData(cboBanco.ListIndex), feAnio.Text, feMes.Text)


Me.MousePointer = vbDefault


MsgBox "Actualización realizada satisfactoriamente!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnCasos_Buscar_Click(Index As Integer)

Select Case Index
    Case 0 'Buscar
        Call sbResultados_Consulta
    Case 1 'Exportar
        Dim vHeaders As vGridHeaders
            vHeaders.Columnas = 11
            vHeaders.Headers(3) = "Id"
            vHeaders.Headers(4) = "Fecha"
            vHeaders.Headers(5) = "Tipo"
            vHeaders.Headers(6) = "Documento"
            vHeaders.Headers(7) = "Importe"
            vHeaders.Headers(8) = "Descripción"
            vHeaders.Headers(9) = "Estado"
            vHeaders.Headers(10) = "Id Concilia Ref"
            Call sbSIFGridExportar(gResultados, vHeaders, "Bancos_Concilia_Resultados_Cta Id-" & cboBanco.ItemData(cboBanco.ListIndex) & feAnio.Text & "_" & feMes.Text)
    
End Select

End Sub

Private Sub btnConcilia_Aplicar_Click()
Dim strSQL As String, i As Long
Dim pId_Bancos As Long, pId_Libros As String

If Mid(fePeriodoEstado.Text, 1, 1) = "C" Then Exit Sub

If CCur(feMov_SelMonto.Text) = 0 Then
    MsgBox "No se ha indicado ningún movimiento para conciliar?", vbExclamation
    Exit Sub
End If

If Abs(CCur(feMov_Pendiente.Text)) > 5 Then
    MsgBox "La cantidad de casos marcada sobre pasa el pendiente a conciliar, revise!", vbExclamation
    Exit Sub
End If


'Inicio
On Error GoTo vError

Me.MousePointer = vbHourglass


With gConcilia
    strSQL = ""

    For i = 1 To .MaxRows
        .Row = i
        .col = 1
        If .Value = vbChecked Then
            .col = 2
            
            If feMov_Id.Tag = "B" Then
               pId_Bancos = feMov_Id.Text
               pId_Libros = .Text
            Else
               pId_Libros = feMov_Id.Text
               pId_Bancos = .Text
            End If
            
            
            If Mid(cboMovFiltro.Text, 1, 1) = "T" Then
                strSQL = strSQL & Space(10) & "exec spTes_Concilia_Aplicacion " & cboBanco.ItemData(cboBanco.ListIndex) _
                       & "," & feAnio.Text & "," & feMes.Text & "," & pId_Bancos & "," & pId_Libros & ",'" & glogon.Usuario & "'"
            Else
                strSQL = strSQL & Space(10) & "exec spTes_Concilia_Aplicacion_Lote " & cboBanco.ItemData(cboBanco.ListIndex) _
                       & "," & feAnio.Text & "," & feMes.Text & "," & pId_Bancos & ",'" & pId_Libros & "','" & glogon.Usuario & "'"
            End If
     
        End If
        
        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If
        
    Next i

End With

'Revisa Ultimo Lote
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If


''Actualiza Resumen
'strSQL = "exec spTes_Concilia_Periodo_Actualiza " & cboBanco.ItemData(cboBanco.ListIndex) _
'       & "," & feAnio.Text & "," & feMes.Text & ",'" & glogon.Usuario & "'"
'Call ConectionExecute(strSQL)

'Refesca
Call sbConcilia(feMov_Id.Text, feMov_Id.Tag)


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnConcilia_Buscar_Click()
If vPaso Then Exit Sub
If Not IsNumeric(feMov_Id.Text) Then Exit Sub

Call sbConcilia(feMov_Id.Text, feMov_Id.Tag)

End Sub

Private Sub btnConcilia1a1_Click(Index As Integer)
Dim strSQL As String

On Error GoTo vError

If Mid(fePeriodoEstado.Text, 1, 1) = "C" Then Exit Sub

Me.MousePointer = vbHourglass

Select Case Index
    Case 0 'Concilia Movimientos (E/S) Banco
        strSQL = "exec spTes_Concilia_Bancos_EntreSi " & cboBanco.ItemData(cboBanco.ListIndex) _
               & "," & feAnio.Text & "," & feMes.Text & ",'" & glogon.Usuario & "'"
    
    Case 1 'Concilia Movimientos (E/S) Libros
        strSQL = "exec spTes_Concilia_Libros_EntreSi " & cboBanco.ItemData(cboBanco.ListIndex) _
               & "," & feAnio.Text & "," & feMes.Text & ",'" & glogon.Usuario & "'"
End Select

Call ConectionExecute(strSQL)


Me.MousePointer = vbDefault


MsgBox "Conciliación 1:1 Movimientos Entre Sí, Realizada Satisfactoriamente! ", vbInformation

Call btnPeriodoActualiza_Click


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnConciliaReversa_Click()
Dim strSQL As String, i As Long
Dim pId_Bancos As Long, pId_Libros As String

If Mid(fePeriodoEstado.Text, 1, 1) = "C" Or lswConcilia.ListItems.Count = 0 Then Exit Sub

'Inicio
On Error GoTo vError

Me.MousePointer = vbHourglass


With lswConcilia.ListItems
    strSQL = ""

    For i = 1 To .Count

        If feMov_Id.Tag = "B" Then
           pId_Bancos = feMov_Id.Text
           pId_Libros = .Item(i).Text
        Else
           pId_Libros = feMov_Id.Text
           pId_Bancos = .Item(i).Text
        End If


        If Mid(cboMovFiltro.Text, 1, 1) = "T" Then
            strSQL = strSQL & Space(10) & "exec spTes_Concilia_Reversa " & cboBanco.ItemData(cboBanco.ListIndex) _
                   & "," & feAnio.Text & "," & feMes.Text & "," & pId_Bancos & "," & pId_Libros & ",'" & glogon.Usuario & "'"
        Else
            strSQL = strSQL & Space(10) & "exec spTes_Concilia_Reversa " & cboBanco.ItemData(cboBanco.ListIndex) _
                   & "," & feAnio.Text & "," & feMes.Text & "," & pId_Bancos & ",'" & pId_Libros & "','" & glogon.Usuario & "'"
        End If

        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If

    Next i

End With

'Revisa Ultimo Lote
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If


'Actualiza Resumen
strSQL = "exec spTes_Concilia_Periodo_Actualiza " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

'Refesca
Call sbConcilia(feMov_Id.Text, feMov_Id.Tag)


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnGuardar_Click()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

'Guarda y Actualiza
strSQL = "exec spTes_Concilia_Periodo_Add " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'A','" & feNotas.Text & "','" & glogon.Usuario _
       & "'," & CCur(feLibrosSaldoActual.Text) & "," & CCur(feBancosSaldoActual.Text)
Call ConectionExecute(strSQL)

'Consulta
Call sbPeriodo_Consulta(cboBanco.ItemData(cboBanco.ListIndex), feAnio.Text, feMes.Text)

Me.MousePointer = vbDefault


MsgBox "Cambios y actualización realizada satisfactoriamente!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnPendiente_Click()
Dim strSQL As String, i As Long, pTipo As String

If Mid(fePeriodoEstado.Text, 1, 1) = "C" Then Exit Sub

If CCur(feAR_Monto.Text) = 0 Then
    MsgBox "No se ha indicado ningún caso para poner pendiente!", vbExclamation
    Exit Sub
End If


'Inicio
On Error GoTo vError

Me.MousePointer = vbHourglass


With gResultados
    strSQL = ""
    
    For i = 1 To .MaxRows
        .Row = i
        .col = 2
        If .Value = vbChecked Then
            .col = 3
            strSQL = strSQL & Space(10) & "exec spTes_Concilia_Pendiente " & cboBanco.ItemData(cboBanco.ListIndex) _
                   & "," & feAnio.Text & "," & feMes.Text & "," & .Text _
                   & ",'" & cboUbicacion.ItemData(cboUbicacion.ListIndex) & "','" & glogon.Usuario & "'"
        End If
        
        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If
        
    Next i

End With

'Revisa Ultimo Lote
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If


'Actualiza Resumen> No Aplica porque Los pendientes se tienen que reflejar
'strSQL = "exec spTes_Concilia_Periodo_Actualiza " & cboBanco.ItemData(cboBanco.ListIndex) _
'       & "," & feAnio.Text & "," & feMes.Text & ",'" & glogon.Usuario & "'"
'Call ConectionExecute(strSQL)

'Refresca
Call sbResultados_Consulta

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnPeriodoActualiza_Click()
Dim strSQL As String

On Error GoTo vError

If Mid(fePeriodoEstado.Text, 1, 1) = "C" Then Exit Sub

Me.MousePointer = vbHourglass

'Concilia y Actualiza
strSQL = "exec spTes_Concilia_Automatica " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)
                
strSQL = "exec spTes_Concilia_Periodo_Actualiza " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


'Consulta
Call sbPeriodo_Consulta(cboBanco.ItemData(cboBanco.ListIndex), feAnio.Text, feMes.Text)


Me.MousePointer = vbDefault


MsgBox "Actualización realizada satisfactoriamente!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnPeriodoCerrar_Click()
Dim strSQL As String

If Mid(fePeriodoEstado.Text, 1, 1) = "C" Then Exit Sub

On Error GoTo vError

strSQL = "exec spTes_Concilia_Periodo_Cierra " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call sbPeriodo_Consulta(cboBanco.ItemData(cboBanco.ListIndex), feAnio.Text, feMes.Text)

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnPeriodoInforme_Click()
Dim strSQL As String, strBusqueda As String
Dim strEstado As String, blnDesde As Boolean, blnHasta As Boolean

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = ""

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Sub-Módulo de Conciliación Bancaria"
    .WindowShowGroupTree = True
    
    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "Fecha='" & Format(fxFechaServidor, ("dd/mm/yyyy")) & "'"
    .Formulas(2) = "Usuario='" & glogon.Usuario & "'"
    
    .Connect = glogon.ConectRPT
     
     
    strSQL = "{TES_CONCILIA_PERIODO.ID_BANCO} = " & cboBanco.ItemData(cboBanco.ListIndex) _
           & " AND {TES_CONCILIA_PERIODO.ANIO} = " & feAnio.Text _
           & " AND {TES_CONCILIA_PERIODO.MES} = " & feMes.Text
     
'        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'        strSQL = strSQL & "({TES_TRANSACCIONES.FECHA_EMISION} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
'               & ") To Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & "))"
     
    Select Case True
      Case rbInformes.Item(0).Value 'Informe General
           .ReportFileName = SIFGlobal.fxPathReportes("Banking_Concilia_Informe.rpt")
      
           .Formulas(3) = "Titulo='Informe General'"
      
      Case rbInformes.Item(1).Value 'Pendientes en Bancos
           .ReportFileName = SIFGlobal.fxPathReportes("Banking_Concilia_Pendientes.rpt")
           strSQL = strSQL & " AND {vTes_Conciliacion_Pendientes.UBICACION_ID} = 'B'"
           strSQL = strSQL & " AND ({vTes_Conciliacion_Pendientes.FECHA} in Date(" & Format(dtpInforme(0).Value, "yyyy,mm,dd") _
               & ") To Date(" & Format(dtpInforme(1).Value, "yyyy,mm,dd") & "))"
            
           
           .Formulas(3) = "Titulo='Pendientes de Conciliar: En Bancos'"
            
      Case rbInformes.Item(2).Value 'Movimientos en Bancos
           .ReportFileName = SIFGlobal.fxPathReportes("Banking_Concilia_Movimientos.rpt")
           strSQL = strSQL & " AND {vTes_Conciliacion_Informe_Movimientos.UBICACION_ID} = 'B'"
           strSQL = strSQL & " AND ({vTes_Conciliacion_Informe_Movimientos.FECHA} in Date(" & Format(dtpInforme(0).Value, "yyyy,mm,dd") _
               & ") To Date(" & Format(dtpInforme(1).Value, "yyyy,mm,dd") & "))"
               
           .Formulas(3) = "Titulo='Movimientos: En Bancos'"
            
            
      Case rbInformes.Item(3).Value 'Pendientes en Libros
           .ReportFileName = SIFGlobal.fxPathReportes("Banking_Concilia_Pendientes.rpt")
           strSQL = strSQL & " AND {vTes_Conciliacion_Pendientes.UBICACION_ID} = 'L'"
           strSQL = strSQL & " AND ({vTes_Conciliacion_Pendientes.FECHA} in Date(" & Format(dtpInforme(0).Value, "yyyy,mm,dd") _
               & ") To Date(" & Format(dtpInforme(1).Value, "yyyy,mm,dd") & "))"
               
           .Formulas(3) = "Titulo='Pendientes de Conciliar: En Libros'"
    
      Case rbInformes.Item(4).Value 'Movimientos en Libros
           .ReportFileName = SIFGlobal.fxPathReportes("Banking_Concilia_Movimientos.rpt")
           strSQL = strSQL & " AND {vTes_Conciliacion_Informe_Movimientos.UBICACION_ID} = 'L'"
           strSQL = strSQL & " AND ({vTes_Conciliacion_Informe_Movimientos.FECHA} in Date(" & Format(dtpInforme(0).Value, "yyyy,mm,dd") _
               & ") To Date(" & Format(dtpInforme(1).Value, "yyyy,mm,dd") & "))"
           
           .Formulas(3) = "Titulo='Movimientos: En Libros'"
      
      Case rbInformes.Item(5).Value 'Transacciones en Lotes
           .ReportFileName = SIFGlobal.fxPathReportes("Banking_Concilia_Movimientos_Lotes.rpt")
    
           .Formulas(3) = "Titulo='Movimientos en Lotes: Transferencias'"
    End Select
     
           
    .SelectionFormula = strSQL
    .Action = 1
'    .PrintReport
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnPeriodoInicializa_Click()
Dim strSQL As String

If Mid(fePeriodoEstado.Text, 1, 1) = "C" Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spTes_Concilia_Periodo_Inicializa " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

strSQL = "exec spTes_Concilia_Periodo_Actualiza " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call sbPeriodo_Consulta(cboBanco.ItemData(cboBanco.ListIndex), feAnio.Text, feMes.Text)

Me.MousePointer = vbDefault

MsgBox "Conciliación del Periodo - Inicializada Satisfactoriamente!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnPeriodoNuevo_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spTes_Concilia_Periodo_Add " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'A','','" & glogon.Usuario _
       & "',0,0"
Call ConectionExecute(strSQL)

Call sbPeriodo_Consulta(cboBanco.ItemData(cboBanco.ListIndex), feAnio.Text, feMes.Text)

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub

Private Sub sbPeriodo_Consulta(pBancoId As Long, pAnio As Long, pMes As Integer)

Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vPaso Then Exit Sub

strSQL = "select *" _
       & " from vTES_CONCILIA_PERIODO" _
       & " where id_Banco = " & pBancoId _
       & "   and Anio = " & pAnio _
       & "   and Mes = " & pMes
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    tcMain.Item(1).Selected = True
    feAnio.Text = rs!Anio
    feMes.Text = rs!Mes
    fePeriodo.Text = ""
    fePeriodoEstado.Text = IIf(rs!Estado = "A", "Abierto", "Cerrado")
    
    feLibrosSaldoActual.Text = Format(rs!LIBROS_SALDO, "Standard")
    feBancosNC.Text = Format(rs!LIBROS_NC, "Standard")
    feBancosND.Text = Format(rs!LIBROS_ND, "Standard")
    feLibrosSaldoConcilia.Text = Format(rs!LIBROS_SALDO_CONCILIA, "Standard")
    
    feBancosSaldoActual.Text = Format(rs!CTA_SALDO, "Standard")
    feDepositosTransito.Text = Format(rs!DEPOSITOS_TRANSITO, "Standard")
    feChequesNoCobrados.Text = Format(rs!CHEQUES_NO_COBRADOS, "Standard")
    feBancosSaldoConcilia.Text = Format(rs!CTA_SALDO_CONCILIA, "Standard")
    
    feSaldoDiferencia.Text = Format(rs!LIBROS_SALDO_CONCILIA - rs!CTA_SALDO_CONCILIA, "Standard")
    
    feNotas.Text = rs!NOTAS & ""
        
    feBancosSaldoActual.ToolTipText = ""
    If rs!CTA_SALDO_UPD_IND = 1 Then
        feBancosSaldoActual.ToolTipText = "Modificado: " & rs!CTA_SALDO_UPD_FECHA & ", " & rs!CTA_SALDO_UPD_USUARIO & ""
    End If
    
    
    dtpInforme(0).Value = rs!Periodo_Inicio
    dtpInforme(1).Value = rs!Periodo_Corte
    
    
Else
    tcMain.Item(0).Selected = True
    feAnio.Text = feAnio.Text
    feMes.Text = feMes.Text
    fePeriodo.Text = ""
    fePeriodoEstado.Text = "Abierto"
    
    feLibrosSaldoActual.Text = Format(0, "Standard")
    feBancosNC.Text = Format(0, "Standard")
    feBancosND.Text = Format(0, "Standard")
    
    feLibrosSaldoConcilia.Text = Format(0, "Standard")
    
    feBancosSaldoActual.Text = Format(0, "Standard")
    
    feBancosSaldoActual.ToolTipText = ""
    
    feDepositosTransito.Text = Format(0, "Standard")
    feChequesNoCobrados.Text = Format(0, "Standard")
    feBancosSaldoConcilia.Text = Format(0, "Standard")
    
    feSaldoDiferencia.Text = Format(0, "Standard")
    
    feNotas.Text = ""
    
End If

Call sbRefrescaInformacion

rs.Close



Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub cboBanco_Click()
If vPaso Then Exit Sub

Dim strSQL As String

On Error GoTo vError

gHistorial.MaxRows = 0

If cboBanco.ListCount = 0 Then Exit Sub
If cboBanco.ItemData(cboBanco.ListIndex) = Empty Then Exit Sub

vPaso = True
    strSQL = "exec spTes_Concilia_Periodo_Consulta " & cboBanco.ItemData(cboBanco.ListIndex) _
            & ",'" & glogon.Usuario & "'"
    Call sbCargaGrid(gHistorial, 8, strSQL)
    
    gHistorial.MaxRows = gHistorial.MaxRows - 1
vPaso = False

tcMain.Item(0).Selected = True

If gHistorial.MaxRows = 0 Then
    feAnio.Text = gCntX_Parametros.PeriodoAnio
    feMes.Text = gCntX_Parametros.PeriodoMes
    fePeriodoEstado.Text = "Abierto"
    Call sbRefrescaInformacion
End If


Exit Sub

vError:



End Sub

Private Sub cboCasos_Estados_Click()
If vPaso Then Exit Sub

feResTotalCasos.Text = "0"
feResTotalMonto.Text = "0"

gResultados.MaxRows = 0


'--- Estado de los Botones de Auto Registro y Pendientes
btnPendiente.Visible = False
btnAutoRegistro.Visible = False
    
If Mid(cboCasos_Estados.Text, 1, 1) = "P" Then
   btnPendiente.Visible = True
End If

If Mid(cboCasos_Estados.Text, 1, 1) = "P" And Mid(cboTipoDoc.Text, 1, 1) <> "T" _
        And cboUbicacion.ItemData(cboUbicacion.ListIndex) = "B" Then
    btnAutoRegistro.Visible = True
End If


End Sub

Private Sub cboMovFiltro_Click()
If vPaso Then Exit Sub
If Not IsNumeric(feMov_Id.Text) Then Exit Sub

Call sbConcilia(feMov_Id.Text, feMov_Id.Tag)

End Sub

Private Sub cboTipoDoc_Click()
If vPaso Then Exit Sub

feResTotalCasos.Text = "0"
feResTotalMonto.Text = "0"
gResultados.MaxRows = 0

End Sub

Private Sub cboUbicacion_Click()

If cboUbicacion.ListCount = 0 Or vPaso Then Exit Sub

vPaso = True

gResultados.MaxRows = 0
feResTotalCasos.Text = "0"
feResTotalMonto.Text = "0"


cboTipoDoc.Clear

If cboUbicacion.ItemData(cboUbicacion.ListIndex) = "L" Then
    cboTipoDoc.AddItem "Cheques No Cobrados"
    cboTipoDoc.ItemData(cboTipoDoc.ListCount - 1) = "CK"
    cboTipoDoc.AddItem "Depósitos en Transito"
    cboTipoDoc.ItemData(cboTipoDoc.ListCount - 1) = "DP"
    cboTipoDoc.Text = "Cheques No Cobrados"
Else
    cboTipoDoc.AddItem "Débitos"
    cboTipoDoc.ItemData(cboTipoDoc.ListCount - 1) = "D"
    cboTipoDoc.AddItem "Créditos"
    cboTipoDoc.ItemData(cboTipoDoc.ListCount - 1) = "C"
    cboTipoDoc.Text = "Débitos"
End If

cboTipoDoc.AddItem "TODOS"

cboCasos_Estados.Clear
cboCasos_Estados.AddItem "Pendientes"
cboCasos_Estados.AddItem "Conciliado"
cboCasos_Estados.AddItem "Pendientes"

cboCasos_Estados.Text = "Pendientes"

vPaso = False

'--- Estado de los Botones de Auto Registro y Pendientes
btnPendiente.Visible = False
btnAutoRegistro.Visible = False
    
If Mid(cboCasos_Estados.Text, 1, 1) = "P" Then
   btnPendiente.Visible = True
End If

If Mid(cboCasos_Estados.Text, 1, 1) = "P" And Mid(cboTipoDoc.Text, 1, 1) <> "T" _
        And cboUbicacion.ItemData(cboUbicacion.ListIndex) = "B" Then
    btnAutoRegistro.Visible = True
End If

End Sub

Private Sub chkConciliaFiltroFechas_Click()


dtpConciliaCorte.Visible = dtpConciliaInicio.Visible

Call btnConcilia_Buscar_Click
End Sub

Private Sub chkConciliaFiltroMontos_Click()
Call btnConcilia_Buscar_Click
End Sub

Private Sub chkConciliaPendientes_Click()
Call btnConcilia_Buscar_Click
End Sub

Private Sub chkConciliaTodos_Click()


If vPaso Then Exit Sub
If gConcilia.MaxRows = 0 Then Exit Sub

Dim i As Long


On Error GoTo vError

feMov_SelCasos.Text = Format(0, "###,###,##0")
feMov_SelMonto.Text = Format(0, "Standard")

feMov_Conciliado.Text = feMov_Conciliado.Tag
feMov_Pendiente.Text = feMov_Pendiente.Tag


With gConcilia
    vPaso = True
    For i = 1 To .MaxRows
        .Row = i
        .col = 1
        .Value = chkConciliaTodos
        
        If .Value = vbChecked Then
            .col = 6
            feMov_SelCasos.Text = Format(CLng(feMov_SelCasos.Text) + 1, "###,###,##0")
            feMov_SelMonto.Text = Format(CCur(feMov_SelMonto.Text) + CCur(.Text), "Standard")
            
            feMov_Conciliado.Text = Format(CCur(feMov_Conciliado.Text) + CCur(.Text), "Standard")
            feMov_Pendiente.Text = Format(CCur(feMov_Pendiente.Text) - CCur(.Text), "Standard")
        End If
            
    Next i
    vPaso = False
End With

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub feAnio_Change()
Call sbRefrescaInformacion
End Sub



Private Sub feAR_Cuenta_GotFocus()
On Error GoTo vError
feAR_Cuenta.Text = fxgCntCuentaFormato(False, feAR_Cuenta.Text, 0)
vError:
End Sub

Private Sub feAR_Cuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then feAR_CuentaDesc.SetFocus

If KeyCode = vbKeyF4 Then
   Call sbgCntCuentaConsulta("C")
   feAR_Cuenta.Text = gBusquedas.Resultado
   feAR_CuentaDesc.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub feAR_Cuenta_LostFocus()
On Error GoTo vError

feAR_CuentaDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, feAR_Cuenta.Text))
feAR_Cuenta.Text = fxgCntCuentaFormato(True, feAR_Cuenta.Text)

vError:
End Sub

Private Sub feBancosSaldoActual_GotFocus()
On Error GoTo vError

feBancosSaldoActual.Text = CCur(feBancosSaldoActual.Text)

Exit Sub

vError:
End Sub

Private Sub feBancosSaldoActual_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyF4 Then btnBancos_Saldo_Update.SetFocus
End Sub

Private Sub feBancosSaldoActual_LostFocus()
On Error GoTo vError

feBancosSaldoActual.Text = Format(CCur(feBancosSaldoActual.Text), "Standard")

Exit Sub

vError:
End Sub

Private Sub feMes_Change()
Call sbRefrescaInformacion
End Sub

Private Sub Form_Load()

Dim strSQL As String

vModulo = 9

Set Me.imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

tcMain.Item(0).Selected = True

gHistorial.MaxRows = 0

With lswConcilia.ColumnHeaders
    .Clear
    .Add , , "Id", 1200
    .Add , , "Fecha", 1400, vbCenter
    .Add , , "Tipo", 1200, vbCenter
    .Add , , "Documento", 2100, vbCenter
    .Add , , "Importe", 1800, vbRightJustify
    .Add , , "Descripción", 3500
    .Add , , "Id Referencia", 1400, vbCenter
    .Add , , "Tipo Concilia", 1400, vbCenter
    
End With

With lswConcilia_Lote.ColumnHeaders
    .Clear
    .Add , , "Id", 1200
    .Add , , "Fecha", 1400, vbCenter
    .Add , , "Tipo", 1200, vbCenter
    .Add , , "Documento", 2100, vbCenter
    .Add , , "Importe", 1800, vbRightJustify
    .Add , , "Descripción", 3500
    .Add , , "Id Referencia", 1400, vbCenter
    .Add , , "Tipo Concilia", 1400, vbCenter
End With



vPaso = True

strSQL = "exec spTes_Cuenta_Bancaria_Acceso_General '" & glogon.Usuario & "','ASI'"

Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

cboUbicacion.Clear

cboUbicacion.AddItem "Movimientos Según Bancos"
cboUbicacion.ItemData(cboUbicacion.ListCount - 1) = "B"
cboUbicacion.AddItem "Movimientos Según Libros"
cboUbicacion.ItemData(cboUbicacion.ListCount - 1) = "L"
cboUbicacion.Text = "Movimientos Según Bancos"


cboMovFiltro.Clear
cboMovFiltro.AddItem "Transacciones"
cboMovFiltro.AddItem "Lotes"
cboMovFiltro.Text = "Transacciones"

vPaso = False


End Sub


Private Sub sbConcilia(pCaso As Long, pTipo As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, pIdActual As Long

Dim curConciliado As Currency

On Error GoTo vError

tcMain.Item(3).Selected = True

curConciliado = 0

If chkConciliaFiltroFechas.Value = xtpChecked Then
    dtpConciliaInicio.Visible = True
Else
    dtpConciliaInicio.Visible = False
End If
dtpConciliaCorte.Visible = dtpConciliaInicio.Visible


If IsNumeric(feMov_Id.Text) Then
    pIdActual = feMov_Id.Text
Else
    pIdActual = 0
End If

feMov_Id.Text = "0"
feMov_Id.Tag = ""

feMov_Fecha.Text = ""
feMov_Descripcion.Text = ""
feMov_Documento.Text = ""
feMov_Estado.Text = ""
feMov_Importe.Text = "0.00"

feMov_Pendiente.Text = "0"
feMov_Conciliado.Text = "0"
feMov_SelCasos.Text = "0"
feMov_SelMonto.Text = Format(0, "Standard")

strSQL = "exec spTes_Concilia_Periodo_Resultados_Caso " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'" & cboUbicacion.ItemData(cboUbicacion.ListIndex) & "'," & pCaso
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    
    If pIdActual <> rs!Id Then
            If Day(rs!fecha) = 1 Then
                dtpConciliaInicio.Value = rs!fecha
            Else
                dtpConciliaInicio.Value = DateAdd("d", -1, rs!fecha)
            
            End If
            dtpConciliaCorte.Value = DateAdd("d", 1, dtpConciliaInicio.Value)
    End If
    
    feMov_Id.Text = rs!Id
    feMov_Id.Tag = cboUbicacion.ItemData(cboUbicacion.ListIndex)
    
    feMov_Tipo.Text = rs!Tipo_Desc
    feMov_Tipo.Tag = rs!Tipo
    
    feMov_Descripcion.Text = rs!Descripcion
    feMov_Fecha.Text = rs!fecha
    feMov_Documento.Text = rs!Documento
    feMov_Estado.Text = rs!Estado
    
    feMov_Importe.Text = Format(rs!Importe, "Standard")
    

End If
rs.Close

'Detalle de Transacciones Vinculadas
strSQL = "exec spTes_Concilia_Periodo_Resultados_Caso_Detalle " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'" & cboUbicacion.ItemData(cboUbicacion.ListIndex) & "'," & pCaso
Call OpenRecordSet(rs, strSQL)

lswConcilia.ListItems.Clear

Do While Not rs.EOF
  Set itmX = lswConcilia.ListItems.Add(, , rs!Id)
      itmX.SubItems(1) = rs!fecha
      itmX.SubItems(2) = rs!Tipo_Desc
      itmX.SubItems(3) = rs!Documento
      itmX.SubItems(4) = Format(rs!Importe, "Standard")
      itmX.SubItems(5) = rs!Descripcion
      itmX.SubItems(6) = rs!CONCILIA_ID_REF
      itmX.SubItems(7) = rs!CONCILIA_DESC
      
      curConciliado = curConciliado + rs!Importe
  rs.MoveNext
Loop
rs.Close

feMov_Pendiente.Text = Format(CCur(feMov_Importe.Text) - curConciliado, "Standard")
feMov_Conciliado.Text = Format(curConciliado, "Standard")

feMov_Pendiente.Tag = Format(CCur(feMov_Importe.Text) - curConciliado, "Standard")
feMov_Conciliado.Tag = Format(curConciliado, "Standard")



'Detalle de Transacciones Vinculadas
strSQL = "exec spTes_Concilia_Periodo_Resultados_Caso_Lote " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'" & cboUbicacion.ItemData(cboUbicacion.ListIndex) & "'," & pCaso
Call OpenRecordSet(rs, strSQL)

lswConcilia_Lote.ListItems.Clear

Do While Not rs.EOF
  Set itmX = lswConcilia_Lote.ListItems.Add(, , rs!Id)
      itmX.SubItems(1) = rs!fecha
      itmX.SubItems(2) = rs!Tipo_Desc
      itmX.SubItems(3) = rs!Documento
      itmX.SubItems(4) = Format(rs!Importe, "Standard")
      itmX.SubItems(5) = rs!Descripcion
      itmX.SubItems(6) = rs!CONCILIA_ID_REF
      itmX.SubItems(7) = rs!CONCILIA_DESC
  rs.MoveNext
Loop
rs.Close


'Disponibles
strSQL = "exec spTes_Concilia_Periodo_Disponibles " & cboBanco.ItemData(cboBanco.ListIndex) _
       & "," & feAnio.Text & "," & feMes.Text & ",'" & cboUbicacion.ItemData(cboUbicacion.ListIndex) _
       & "','" & feMov_Tipo.Tag & "'," & CCur(feMov_Importe.Text) & ",'" & Mid(cboMovFiltro.Text, 1, 1) & "'" _
       & "," & chkConciliaPendientes.Value & "," & chkConciliaFiltroMontos.Value & "," & chkConciliaFiltroFechas.Value _
       & ",'" & Format(dtpConciliaInicio.Value, "yyyy/mm/dd") & "','" & Format(dtpConciliaCorte.Value, "yyyy/mm/dd") & " 23:59'"
       
       

Call OpenRecordSet(rs, strSQL)

With gConcilia
   .MaxRows = 0

    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      
      .col = 2
      .Text = CStr(rs!Id)
      .col = 3
      .Text = Format(rs!fecha, "dd/mm/yyyy")
      .col = 4
      .Text = rs!Tipo_Desc & ""
      .col = 5
      .Text = rs!Documento
      .col = 6
      .Text = Format(rs!Importe, "Standard")
      .col = 7
      .Text = rs!Descripcion & ""
      .col = 8
      .Text = rs!Estado & ""
      .col = 9
      .Text = CStr(rs!CONCILIA_ID_REF & "")

      rs.MoveNext
    Loop
    rs.Close

End With

If lswConcilia.ListItems.Count > 0 Then
  tcConcilia.Item(1).Selected = True
Else
  tcConcilia.Item(0).Selected = True
End If

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub gConcilia_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If vPaso Then Exit Sub
If gConcilia.MaxRows = 0 Then Exit Sub

On Error GoTo vError

With gConcilia
    .Row = Row
    .col = col
    If .Value = vbChecked Then
        .col = 6
        feMov_SelCasos.Text = Format(CLng(feMov_SelCasos.Text) + 1, "###,###,##0")
        feMov_SelMonto.Text = Format(CCur(feMov_SelMonto.Text) + CCur(.Text), "Standard")
        
        feMov_Conciliado.Text = Format(CCur(feMov_Conciliado.Text) + CCur(.Text), "Standard")
        feMov_Pendiente.Text = Format(CCur(feMov_Pendiente.Text) - CCur(.Text), "Standard")
        
    Else
        .col = 6
        feMov_SelCasos.Text = Format(CLng(feMov_SelCasos.Text) - 1, "###,###,##0")
        feMov_SelMonto.Text = Format(CCur(feMov_SelMonto.Text) - CCur(.Text), "Standard")
    
        feMov_Conciliado.Text = Format(CCur(feMov_Conciliado.Text) - CCur(.Text), "Standard")
        feMov_Pendiente.Text = Format(CCur(feMov_Pendiente.Text) + CCur(.Text), "Standard")
    
    End If
        
        
End With

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub gResultados_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub
If gResultados.MaxRows = 0 Then Exit Sub

Dim pId As Long, frm

On Error Resume Next

With gResultados
    .Row = Row
    .col = col
    Select Case col
        Case 1 'Conciliacion
            .col = 3
            If IsNumeric(.Text) Then
                Call sbConcilia(.Text, Mid(cboUbicacion.Text, 1, 1))
                    
            End If
        
        Case 2 'Auto Registro
            .col = 2
            If .Value = vbChecked Then
                .col = 7
                feAR_Casos.Text = Format(CLng(feAR_Casos.Text) + 1, "###,###,##0")
                feAR_Monto.Text = Format(CCur(feAR_Monto.Text) + CCur(.Text), "Standard")
            Else
                .col = 7
                feAR_Casos.Text = Format(CLng(feAR_Casos.Text) - 1, "###,###,##0")
                feAR_Monto.Text = Format(CCur(feAR_Monto.Text) - CCur(.Text), "Standard")
            End If
        
        
        Case 11 'Transacciones
            .col = 10
            If IsNumeric(.Text) Then
                Call sbFormsCall("frmTES_Transacciones")
                For Each frm In Forms
                  If UCase(frm.Name) = UCase("frmTES_Transacciones") Then
                    Call frm.sbTESDocConsulta(.Text)
                    Exit For
                  End If
                Next frm
            End If
    End Select
End With

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

On Error GoTo vError

If Item.Index > 0 And Trim(fePeriodo.Text) = "" Then
    tcMain.Item(0).Selected = True
    Exit Sub
End If

Select Case Item.Index
 Case 0 'Historial
 Case 1 'Resumen
    Call sbPeriodo_Consulta(cboBanco.ItemData(cboBanco.ListIndex), feAnio.Text, feMes.Text)

 Case 2 'Resultados
    Call cboUbicacion_Click
     
 Case 3 'Conciliacion
    If Not IsNumeric(feMov_Id.Text) Then
        gConcilia.MaxRows = 0
    End If
End Select


Exit Sub

vError:

End Sub


Private Sub gHistorial_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

Dim pAnio As Long, pMes As Integer

If col = 1 Then
    gHistorial.Row = Row
    gHistorial.col = 2
    pAnio = gHistorial.Text
    
    gHistorial.col = 3
    pMes = gHistorial.Text
    

    Call sbPeriodo_Consulta(cboBanco.ItemData(cboBanco.ListIndex), pAnio, pMes)
    tcMain.Item(1).Selected = True
    
End If
End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call cboBanco_Click

End Sub
