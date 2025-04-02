VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAH_ExcedentesMensuales 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Excedentes"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   11295
   HelpContextID   =   2003
   Icon            =   "ExcedentesMensuales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   195
      Left            =   0
      TabIndex        =   79
      Top             =   7695
      Width           =   11415
      _Version        =   1441793
      _ExtentX        =   20135
      _ExtentY        =   344
      _StockProps     =   93
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   315
      Left            =   9960
      TabIndex        =   39
      Top             =   1200
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Aplicar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9960
      Top             =   720
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6132
      Left            =   0
      TabIndex        =   0
      Top             =   1224
      Width           =   11292
      _Version        =   1441793
      _ExtentX        =   19918
      _ExtentY        =   10816
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
      Item(0).Caption =   "Resumen"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "GroupBox1"
      Item(1).Caption =   "Mensual"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "GroupBox2"
      Item(1).Control(1)=   "GroupBox3"
      Item(1).Control(2)=   "GroupBox4"
      Item(2).Caption =   "Cierre"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "GroupBox5"
      Item(2).Control(1)=   "GroupBox6"
      Item(3).Caption =   "Aplicaciones"
      Item(3).ControlCount=   5
      Item(3).Control(0)=   "chkPosCierreCargaCero"
      Item(3).Control(1)=   "chkPosCierreLimpiar"
      Item(3).Control(2)=   "lblUltimoCierre"
      Item(3).Control(3)=   "Label4(0)"
      Item(3).Control(4)=   "GroupBox7"
      Item(4).Caption =   "Parámetros"
      Item(4).ControlCount=   4
      Item(4).Control(0)=   "cmdAjustes"
      Item(4).Control(1)=   "cmdCapitalizacionIndividual"
      Item(4).Control(2)=   "lswParametros"
      Item(4).Control(3)=   "btnCierreParametros(3)"
      Begin XtremeSuiteControls.ListView lswParametros 
         Height          =   5535
         Left            =   -69880
         TabIndex        =   62
         Top             =   480
         Visible         =   0   'False
         Width           =   8535
         _Version        =   1441793
         _ExtentX        =   15055
         _ExtentY        =   9763
         _StockProps     =   77
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.GroupBox GroupBox7 
         Height          =   5652
         Left            =   -67600
         TabIndex        =   65
         Top             =   1080
         Visible         =   0   'False
         Width           =   8652
         _Version        =   1441793
         _ExtentX        =   15261
         _ExtentY        =   9970
         _StockProps     =   79
         Caption         =   "Aplicaciones"
         ForeColor       =   8421504
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
         Begin XtremeSuiteControls.RadioButton optAplicaciones 
            Height          =   252
            Index           =   0
            Left            =   360
            TabIndex        =   66
            Top             =   360
            Width           =   4212
            _Version        =   1441793
            _ExtentX        =   7429
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Carga y Aplica Donaciones al Excedente"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optAplicaciones 
            Height          =   252
            Index           =   1
            Left            =   360
            TabIndex        =   67
            Top             =   720
            Width           =   4212
            _Version        =   1441793
            _ExtentX        =   7429
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Carga y Aplica Ajuste (+,-) al Excedente"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optAplicaciones 
            Height          =   252
            Index           =   2
            Left            =   360
            TabIndex        =   68
            Top             =   1080
            Width           =   4212
            _Version        =   1441793
            _ExtentX        =   7429
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Carga/Apl Saldos Garantia con Excedentes"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optAplicaciones 
            Height          =   252
            Index           =   3
            Left            =   360
            TabIndex        =   69
            Top             =   1440
            Width           =   4212
            _Version        =   1441793
            _ExtentX        =   7429
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Carga/Apl/Act Mora al Excedente"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optAplicaciones 
            Height          =   252
            Index           =   4
            Left            =   360
            TabIndex        =   70
            Top             =   1800
            Width           =   4212
            _Version        =   1441793
            _ExtentX        =   7429
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Carga/Apl/Act Mora OPCF al Excedente"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optAplicaciones 
            Height          =   252
            Index           =   5
            Left            =   360
            TabIndex        =   71
            Top             =   2160
            Width           =   4212
            _Version        =   1441793
            _ExtentX        =   7429
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aplica Capitalización Extraodrinaria"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optAplicaciones 
            Height          =   252
            Index           =   6
            Left            =   360
            TabIndex        =   72
            Top             =   2760
            Width           =   4812
            _Version        =   1441793
            _ExtentX        =   8488
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Actualiza Ahorros con Capitalizaciones Extraordinarias"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optAplicaciones 
            Height          =   252
            Index           =   7
            Left            =   360
            TabIndex        =   73
            Top             =   3120
            Width           =   4812
            _Version        =   1441793
            _ExtentX        =   8488
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Actualiza Ahorros con Capitalización General"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optAplicaciones 
            Height          =   252
            Index           =   8
            Left            =   360
            TabIndex        =   74
            Top             =   3480
            Width           =   4812
            _Version        =   1441793
            _ExtentX        =   8488
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Actualiza Información de Ajustes"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optAplicaciones 
            Height          =   252
            Index           =   9
            Left            =   360
            TabIndex        =   75
            Top             =   3840
            Width           =   4812
            _Version        =   1441793
            _ExtentX        =   8488
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Actualiza Créditos con Abonos a Créditos Excedentes"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optAplicaciones 
            Height          =   252
            Index           =   10
            Left            =   360
            TabIndex        =   76
            Top             =   4200
            Width           =   4812
            _Version        =   1441793
            _ExtentX        =   8488
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Crear Asiento General de Excedentes"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optAplicaciones 
            Height          =   252
            Index           =   11
            Left            =   5400
            TabIndex        =   77
            Top             =   360
            Width           =   4812
            _Version        =   1441793
            _ExtentX        =   8488
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Separar las Salidas de los Casos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optAplicaciones 
            Height          =   252
            Index           =   12
            Left            =   5400
            TabIndex        =   78
            Top             =   720
            Width           =   4812
            _Version        =   1441793
            _ExtentX        =   8488
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Traslado de  Excedentes a Fondos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
         End
         Begin VB.Image imgPass 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   12
            Left            =   5160
            Picture         =   "ExcedentesMensuales.frx":030A
            Top             =   720
            Width           =   240
         End
         Begin VB.Image imgPass 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   11
            Left            =   5160
            Picture         =   "ExcedentesMensuales.frx":0A21
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgPass 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   10
            Left            =   120
            Picture         =   "ExcedentesMensuales.frx":1138
            Top             =   4200
            Width           =   240
         End
         Begin VB.Image imgPass 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   9
            Left            =   120
            Picture         =   "ExcedentesMensuales.frx":184F
            Top             =   3840
            Width           =   240
         End
         Begin VB.Image imgPass 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   8
            Left            =   120
            Picture         =   "ExcedentesMensuales.frx":1F66
            Top             =   3480
            Width           =   240
         End
         Begin VB.Image imgPass 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   7
            Left            =   120
            Picture         =   "ExcedentesMensuales.frx":267D
            Top             =   3120
            Width           =   240
         End
         Begin VB.Image imgPass 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   6
            Left            =   120
            Picture         =   "ExcedentesMensuales.frx":2D94
            Top             =   2760
            Width           =   240
         End
         Begin VB.Image imgPass 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   5
            Left            =   120
            Picture         =   "ExcedentesMensuales.frx":34AB
            Top             =   2160
            Width           =   240
         End
         Begin VB.Image imgPass 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   4
            Left            =   120
            Picture         =   "ExcedentesMensuales.frx":3BC2
            Top             =   1800
            Width           =   240
         End
         Begin VB.Image imgPass 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   3
            Left            =   120
            Picture         =   "ExcedentesMensuales.frx":42D9
            Top             =   1440
            Width           =   240
         End
         Begin VB.Image imgPass 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   2
            Left            =   120
            Picture         =   "ExcedentesMensuales.frx":49F0
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgPass 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   1
            Left            =   120
            Picture         =   "ExcedentesMensuales.frx":5107
            Top             =   720
            Width           =   240
         End
         Begin VB.Image imgPass 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   0
            Left            =   120
            Picture         =   "ExcedentesMensuales.frx":581E
            Top             =   360
            Width           =   240
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   6012
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   10692
         _Version        =   1441793
         _ExtentX        =   18860
         _ExtentY        =   10604
         _StockProps     =   79
         Caption         =   "Periodo a consultar:"
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
         Begin XtremeSuiteControls.ListView lsw 
            Height          =   3975
            Left            =   0
            TabIndex        =   13
            Top             =   1080
            Width           =   10695
            _Version        =   1441793
            _ExtentX        =   18865
            _ExtentY        =   7011
            _StockProps     =   77
            BackColor       =   -2147483643
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
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkReporteResumen 
            Height          =   252
            Left            =   6960
            TabIndex        =   12
            Top             =   240
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Resumen?"
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
            Value           =   1
            Alignment       =   1
         End
         Begin XtremeSuiteControls.PushButton btnInforme 
            Height          =   492
            Left            =   8520
            TabIndex        =   9
            Top             =   480
            Width           =   1092
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Informe"
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
         Begin XtremeSuiteControls.ComboBox cboRepCorte 
            Height          =   312
            Left            =   3720
            TabIndex        =   7
            Top             =   600
            Width           =   2412
            _Version        =   1441793
            _ExtentX        =   4260
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
         Begin XtremeSuiteControls.ComboBox cboRepPeriodo 
            Height          =   312
            Left            =   0
            TabIndex        =   8
            Top             =   600
            Width           =   3732
            _Version        =   1441793
            _ExtentX        =   6588
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
         Begin XtremeSuiteControls.PushButton btnConsulta 
            Height          =   492
            Left            =   9600
            TabIndex        =   10
            Top             =   480
            Width           =   1092
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Consultar"
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
         Begin XtremeSuiteControls.ComboBox cboRepInforme 
            Height          =   312
            Left            =   6120
            TabIndex        =   11
            Top             =   600
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.PushButton btnExport 
            Height          =   255
            Index           =   0
            Left            =   10440
            TabIndex        =   80
            Top             =   5160
            Width           =   255
            _Version        =   1441793
            _ExtentX        =   444
            _ExtentY        =   444
            _StockProps     =   79
            Appearance      =   16
            Picture         =   "ExcedentesMensuales.frx":5F35
         End
         Begin VB.Label Label6 
            Caption         =   "Corte"
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
            Index           =   0
            Left            =   3720
            TabIndex        =   6
            Top             =   360
            Width           =   852
         End
         Begin VB.Label Label6 
            Caption         =   "Periodo"
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
            Index           =   1
            Left            =   0
            TabIndex        =   5
            Top             =   360
            Width           =   1092
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1815
         Left            =   -68920
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   9015
         _Version        =   1441793
         _ExtentX        =   15901
         _ExtentY        =   3201
         _StockProps     =   79
         Caption         =   "Distribuir Excedentes del Periodo [Mensual]:"
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
         Begin XtremeSuiteControls.ComboBox cboAPL_Corte 
            Height          =   312
            Left            =   4800
            TabIndex        =   21
            Top             =   720
            Width           =   2412
            _Version        =   1441793
            _ExtentX        =   4260
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
         Begin XtremeSuiteControls.ComboBox cboAPL_Periodo 
            Height          =   312
            Left            =   1080
            TabIndex        =   22
            Top             =   720
            Width           =   3732
            _Version        =   1441793
            _ExtentX        =   6588
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
         Begin XtremeSuiteControls.FlatEdit txtMonto 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   4800
            TabIndex        =   26
            Top             =   1080
            Width           =   2412
            _Version        =   1441793
            _ExtentX        =   4254
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
         Begin XtremeSuiteControls.FlatEdit txtTipoDistribucion 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   4800
            TabIndex        =   82
            Top             =   1440
            Width           =   2415
            _Version        =   1441793
            _ExtentX        =   4254
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777152
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Transparent     =   -1  'True
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo distribución:   "
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
            Index           =   19
            Left            =   2760
            TabIndex        =   81
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Monto a distribuir:   "
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
            Left            =   2760
            TabIndex        =   25
            Top             =   1080
            Width           =   1932
         End
         Begin VB.Label Label6 
            Caption         =   "Periodo"
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
            Left            =   1080
            TabIndex        =   24
            Top             =   480
            Width           =   1092
         End
         Begin VB.Label Label6 
            Caption         =   "Corte"
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
            Left            =   4800
            TabIndex        =   23
            Top             =   480
            Width           =   852
         End
      End
      Begin XtremeSuiteControls.CheckBox chkPosCierreCargaCero 
         Height          =   372
         Left            =   -69640
         TabIndex        =   15
         Top             =   1920
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Carga Info en Cero?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkPosCierreLimpiar 
         Height          =   612
         Left            =   -69640
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Limpia Aplicación Anterior?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton cmdAjustes 
         Height          =   615
         Left            =   -61240
         TabIndex        =   17
         Top             =   4560
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Ajustes a Excedentes"
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
      Begin XtremeSuiteControls.PushButton cmdCapitalizacionIndividual 
         Height          =   615
         Left            =   -61240
         TabIndex        =   18
         Top             =   5280
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Capitalización Extraordinaria"
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
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1815
         Left            =   -68800
         TabIndex        =   19
         Top             =   2520
         Visible         =   0   'False
         Width           =   9015
         _Version        =   1441793
         _ExtentX        =   15896
         _ExtentY        =   3196
         _StockProps     =   79
         Caption         =   "Información General:"
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
         Begin XtremeSuiteControls.FlatEdit txtAPLTotalAhorros 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   312
            Left            =   2520
            TabIndex        =   29
            Top             =   840
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
         Begin XtremeSuiteControls.FlatEdit txtAPLTotalAportes 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   312
            Left            =   6480
            TabIndex        =   30
            Top             =   840
            Width           =   2412
            _Version        =   1441793
            _ExtentX        =   4254
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
         Begin XtremeSuiteControls.FlatEdit txtAPLFactor 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   0
            EndProperty
            Height          =   312
            Left            =   2520
            TabIndex        =   34
            Top             =   1320
            Width           =   6372
            _Version        =   1441793
            _ExtentX        =   11239
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
         Begin XtremeSuiteControls.FlatEdit txtAPLCasosGeneral 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   312
            Left            =   2520
            TabIndex        =   28
            Top             =   360
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Factor de distribución:   "
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
            Index           =   8
            Left            =   0
            TabIndex        =   33
            Top             =   1320
            Width           =   2292
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Aporte Patronal:   "
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
            Index           =   7
            Left            =   4800
            TabIndex        =   32
            Top             =   720
            Width           =   1572
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Obrero + Capitalización:   "
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
            Index           =   6
            Left            =   360
            TabIndex        =   31
            Top             =   720
            Width           =   1932
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Casos:   "
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
            Left            =   360
            TabIndex        =   27
            Top             =   360
            Width           =   1932
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   1215
         Left            =   -68800
         TabIndex        =   20
         Top             =   4680
         Visible         =   0   'False
         Width           =   9015
         _Version        =   1441793
         _ExtentX        =   15896
         _ExtentY        =   2138
         _StockProps     =   79
         Caption         =   "Información aplicada:"
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
         Begin XtremeSuiteControls.FlatEdit txtAPLCasosProceso 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   312
            Left            =   2520
            TabIndex        =   36
            Top             =   360
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
         Begin XtremeSuiteControls.FlatEdit txtAPLTotalDistribuido 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   312
            Left            =   2520
            TabIndex        =   37
            Top             =   840
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Distribuido:   "
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
            Index           =   10
            Left            =   360
            TabIndex        =   38
            Top             =   840
            Width           =   1932
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Casos:   "
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
            Left            =   360
            TabIndex        =   35
            Top             =   360
            Width           =   1932
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox5 
         Height          =   3252
         Left            =   -69520
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   9372
         _Version        =   1441793
         _ExtentX        =   16531
         _ExtentY        =   5736
         _StockProps     =   79
         Caption         =   "Aplicación de Excedentes del Periodo:"
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
         Begin XtremeSuiteControls.ListView lswCRE_Renta 
            Height          =   1692
            Left            =   1560
            TabIndex        =   45
            Top             =   1080
            Width           =   6852
            _Version        =   1441793
            _ExtentX        =   12086
            _ExtentY        =   2984
            _StockProps     =   77
            BackColor       =   -2147483643
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
            Appearance      =   17
            UseVisualStyle  =   0   'False
            ShowBorder      =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnCierreParametros 
            Height          =   312
            Index           =   0
            Left            =   8640
            TabIndex        =   57
            ToolTipText     =   "Abrir Configuración del Parámetro"
            Top             =   600
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "..."
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.ComboBox cboCRE_Periodo 
            Height          =   312
            Left            =   1560
            TabIndex        =   41
            Top             =   600
            Width           =   3732
            _Version        =   1441793
            _ExtentX        =   6588
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
         Begin XtremeSuiteControls.FlatEdit txtCREPorcentajeCapitalizacion 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   312
            Left            =   7200
            TabIndex        =   42
            Top             =   600
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnCierreParametros 
            Height          =   312
            Index           =   1
            Left            =   8640
            TabIndex        =   58
            ToolTipText     =   "Abrir Configuración del Parámetro"
            Top             =   1080
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "..."
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.PushButton btnCierreParametros 
            Height          =   312
            Index           =   2
            Left            =   8640
            TabIndex        =   59
            ToolTipText     =   "Abrir Configuración del Parámetro"
            Top             =   2880
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "..."
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.CheckBox chkCRE_RentaAplCap 
            Height          =   252
            Left            =   3480
            TabIndex        =   60
            Top             =   2880
            Width           =   4932
            _Version        =   1441793
            _ExtentX        =   8700
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Renta aplica sobre la Capitalización?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            TextAlignment   =   1
            Appearance      =   16
            Alignment       =   1
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Tabla de Renta:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   612
            Index           =   18
            Left            =   240
            TabIndex        =   61
            Top             =   1080
            Width           =   1092
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Periodo:"
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
            Index           =   12
            Left            =   240
            TabIndex        =   44
            Top             =   600
            Width           =   1092
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "[%] Capitalización:   "
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
            Index           =   11
            Left            =   5160
            TabIndex        =   43
            Top             =   600
            Width           =   1932
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox6 
         Height          =   2292
         Left            =   -69280
         TabIndex        =   46
         Top             =   3960
         Visible         =   0   'False
         Width           =   9252
         _Version        =   1441793
         _ExtentX        =   16319
         _ExtentY        =   4043
         _StockProps     =   79
         Caption         =   "Cálculo para aplicación:"
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
         Begin XtremeSuiteControls.FlatEdit txtCRECasos 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   312
            Left            =   6240
            TabIndex        =   47
            Top             =   360
            Width           =   972
            _Version        =   1441793
            _ExtentX        =   1714
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
         Begin XtremeSuiteControls.FlatEdit txtCREExcedenteBruto 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   312
            Left            =   2760
            TabIndex        =   48
            Top             =   360
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
         Begin XtremeSuiteControls.FlatEdit txtCREExcedenteNeto 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   312
            Left            =   2760
            TabIndex        =   51
            Top             =   1440
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
         Begin XtremeSuiteControls.FlatEdit txtCRERenta 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   312
            Left            =   2760
            TabIndex        =   53
            Top             =   1080
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
         Begin XtremeSuiteControls.FlatEdit txtCRECapitalizado 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   312
            Left            =   2760
            TabIndex        =   55
            Top             =   720
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4043
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "[-] Capitalización:   "
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
            Index           =   17
            Left            =   600
            TabIndex        =   56
            Top             =   720
            Width           =   1932
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "[-] Impuesto Renta:   "
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
            Index           =   16
            Left            =   600
            TabIndex        =   54
            Top             =   1080
            Width           =   1932
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Excedente Neto:   "
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
            Index           =   15
            Left            =   600
            TabIndex        =   52
            Top             =   1440
            Width           =   1932
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Casos:   "
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
            Index           =   14
            Left            =   4080
            TabIndex        =   50
            Top             =   360
            Width           =   1932
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Excedente Bruto:   "
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
            Index           =   13
            Left            =   600
            TabIndex        =   49
            Top             =   360
            Width           =   1932
         End
      End
      Begin XtremeSuiteControls.PushButton btnCierreParametros 
         Height          =   312
         Index           =   3
         Left            =   -61240
         TabIndex        =   63
         ToolTipText     =   "Abrir Configuración del Parámetro"
         Top             =   480
         Visible         =   0   'False
         Width           =   372
         _Version        =   1441793
         _ExtentX        =   656
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "..."
         Transparent     =   -1  'True
         Appearance      =   16
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Ultimo Cierre"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   252
         Index           =   0
         Left            =   -69640
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1932
      End
      Begin VB.Label lblUltimoCierre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   -67600
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   8652
      End
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   6120
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExcedentesMensuales.frx":609F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExcedentesMensuales.frx":63BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExcedentesMensuales.frx":6C97
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExcedentesMensuales.frx":7573
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption lblStatus 
      Height          =   252
      Left            =   0
      TabIndex        =   64
      Top             =   7440
      Width           =   11292
      _Version        =   1441793
      _ExtentX        =   19918
      _ExtentY        =   444
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.41
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cálculo y Distribución de Excedentes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   11
      Left            =   1920
      TabIndex        =   14
      Top             =   360
      Width           =   9252
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frmAH_ExcedentesMensuales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsParametros As New ADODB.Recordset
Dim vReporte As String, vPaso As Boolean

Private Sub btnAplicar_Click()
    Select Case tcMain.Selected.Index
      Case 1 'Aplica
        Call sbExcedentes_Corte_Aplica
      Case 2 'Cierre
        Call sbExcedentes_Cierre_Aplica
      Case 3 'PosCierre
        Call sbExcedentes_PosCierre_Aplica
    End Select
End Sub

Private Sub btnCierreParametros_Click(Index As Integer)
Select Case Index
 Case 0 'Periodo
    frmAH_ExcedentesPeriodos.Show vbModal, Me
 Case 1 'Tabla de Renta
    frmAH_ExcedentesRenta_Tabla.Show vbModal, Me
 Case 2, 3 'Parametros Generales
    frmAH_ExcedentesParametros.Show vbModal, Me
End Select

Call cboCRE_Periodo_Click

End Sub

Private Sub btnConsulta_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curBruto As Currency


On Error GoTo vError

With lsw.ColumnHeaders
    .Clear
    .Add , , "Corte", 2200
    .Add , , "Casos", 1100, vbRightJustify
    .Add , , "Base Cálculo", 2200, vbRightJustify
    .Add , , "Exc. Bruto", 2200, vbRightJustify
    .Add , , "Factor de Distribución", 3000
End With


curBruto = 0

strSQL = "select * " _
       & "      From vExc_Periodos_Cortes_Resumen" _
       & "      Where id_periodo = " & cboRepPeriodo.ItemData(cboRepPeriodo.ListIndex) _
       & "      order by corte desc"

Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Corte)
     itmX.SubItems(1) = Format(rs!Casos, "###,###,##0")
     itmX.SubItems(2) = Format(rs!Base, "Standard")
     itmX.SubItems(3) = Format(rs!Bruto, "Standard")
     itmX.SubItems(4) = Space(10) & (rs!Bruto / rs!Base)
     curBruto = curBruto + rs!Bruto
     
 rs.MoveNext
Loop
rs.Close

 Set itmX = lsw.ListItems.Add(, , "")
     itmX.SubItems(3) = "_____________"

 Set itmX = lsw.ListItems.Add(, , "")
     itmX.SubItems(3) = Format(curBruto, "Standard")
     itmX.Bold = True

Exit Sub

vError:

  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnExport_Click(Index As Integer)
Call Excel_Exportar_Lsw(lsw)
End Sub

Private Sub btnInforme_Click()
Dim strSQL As String
Dim pCorte As Date, pCorteFiltro As String


On Error GoTo vError


Me.MousePointer = vbHourglass

If cboRepCorte.Text = "TODOS" Then
    pCorte = Date
Else
    pCorte = cboRepCorte.ItemData(cboRepCorte.ListIndex)
End If
pCorteFiltro = " in Date (" & Format(pCorte, "yyyy,mm,dd") & ") to Date (" & Format(pCorte, "yyyy,mm,dd") & ")"

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Excedentes - Reportes"
    
    .Connect = glogon.ConectRPT
     
    .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fecha='" & Format(fxFechaServidor, "DD/MM/YYYY") & "'"
    .Formulas(3) = "usuario='" & UCase(glogon.Usuario) & "'"
    
    

    Select Case cboRepInforme.ItemData(cboRepInforme.ListIndex)
     Case "CARMES"
       .SelectionFormula = "{EXC_CARGA.CORTE} " & pCorteFiltro _
              & " and {EXC_CARGA.ID_PERIODO} = " & cboRepPeriodo.ItemData(cboRepPeriodo.ListIndex)
       
       .Formulas(4) = "subtitulo='PERIODO: " & cboRepPeriodo.Text _
                    & " ¦ CORTE : " & cboRepCorte.Text & "'"
       .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_CARGADOS.rpt")

       Call Bitacora("Imprime", "Listado Exc. Cargado Per." & cboRepPeriodo.Text)

     Case "APLMES"
       If chkReporteResumen.Value = vbChecked Then
              '.SelectionFormula = "{EXC_CARGA.ID_PERIODO} = " & cboRepPeriodo.ItemData(cboRepPeriodo.ListIndex)
              
               .SelectionFormula = "{EXC_CARGA.CORTE} " & pCorteFiltro _
                       & " and {EXC_CARGA.ID_PERIODO} = " & cboRepPeriodo.ItemData(cboRepPeriodo.ListIndex)
                
                .Formulas(4) = "subtitulo='PERIODO: " & cboRepPeriodo.Text _
                           & " --> CORTE APLICADO'"
              .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_AplicadoPeriodo.rpt")
              
              Call Bitacora("Imprime", "Resumen Exc. Aplicado Per." & cboRepPeriodo.Text)

       Else
                .SelectionFormula = "{EXC_CARGA.CORTE} " & pCorteFiltro _
                       & " and {EXC_CARGA.ID_PERIODO} = " & cboRepPeriodo.ItemData(cboRepPeriodo.ListIndex)
                
                .Formulas(4) = "subtitulo='PERIODO: " & cboRepPeriodo.Text _
                             & " ¦ CORTE : " & cboRepCorte.Text & "'"
                           
              .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_APLICADO.rpt")

              Call Bitacora("Imprime", "Listado Exc. Aplicado Per." & cboRepPeriodo.Text)
       End If
       
     Case "APLPERIODO"
     
              .SelectionFormula = "{EXC_CARGA.ID_PERIODO} = " & cboRepPeriodo.ItemData(cboRepPeriodo.ListIndex)
              
              .Formulas(4) = "subtitulo='PERIODO: " & cboRepPeriodo.Text _
                           & " --> CORTE APLICADO'"
              .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_AplicadoPeriodo_Total.rpt")
              
              Call Bitacora("Imprime", "Resumen Exc. Aplicado TOTAL Per." & cboRepPeriodo.Text)
     
     
     Case "CIERREORDINARIO"
      
        .SelectionFormula = "{EXC_CIERRE.ID_PERIODO} = " & cboRepPeriodo.ItemData(cboRepPeriodo.ListIndex)
        .Formulas(4) = "subtitulo='PERIODO: " & cboRepPeriodo.Text & "'"

      If chkReporteResumen.Value = 1 Then
       .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_CIERRERESUMEN.rpt")
      Else
       .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_CIERREORDINARIO.rpt")
      End If

       Call Bitacora("Imprime", "Listado Exc. Cierre Per." & cboRepPeriodo.Text)

     Case "CIERRERESUMEN" 'POSCIERRE
        .SelectionFormula = "{EXC_CIERRE.ID_PERIODO} = " & cboRepPeriodo.ItemData(cboRepPeriodo.ListIndex)
        .Formulas(4) = "subtitulo='PERIODO: " & cboRepPeriodo.Text & "'"
      
      If chkReporteResumen.Value = 1 Then
       .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_POS_CIERRE_RESUMEN.rpt")
      Else
       .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_POS_CIERRE.rpt")
      End If

      Call Bitacora("Imprime", "Listado Exc. Pos-Cierre Per." & cboRepPeriodo.Text)

     Case "CAPIND"
       .Formulas(4) = "subtitulo='LISTADO GENERAL'"
       .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_CAP_IND.rpt")

     Case "AJUSTES"
       .Formulas(4) = "subtitulo='AJUSTES PENDIENTES DE APLICAR'"
       .SelectionFormula = "{EXC_AJUSTE.ESTADO}='P'"
       .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_LISTADOAJUSTES.rpt")

    End Select

    .Action = 1
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboAPL_Corte_Click()

If vPaso Then Exit Sub
If cboAPL_Corte.ListCount = 0 Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset
Dim pAnio As Long, pMes As Integer, pFecha As Date

On Error GoTo vError

Me.MousePointer = vbHourglass

pFecha = CDate(cboAPL_Corte.ItemData(cboAPL_Corte.ListIndex))
pAnio = Year(pFecha)
pMes = Month(pFecha)

'Resultado Contable o Manual
If txtTipoDistribucion.Tag = "R" Or txtTipoDistribucion.Tag = "M" Then
    strSQL = "select dbo.fxCntX_Utilidad_Mes_SinCF(" & pAnio & "," & pMes & "," & GLOBALES.gEnlace & ", '','') as 'Excedente'"
    Call OpenRecordSet(rs, strSQL)
    If Not glogon.error Then
      txtMonto.Text = Format(rs!Excedente, "Standard")
    End If
End If

'spExc_Mnt_Distribuir_Dato(@PeriodoId int, @Corte datetime, @Tipo char(1))

If txtTipoDistribucion.Tag <> "R" Then
    strSQL = "exec spExc_Mnt_Distribuir_Dato " & cboAPL_Periodo.ItemData(cboAPL_Periodo.ListIndex) _
           & ", '" & Format(pFecha, "yyyy-mm-dd") & " 23:59', '" & txtTipoDistribucion.Tag & "'"

    Call OpenRecordSet(rs, strSQL)
    If Not glogon.error And Not rs.EOF And Not rs.BOF Then
      txtMonto.Text = Format(rs!Monto, "Standard")
    End If
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub cboAPL_Periodo_Click()
If vPaso Then Exit Sub
If cboAPL_Periodo.ListCount = 0 Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vPaso = True

strSQL = "select CORTE_DATETIME_STR as 'IdX', CORTE_DATE_STR as 'ItmX' " _
       & " from vExc_Periodos_Cortes" _
       & " Where id_periodo = " & cboAPL_Periodo.ItemData(cboAPL_Periodo.ListIndex) _
       & " order by CORTE desc"

Call sbCbo_Llena_New(cboAPL_Corte, strSQL, False, True)

vPaso = False

strSQL = "select ESTADO, TIPO_APL_MENSUAL, TIPO_APL_MENSUAL_DESC" _
       & "  From vExc_Periodos_Consulta" _
       & "  Where ID_PERIODO = " & cboAPL_Periodo.ItemData(cboAPL_Periodo.ListIndex)
Call OpenRecordSet(rs, strSQL)

txtTipoDistribucion.Tag = rs!TIPO_APL_MENSUAL
txtTipoDistribucion.Text = rs!TIPO_APL_MENSUAL_DESC

txtMonto.Locked = True
txtMonto.Text = Format(0, "Standard")

If rs!TIPO_APL_MENSUAL = "M" Then
    txtMonto.Locked = False
End If
rs.Close


If cboAPL_Corte.ListCount > 0 Then
    strSQL = "exec spExc_Mnt_Distribuir_Dato " & cboAPL_Periodo.ItemData(cboAPL_Periodo.ListIndex) _
           & ", '" & cboAPL_Corte.Text & "', '" & txtTipoDistribucion.Tag & "'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
        txtMonto.Text = Format(rs!Monto, "Standard")
    End If
    
    rs.Close
End If
       
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub

Private Sub cboCRE_Periodo_Click()

If vPaso Then Exit Sub
If cboCRE_Periodo.ListCount = 0 Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

On Error GoTo vError

vPaso = True

'Capitalizacion
strSQL = "select CAPITALIZA_PORC , CAPITALIZA_RENTA_APLICA" _
       & " From EXC_PERIODOS" _
       & " Where id_periodo = " & cboCRE_Periodo.ItemData(cboCRE_Periodo.ListIndex)
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtCREPorcentajeCapitalizacion.Text = Format(rs!Capitaliza_Porc, "Standard")
End If
rs.Close

'Cargar Tabla de Renta
strSQL = "select DESDE, HASTA, PORCENTAJE " _
       & " From EXC_RENTA_TABLA" _
       & " order by desde"
Call OpenRecordSet(rs, strSQL)

With lswCRE_Renta.ListItems
.Clear
Do While Not rs.EOF
  Set itmX = .Add(, , Format(rs!Desde, "Standard"))
      itmX.SubItems(1) = Format(rs!Hasta, "Standard")
      itmX.SubItems(2) = Format(rs!Porcentaje, "Standard")
  rs.MoveNext
Loop
rs.Close
       
End With

'Pre-Calculo de aplicacion

txtCRECasos.Text = "0"
txtCREExcedenteBruto.Text = Format(0, "Standard")
txtCRECapitalizado.Text = Format(0, "Standard")
txtCRERenta.Text = Format(0, "Standard")
txtCREExcedenteNeto.Text = Format(0, "Standard")

vPaso = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  


End Sub

Private Sub cboRepPeriodo_Click()
If vPaso Then Exit Sub
If cboRepPeriodo.ListCount = 0 Then Exit Sub

Dim strSQL As String

On Error GoTo vError

vPaso = True

strSQL = "select CORTE_DATETIME_STR as 'IdX', CORTE_DATE_STR as 'ItmX' " _
       & " from vExc_Periodos_Cortes" _
       & " Where id_periodo = " & cboRepPeriodo.ItemData(cboRepPeriodo.ListIndex) _
       & " order by CORTE desc"

Call sbCbo_Llena_New(cboRepCorte, strSQL, False, True)

vPaso = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub cmdAjustes_Click()
 Call sbFormsCall("frmAH_ExcedentesAjuste", , , , False, Me, True)
End Sub

Private Sub cmdCapitalizacionIndividual_Click()
 Call sbFormsCall("frmAH_ExcedentesCapInd", , , , False, Me, True)
End Sub



Private Sub Form_Load()
Dim strSQL As String

 vModulo = 2

 Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

With lswParametros.ColumnHeaders
    .Clear
    .Add , , "Código", 1000
    .Add , , "Descripción", 5500
    .Add , , "Valor", 2000, vbCenter
End With
 
 With lswCRE_Renta.ColumnHeaders
    .Clear
    .Add , , "Inicio", 1800, vbRightJustify
    .Add , , "Corte", 1800, vbRightJustify
    .Add , , "[%] Renta", 1100, vbRightJustify
 End With
 
 With cboRepInforme
  .Clear
  .AddItem "Cargado del Mes"
  .ItemData(.ListCount - 1) = CStr("CARMES")
  
  .AddItem "Aplicado del Mes"
  .ItemData(.ListCount - 1) = CStr("APLMES")
  
  .AddItem "Aplicado Total Periodo"
  .ItemData(.ListCount - 1) = CStr("APLPERIODO")
  
  .AddItem "Cierre Excedentes"
  .ItemData(.ListCount - 1) = CStr("CIERREORDINARIO")
  
  .AddItem "Cierre + Aplicaciones"
  .ItemData(.ListCount - 1) = CStr("CIERRERESUMEN")
  
  .AddItem "Capitalización Extra"
  .ItemData(.ListCount - 1) = CStr("CAPIND")
  
  .AddItem "Ajustes Realizados"
  .ItemData(.ListCount - 1) = CStr("AJUSTES")
  
  
  
  .Text = "Cargado del Mes"
End With

 
 tcMain.Item(0).Selected = True
 

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
End Sub


Function fxMesCargado(i As Integer) As String
Select Case i
  Case 1
   fxMesCargado = "Enero"
  Case 2
   fxMesCargado = "Febrero"
  Case 3
   fxMesCargado = "Marzo"
  Case 4
   fxMesCargado = "Abril"
  Case 5
   fxMesCargado = "Mayo"
  Case 6
   fxMesCargado = "Junio"
  Case 7
   fxMesCargado = "Julio"
  Case 8
   fxMesCargado = "Agosto"
  Case 9
   fxMesCargado = "Setiembre"
  Case 10
   fxMesCargado = "Octubre"
  Case 11
   fxMesCargado = "Noviembre"
  Case 12
   fxMesCargado = "Diciembre"
  Case Else
   fxMesCargado = "Ninguno"
End Select
End Function



Private Sub sbSalidas_Separa(pPeriodoId As Long)
Dim strSQL As String


On Error GoTo vError


lblStatus.Caption = "Procesando Separación de Salida [Espere!]"
DoEvents

Me.MousePointer = vbHourglass

strSQL = "exec spExc_Procesos_Salidas_Separa " & pPeriodoId & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


Me.MousePointer = vbDefault

lblStatus.Caption = ""

MsgBox "Proceso Finalizado...", vbInformation


Exit Sub

vError:
    lblStatus.Caption = ""
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbInformation
End Sub




Private Sub sbSalidas_Fondos(pPeriodoId As Long)
Dim rs As New ADODB.Recordset, strSQL As String

Dim pFndOperadora As Integer, pFndPlan As String, pSalida As String
Dim rsMain As New ADODB.Recordset

MousePointer = vbHourglass

On Error GoTo vError


lblStatus.Caption = "Cargando Información Base..."
DoEvents


strSQL = " select COD_SALIDA, DESCRIPCION, DESTINO_OPERADORA, DESTINO_PLAN " _
       & " From EXC_TIPOS_SALIDAS " _
       & " WHERE DESTINO_PLAN <> ''" _
       & " AND ('Salida: ' + COD_SALIDA) NOT IN(SELECT DETALLE FROM EXC_PERIODOS_BITACORA WHERE COD_PROCESO = '12' and ID_PERIODO = " & pPeriodoId & ")" _
       & " AND COD_SALIDA IN(SELECT COD_SALIDA FROM vExc_Cierre_Salida_Rsm WHERE ID_PERIODO = " & pPeriodoId _
       & " AND EXCEDENTE_FINAL > 0)"
Call OpenRecordSet(rsMain, strSQL)

Do While Not rsMain.EOF
  pSalida = rsMain!Cod_Salida
  pFndOperadora = rsMain!Destino_Operadora
  pFndPlan = rsMain!Destino_Plan
  
   lblStatus.Caption = "Procesando Salidas: " & rsMain!Descripcion
   DoEvents
   
   strSQL = "exec spExc_Procesos_Salidas_Fondos " & pPeriodoId & ", '" & pSalida & "', '" & glogon.Usuario & "'"
   Call OpenRecordSet(rs, strSQL)
        'Comprobante de aplicacion
        Call sbExcedenteDocumento("Salida", rs!TipoDoc, rs!NumDoc, rs!Concepto, rs!Cuenta, pPeriodoId, pSalida)
   rs.Close
       
  rsMain.MoveNext
Loop
rsMain.Close

  
    
prgBar.Value = 1
lblStatus.Caption = ""
MousePointer = vbDefault


MsgBox "Fondos Acreditados Satisfactoriamente...", vbInformation

Exit Sub

vError:
  MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbExcedentes_PosCierre_Aplica()
Dim rs As New ADODB.Recordset, strSQL As String
Dim pPeriodoId As Long, pPeriodoDesc As String

On Error GoTo vError


pPeriodoId = lblUltimoCierre.Tag
pPeriodoDesc = Mid(lblUltimoCierre.Caption, 1, 17)

glogon.Conection.CommandTimeout = 1300

Select Case True
  
  Case optAplicaciones(0).Value 'Carga y Aplica Donaciones
    If Not fxExcedente_Bitacora_Valida(pPeriodoId, "04", "Actualiza") Then
        MsgBox "Este proceso no se puede realizar porque ya se aplicó la morosidad!", vbExclamation
        Exit Sub
    End If
       
       Call sbExcedentesAplicaDonacion(pPeriodoId, Me)
       
  Case optAplicaciones(1).Value 'Carga y Aplica Ajustes
    If Not fxExcedente_Bitacora_Valida(pPeriodoId, "04", "Actualiza") Then
        MsgBox "Este proceso no se puede realizar porque ya se aplicó la morosidad!", vbExclamation
        Exit Sub
    End If
       
       Call sbExcedentesCargaAjuste(pPeriodoId, Me)
       
       Call sbExcedentesAplicaAjuste(pPeriodoId, Me)
  
  Case optAplicaciones(2).Value 'Carga y Aplica Creditos Sobre Excedentes
    If Not fxExcedente_Bitacora_Valida(pPeriodoId, "04", "Actualiza") Then
        MsgBox "Este proceso no se puede realizar porque ya se aplicó la morosidad!", vbExclamation
        Exit Sub
    End If
       
       Call sbExcedentesCargaAplCreditosASE(pPeriodoId, Me)

  Case optAplicaciones(3).Value 'Carga Mora/Aplica y Actualiza
    
    If Not fxExcedente_Bitacora_Valida(pPeriodoId, "04", "Actualiza") Then
        MsgBox "Este proceso no se puede realizar porque ya se aplicó la morosidad!", vbExclamation
        Exit Sub
    End If
      
       Call sbExcedentesCargaAplMora(pPeriodoId, Me)
  
       If chkPosCierreCargaCero.Value = vbUnchecked Then 'Si se cargo información aplicarla
           
           If GLOBALES.SysPlanPagos = 1 Then
               Call sbExcedentesActualizaMora_PlanPagos(pPeriodoId, Me)
           Else
               Call sbExcedentesActualizaMora(pPeriodoId, Me)
           End If

       Else
            Call sbExcedente_Bitacora(pPeriodoId, "04", "Actualiza")

       End If

  
  Case optAplicaciones(4).Value 'Carga/Aplica/Actualiza Operaciones OPCF y su Mora Como prioridad
    
    If Not fxExcedente_Bitacora_Valida(pPeriodoId, "05", "Actualiza") Then
        MsgBox "Este proceso no se puede realizar porque ya se aplicó anteriormente!", vbExclamation
        Exit Sub
    End If
       
       Call sbExcedentesCargaAplMoraOPCF(pPeriodoId, Me)
  
       If chkPosCierreCargaCero.Value = vbUnchecked Then 'Si se cargo información aplicarla
           
           If GLOBALES.SysPlanPagos = 1 Then
                Call sbExcedentesActualizaMoraOPCF_PlanPagos(pPeriodoId, Me)
           Else
                Call sbExcedentesActualizaMoraOPCF(pPeriodoId, Me)
           End If
       Else
            Call sbExcedente_Bitacora(pPeriodoId, "05", "Actualiza")
       
       End If
  
  Case optAplicaciones(5).Value 'Aplica Capitalizacion Individual
    If Not fxExcedente_Bitacora_Valida(pPeriodoId, "09", "Actualiza") Then
        MsgBox "Este proceso no se puede realizar porque ya se aplicó anteriormente!", vbExclamation
        Exit Sub
    End If
       
       Call sbExcedentesAplicaCapInd(pPeriodoId, Me)

  
  Case optAplicaciones(6).Value 'Actualiza Capitalizacion Individual
    If Not fxExcedente_Bitacora_Valida(pPeriodoId, "09", "Actualiza") Then
        MsgBox "Este proceso no se puede realizar porque ya se aplicó anteriormente!", vbExclamation
        Exit Sub
    End If
       Call sbExcedentesActualizaCapIndFND(pPeriodoId, Me)
       

  Case optAplicaciones(7).Value 'Actualiza Capitalizacion General
    If Not fxExcedente_Bitacora_Valida(pPeriodoId, "08", "Actualiza") Then
        MsgBox "Este proceso no se puede realizar porque ya se aplicó anteriormente!", vbExclamation
        Exit Sub
    End If
       Call sbExcedentesActualizaCapGeneral(pPeriodoId, Me)
  
  Case optAplicaciones(8).Value 'Actualiza Ajustes
    If Not fxExcedente_Bitacora_Valida(pPeriodoId, "03", "Actualiza") Then
        MsgBox "Este proceso no se puede realizar porque ya se aplicó anteriormente!", vbExclamation
        Exit Sub
    End If
       Call sbExcedentesActualizaAjustes(pPeriodoId, Me)

  Case optAplicaciones(9).Value 'Actualiza Creditos con Garantia Sobre Excedentes
    If Not fxExcedente_Bitacora_Valida(pPeriodoId, "06", "Actualiza") Then
        MsgBox "Este proceso no se puede realizar porque ya se aplicó anteriormente!", vbExclamation
        Exit Sub
    End If
       Call sbExcedentesActualizaASE(pPeriodoId, Me)

  Case optAplicaciones(10).Value 'Genera Asiento General de Excedentes
    If Not fxExcedente_Bitacora_Valida(pPeriodoId, "11", "") Then
        MsgBox "Este proceso no se puede realizar porque ya se aplicó anteriormente!", vbExclamation
        Exit Sub
    End If
       Call sbExcedentesAsientoGeneral(pPeriodoId, Me)

  Case optAplicaciones(11).Value 'Salidas Separa
    If Not fxExcedente_Bitacora_Valida(pPeriodoId, "12", "Actualiza") Then
        MsgBox "Este proceso no se puede realizar porque ya se aplicó anteriormente!", vbExclamation
        Exit Sub
    End If
       Call sbSalidas_Separa(pPeriodoId)

  Case optAplicaciones(12).Value 'Acredita los Fondos para Pago de Excedentes
    If Not fxExcedente_Bitacora_Valida(pPeriodoId, "12", "Actualiza") Then
        MsgBox "Este proceso no se puede realizar porque ya se aplicó anteriormente!", vbExclamation
        Exit Sub
    End If
       Call sbSalidas_Fondos(pPeriodoId)
       Call sbExcedente_Bitacora(pPeriodoId, "12", "Actualiza")



End Select

lblStatus.Caption = ""

'Actualiza Seguimiento
Call sbAplicacion_Pass

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbAplicacion_Pass()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
 
 
For i = 0 To imgPass.Count - 1
    imgPass.Item(i).Visible = False
Next i

strSQL = "select * from EXC_PERIODOS_BITACORA where id_periodo = " & lblUltimoCierre.Tag
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Select Case rs!cod_Proceso
    Case "02" 'Donacion
        imgPass.Item(0).Visible = True
    
    Case "03" '[Aplica] Ajustes
        Select Case rs!Detalle
            Case "Aplica"
                imgPass.Item(1).Visible = True
            Case "Actualiza"
                imgPass.Item(8).Visible = True
        End Select
        
    
    Case "04" 'Morosidad
        Select Case rs!Detalle
            Case "Actualiza"
                imgPass.Item(3).Visible = True
            Case Else
        End Select
    
    Case "05" 'OPCF
        Select Case rs!Detalle
            Case "Actualiza"
                imgPass.Item(4).Visible = True
            Case Else
        End Select
    
    Case "06" 'Credito s/Excedentes
        Select Case rs!Detalle
            Case "Carga/Aplica"
                imgPass.Item(2).Visible = True
            Case "Actualiza"
                imgPass.Item(9).Visible = True
        End Select
    
    Case "07" 'No se Utiliza
    
    Case "08" 'Capitalizacion General
        imgPass.Item(7).Visible = True
    
    Case "09" 'Capitalizacion Extraordinaria
        Select Case rs!Detalle
            Case "Aplica"
                imgPass.Item(5).Visible = True
            Case "Actualiza"
                imgPass.Item(6).Visible = True
        End Select
    
    Case "10" 'No se Utiliza
    Case "11" 'Asiento de Excedentes
        imgPass.Item(10).Visible = True
        
    Case "12" 'Salidas
           Select Case rs!Detalle
            Case "" 'Salida sin Descripcion es que ya se proceso la separacion
                imgPass.Item(11).Visible = True
            Case "Actualiza" 'Traslado al Fondo
                imgPass.Item(12).Visible = True
        End Select
        
 End Select
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Select Case Item.Index
    Case 3 'Pos Cierre
       strSQL = "select *  From vExc_Periodos" _
              & " where idx in(select max(idx) from vExc_Periodos where estado = 'C')"
       Call OpenRecordSet(rs, strSQL)
       If Not rs.BOF And Not rs.BOF Then
            lblUltimoCierre.Tag = rs!IdX
            lblUltimoCierre.Caption = rs!itmX
       Else
            lblUltimoCierre.Tag = 0
            lblUltimoCierre.Caption = "No Existen Periodos Cerrados!"
      
       End If
       rs.Close
    
       optAplicaciones(0).Value = True
       Call sbAplicacion_Pass
       
    Case 4 'Parametros
        lswParametros.ListItems.Clear
        
       strSQL = "select * from exc_Parametros order by cod_Parametro"
       Call OpenRecordSet(rs, strSQL)
       Do While Not rs.EOF
         Set itmX = lswParametros.ListItems.Add(, , rs!Cod_Parametro)
             itmX.SubItems(1) = rs!Descripcion
             itmX.SubItems(2) = rs!Valor
         rs.MoveNext
       Loop
       rs.Close

End Select

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset

TimerX.Interval = 0
TimerX.Enabled = False

Me.MousePointer = vbHourglass

vPaso = True

strSQL = "select IdX, ItmX from vExc_Periodos order by Idx desc"
Call sbCbo_Llena_New(cboRepPeriodo, strSQL, False, True)


strSQL = "select IdX, ItmX from vExc_Periodos where estado in('P','A') order by Idx desc"
Call sbCbo_Llena_New(cboAPL_Periodo, strSQL, False, True)

strSQL = "select IdX, ItmX from vExc_Periodos where estado in('P','A') order by Idx desc"
Call sbCbo_Llena_New(cboCRE_Periodo, strSQL, False, True)

vPaso = False


Call cboRepPeriodo_Click
Call cboAPL_Periodo_Click
Call cboCRE_Periodo_Click

Me.MousePointer = vbDefault

End Sub


Private Sub txtCREPorcentajeCapitalizacion_KeyPress(KeyAscii As Integer)
' Set GLOBALES.gCajaTxt = txtCREPorcentajeCapitalizacion
' ValidaMonto
End Sub

Private Sub txtCREPorcentajeRenta_KeyPress(KeyAscii As Integer)
' Set GLOBALES.gCajaTxt = txtCREPorcentajeRenta
' ValidaMonto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
' Set GLOBALES.gCajaTxt = txtMonto
' ValidaMonto
End Sub


Private Sub sbExcedentes_Corte_Aplica()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curAhorro As Currency, curAporte As Currency
Dim pPeriodoId As Long, pCorte As Date
Dim dlbFactor As Double, curMontoDistribuido As Currency, curMonto As Currency

If cboAPL_Periodo.ListCount = 0 Then Exit Sub
If cboAPL_Corte.ListCount = 0 Then Exit Sub


On Error GoTo vError

If (MsgBox("Está seguro que desea aplicar distribución de Excedentes Brutos a este mes... " & cboAPL_Corte.Text & "", vbYesNo)) = vbNo Then Exit Sub
  
  
 Me.MousePointer = vbDefault


 lblStatus.Caption = "Cargando Parametros e Información General ..."
 
 
 pPeriodoId = cboAPL_Periodo.ItemData(cboAPL_Periodo.ListIndex)
 pCorte = CDate(cboAPL_Corte.ItemData(cboAPL_Corte.ListIndex))

 If fxValidaAplicacion(pPeriodoId, pCorte) Then
      
      lblStatus.Caption = "Procesando Espere..."
      'spExc_Cierre_Aplicacion_Mensual(@PeriodoId int,  @Usuario varchar(30), @Corte datetime, @Monto dec(18,2), @TipoApl varchar(10) = 'M')
          
      strSQL = "exec spExc_Cierre_Aplicacion_Mensual " & pPeriodoId & ", '" & glogon.Usuario & "', '" & Format(pCorte, "YYYY-MM-DD hh:mm:ss") _
             & "', " & CCur(txtMonto.Text) & ", '" & txtTipoDistribucion.Tag & "'"
      Call OpenRecordSet(rs, strSQL)

       curAhorro = IIf(IsNull(rs!ahorro), 0, rs!ahorro) + IIf(IsNull(rs!aholiq), 0, rs!aholiq)
       curAporte = IIf(IsNull(rs!Aporte), 0, rs!Aporte) + IIf(IsNull(rs!apoliq), 0, rs!apoliq)
       curAporte = curAporte + IIf(IsNull(rs!capitaliza), 0, rs!capitaliza) + IIf(IsNull(rs!capliq), 0, rs!capliq)
    
    
        txtAPLCasosGeneral.Text = IIf(IsNull(rs!total), 0, rs!total)
        txtAPLTotalAhorros.Text = Format(curAhorro, "Standard")
        txtAPLTotalAportes.Text = Format(curAporte, "Standard")
        txtAPLFactor.Text = rs!Factor
        txtAPLTotalDistribuido = Format(rs!Excedente, "Standard")
        txtAPLCasosProceso = rs!total
     
     rs.Close
     
     
 End If 'Fin de la Validacion de la Aplicacion
 
 
 MsgBox "Aplicacion Finalizada Satisfactoriamente...", vbInformation
 
 lblStatus.Caption = ""
 prgBar.Value = 0
 
 Me.MousePointer = vbDefault


txtMonto.Text = "0"
txtMonto.SetFocus

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Function fxMesesCierre(intTipoCierre As Integer, intNumeroCierre As Integer) As String
Dim strMeses As String
Select Case intTipoCierre
 Case 1
    Select Case intNumeroCierre
      Case 1
        fxMesesCierre = "10"
      Case 2
        fxMesesCierre = "11"
      Case 3
        fxMesesCierre = "12"
      Case 4
        fxMesesCierre = "1"
      Case 5
        fxMesesCierre = "2"
      Case 6
        fxMesesCierre = "3"
      Case 7
        fxMesesCierre = "4"
      Case 8
        fxMesesCierre = "5"
      Case 9
        fxMesesCierre = "6"
      Case 10
        fxMesesCierre = "7"
      Case 11
        fxMesesCierre = "8"
      Case 12
        fxMesesCierre = "9"
    End Select
 
 Case 2
    Select Case intNumeroCierre
      Case 1
        fxMesesCierre = "10,11"
      Case 2
        fxMesesCierre = "12,1"
      Case 3
        fxMesesCierre = "2,3"
      Case 4
        fxMesesCierre = "4,5"
      Case 5
        fxMesesCierre = "6,7"
      Case 6
        fxMesesCierre = "8,9"
    End Select
 
 Case 3
    Select Case intNumeroCierre
      Case 1
        fxMesesCierre = "10,11,12"
      Case 2
        fxMesesCierre = "1,2,3"
      Case 3
        fxMesesCierre = "4,5,6"
      Case 4
        fxMesesCierre = "7,8,9"
    End Select
 Case 4
    Select Case intNumeroCierre
      Case 1
        fxMesesCierre = "10,11,12,1"
      Case 2
        fxMesesCierre = "2,3,4,5"
      Case 3
        fxMesesCierre = "6,7,8,9"
    End Select
 
 Case 6
    Select Case intNumeroCierre
      Case 1
        fxMesesCierre = "10,11,12,1,2,3"
      Case 2
        fxMesesCierre = "4,5,6,7,8,9"
    End Select

 Case 12
    fxMesesCierre = "10,11,12,1,2,3,4,5,6,7,8,9"
 Case Else
    fxMesesCierre = ""
End Select

End Function



'Private Sub ActualizaParametrosDespuesCierre(intPeriodoDesde As Integer, intPeriodoHasta As Integer, intTipoCierre As Integer, intNumeroCierre As Integer)
'Dim strSQL As String, vPeriodoDe As Integer, vPeriodoHasta As Integer
'
'vPeriodoDe = intPeriodoDesde
'vPeriodoHasta = intPeriodoHasta
'
'
'Select Case intTipoCierre
'  Case 1
'    If intNumeroCierre < 12 Then
'      intNumeroCierre = intNumeroCierre + 1
'    Else
'      vPeriodoDe = vPeriodoDe + 1
'      vPeriodoHasta = vPeriodoHasta + 1
'      intNumeroCierre = 1
'    End If
'
'  Case 2
'    If intNumeroCierre < 6 Then
'      intNumeroCierre = intNumeroCierre + 1
'    Else
'      vPeriodoDe = vPeriodoDe + 1
'      vPeriodoHasta = vPeriodoHasta + 1
'      intNumeroCierre = 1
'    End If
'
'  Case 3
'    If intNumeroCierre < 4 Then
'      intNumeroCierre = intNumeroCierre + 1
'    Else
'      vPeriodoDe = vPeriodoDe + 1
'      vPeriodoHasta = vPeriodoHasta + 1
'      intNumeroCierre = 1
'    End If
'
'  Case 4
'
'    If intNumeroCierre < 3 Then
'      intNumeroCierre = intNumeroCierre + 1
'    Else
'      vPeriodoDe = vPeriodoDe + 1
'      vPeriodoHasta = vPeriodoHasta + 1
'      intNumeroCierre = 1
'    End If
'
'  Case 6
'    If intNumeroCierre < 2 Then
'      intNumeroCierre = intNumeroCierre + 1
'    Else
'      vPeriodoDe = vPeriodoDe + 1
'      vPeriodoHasta = vPeriodoHasta + 1
'      intNumeroCierre = 1
'    End If
'
'  Case 12
'    If intNumeroCierre < 1 Then
'      intNumeroCierre = intNumeroCierre + 1
'    Else
'      vPeriodoDe = vPeriodoDe + 1
'      vPeriodoHasta = vPeriodoHasta + 1
'      intNumeroCierre = 1
'    End If
'
'End Select
'
'strSQL = "update excedentes_parametros set periodo_de = " & vPeriodoDe _
'       & ",periodo_hasta =" & vPeriodoHasta & ",cierre_pendiente =" & intNumeroCierre
'Call ConectionExecute(strSQL)
'
'End Sub



Private Sub sbExcedentes_Cierre_Aplica()
Dim strSQL As String, rs As New ADODB.Recordset
Dim intTipoCierre As Integer, intNumeroCierre As Integer
Dim iRespuesta As Integer, rs2 As New ADODB.Recordset


Dim pPeriodoId As Long

On Error GoTo vError


pPeriodoId = cboCRE_Periodo.ItemData(cboCRE_Periodo.ListIndex)

Me.MousePointer = vbHourglass


'Verifica el Cierre
strSQL = "select dbo.fxExc_Cierre_Valida(" & pPeriodoId & ") as 'Mensaje'"
Call OpenRecordSet(rs, strSQL)

If Len(rs!Mensaje) > 0 Then
    Me.MousePointer = vbDefault
    MsgBox rs!Mensaje, vbExclamation
    Exit Sub
End If 'Verifica

   
lblStatus.Caption = "Procesando Cierre de Excedentes [Espere]"

      
strSQL = "exec spExc_Cierre " & pPeriodoId & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

    txtCRECasos = rs!Casos
    txtCREExcedenteBruto = Format(rs!Excedente_Bruto, "Standard")
    txtCREExcedenteNeto = Format(rs!Excedente_Neto, "Standard")
    txtCRERenta = Format(rs!Renta, "Standard")
    txtCRECapitalizado = Format(rs!Capitalizacion, "Standard")

rs.Close

lblStatus.Caption = ""

'Refresca combos
Call TimerX_Timer

Me.MousePointer = vbDefault
MsgBox "Cierre de Excedentes se aplicó satisfactoriamente...", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Function fxValidaAplicacion(pPeriodoId As Long, pCorte As Date) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset


fxValidaAplicacion = True
 
strSQL = "select dbo.fxEXC_Periodo_Corte_Valida(" & pPeriodoId & ",'" & Format(pCorte, "YYYY-MM-DD hh:mm:ss") & "') as 'Resultado'"
Call OpenRecordSet(rs, strSQL)
If Len(rs!Resultado) > 0 Then
    fxValidaAplicacion = False
    MsgBox rs!Resultado, vbCritical
End If
rs.Close
 
End Function

Private Sub txtMonto_LostFocus()
On Error GoTo vError

txtMonto.Text = Format(CCur(txtMonto.Text), "Standard")

vError:
End Sub
