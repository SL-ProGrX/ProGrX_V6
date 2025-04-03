VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmAH_Excedentes_Pago 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Excedentes: Auxiliar de Pago"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   720
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   8055
      _Version        =   1572864
      _ExtentX        =   14208
      _ExtentY        =   10821
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
      ItemCount       =   2
      Item(0).Caption =   "Pagos"
      Item(0).ControlCount=   15
      Item(0).Control(0)=   "Label3(0)"
      Item(0).Control(1)=   "Label3(1)"
      Item(0).Control(2)=   "Label3(2)"
      Item(0).Control(3)=   "Label3(3)"
      Item(0).Control(4)=   "Label3(4)"
      Item(0).Control(5)=   "rbProceso(0)"
      Item(0).Control(6)=   "rbProceso(1)"
      Item(0).Control(7)=   "rbProceso(2)"
      Item(0).Control(8)=   "rbProceso(3)"
      Item(0).Control(9)=   "rbProceso(4)"
      Item(0).Control(10)=   "rbProceso(5)"
      Item(0).Control(11)=   "gbSep1(0)"
      Item(0).Control(12)=   "rbProceso(6)"
      Item(0).Control(13)=   "rbProceso(7)"
      Item(0).Control(14)=   "gbLote"
      Item(1).Caption =   "Reportes"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "gbSep1(1)"
      Item(1).Control(1)=   "gbSep1(2)"
      Begin XtremeSuiteControls.GroupBox gbLote 
         Height          =   1695
         Left            =   360
         TabIndex        =   31
         Top             =   3480
         Visible         =   0   'False
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12938
         _ExtentY        =   2990
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton btnTI_Procesa 
            Height          =   495
            Left            =   5040
            TabIndex        =   33
            Top             =   1080
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Procesar"
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
            Picture         =   "frmAH_Excedentes_Pago.frx":0000
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.PushButton btnTI_Cierra 
            Height          =   495
            Left            =   6360
            TabIndex        =   35
            Top             =   1080
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1508
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Cerrar"
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
            Picture         =   "frmAH_Excedentes_Pago.frx":0719
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.FlatEdit txtLote 
            Height          =   375
            Left            =   3240
            TabIndex        =   36
            Top             =   720
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1508
            _ExtentY        =   661
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
            Text            =   "5000"
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeShortcutBar.ShortcutCaption scPendientes 
            Height          =   375
            Left            =   2040
            TabIndex        =   37
            Top             =   0
            Width           =   5295
            _Version        =   1572864
            _ExtentX        =   9340
            _ExtentY        =   661
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
            Alignment       =   2
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   495
            Left            =   240
            TabIndex        =   34
            Top             =   600
            Width           =   2775
            _Version        =   1572864
            _ExtentX        =   4895
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Indique el Tamaño de Lote a Procesar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            WordWrap        =   -1  'True
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   375
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   4095
            _Version        =   1572864
            _ExtentX        =   7223
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Tamaño del Lote a Procesar"
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
      End
      Begin XtremeSuiteControls.GroupBox gbSep1 
         Height          =   735
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   5280
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12938
         _ExtentY        =   1296
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnAplicar 
            Height          =   495
            Left            =   5880
            TabIndex        =   16
            Top             =   240
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   873
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmAH_Excedentes_Pago.frx":0D57
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.Label lblStatus 
            Height          =   495
            Left            =   240
            TabIndex        =   26
            Top             =   240
            Width           =   5535
            _Version        =   1572864
            _ExtentX        =   9763
            _ExtentY        =   873
            _StockProps     =   79
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
      End
      Begin XtremeSuiteControls.RadioButton rbProceso 
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   9
         Top             =   960
         Width           =   4215
         _Version        =   1572864
         _ExtentX        =   7435
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Separar Ex-Empleados y Excedentes Cero"
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
      Begin XtremeSuiteControls.RadioButton rbProceso 
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   10
         Top             =   2400
         Width           =   4215
         _Version        =   1572864
         _ExtentX        =   7435
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Asignacion de Casos Especiales"
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
      Begin XtremeSuiteControls.RadioButton rbProceso 
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   11
         Top             =   3120
         Width           =   4215
         _Version        =   1572864
         _ExtentX        =   7435
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Acreditar Cuentas de Ahorros ASECCSS"
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
      Begin XtremeSuiteControls.RadioButton rbProceso 
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   12
         Top             =   3840
         Width           =   4215
         _Version        =   1572864
         _ExtentX        =   7435
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Enviar a Tesorería"
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
      Begin XtremeSuiteControls.RadioButton rbProceso 
         Height          =   255
         Index           =   4
         Left            =   960
         TabIndex        =   13
         Top             =   4920
         Width           =   4215
         _Version        =   1572864
         _ExtentX        =   7435
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Enviar a Fondos de ahorros [fx]"
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
      Begin XtremeSuiteControls.RadioButton rbProceso 
         Height          =   255
         Index           =   5
         Left            =   960
         TabIndex        =   14
         Top             =   4200
         Width           =   4215
         _Version        =   1572864
         _ExtentX        =   7435
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Reclasificar Salidas"
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
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.GroupBox gbSep1 
         Height          =   855
         Index           =   1
         Left            =   -69640
         TabIndex        =   17
         Top             =   4920
         Visible         =   0   'False
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12938
         _ExtentY        =   1508
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnInforme 
            Height          =   495
            Left            =   5880
            TabIndex        =   18
            Top             =   240
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Reporte"
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
            Picture         =   "frmAH_Excedentes_Pago.frx":147E
            ImageAlignment  =   4
         End
      End
      Begin XtremeSuiteControls.GroupBox gbSep1 
         Height          =   4215
         Index           =   2
         Left            =   -69760
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   7455
         _Version        =   1572864
         _ExtentX        =   13150
         _ExtentY        =   7435
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.RadioButton rbInforme 
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   20
            Top             =   720
            Width           =   4215
            _Version        =   1572864
            _ExtentX        =   7435
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Resumen de Salidas"
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
         Begin XtremeSuiteControls.RadioButton rbInforme 
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   21
            Top             =   1200
            Width           =   4215
            _Version        =   1572864
            _ExtentX        =   7435
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Detalle de Excedentes"
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
         Begin XtremeSuiteControls.RadioButton rbInforme 
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   22
            Top             =   1680
            Width           =   4215
            _Version        =   1572864
            _ExtentX        =   7435
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Dimex Inactivos"
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
         Begin XtremeSuiteControls.RadioButton rbInforme 
            Height          =   255
            Index           =   3
            Left            =   1200
            TabIndex        =   23
            Top             =   2160
            Width           =   4215
            _Version        =   1572864
            _ExtentX        =   7435
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Más de una Cuenta (Ahorros Interna)"
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
         Begin XtremeSuiteControls.RadioButton rbInforme 
            Height          =   255
            Index           =   4
            Left            =   1200
            TabIndex        =   24
            Top             =   2640
            Width           =   4215
            _Version        =   1572864
            _ExtentX        =   7435
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Detalle de Pago de Excedentes"
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
         Begin XtremeSuiteControls.RadioButton rbInforme 
            Height          =   255
            Index           =   5
            Left            =   1200
            TabIndex        =   29
            Top             =   3120
            Width           =   4215
            _Version        =   1572864
            _ExtentX        =   7435
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Boleta PosCierre Pago Excedentes"
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
         Begin XtremeSuiteControls.RadioButton rbInforme 
            Height          =   255
            Index           =   6
            Left            =   1200
            TabIndex        =   30
            Top             =   3600
            Width           =   4215
            _Version        =   1572864
            _ExtentX        =   7435
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Boleta Aportes Activos y Liquidados por Mes"
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
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   0
            Width           =   7335
            _Version        =   1572864
            _ExtentX        =   12938
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Informes de Pago de Excedentes"
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
      End
      Begin XtremeSuiteControls.RadioButton rbProceso 
         Height          =   255
         Index           =   6
         Left            =   960
         TabIndex        =   27
         Top             =   1320
         Width           =   6135
         _Version        =   1572864
         _ExtentX        =   10821
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Separar Casos con Cuentas Bancarias (Transferencias SINPE)"
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
      Begin XtremeSuiteControls.RadioButton rbProceso 
         Height          =   255
         Index           =   7
         Left            =   960
         TabIndex        =   28
         Top             =   1680
         Width           =   5175
         _Version        =   1572864
         _ExtentX        =   9128
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Separar Casos Sin Cuenta y Dimex Inactivo"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   8
         Top             =   4560
         Width           =   4935
         _Version        =   1572864
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Paso 5: Traslado de Fondos de Ahorros"
         ForeColor       =   16711680
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   3480
         Width           =   4935
         _Version        =   1572864
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Paso 4: Traslado a Bancos (Tesorería)"
         ForeColor       =   16711680
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   2760
         Width           =   4935
         _Version        =   1572864
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Paso 3: Transferencias Internas"
         ForeColor       =   16711680
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   4935
         _Version        =   1572864
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Paso 2: Casos Especiales"
         ForeColor       =   16711680
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   4935
         _Version        =   1572864
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Paso 1: Separar Casos"
         ForeColor       =   16711680
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
      End
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodo 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   3735
      _Version        =   1572864
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Periodo"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Auxiliar de Pagos de Excedentes"
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
      TabIndex        =   0
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
Attribute VB_Name = "frmAH_Excedentes_Pago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean


Private Sub sbPaso_1_Salidas(Optional Paso As Integer = 0)

On Error GoTo vError

Me.MousePointer = vbHourglass

'spExc_ASECCSS_SepararCasos(
'                                        @pPeriodoId     int,
'                                        @pEnviarSinpe   smallint,
'                                        @pPaso          smallint,
'                                        @pUsuario       varchar(35)
'                                        )

Select Case Paso
    Case 1
        lblStatus.Caption = "Separar Ex Empleados y Excedenes en Cero"
        DoEvents
        
        strSQL = "exec spExc_ASECCSS_SepararCasos " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
               & ", 0, 1, '" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)

    Case 2
        Me.MousePointer = vbDefault

        Dim respuesta As Integer
        
        respuesta = MsgBox("¿Deseas Enviar las salidas por Banco SINPE?", vbYesNo + vbQuestion, "Confirmación")
        
        Me.MousePointer = vbHourglass
        
        
        lblStatus.Caption = "Separar Casos con Cuentas Bancarias (Transferencias SINPE)"
        DoEvents
        
        If respuesta = vbYes Then
                strSQL = "exec spExc_ASECCSS_SepararCasos " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
                       & ", 1, 2, '" & glogon.Usuario & "'"
        Else
                strSQL = "exec spExc_ASECCSS_SepararCasos " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
                       & ", 0, 2, '" & glogon.Usuario & "'"
        End If
        
        Call ConectionExecute(strSQL)

    Case 3
        lblStatus.Caption = "Separar Casos Sin Cuenta y Dimex Inactivo"
        DoEvents
        
        strSQL = "exec spExc_ASECCSS_SepararCasos " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
               & ", 0, 3, '" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)

End Select

lblStatus.Caption = ""

Me.MousePointer = vbDefault

MsgBox "Separación de Casos Realizada Satisfactoriamente!", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbPaso_2_CasosEspeciales()

On Error GoTo vError

Me.MousePointer = vbHourglass


lblStatus.Caption = "Asignación de Casos Especiales"
DoEvents

strSQL = "exec spExc_ASECCSS_CasosEspeciales " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
       & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

lblStatus.Caption = ""

MsgBox "Asignación de Casos Especiales Realizada Satisfactoriamente!", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbPaso_3_Acreditar_Cuentas_Internas()

On Error GoTo vError

If Not IsNumeric(txtLote.Text) Then
    MsgBox "El Lote no es válido!", vbExclamation
    Exit Sub
End If

Me.MousePointer = vbHourglass

'    spEXC_ValidaCuentasTransfInterna
'    spEXC_EnviaFondoTransfInterna
'    spEXC_FNDDocumento

lblStatus.Caption = "Revisando Cuentas Internas"
DoEvents

strSQL = "exec spExc_ASECCSS_ValidaCuentasTransfInterna " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
       & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


lblStatus.Caption = "Acreditar Cuentas Internas"
DoEvents

'spExc_ASECCSS_EnviaFondoTransfInterna(
'                                                    @pPeriodoId As Numeric,
'                                                    @pOperadora As Numeric ,
'                                                    @pSalidaOrigen As Varchar(10),
'                                                    @pPlan As Varchar(10),
'                                                    @pUsuario  As Varchar(35),
'                                                    @pConcepto As Varchar(10) = 'FND001',
'                                                    @pNumDoc as varchar(30),
'                                                    @pTipoDoc as varchar(30),
'                                                    @pTop As INT = 5000
'                                                    )
Dim NumDoc As String, pCuenta As String

NumDoc = "Exc_[" & cboPeriodo.ItemData(cboPeriodo.ListIndex) & "]_TI"

strSQL = "SELECT  VALOR FROM  SIF_PARAMETROS WHERE COD_PARAMETRO = 'CCEX'"
Call OpenRecordSet(rs, strSQL)
    pCuenta = Trim(rs!Valor)
rs.Close

'Lote Inicial
strSQL = "exec spExc_ASECCSS_EnviaFondoTransfInterna " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
       & ", 1, 'TI', 'SINPE', '" & glogon.Usuario & "', 'FND001', '" & NumDoc & "', 'FND', " & txtLote.Text
Call OpenRecordSet(rs, strSQL)
Do While rs!Pendientes > 0

    scPendientes.Caption = "Pendientes de Procesar: " & Format(rs!Pendientes, "###,##0")
    
    MsgBox "Lote Procesado Satisfactoriamente! Continuar con el lote siguiente:", vbInformation
    
    strSQL = "exec spExc_ASECCSS_EnviaFondoTransfInterna " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
           & ", 1, 'TI', 'SINPE', '" & glogon.Usuario & "', 'FND001', '" & NumDoc & "', 'FND', " & txtLote.Text
    Call OpenRecordSet(rs, strSQL)

Loop


lblStatus.Caption = "Creando Comprobante..."
DoEvents

strSQL = "exec spExc_FNDDocumento 'FND', '" & NumDoc & "', 'FND001', '" & pCuenta & "', " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
       & ", '" & glogon.Usuario & "', 'TI'"
Call ConectionExecute(strSQL)


'spExc_FNDDocumento(@pTipoDoc As varchar(10),
'                                        @pNumDoc As varchar(30),
'                                        @pConcepto As varchar(10),
'                                        @pCuenta As varchar(60) ,
'                                        @pPeriodoId As Integer,
'                                        @pUsuario as Varchar(35),
'                                        @pSalida As varchar(50) = '')

Me.MousePointer = vbDefault

lblStatus.Caption = ""

MsgBox "Transferencia Interna Realizada Satisfactoriamente!", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbPaso_4_Envio_Tesoreria()

On Error GoTo vError

Me.MousePointer = vbHourglass


lblStatus.Caption = "Traslado de Excedentes a Bancos"
DoEvents

'spEXC_TrasladoExcedentesTesoreria](
'                                                      @pPeriodoId As Numeric,
'                                                      @pOficina As Varchar(10) ,
'                                                      @pUsuario  As Varchar(35)
'                                                     )

strSQL = "exec spEXC_TrasladoExcedentesTesoreria " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
       & ", 'AOC', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

lblStatus.Caption = ""

MsgBox "Traslado de Excedentes a Tesorería -> Realizada Satisfactoriamente!", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





Private Sub sbPaso_5_Envio_Fondos()

On Error GoTo vError

Me.MousePointer = vbHourglass


lblStatus.Caption = "Traslado de Excedentes a Fondos de Ahorro"
DoEvents

'spEXC_ASECCSS_TrasladoAFondosAhorros(
'                                                @pPeriodoId As Numeric,
'                                                @pOperadora As Numeric ,
'                                                @pUsuario  As Varchar(35),
'                                                @pConcepto As Varchar(10) = 'FND001'
'                                                )

strSQL = "exec spEXC_ASECCSS_TrasladoAFondosAhorros " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
       & ", 1, '" & glogon.Usuario & "', 'FND001'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

lblStatus.Caption = ""

MsgBox "Traslado de Fondos -> Realizada Satisfactoriamente!", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbPaso_6_Reclasificaciones()

On Error GoTo vError

Me.MousePointer = vbHourglass


lblStatus.Caption = "Procesando Reclasificaciones"
DoEvents

strSQL = "exec spEXC_CasosReclasificacionesDevoluciones " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
       & ", 'AOC', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

lblStatus.Caption = ""

MsgBox "Reclasificaciones Realizada Satisfactoriamente!", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub btnAplicar_Click()

Select Case True
    Case rbProceso(0) 'Separar Salidas
        Call sbPaso_1_Salidas(1)
    Case rbProceso(6) 'Separar Salidas
        Call sbPaso_1_Salidas(2)
    Case rbProceso(7) 'Separar Salidas
        Call sbPaso_1_Salidas(3)
        
        
    Case rbProceso(1) 'Asignacion de Casos Especiales
        Call sbPaso_2_CasosEspeciales

    Case rbProceso(2) 'Acreditar Cuentas Internas
        gbLote.Visible = True
        
    Case rbProceso(3) 'Enviar a Tesoreria
        Call sbPaso_4_Envio_Tesoreria

    
    Case rbProceso(4) 'Traslado a Fondos de Ahorros
        Call sbPaso_5_Envio_Fondos

    
    Case rbProceso(5) 'Reclasificar Salidas
        Call sbPaso_6_Reclasificaciones
End Select

End Sub

Private Sub btnInforme_Click()
Dim strSQL As String
Dim pCorte As Date, pCorteFiltro As String


On Error GoTo vError


Me.MousePointer = vbHourglass


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
    .Formulas(4) = "subtitulo='PERIODO: " & cboPeriodo.Text & "'"
    
    

    Select Case True
      Case rbInforme(0).Value 'Resumen de Salidas
        .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_Salidas_Resumen.rpt")
        .StoredProcParam(0) = cboPeriodo.ItemData(cboPeriodo.ListIndex)
        .StoredProcParam(1) = glogon.Usuario
        
      Case rbInforme(1).Value 'Detalle
        .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_Salidas_Pago_Detalle_Info.rpt")
        .StoredProcParam(0) = cboPeriodo.ItemData(cboPeriodo.ListIndex)
        .StoredProcParam(1) = glogon.Usuario
        
      Case rbInforme(2).Value 'Casos con Dimex inactivos
        .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_Dimex_Inactivo_Lista.rpt")
        .StoredProcParam(0) = cboPeriodo.ItemData(cboPeriodo.ListIndex)
        .StoredProcParam(1) = glogon.Usuario

      Case rbInforme(3).Value 'Personas con mas de una cuenta Sinpe
        .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_Personas_ConMasCtasSinpe.rpt")
        .StoredProcParam(0) = glogon.Usuario

      Case rbInforme(4).Value 'Detalle de Pago
        .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_Salidas_Pago_Detalle.rpt")
        .StoredProcParam(0) = cboPeriodo.ItemData(cboPeriodo.ListIndex)
        .StoredProcParam(1) = glogon.Usuario
        
      Case rbInforme(5).Value 'Boleta de PosCierre Pago de Excedentes
        .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_Salidas_Pago_Boleta.rpt")
        .StoredProcParam(0) = cboPeriodo.ItemData(cboPeriodo.ListIndex)
        .StoredProcParam(1) = glogon.Usuario
        
      Case rbInforme(6).Value 'Boleta de Aportes Activos y Liquidados por Mes
        .ReportFileName = SIFGlobal.fxPathReportes("Excedentes_AplicadoPeriodo_TOTAL_Resumen.rpt")
        .StoredProcParam(0) = cboPeriodo.ItemData(cboPeriodo.ListIndex)
        .StoredProcParam(1) = glogon.Usuario
        
        
    End Select

    .Action = 1
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnTI_Cierra_Click()
gbLote.Visible = False
End Sub

Private Sub btnTI_Procesa_Click()

Call sbPaso_3_Acreditar_Cuentas_Internas

gbLote.Visible = False

End Sub

Private Sub Form_Load()

vModulo = 2

glogon.Conection.CommandTimeout = 5200


Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

tcMain.Item(0).Selected = True

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)
glogon.Conection.CommandTimeout = 360

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Me.MousePointer = vbHourglass

vPaso = True


strSQL = "select IdX, ItmX from vExc_Periodos where estado in('C') order by Idx desc"
Call sbCbo_Llena_New(cboPeriodo, strSQL, False, True)

vPaso = False

Me.MousePointer = vbDefault

End Sub
