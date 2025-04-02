VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCC_ProcesoMensual 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Proceso: Deducciones de Planillas"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   9855
      _Version        =   1441793
      _ExtentX        =   17383
      _ExtentY        =   4048
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
   Begin XtremeSuiteControls.GroupBox fraAplicacion 
      Height          =   3255
      Left            =   7920
      TabIndex        =   15
      Top             =   1920
      Width           =   9495
      _Version        =   1441793
      _ExtentX        =   16743
      _ExtentY        =   5736
      _StockProps     =   79
      Caption         =   "Aplicación"
      ForeColor       =   8421504
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.GroupBox fraAhorros 
         Height          =   2295
         Left            =   0
         TabIndex        =   34
         Top             =   480
         Width           =   4575
         _Version        =   1441793
         _ExtentX        =   8064
         _ExtentY        =   4043
         _StockProps     =   79
         Caption         =   "Aplicación de Cuotas Obrero Patronal"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.RadioButton optAhorros 
            Height          =   252
            Index           =   0
            Left            =   480
            TabIndex        =   36
            Top             =   480
            Width           =   3012
            _Version        =   1441793
            _ExtentX        =   5313
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aplicación de Aportes"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optAhorros 
            Height          =   252
            Index           =   1
            Left            =   480
            TabIndex        =   37
            Top             =   960
            Width           =   3012
            _Version        =   1441793
            _ExtentX        =   5313
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Informe de Inconsistencias"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optAhorros 
            Height          =   252
            Index           =   2
            Left            =   480
            TabIndex        =   38
            Top             =   1440
            Width           =   3012
            _Version        =   1441793
            _ExtentX        =   5313
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Informe de Devoluciones"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optAhorros 
            Height          =   252
            Index           =   3
            Left            =   480
            TabIndex        =   43
            Top             =   1920
            Width           =   3012
            _Version        =   1441793
            _ExtentX        =   5313
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Informe de Traslado a Fondos"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin VB.Image imgRepAhFondos 
            Height          =   252
            Left            =   3600
            Picture         =   "frmCC_ProcesoMensual.frx":0000
            Stretch         =   -1  'True
            ToolTipText     =   "Reporte de las Devoluciones"
            Top             =   1920
            Width           =   252
         End
         Begin VB.Image imgRepAhDevolucion 
            Height          =   252
            Left            =   3600
            Picture         =   "frmCC_ProcesoMensual.frx":07AC
            Stretch         =   -1  'True
            ToolTipText     =   "Reporte de las Devoluciones"
            Top             =   1440
            Width           =   252
         End
         Begin VB.Image imgAplAhDevolucion 
            Height          =   228
            Left            =   4080
            Picture         =   "frmCC_ProcesoMensual.frx":0F58
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   228
         End
         Begin VB.Image imgRepAhIncon 
            Height          =   252
            Left            =   3600
            Picture         =   "frmCC_ProcesoMensual.frx":16BA
            Stretch         =   -1  'True
            ToolTipText     =   "Reporte de Inconsistencias"
            Top             =   960
            Width           =   252
         End
         Begin VB.Image imgAplAhIncon 
            Height          =   228
            Left            =   4080
            Picture         =   "frmCC_ProcesoMensual.frx":1E66
            Stretch         =   -1  'True
            Top             =   960
            Width           =   228
         End
         Begin VB.Image imgRepAhAplica 
            Height          =   252
            Left            =   3600
            Picture         =   "frmCC_ProcesoMensual.frx":25C8
            Stretch         =   -1  'True
            ToolTipText     =   "Reporte de la Aplicación de los Ahorros"
            Top             =   480
            Width           =   252
         End
         Begin VB.Image imgAplAhAplica 
            Height          =   228
            Left            =   4080
            Picture         =   "frmCC_ProcesoMensual.frx":2D74
            Stretch         =   -1  'True
            Top             =   480
            Width           =   228
         End
      End
      Begin XtremeSuiteControls.GroupBox fraCreditos 
         Height          =   2775
         Left            =   4680
         TabIndex        =   35
         Top             =   480
         Width           =   4815
         _Version        =   1441793
         _ExtentX        =   8488
         _ExtentY        =   4890
         _StockProps     =   79
         Caption         =   "Aplicación de Créditos y otros "
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.RadioButton optCreditos 
            Height          =   252
            Index           =   0
            Left            =   360
            TabIndex        =   39
            Top             =   480
            Width           =   3012
            _Version        =   1441793
            _ExtentX        =   5313
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aplicación de Abonos y Recaudo"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optCreditos 
            Height          =   252
            Index           =   1
            Left            =   360
            TabIndex        =   40
            Top             =   960
            Width           =   3012
            _Version        =   1441793
            _ExtentX        =   5313
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Informe de Inconsistencias"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optCreditos 
            Height          =   252
            Index           =   2
            Left            =   360
            TabIndex        =   41
            Top             =   1440
            Width           =   3012
            _Version        =   1441793
            _ExtentX        =   5313
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cálculo de Intereses Moratorios"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optCreditos 
            Height          =   252
            Index           =   3
            Left            =   360
            TabIndex        =   42
            Top             =   1920
            Width           =   3732
            _Version        =   1441793
            _ExtentX        =   6583
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "[Opcional] Recalculo de Saldo del Mes"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin VB.Image imgAplCrRecalculo 
            Height          =   228
            Left            =   4320
            Picture         =   "frmCC_ProcesoMensual.frx":34D6
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   228
         End
         Begin VB.Image imgRepCrIncon 
            Height          =   252
            Left            =   3840
            Picture         =   "frmCC_ProcesoMensual.frx":3C38
            Stretch         =   -1  'True
            ToolTipText     =   "Reporte de Inconsistencias"
            Top             =   960
            Width           =   252
         End
         Begin VB.Image imgAplCrIncon 
            Height          =   228
            Left            =   4320
            Picture         =   "frmCC_ProcesoMensual.frx":43E4
            Stretch         =   -1  'True
            Top             =   960
            Width           =   228
         End
         Begin VB.Image imgRepCrAplica 
            Height          =   252
            Left            =   3840
            Picture         =   "frmCC_ProcesoMensual.frx":4B46
            Stretch         =   -1  'True
            ToolTipText     =   "Reporte de la Aplicación de los Abonos"
            Top             =   480
            Width           =   252
         End
         Begin VB.Image imgAplCrAplica 
            Height          =   228
            Left            =   4320
            Picture         =   "frmCC_ProcesoMensual.frx":52F2
            Stretch         =   -1  'True
            Top             =   480
            Width           =   228
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   9855
         _Version        =   1441793
         _ExtentX        =   17383
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "Aplicación de la información recibida:"
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
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   1095
      Left            =   9240
      TabIndex        =   51
      Top             =   6600
      Visible         =   0   'False
      Width           =   5535
      _Version        =   524288
      _ExtentX        =   9763
      _ExtentY        =   1931
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      SpreadDesigner  =   "frmCC_ProcesoMensual.frx":5A54
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.GroupBox fraRecepcion 
      Height          =   2895
      Left            =   5520
      TabIndex        =   10
      Top             =   2040
      Width           =   9495
      _Version        =   1441793
      _ExtentX        =   16743
      _ExtentY        =   5101
      _StockProps     =   79
      Caption         =   "Recepción"
      ForeColor       =   8421504
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.RadioButton optGeneral 
         Height          =   252
         Index           =   2
         Left            =   600
         TabIndex        =   11
         Top             =   480
         Width           =   3012
         _Version        =   1441793
         _ExtentX        =   5313
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Carga Deducciones"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.RadioButton optGeneral 
         Height          =   252
         Index           =   3
         Left            =   600
         TabIndex        =   12
         Top             =   960
         Width           =   3012
         _Version        =   1441793
         _ExtentX        =   5313
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Detallar Aportes y Créditos"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Width           =   9855
         _Version        =   1441793
         _ExtentX        =   17383
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "Recepción y detalle de la información:"
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
      Begin VB.Image imgRepDesgloce 
         Height          =   252
         Left            =   4680
         Picture         =   "frmCC_ProcesoMensual.frx":5FDA
         Stretch         =   -1  'True
         ToolTipText     =   "Reporte del desgloce de la Carga"
         Top             =   960
         Width           =   252
      End
      Begin VB.Image imgRepCarga 
         Height          =   252
         Left            =   4680
         Picture         =   "frmCC_ProcesoMensual.frx":6786
         Stretch         =   -1  'True
         ToolTipText     =   "Reporte de la Carga de la Planilla"
         Top             =   480
         Width           =   252
      End
      Begin VB.Image imgAplCarga 
         Height          =   228
         Left            =   5160
         Picture         =   "frmCC_ProcesoMensual.frx":6F32
         Stretch         =   -1  'True
         Top             =   480
         Width           =   228
      End
      Begin VB.Image imgAplDesgloce 
         Height          =   228
         Left            =   5160
         Picture         =   "frmCC_ProcesoMensual.frx":7694
         Stretch         =   -1  'True
         Top             =   960
         Width           =   228
      End
   End
   Begin XtremeSuiteControls.GroupBox fraEnvio 
      Height          =   3615
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   9855
      _Version        =   1441793
      _ExtentX        =   17383
      _ExtentY        =   6376
      _StockProps     =   79
      Caption         =   "Envio"
      ForeColor       =   8421504
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.RadioButton optGeneral 
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   480
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5313
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cambio Fecha de Proceso"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.RadioButton optGeneral 
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   19
         Top             =   840
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5313
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Genera Deducciones"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.GroupBox fraGenera 
         Height          =   2175
         Left            =   360
         TabIndex        =   20
         Top             =   1440
         Visible         =   0   'False
         Width           =   7335
         _Version        =   1441793
         _ExtentX        =   12933
         _ExtentY        =   3831
         _StockProps     =   79
         Caption         =   "Genera Deducciones"
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkPlanillaTransito 
            Height          =   252
            Left            =   1320
            TabIndex        =   21
            Top             =   720
            Width           =   3852
            _Version        =   1441793
            _ExtentX        =   6794
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Tomar en Cuenta Planilla Pendiente"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtPlAnio 
            Height          =   330
            Left            =   1200
            TabIndex        =   22
            Top             =   360
            Width           =   732
            _Version        =   1441793
            _ExtentX        =   1291
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.FlatEdit txtPlMes 
            Height          =   330
            Left            =   1920
            TabIndex        =   23
            Top             =   360
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.PushButton btnGenera 
            Height          =   420
            Index           =   0
            Left            =   2160
            TabIndex        =   24
            Top             =   1680
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   741
            _StockProps     =   79
            Caption         =   "Aplicar"
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
            Appearance      =   16
            Picture         =   "frmCC_ProcesoMensual.frx":7DF6
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.PushButton btnGenera 
            Height          =   420
            Index           =   1
            Left            =   3600
            TabIndex        =   25
            Top             =   1680
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   741
            _StockProps     =   79
            Caption         =   "Cancelar"
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
            Appearance      =   16
            Picture         =   "frmCC_ProcesoMensual.frx":851D
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.CheckBox chkRedondeo 
            Height          =   252
            Left            =   1320
            TabIndex        =   26
            Top             =   960
            Width           =   3852
            _Version        =   1441793
            _ExtentX        =   6794
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Redondear a Un Decimal los Montos"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkCambioDeducciones 
            Height          =   252
            Left            =   1320
            TabIndex        =   27
            Top             =   1200
            Width           =   3852
            _Version        =   1441793
            _ExtentX        =   6794
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Considera cambios manuales anteriores"
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
            Appearance      =   16
            Value           =   1
         End
         Begin XtremeSuiteControls.ComboBox cboFrecuencia 
            Height          =   330
            Left            =   2280
            TabIndex        =   44
            Top             =   360
            Width           =   1812
            _Version        =   1441793
            _ExtentX        =   3201
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
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "[ (Año/Mes)  ej. 2021-11] + Frecuencia"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   1
            Left            =   4200
            TabIndex        =   29
            Top             =   360
            Width           =   3612
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Planilla"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   0
            Left            =   480
            TabIndex        =   28
            Top             =   360
            Width           =   732
         End
      End
      Begin XtremeSuiteControls.GroupBox fraFechaProceso 
         Height          =   1095
         Left            =   360
         TabIndex        =   30
         Top             =   1440
         Visible         =   0   'False
         Width           =   5415
         _Version        =   1441793
         _ExtentX        =   9546
         _ExtentY        =   1926
         _StockProps     =   79
         Caption         =   "Cambio de fecha de proceso [Corte]"
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtAno 
            Height          =   312
            Left            =   1200
            TabIndex        =   31
            Top             =   360
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
         Begin XtremeSuiteControls.ComboBox cboMes 
            Height          =   312
            Left            =   1920
            TabIndex        =   32
            Top             =   360
            Width           =   2292
            _Version        =   1441793
            _ExtentX        =   4048
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
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboFrecuenciaCambia 
            Height          =   312
            Left            =   1920
            TabIndex        =   45
            Top             =   720
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
         Begin VB.Image imgCambiaFecha 
            Height          =   480
            Left            =   4560
            Picture         =   "frmCC_ProcesoMensual.frx":8C33
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Periodo: "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   360
            TabIndex        =   33
            Top             =   360
            Width           =   852
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Width           =   9855
         _Version        =   1441793
         _ExtentX        =   17383
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "Envío de Información:"
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
      Begin VB.Image imgModificaCtas 
         Height          =   255
         Left            =   3480
         Picture         =   "frmCC_ProcesoMensual.frx":9401
         Stretch         =   -1  'True
         ToolTipText     =   "Revisar y Modificar Cuotas"
         Top             =   840
         Width           =   255
      End
      Begin VB.Image imgAplGenera 
         Height          =   225
         Left            =   4920
         Picture         =   "frmCC_ProcesoMensual.frx":9D56
         Stretch         =   -1  'True
         Top             =   840
         Width           =   225
      End
      Begin VB.Image imgRepGenera 
         Height          =   255
         Left            =   4440
         Picture         =   "frmCC_ProcesoMensual.frx":A4B8
         Stretch         =   -1  'True
         ToolTipText     =   "Reporte de la Generación"
         Top             =   840
         Width           =   255
      End
      Begin VB.Image imgAplFecha 
         Height          =   225
         Left            =   4920
         Picture         =   "frmCC_ProcesoMensual.frx":AC64
         Stretch         =   -1  'True
         Top             =   480
         Width           =   225
      End
      Begin VB.Image imgGeneraArchivo 
         Height          =   255
         Left            =   3960
         Picture         =   "frmCC_ProcesoMensual.frx":B3C6
         Stretch         =   -1  'True
         ToolTipText     =   "Genera Archivo"
         Top             =   840
         Width           =   255
      End
   End
   Begin XtremeSuiteControls.PushButton btnProceso 
      Height          =   420
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Envío"
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
      Picture         =   "frmCC_ProcesoMensual.frx":BAAD
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnProceso 
      Height          =   420
      Index           =   1
      Left            =   1680
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Recepción"
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
      Picture         =   "frmCC_ProcesoMensual.frx":C0CB
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnProceso 
      Height          =   420
      Index           =   2
      Left            =   3240
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Aplicación"
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
      Picture         =   "frmCC_ProcesoMensual.frx":C6E9
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnEjecucion 
      Height          =   420
      Index           =   0
      Left            =   7200
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Ejecutar"
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
      Picture         =   "frmCC_ProcesoMensual.frx":CD05
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnEjecucion 
      Height          =   420
      Index           =   1
      Left            =   8760
      TabIndex        =   8
      ToolTipText     =   "Bitácoras"
      Top             =   1320
      Width           =   615
      _Version        =   1441793
      _ExtentX        =   1080
      _ExtentY        =   741
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
      Appearance      =   16
      Picture         =   "frmCC_ProcesoMensual.frx":D5B7
   End
   Begin XtremeSuiteControls.PushButton btnEjecucion 
      Height          =   420
      Index           =   2
      Left            =   9360
      TabIndex        =   9
      ToolTipText     =   "Informes"
      Top             =   1320
      Width           =   615
      _Version        =   1441793
      _ExtentX        =   1080
      _ExtentY        =   741
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
      Appearance      =   16
      Picture         =   "frmCC_ProcesoMensual.frx":DE63
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1092
      Left            =   8520
      TabIndex        =   13
      Top             =   6120
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   1926
      _StockProps     =   79
      Caption         =   "Envio"
      ForeColor       =   8421504
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
      Appearance      =   16
      BorderStyle     =   1
   End
   Begin XtremeSuiteControls.GroupBox fra 
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   8280
      Width           =   9855
      _Version        =   1441793
      _ExtentX        =   17383
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "Estado:"
      ForeColor       =   8421504
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.ProgressBar prgProcesoMensual 
         Height          =   135
         Left            =   0
         TabIndex        =   46
         Top             =   600
         Width           =   9855
         _Version        =   1441793
         _ExtentX        =   17383
         _ExtentY        =   238
         _StockProps     =   93
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Height          =   252
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   9372
      End
   End
   Begin VB.TextBox txtInstitucion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   7200
      Visible         =   0   'False
      Width           =   9492
   End
   Begin XtremeSuiteControls.ComboBox cboAplicacion 
      Height          =   330
      Left            =   5160
      TabIndex        =   52
      Top             =   1320
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3201
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
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   5520
      Width           =   9855
      _Version        =   1441793
      _ExtentX        =   17383
      _ExtentY        =   450
      _StockProps     =   14
      Caption         =   "Bitácora del Proceso:"
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
   Begin VB.Label lblFechaProceso 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Proceso: "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   1884
      TabIndex        =   2
      Top             =   600
      Width           =   8052
   End
   Begin VB.Label lblInstitucion 
      BackStyle       =   0  'Transparent
      Caption         =   "Institucion"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1880
      TabIndex        =   1
      Top             =   120
      Width           =   7935
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmCC_ProcesoMensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mFechaSistema As String, mFrecuencia As String

Private Sub sbFechaProcesoSiguiente()
Dim Mes As Integer, pClean As Long, pQuincena As Currency

pClean = GLOBALES.glngFechaCR
pQuincena = GLOBALES.glngFechaCR - pClean

Mes = fxConvierteMES(cboMes.Text)

Select Case pQuincena
    Case 0   'M
        If Mes = 12 Then
         txtAno.Text = Val(txtAno) + 1
         Mes = 0
        End If
        Mes = Mes + 1
        
        cboFrecuenciaCambia.Text = "Mensual"
            
    Case 0.1 'Q1
        cboFrecuenciaCambia.Text = "2da Quincena"
    
    Case 0.2 'Q2
        If Mes = 12 Then
         txtAno.Text = Val(txtAno) + 1
         Mes = 0
        End If
        Mes = Mes + 1
        
        cboFrecuenciaCambia.Text = "1er Quincena"

End Select
 
'Selecciona el mes
cboMes.Text = fxConvierteMES(Mes)

End Sub



Private Sub sbAhAplicaAhorroRep(vFecha As Currency)
Dim strSQL As String, rs As New ADODB.Recordset
Dim dbPorcentaje As Double, dbPorcAhorro As Double

strSQL = "Select porc_aporte,porc_ahorro,frecuencia from instituciones where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
  dbPorcentaje = IIf(IsNull(rs!PORC_APORTE), 0, rs!PORC_APORTE) / 100
  dbPorcAhorro = IIf(IsNull(rs!porc_ahorro), 0, rs!porc_ahorro) / 100
rs.Close


With frmContenedor.Crt
 .Reset
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Reportes Módulo de Ahorros"
 
 .Connect = glogon.ConectRPT
 
 .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_PatAplicados.rpt")
 .Formulas(0) = "Fecha='" & fxFechaProcesoFormat(vFecha) & "'"
 .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "usuario='" & glogon.Usuario & "'"
 .Formulas(3) = "Porcentaje=" & dbPorcentaje
 .Formulas(4) = "PorcAhorro=" & dbPorcAhorro
 .Formulas(5) = "institucion = '" & GLOBALES.gNombreInstitucion & "'"
 .SelectionFormula = "{SOCIOSTEMP.EXISTE} = 'S' AND {SOCIOSTEMP.FECHAPROC} = " & vFecha _
                   & " AND {SOCIOSTEMP.COD_INSTITUCION} = " & GLOBALES.gInstitucion
 .PrintReport
End With

End Sub


Private Sub sbAhAplicaAhorro()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PARAMETRO:strfpro fecha de procesamiento del combo
'cargo a ahorro_detallado y consolidado los socios, montos
'cuando los socios del archivo existen en la tabla de socios.
'SE PARTE DEL SUPUESTO DE QUE EXISTE LA CONFIGURACION DE CUENTAS CONTABLES EN PAR_AFAH
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSQL As String, rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
Dim vFecha As Date, vFechaDias As Date, curDevoluciones As Currency, curDevAporte As Currency
Dim vTemp As String

vFecha = fxFechaServidor

On Error GoTo vError
prgProcesoMensual.Value = 1
Me.MousePointer = vbHourglass

lblStatus.Caption = "Aplicado Aportes (Espere...!)"
DoEvents

strSQL = "exec spPrmAporteAplica " & GLOBALES.glngFechaCR & "," & GLOBALES.gInstitucion & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


'Aplicar Devoluciones Al Fondo de Ahorros e Inversiones
'Solo Personas que se encuentran registradas
strSQL = "select * from instituciones where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs1, strSQL)

'Comprobante de Planilla de Fondos (Proceso + Institucion + Consecutivo)
vTemp = GLOBALES.glngFechaCR & "." & GLOBALES.gInstitucion & ".PAT.01"


If rs1!fnd_ap_Aplica = 1 Then
  strSQL = "select * from sociostemp where existe = 'D' and cod_institucion = " & GLOBALES.gInstitucion _
         & " and fechaproc = " & GLOBALES.glngFechaCR & " and (monto + aporte) > 0"
  Call OpenRecordSet(rs, strSQL)
  
  Do While Not rs.EOF
  
        'Inserta Ahorro Obrero
        strSQL = "exec spPrmDevFondos " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR & "," & rs1!fnd_ap_operadora _
               & ",'" & Trim(rs1!fnd_ap_plan) & "','" & Trim(rs!Cedula) & "'," & rs!Monto & ",'" & vTemp & "','" & Trim(rs1!cta_inconsistencia) _
               & "','A','" & Format(vFecha, "yyyy/mm/dd") & "'"
        Call ConectionExecute(strSQL)
  
        'Inserta Aporte Patronal
        strSQL = "exec spPrmDevFondos " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR & "," & rs1!fnd_ap_operadora _
               & ",'" & Trim(rs1!fnd_ap_planP) & "','" & Trim(rs!Cedula) & "'," & rs!Aporte & ",'" & vTemp & "','" & Trim(rs1!cta_inconsistencia) _
               & "','P','" & Format(vFecha, "yyyy/mm/dd") & "'"
        Call ConectionExecute(strSQL)
  
  
   
    rs.MoveNext
  Loop
  rs.Close
  
  'Borra Inconsistencias y Devoluciones
  strSQL = "delete sociosTemp where existe = 'D' and fechaProc = " & GLOBALES.glngFechaCR _
         & " and cod_institucion = " & GLOBALES.gInstitucion
  Call ConectionExecute(strSQL)
  
  
  'Asiento Masivo: Por Tipo de Documento
  Call sbFndAsiento(GLOBALES.glngFechaCR, rs1!fnd_ap_operadora, rs1!fnd_ap_plan, rs1!cta_inconsistencia, vTemp)


End If

rs1.Close


'Finaliza Proceso
Call Bitacora("Aplica", "PRM - Aplicación de Aportes Inst:" & GLOBALES.gInstitucion)
Call sbBitacoraPlanilla("05", GLOBALES.gInstitucion, GLOBALES.glngFechaCR, "R")

strSQL = "update instituciones set pr_apAplica = 1 where cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)

lblStatus.Caption = "Estatus..."
DoEvents

Call sbEstadoActualProceso

Me.MousePointer = vbDefault

'Reporte
Call sbAhAplicaAhorroRep(GLOBALES.glngFechaCR)

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub sbBitacoraConsulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

On Error GoTo vError

lsw.ListItems.Clear

strSQL = "select * from prm_bitacora where cod_institucion = " & GLOBALES.gInstitucion _
       & " and proceso = " & GLOBALES.glngFechaCR & " Order by id_seq"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Id_seq)
     itmX.SubItems(1) = IIf(rs!Gestion = "R", "Recepción", "Envio")
     itmX.SubItems(2) = fxPlanillaTipoTransac(rs!Transaccion)
     itmX.SubItems(3) = rs!Documento
     itmX.SubItems(4) = rs!Usuario
     itmX.SubItems(5) = rs!fecha


 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault


Exit Sub

vError:
 Me.MousePointer = vbDefault
 lsw.ListItems.Clear
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub


Private Sub sbReporte_CargadoNoLocalizado()

On Error GoTo vError

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "PROCESO MENSUAL - CARGADO DE INFORMACION"
     
    .Connect = glogon.ConectRPT
     
    .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fecha='" & Format(GLOBALES.glngFechaCR, "####-##") & "'"
    .Formulas(3) = "usuario='" & glogon.Usuario & "'"
    .Formulas(4) = "institucion='" & GLOBALES.gNombreInstitucion & "'"
    .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_Cargada_NoLocalizado.rpt")
    
    .SelectionFormula = "{vPrmCargadoPersonasNoEncontradas.FECHA_PROCESO} = " & GLOBALES.glngFechaCR _
              & " AND {vPrmCargadoPersonasNoEncontradas.COD_INSTITUCION} = " & GLOBALES.gInstitucion
    

    .PrintReport
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnEjecucion_Click(Index As Integer)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DESCRIPCION
'Funcion principal que controla los botones del toolbar y ademas llama a las funciones
'que ejecutan cada uno de los procesos
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSQL As String, rs As New ADODB.Recordset
 
 
prgProcesoMensual.Visible = True
lblStatus.Visible = True

  Select Case Index
    Case 0 'Ejecutar
       
       Select Case True
          Case fraEnvio.Visible  'Envío
              
              Select Case True
                 Case optGeneral(0).Value 'Cambia Fecha de Proceso
                      fraGenera.Visible = False
                      
                      fraFechaProceso.Visible = True
                      fraFechaProceso.Enabled = True
                      imgCambiaFecha.Enabled = True
                      
                      Call sbFechaProcesoSiguiente
                 
                 Case optGeneral(1).Value 'Genera Deducciones
                    If fxValidaPaso("01") Then
                       fraFechaProceso.Visible = False
                       fraGenera.Visible = True
                       fraGenera.Top = fraFechaProceso.Top
                       txtPlAnio = txtAno
                       txtPlMes = Format(fxConvierteMES(cboMes.Text), "00")
                    End If
               End Select
                
           Case fraRecepcion.Visible 'Recepcion
               
               Select Case True
                        Case optGeneral(2).Value 'Carga Deducciones
                           If fxValidaPaso("03") Then
                               Call sbProcesosAdd("03", "PRE")
                               
                               strSQL = "select planilla from instituciones" _
                                    & " where cod_institucion = " & GLOBALES.gInstitucion
                               Call OpenRecordSet(rs, strSQL)
                               Select Case Trim(rs!planilla)
                                 Case "00", "03"
                                    Call sbCargaDeduc_Excel
                                 Case "02" 'Integra Nuevo Csv
                                    Call sbCargaDeduc_Csv_Integra
                                 Case "30"
                                    Call sbCargaDeduc_ExcelNew
                                 Case "28" 'Tek Experts
                                    Call sbCargaDeduc_Excel_Tek_Experts
                                 Case "32" 'DxC Costa Rica
                                    Call sbCargaDeduc_Excel_DxC_Costa_Rica
                                 Case "33" 'DxC Centroamerica
                                    Call sbCargaDeduc_Excel_DxC_CentroAmerica
                                 Case Else
                                    Call sbCargaDeducciones
                               End Select
                               rs.Close
                               
                               Call sbProcesosAdd("03", "POS")
                               
                               'Revisa si existen casos sin Localizar
                               strSQL = "exec spPrmCargadoPersonasNoEncontradas " & GLOBALES.gInstitucion _
                                      & "," & GLOBALES.glngFechaCR
                               Call OpenRecordSet(rs, strSQL)
                               If rs!Existen > 1 Then
                                    MsgBox "Existen casos no registrados en la base de datos, verifiquelos!", vbExclamation
                                    Call sbReporte_CargadoNoLocalizado
                               End If
                               rs.Close
                               
                           End If
                           
                        Case optGeneral(3).Value 'Desglose de Creditos y Ahorros
                           If fxValidaPaso("04") Then
                              Call sbProcesosAdd("04", "PRE")
                              Call sbDesglocePlanilla
                              Call sbProcesosAdd("04", "POS")
                           End If
                    
              End Select
              
          Case fraAplicacion.Visible
              Select Case True
                 Case optAhorros(0).Value And fraAhorros.Enabled 'Aplica Ahorros y Genera Asiento
                    If fxValidaPaso("05") Then
                       Call sbProcesosAdd("05", "PRE")
                       Call sbAhAplicaAhorro
                       Call sbProcesosAdd("05", "POS")
                       
                       MsgBox "La Informacion se aplicó Satisfactoriamente", vbInformation

                    End If
                    
                 Case optAhorros(1).Value And fraAhorros.Enabled 'Reporte de Inconsistencias
                    If fxValidaPaso("06") Then
                       Call sbProcesosAdd("06", "PRE")
                       Call sbAhInconsistencias(GLOBALES.glngFechaCR)
                       glogon.Conection.Execute "update instituciones set pr_apInco = 1 where cod_institucion = " & GLOBALES.gInstitucion
                      
                       Call Bitacora("Aplica", "PRM-AHORRO Reporte Inconsistencias Inst:" & GLOBALES.gInstitucion)
                       Call sbBitacoraPlanilla("06", GLOBALES.gInstitucion, GLOBALES.glngFechaCR, "R")
                       Call sbEstadoActualProceso
                       Call sbProcesosAdd("06", "POS")
                    End If
                    
                 Case optAhorros(2).Value And fraAhorros.Enabled 'Reporte de Devoluciones
                     If fxValidaPaso("07") Then
                       Call sbProcesosAdd("07", "PRE")
                       Call sbAhDevoluciones(GLOBALES.glngFechaCR)
                       Call Bitacora("Aplica", "PRM-AHORRO Reporte Devoluciones Inst:" & GLOBALES.gInstitucion)
                       glogon.Conection.Execute "update instituciones set PR_APDev = 1 where cod_institucion = " & GLOBALES.gInstitucion
                      
                       Call sbBitacoraPlanilla("07", GLOBALES.gInstitucion, GLOBALES.glngFechaCR, "R")
                       Call sbEstadoActualProceso
                       Call sbProcesosAdd("07", "POS")
                    
                    End If
                 Case optAhorros(3).Value And fraAhorros.Enabled 'Reporte de Traslados al Fondo
                      Call sbRepFondo(GLOBALES.glngFechaCR)
                    
                 'Creditos (Aplicación)
                 Case optCreditos(0).Value 'Aplica Abonos y Genera Asiento
                   If fxValidaPaso("08") Then
                      Call sbProcesosAdd("08", "PRE")
                      Call sbCrAplicaAbonos
                      Call sbProcesosAdd("08", "POS")
                      
                      MsgBox "Información Aplicada ...", vbInformation
                   End If
                   
                 Case optCreditos(1).Value 'Reporte de Inconsistencias
                   If fxValidaPaso("09") Then
                      Call sbProcesosAdd("09", "PRE")
                   
                      Call Bitacora("Aplica", "PRM-CREDITO Reporte Inconsistencias Inst:" & GLOBALES.gInstitucion)
                      glogon.Conection.Execute "update instituciones set pr_crInco = 1 where cod_institucion = " & GLOBALES.gInstitucion
                     
                      Call sbCrReporteInconsistencias(GLOBALES.glngFechaCR)
                     
                      Call sbBitacoraPlanilla("09", GLOBALES.gInstitucion, GLOBALES.glngFechaCR, "R")
                      Call sbEstadoActualProceso
                      Call sbProcesosAdd("09", "POS")
                   End If
                   
                 Case optCreditos(2).Value 'Recalculo de Intereses Moratorios
                   If fxValidaPaso("10") Then
                      Call sbProcesosAdd("10", "PRE")
                      Call Bitacora("Aplica", "PRM-CREDITO Recalcula Mora Inst:" & GLOBALES.gInstitucion)
                      glogon.Conection.Execute "update instituciones set pr_crMora = 1 where cod_institucion = " & GLOBALES.gInstitucion
                      
                      Call sbBitacoraPlanilla("10", GLOBALES.gInstitucion, GLOBALES.glngFechaCR, "R")
                      Call sbCrRecalculaCuotaEnMora
                   
                      Call sbProcesosAdd("10", "POS")
                   
                      MsgBox "Actualización de Intereses Moratorios Realizado ...", vbInformation
                   End If
                 
                 Case optCreditos(3).Value 'Opcional - Recalculo del Saldo del Mes
                    Me.MousePointer = vbHourglass
                      Call sbProcesosAdd("11", "PRE")
                      Call sbCrCalculaSaldoMes(False, GLOBALES.glngFechaCR)
                      Call sbBitacoraPlanilla("11", GLOBALES.gInstitucion, GLOBALES.glngFechaCR, "R")
                      Call sbProcesosAdd("11", "POS")
                    Me.MousePointer = vbDefault
              
              End Select
          
        End Select
        
     Case 1 'Bitacora
            Call sbFormsCall("frmCC_PlanillaBitacora", , , , False, Me)
     Case 2 'Reportes"
            Call sbFormsCall("frmCC_PlanillaReportes", , , , False, Me)
  End Select


 prgProcesoMensual.Visible = False
 lblStatus.Visible = False
End Sub

Private Sub btnGenera_Click(Index As Integer)
Dim FechaProceso As Currency

Select Case Index
  Case 0 'aplicar
      'Los procesos complementarios van a lo interno del procedimiento de generación
      
      FechaProceso = CCur(txtPlAnio & Format(txtPlMes, "00") & "." & cboFrecuencia.ItemData(cboFrecuencia.ListIndex))
      
      Call sbGeneraDeducciones(FechaProceso)
  Case 1 'Cancelar
      fraGenera.Visible = False
End Select

'En cualquier Opcion siempre cierra el marco
fraGenera.Visible = False

End Sub

Private Sub btnProceso_Click(Index As Integer)
Dim i As Integer

For i = 0 To 2
    btnProceso.Item(Index).GreyDisabledPicture = False
Next i
btnProceso.Item(Index).GreyDisabledPicture = True


fraEnvio.Visible = False
fraRecepcion.Visible = False
fraAplicacion.Visible = False

fraEnvio.Left = 120
fraRecepcion.Left = 120
fraAplicacion.Left = 120

fraEnvio.Width = 9855
fraRecepcion.Width = 9855
fraAplicacion.Width = 9855


fraEnvio.Top = 1800
fraRecepcion.Top = 1800
fraAplicacion.Top = 1800

Select Case Index
  Case 0 'Envio"
        fraEnvio.Visible = True
  
  Case 1 'Recepcion"
        fraRecepcion.Visible = True
  
  Case 2 'Aplicacion"
        fraAplicacion.Visible = True
End Select

End Sub


Private Sub cboAplicacion_Click()

Dim pProcesoClean As Long

If cboAplicacion.ListCount = 0 Then Exit Sub
If Not cboAplicacion.Visible Then Exit Sub


pProcesoClean = GLOBALES.glngFechaCR

GLOBALES.glngFechaCR = pProcesoClean & "." & cboAplicacion.ItemData(cboAplicacion.ListIndex)


Select Case (GLOBALES.glngFechaCR - pProcesoClean)
    Case 0.1
        lblFechaProceso.Caption = Format(pProcesoClean, "###-##") & "_Q1"
    Case 0.2
        lblFechaProceso.Caption = Format(pProcesoClean, "###-##") & "_Q2"
    Case 0
        lblFechaProceso.Caption = Format(pProcesoClean, "###-##")
End Select

cboFrecuencia.Text = cboAplicacion.Text
Call sbBitacoraConsulta


End Sub


Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 3

mFechaSistema = fxFechaServidor

glogon.Conection.CommandTimeout = 360

With lsw.ColumnHeaders
   .Clear
   .Add , , "Id", 500, vbCenter
   .Add , , "Gestión", 1010
   .Add , , "Transacción", 3000
   .Add , , "Documento", 1740
   .Add , , "Usuario", 1200
   .Add , , "Fecha", 2240
End With


With cboMes
   .Clear
   .AddItem "Enero"
   .AddItem "Febrero"
   .AddItem "Marzo"
   .AddItem "Abril"
   .AddItem "Mayo"
   .AddItem "Junio"
   .AddItem "Julio"
   .AddItem "Agosto"
   .AddItem "Setiembre"
   .AddItem "Octubre"
   .AddItem "Noviembre"
   .AddItem "Diciembre"
End With

fraEnvio.Visible = True

txtInstitucion = GLOBALES.gNombreInstitucion
txtInstitucion.Tag = GLOBALES.gInstitucion

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture
'-------------------------------------------------------------------------------------
cboAplicacion.Clear

cboAplicacion.AddItem "Mensual"
cboAplicacion.ItemData(cboAplicacion.ListCount - 1) = "0"
cboAplicacion.Text = "Mensual"


strSQL = "Select Portal_Id from sif_Empresa"
Call OpenRecordSet(rs, strSQL)

'Temporal ASOSEJUD
If rs!Portal_Id = 53 Or rs!Portal_Id = 0 Then
    cboAplicacion.AddItem "1er Quincena"
    cboAplicacion.ItemData(cboAplicacion.ListCount - 1) = "1"
    cboAplicacion.AddItem "2da Quincena"
    cboAplicacion.ItemData(cboAplicacion.ListCount - 1) = "2"
Else
    cboAplicacion.Visible = False
End If

rs.Close
'-------------------------------------------------------------------------------------


Call Formularios(Me)

Call sbEstadoActualProceso

'Carga variables que identifican si se procesan Creditos y Aportes
strSQL = "select codigo_aportes,codigo_creditos_env from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
  fraAhorros.Enabled = IIf((Trim(UCase(rs!codigo_aportes)) = "NO"), False, True)
'    ssTab.TabEnabled(1) = IIf((Trim(UCase(rs!Codigo_Aportes_Env)) = "NO"), False, True)
'    ssTab.TabEnabled(2) = IIf((Trim(UCase(rs!codigo_creditos_env)) = "NO"), False, True)
rs.Close


End Sub


Private Function fxDescomponeCadena(iParada As Integer, vCadena As String) As String
Dim x As Integer, x2 As Integer, vResultado As String

x = 0
x2 = 1

For x = 1 To iParada
  vResultado = ""
  Do While Mid(vCadena, x2, 1) <> vbTab
     vResultado = vResultado & Mid(vCadena, x2, 1)
     x2 = x2 + 1
  Loop
  x2 = x2 + 1
Next x

fxDescomponeCadena = vResultado

End Function

Private Function fxDescomponeCsv(iParada As Integer, vCadena As String) As String
Dim x As Integer, x2 As Integer, vResultado As String

x = 0
x2 = 1

For x = 1 To iParada
  vResultado = ""
  Do While Mid(vCadena, x2, 1) <> ";"
     vResultado = vResultado & Mid(vCadena, x2, 1)
     x2 = x2 + 1
  Loop
  x2 = x2 + 1
Next x

fxDescomponeCsv = vResultado

End Function


Private Sub sbCargaDeducciones()
Dim fn, strCadena As String, lng As Long, iPago As Integer
Dim strMonto As String, strSQL As String, rs As New ADODB.Recordset
Dim vPlanilla As String, vTipoAporte As String, vTipoCredito As String
Dim vCedula As String, vTemp  As String, i As Integer
Dim vCodigoDeduccion As String, vCodigoTipo As String

Dim pCadenaExec As String

On Error GoTo vError

fn = FreeFile
pCadenaExec = ""


strSQL = "select planilla,codigo_aportes,codigo_creditos from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
  vPlanilla = Trim(rs!planilla)
  vTipoAporte = Trim(rs!codigo_aportes & "")
  vTipoCredito = Trim(rs!codigo_creditos & "")
rs.Close

With frmContenedor.CD
 .DialogTitle = "Localice archivo con las deducciones de Planilla..."
 .Filter = "*.*"
 .InitDir = "C:\"
 .ShowOpen
End With

If frmContenedor.CD.FileName = "" Then
 MsgBox "Seleccione el Archivo de Deducciones del Proceso " & Format(GLOBALES.glngFechaCR, "####-##"), vbInformation
 Exit Sub
End If


MsgBox "Se procederá a cargar los registros del archivo :" & frmContenedor.CD.FileName, vbInformation

'CAPTURA EL NUMERO DE PAGO DE LA PLANILLA
strSQL = ""
Do While Not IsNumeric(strSQL)
  strSQL = InputBox("Digite el Numero de Planilla del Mes: ", "# Pago en el Mes...", 1)
Loop
iPago = strSQL


Me.MousePointer = vbHourglass

'Limpiando Informacion Anterior
lblStatus = "Borrando Información Anterior ..."
DoEvents

strSQL = "delete prm_cargado where fecha_proceso = " & GLOBALES.glngFechaCR _
       & " and pago = " & iPago & " and cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)


prgProcesoMensual.Value = 1
prgProcesoMensual.Max = 2

lblStatus = "Cargando ..."
DoEvents

Open frmContenedor.CD.FileName For Input As #fn   ' Lee el archivo.
 Do While Not EOF(fn)
   Input #fn, strCadena
   prgProcesoMensual.Max = prgProcesoMensual.Max + 1
 Loop
Close #fn

prgProcesoMensual.Min = 1
DoEvents


'cboX.AddItem "00 - Microsoft Excel"
'cboX.AddItem "01 - [CCSS] Caja Costarricense Seguro Social"
'cboX.AddItem "02 - [INTEGRA] Mecanizada Tesoreria Nacional"
'cboX.AddItem "03 - [ASECCSS] Asociacion Solidarista Emp CCSS"
'cboX.AddItem "04 - [ICE](ACOTEL)Instituto Costarricense Electricidad"
'cboX.AddItem "05 - [COPECAJA] CoopeCaja RL"
'cboX.AddItem "06 - [ICE] Oficinas Centrales"
'cboX.AddItem "07 - [ICE] Proyectos"
'cboX.AddItem "08 - [AYA] Acueductos y Alcantarillados"
'cboX.AddItem "09 - [SPA] Mecanizada Tesoreria Nacional"
'
'cboX.AddItem "10 - [SIF] Sistema SIF [F01.Indefinidos]"
'cboX.AddItem "11 - [SIF] Sistema SIF [F02.Plazo definido]"
'
'cboX.AddItem "12 - [IMAS]Institucto Mixto de Ayuda de Social"
'cboX.AddItem "13 - [INA] Instituto Nacional de Apendizaje"
'cboX.AddItem "14 - [MSJ] Municipalidad de San José"
'cboX.AddItem "15 - [ PJ] Poder Judicial"
'cboX.AddItem "16 - [StarH] PriceWaterHouseCoopers"


Open frmContenedor.CD.FileName For Input As #fn   'Lee el Archivo y lo compara
Do While Not EOF(fn)
   Input #fn, strCadena
   
   Select Case vPlanilla
     Case "01" 'Planilla de la Caja Costarricense del Seguro Social
        'Preguntar si es credito o si es aportes
          strMonto = Format(Mid(strCadena, 28, 13), "###########")
          strMonto = LTrim(RTrim(strMonto))
          If Len(strMonto) > 2 Then
           strMonto = Mid(strMonto, 1, Len(strMonto) - 2) & "." & Mid(strMonto, Len(strMonto) - 1, Len(strMonto))
          Else
           strMonto = "0" & "." & strMonto
          End If
             
             strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion) values(" _
                    & GLOBALES.gInstitucion & "," & iPago & "," & GLOBALES.glngFechaCR
             Select Case Mid(strCadena, 12, 5)
                Case vTipoAporte
                    vCodigoTipo = "1"
                Case vTipoCredito
                    vCodigoTipo = "3"
                Case Else
                    vCodigoTipo = "3"
             End Select
             strSQL = strSQL & "," & vCodigoTipo & ",'" & Trim(Format(Mid(strCadena, 1, 11), "###########")) & "',"
             
             strSQL = strSQL & strMonto & ",'" & Mid(strCadena, 12, 5) & "')"
             
             If Len(strCadena) > 100 Then
'               Call ConectionExecute(strSQL)
                pCadenaExec = pCadenaExec & Space(10) & strSQL
             End If
     
     Case "02", "15", "18" 'INTEGRA: Tesorería Nacional , Poder Judicial, CONAVI
        'Nuevo Carga Formato de MECANIZADA
        
        
        strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion) values(" _
               & GLOBALES.gInstitucion & "," & iPago & "," & GLOBALES.glngFechaCR
        
        If UCase(Right(frmContenedor.CD.FileName, 3)) = "CSV" Then
                vCodigoDeduccion = fxDescomponeCsv(3, strCadena)
                Select Case vCodigoDeduccion
                        Case vTipoAporte
                            vCodigoTipo = "1"
                        Case vTipoCredito
                            vCodigoTipo = "3"
                        Case Else
                            vCodigoTipo = "3"
                End Select
                strSQL = strSQL & "," & vCodigoTipo & ",'" & Val(fxDescomponeCsv(1, strCadena)) _
                      & "'," & fxDescomponeCsv(4, strCadena) & ",'" & vCodigoDeduccion & "')"
        
        Else
                vCodigoDeduccion = fxDescomponeCadena(2, strCadena)
                Select Case vCodigoDeduccion
                        Case vTipoAporte
                            vCodigoTipo = "1"
                        Case vTipoCredito
                            vCodigoTipo = "3"
                        Case Else
                            vCodigoTipo = "3"
                End Select
        
        
                strSQL = strSQL & "," & vCodigoTipo & ",'" & Val(fxDescomponeCadena(1, strCadena)) _
                      & "'," & fxDescomponeCadena(3, strCadena) & ",'" & vCodigoDeduccion & "')"
        
        End If
        If vCodigoDeduccion <> "codigodeduccion" Then
            pCadenaExec = pCadenaExec & Space(10) & strSQL
        End If

      Case "03" 'ASECCSS (Para ASE-ASECCSS)
         
      Case "04" ' - [ICE](ACOTEL)Instituto Costarricense Electricidad"
      Case "05" ' - [COPECAJA] CoopeCaja RL"
      Case "06" ' - [ICE] Oficinas Centrales"
          vCedula = Trim(Mid(strCadena, 8, 9))
          strMonto = Mid(strCadena, 47, 10)
          
          'Elimina Caracter en Blanco
          vTemp = strMonto
'          For i = 1 To Len(strMonto)
'            If Mid(strMonto, i, 1) <> " " Then
'               vTemp = vTemp & Mid(strMonto, i, 1)
'            End If
'          Next i
          
          'Divide base 100, para indicar decimales
          If IsNumeric(vTemp) Then
              strMonto = CCur(vTemp) / 100
          End If
             
          strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto) values(" _
                 & GLOBALES.gInstitucion & "," & iPago & "," & GLOBALES.glngFechaCR & ",3,'" _
                 & vCedula & "'," & strMonto & ")"
          
          'Valida Linea
           If IsNumeric(vCedula) And IsNumeric(strMonto) Then
               'Call ConectionExecute(strSQL)
               pCadenaExec = pCadenaExec & Space(10) & strSQL
           End If
      
      
      Case "07" ' - [ICE] Proyectos"
          
          vCedula = Trim(Mid(strCadena, 37, 9))
          strMonto = Mid(strCadena, 82, 10)
          
          If strMonto = "" Then
            vCedula = Trim(Mid(strCadena, 1, 9))
            strMonto = Mid(strCadena, 47, 10)
          End If
          
          'Elimina Caracter en Blanco
          vTemp = ""
          For i = 1 To Len(strMonto)
            If Mid(strMonto, i, 1) <> " " Then
               vTemp = vTemp & Mid(strMonto, i, 1)
            End If
          Next i
          strMonto = vTemp
             
             
          strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto) values(" _
                 & GLOBALES.gInstitucion & "," & iPago & "," & GLOBALES.glngFechaCR & ",3,'" _
                 & vCedula & "'," & strMonto & ")"
          
          'Valida Linea
           If IsNumeric(vCedula) And IsNumeric(strMonto) Then
'               Call ConectionExecute(strSQL)
                pCadenaExec = pCadenaExec & Space(10) & strSQL
           End If
        
      
      Case "08" ' - [AYA] Acueductos y Alcantarillados"
        
        
      Case "09" ' - [SPA] Mecanizada Tesoreria Nacional"
        If Val(Mid(strCadena, 87, 1)) = 1 Then
             strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto) values(" _
                    & GLOBALES.gInstitucion & "," & iPago & "," & GLOBALES.glngFechaCR
             Select Case Mid(strCadena, 81, 6)
                Case vTipoAporte  'Aporte Obrero
                   strSQL = strSQL & ",1,'" & Val(Mid(strCadena, 1, 10)) & "'," & (CCur(Mid(strCadena, 50, 8)) / 100) & ")"
                   'Call ConectionExecute(strSQL)
                    pCadenaExec = pCadenaExec & Space(10) & strSQL
                Case vTipoCredito, "553551", "533195" 'Abonos a Creditos y Retenciones
                   strSQL = strSQL & ",3,'" & Val(Mid(strCadena, 1, 10)) & "'," & (CCur(Mid(strCadena, 50, 8)) / 100) & ")"
                   'Call ConectionExecute(strSQL)
                   pCadenaExec = pCadenaExec & Space(10) & strSQL
             End Select
        End If


      Case "10", "11" '- [SIF] Formato del Sistema SIF
           If Len(strCadena) >= 77 Then
             strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto) values(" _
                    & GLOBALES.gInstitucion & "," & iPago & "," & GLOBALES.glngFechaCR & ",3,'" _
                    & Mid(strCadena, 1, 15) & "'," & CCur(Mid(strCadena, 67, 12)) & ")"
'             Call ConectionExecute(strSQL)
             pCadenaExec = pCadenaExec & Space(10) & strSQL
           End If
           
      Case "14" 'MSJ
           If Len(strCadena) >= 48 Then
             strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto) values(" _
                    & GLOBALES.gInstitucion & "," & iPago & "," & GLOBALES.glngFechaCR & ",3,'" _
                    & Mid(strCadena, 1, 9) & "'," & (CCur(Mid(strCadena, 40, 10)) / 100) & ")"
'             Call ConectionExecute(strSQL)
             pCadenaExec = pCadenaExec & Space(10) & strSQL
           End If
        
        
     Case "19" 'CGR Contraloría de la Republica
       
       'Si se pudo realizar la deducción
        If Right(strCadena, 1) = "1" And Len(strCadena) >= 66 Then
        
            vCodigoDeduccion = Mid(strCadena, 2, 6)
            
            strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion) values(" _
                   & GLOBALES.gInstitucion & "," & iPago & "," & GLOBALES.glngFechaCR
            
            Select Case vCodigoDeduccion
               Case vTipoAporte
                    vCodigoTipo = "1"
               Case vTipoCredito
                    vCodigoTipo = "3"
               Case Else
                    vCodigoTipo = "3"
            End Select
       
            strSQL = strSQL & "," & vCodigoTipo & ",'" & Val(Mid(strCadena, 38, 10)) & "'," & (CCur(Mid(strCadena, 57, 9)) / 100) & ",'" & vCodigoDeduccion & "')"
            pCadenaExec = pCadenaExec & Space(10) & strSQL
       
       End If
        
        
     Case "35" 'ProGrX: RRGHH
       
'            pIdentificacion = Trim(Mid(strCadena, 1, 20))
'            pCodigo = Trim(Mid(strCadena, 21, 10))
'            pMonto = CCur(Mid(strCadena, 33, 10)) / 100

       'Si se pudo realizar la deducción
        If Len(strCadena) >= 50 Then
        
            vCodigoDeduccion = Trim(Mid(strCadena, 21, 10))
            
            strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion) values(" _
                   & GLOBALES.gInstitucion & "," & iPago & "," & GLOBALES.glngFechaCR
            
            Select Case vCodigoDeduccion
               Case vTipoAporte
                    vCodigoTipo = "1"
               Case vTipoCredito
                    vCodigoTipo = "3"
               Case Else
                    vCodigoTipo = "3"
            End Select
       
            strSQL = strSQL & "," & vCodigoTipo & ",'" & Trim(Mid(strCadena, 1, 20)) & "'," & (CCur(Mid(strCadena, 34, 10)) / 100) & ",'" & vCodigoDeduccion & "')"
            pCadenaExec = pCadenaExec & Space(10) & strSQL
       
       End If
        
   End Select
   
   If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
'   lblStatus.Caption = "Cargando..Registro # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
'   doEvents
   
   If Len(pCadenaExec) > 25000 Then
        lblStatus.Caption = "Subiendo Registros (Espere!)  (" & prgProcesoMensual.Value & " / " & prgProcesoMensual.Max & ")     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
        DoEvents
        Call ConectionExecute(pCadenaExec)
        pCadenaExec = ""
   End If

Loop
Close #fn

If Len(pCadenaExec) > 0 Then
     lblStatus.Caption = "Subiendo Registros (Espere!)"
     DoEvents
     Call ConectionExecute(pCadenaExec)
     pCadenaExec = ""
End If

lblStatus.Caption = "Revisando Ids de las Personas"
DoEvents

'Fix Cedula por Codigo de Empleado
strSQL = "exec spPrmCargado_Revision_Cedulas " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR & "," & iPago
Call ConectionExecute(strSQL)

lblStatus.Caption = ""

Me.MousePointer = vbDefault

prgProcesoMensual.Value = 1

Call Bitacora("Aplica", "PRM-CREDITO Carga Deducciones Inst:" & GLOBALES.gInstitucion)
Call sbBitacoraPlanilla("03", GLOBALES.gInstitucion, GLOBALES.glngFechaCR, "R", "Pla.Num." & iPago)

strSQL = "update instituciones set pr_carga = 1 where cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)

Call sbEstadoActualProceso

lblStatus.Caption = "Estado..."

Call sbReporteCargado(GLOBALES.glngFechaCR)

MsgBox "Información Cargada ...", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaDeduc_Csv_Integra()
Dim fn, strCadena As String, lng As Long, iPago As Integer
Dim strMonto As String, strSQL As String, rs As New ADODB.Recordset
Dim vPlanilla As String, vTipoAporte As String, vTipoCredito As String
Dim vCedula As String, vTemp  As String, i As Integer
Dim vCodigoDeduccion As String, vCodigoTipo As String

Dim pCadenaExec As String

On Error GoTo vError

fn = FreeFile
pCadenaExec = ""


strSQL = "select planilla,codigo_aportes,codigo_creditos from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
  vPlanilla = Trim(rs!planilla)
  vTipoAporte = Trim(rs!codigo_aportes & "")
  vTipoCredito = Trim(rs!codigo_creditos & "")
rs.Close

With frmContenedor.CD
 .DialogTitle = "Localice archivo con las deducciones de Planilla..."
 .Filter = "*.*"
 .InitDir = "C:\"
 .ShowOpen
End With

If frmContenedor.CD.FileName = "" Then
 MsgBox "Seleccione el Archivo de Deducciones del Proceso " & Format(GLOBALES.glngFechaCR, "####-##"), vbInformation
 Exit Sub
End If


MsgBox "Se procederá a cargar los registros del archivo :" & frmContenedor.CD.FileName, vbInformation

'CAPTURA EL NUMERO DE PAGO DE LA PLANILLA
strSQL = ""
Do While Not IsNumeric(strSQL)
  strSQL = InputBox("Digite el Numero de Planilla del Mes: ", "# Pago en el Mes...", 1)
Loop
iPago = strSQL


Me.MousePointer = vbHourglass

'Limpiando Informacion Anterior
lblStatus = "Borrando Información Anterior ..."
DoEvents

strSQL = "delete prm_cargado where fecha_proceso = " & GLOBALES.glngFechaCR _
       & " and pago = " & iPago & " and cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)


prgProcesoMensual.Value = 1
prgProcesoMensual.Max = 2

lblStatus = "Cargando ..."
DoEvents

Open frmContenedor.CD.FileName For Input As #fn   ' Lee el archivo.
 Do While Not EOF(fn)
   Input #fn, strCadena
   prgProcesoMensual.Max = prgProcesoMensual.Max + 1
 Loop
Close #fn

prgProcesoMensual.Min = 1
DoEvents



Open frmContenedor.CD.FileName For Input As #fn   'Lee el Archivo y lo compara
Do While Not EOF(fn)
   Input #fn, strCadena
   
        strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion) values(" _
               & GLOBALES.gInstitucion & "," & iPago & "," & GLOBALES.glngFechaCR
        
        If UCase(Right(frmContenedor.CD.FileName, 3)) = "CSV" Then
                vCodigoDeduccion = fxDescomponeCsv(3, strCadena)
                Select Case vCodigoDeduccion
                        Case vTipoAporte
                            vCodigoTipo = "1"
                        Case vTipoCredito
                            vCodigoTipo = "3"
                        Case Else
                            vCodigoTipo = "3"
                End Select
                strSQL = strSQL & "," & vCodigoTipo & ",'" & Val(fxDescomponeCsv(1, strCadena)) _
                      & "'," & fxDescomponeCsv(4, strCadena) & ",'" & vCodigoDeduccion & "')"
        
        Else
                vCodigoDeduccion = fxDescomponeCadena(2, strCadena)
                Select Case vCodigoDeduccion
                        Case vTipoAporte
                            vCodigoTipo = "1"
                        Case vTipoCredito
                            vCodigoTipo = "3"
                        Case Else
                            vCodigoTipo = "3"
                End Select
        
        
                strSQL = strSQL & "," & vCodigoTipo & ",'" & Val(fxDescomponeCadena(1, strCadena)) _
                      & "'," & fxDescomponeCadena(3, strCadena) & ",'" & vCodigoDeduccion & "')"
        
        End If
        If vCodigoDeduccion <> "codigodeduccion" Then
            pCadenaExec = pCadenaExec & Space(10) & strSQL
        End If

   
   If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
   
   If Len(pCadenaExec) > 25000 Then
        lblStatus.Caption = "Subiendo Registros (Espere!)  (" & prgProcesoMensual.Value & " / " & prgProcesoMensual.Max & ")     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
        DoEvents
        Call ConectionExecute(pCadenaExec)
        pCadenaExec = ""
   End If

Loop
Close #fn

If Len(pCadenaExec) > 0 Then
     lblStatus.Caption = "Subiendo Registros (Espere!)"
     DoEvents
     Call ConectionExecute(pCadenaExec)
     pCadenaExec = ""
End If

lblStatus.Caption = "Revisando Ids de las Personas"
DoEvents

'Fix Cedula por Codigo de Empleado
strSQL = "exec spPrmCargado_Revision_Cedulas " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR & "," & iPago
Call ConectionExecute(strSQL)

lblStatus.Caption = ""

Me.MousePointer = vbDefault

prgProcesoMensual.Value = 1

Call Bitacora("Aplica", "PRM-CREDITO Carga Deducciones Inst:" & GLOBALES.gInstitucion)
Call sbBitacoraPlanilla("03", GLOBALES.gInstitucion, GLOBALES.glngFechaCR, "R", "Pla.Num." & iPago)

strSQL = "update instituciones set pr_carga = 1 where cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)

Call sbEstadoActualProceso

lblStatus.Caption = "Estado..."

Call sbReporteCargado(GLOBALES.glngFechaCR)

MsgBox "Información Cargada ...", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





Private Sub sbCargaDeduc_Excel()
Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset

Dim vArchivo As String, vPasa As Boolean
Dim vPlanilla As String, vTipoAporte As String, vTipoCredito As String
Dim lng As Long, iPago As Integer, vExisteColPat As Boolean, i As Integer


Dim pCadenaExec As String

On Error GoTo vError


pCadenaExec = ""

strSQL = "select planilla,codigo_aportes,codigo_creditos from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
  vPlanilla = Trim(rs!planilla)
  vTipoAporte = Trim(rs!codigo_aportes & "")
  vTipoCredito = Trim(rs!codigo_creditos & "")
rs.Close


With frmContenedor.CD
    .InitDir = "C:\"
    .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]"
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

    vArchivo = .FileName

End With


MsgBox "Se procederá a cargar los registros del archivo :" & frmContenedor.CD.FileName, vbInformation

'CAPTURA EL NUMERO DE PAGO DE LA PLANILLA
strSQL = ""
Do While Not IsNumeric(strSQL)
  strSQL = InputBox("Digite el Numero de Planilla del Mes: ", "# Pago en el Mes...", 1)
Loop
iPago = strSQL

Me.MousePointer = vbHourglass

'Limpiando Informacion Anterior
lblStatus = "Borrando Información Anterior ..."
DoEvents

strSQL = "delete prm_cargado where fecha_proceso = " & GLOBALES.glngFechaCR _
       & " and pago = " & iPago & " and cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)

Set rsExcel = Excel_Load(vArchivo, "Import")

With rsExcel

vExisteColPat = False

For i = 0 To .Fields.Count - 1
   If UCase(.Fields.Item(i).Name) = "PATRONAL" Then
      vExisteColPat = True
      Exit For
   End If
Next i

Do While Not .EOF

  lblStatus.Caption = "Procesando Registro : " & lng & " de " & .RecordCount + 1
  DoEvents
  
  If Trim(!Cedula) <> "" Then
        If UCase(vTipoAporte) <> "NO" Then
          strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                 & " values(" & GLOBALES.gInstitucion & "," & iPago _
                 & "," & GLOBALES.glngFechaCR & ",1,'" & Trim(!Cedula) & "'," & CCur(!Aportes) & ",'O')"
          pCadenaExec = pCadenaExec & Space(10) & strSQL
        
          If vExisteColPat Then
                strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                       & " values(" & GLOBALES.gInstitucion & "," & iPago _
                       & "," & GLOBALES.glngFechaCR & ",1,'" & Trim(!Cedula) & "'," & CCur(!PATRONAL) & ",'P')"
                pCadenaExec = pCadenaExec & Space(10) & strSQL
          End If
        End If
        
        If UCase(vTipoCredito) <> "NO" Then
          strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                 & " values(" & GLOBALES.gInstitucion & "," & iPago _
                 & "," & GLOBALES.glngFechaCR & ",3,'" & Trim(!Cedula) & "'," & CCur(!abonos) & ",'C')"
         If !abonos > 0 Then
            'Call ConectionExecute(strSQL)
             pCadenaExec = pCadenaExec & Space(10) & strSQL
         End If
        
        End If
  End If
  
  lng = lng + 1
  
    
    If Len(pCadenaExec) > 25000 Then
         lblStatus.Caption = "Subiendo Registros (Espere!)  (" & lng & " / " & .RecordCount + 1 & ")"
         DoEvents
         Call ConectionExecute(pCadenaExec)
         pCadenaExec = ""
    End If
  
  
  .MoveNext
Loop
.Close

If Len(pCadenaExec) > 0 Then
     lblStatus.Caption = "Subiendo Registros (Espere!) "
     DoEvents
     Call ConectionExecute(pCadenaExec)
     pCadenaExec = ""
End If
   
End With

'fix para ASEGrupoHolcim
'strSQL = "update C set C.Cedula = S.cedula from prm_cargado C inner join Socios S on C.cedula = S.cedular" _
'       & " where C.cod_Institucion = " & GLOBALES.gInstitucion _
'       & " and C.fecha_Proceso = " & GLOBALES.glngFechaCR _
'       & " and C.Pago = " & iPago
'Call ConectionExecute(strSQL)

'Fix: Reemplaza el Codigo de Empleado por la Cedula de la Persona
strSQL = "exec spPrmCargado_Revision_Cedulas " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR & "," & iPago
Call ConectionExecute(strSQL)


lblStatus.Caption = ""

Me.MousePointer = vbDefault

prgProcesoMensual.Value = 1

Call Bitacora("Aplica", "PRM-CREDITO Carga Deducciones Inst:" & GLOBALES.gInstitucion)
Call sbBitacoraPlanilla("03", GLOBALES.gInstitucion, GLOBALES.glngFechaCR, "R", "Pla.Num." & iPago)

strSQL = "update instituciones set pr_carga = 1 where cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)

Call sbEstadoActualProceso

lblStatus.Caption = "Estado..."

Call sbReporteCargado(GLOBALES.glngFechaCR)

MsgBox "Información Cargada ...", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbCargaDeduc_Excel_Tek_Experts()
Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset

Dim vArchivo As String, vPasa As Boolean
Dim vPlanilla As String, vTipoAporte As String, vTipoCredito As String
Dim lng As Long, iPago As Integer, vExisteColPat As Boolean, i As Integer


Dim pCadenaExec As String

On Error GoTo vError


pCadenaExec = ""

strSQL = "select planilla,codigo_aportes,codigo_creditos from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
  vPlanilla = Trim(rs!planilla)
  vTipoAporte = Trim(rs!codigo_aportes & "")
  vTipoCredito = Trim(rs!codigo_creditos & "")
rs.Close


With frmContenedor.CD
    .InitDir = "C:\"
    .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]"
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

    vArchivo = .FileName

End With


MsgBox "Se procederá a cargar los registros del archivo :" & frmContenedor.CD.FileName, vbInformation

'CAPTURA EL NUMERO DE PAGO DE LA PLANILLA
strSQL = ""
Do While Not IsNumeric(strSQL)
  strSQL = InputBox("Digite el Numero de Planilla del Mes: ", "# Pago en el Mes...", 1)
Loop
iPago = strSQL

Me.MousePointer = vbHourglass

'Limpiando Informacion Anterior
lblStatus = "Borrando Información Anterior ..."
DoEvents

strSQL = "delete prm_cargado where fecha_proceso = " & GLOBALES.glngFechaCR _
       & " and pago = " & iPago & " and cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)

Set rsExcel = Excel_Load(vArchivo, "Import")

With rsExcel


Do While Not .EOF

  lblStatus.Caption = "Procesando Registro : " & lng & " de " & .RecordCount + 1
  DoEvents
  
  If Trim(!Cedula) <> "" Then
  
        '02-A06  [APORTE PATRONAL]
        '02-D30  [APORTE OBRERO]
        '02-D31  AHORRO EXTRAODRINARIO - [%]
        '02-D32  AHORRO NAVIDENO
        '02-D33  AHORRO ESCOLAR
        '02-D34  AHORRO VACACIONAL
        '02-D35  AHORRO MARCHAMO
        '02-D36  PRESTAMOS Y RECAUDOS
        '02-D37  OTRAS DEDUCCIONES
        '02-D38  AHORRO A LA VISTA
        
        '--Aportes Patronal y Obrero
        If ![02-A06] > 0 Then
          strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                 & " values(" & GLOBALES.gInstitucion & "," & iPago _
                 & "," & GLOBALES.glngFechaCR & ",1,'" & Trim(!Cedula) & "'," & ![02-A06] & ",'02-A06')"
          pCadenaExec = pCadenaExec & Space(10) & strSQL
        End If
        
        If ![02-D30] > 0 Then
          strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                 & " values(" & GLOBALES.gInstitucion & "," & iPago _
                 & "," & GLOBALES.glngFechaCR & ",1,'" & Trim(!Cedula) & "'," & ![02-D30] & ",'02-D30')"
          pCadenaExec = pCadenaExec & Space(10) & strSQL
        End If
        
        '--Creditos + Recaudos
        If ![02-D31] > 0 Then
          strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                 & " values(" & GLOBALES.gInstitucion & "," & iPago _
                 & "," & GLOBALES.glngFechaCR & ",3,'" & Trim(!Cedula) & "'," & ![02-D31] & ",'02-D31')"
          pCadenaExec = pCadenaExec & Space(10) & strSQL
        End If
  
        If ![02-D32] > 0 Then
          strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                 & " values(" & GLOBALES.gInstitucion & "," & iPago _
                 & "," & GLOBALES.glngFechaCR & ",3,'" & Trim(!Cedula) & "'," & ![02-D32] & ",'02-D32')"
          pCadenaExec = pCadenaExec & Space(10) & strSQL
        End If
  
        If ![02-D33] > 0 Then
          strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                 & " values(" & GLOBALES.gInstitucion & "," & iPago _
                 & "," & GLOBALES.glngFechaCR & ",3,'" & Trim(!Cedula) & "'," & ![02-D33] & ",'02-D33')"
          pCadenaExec = pCadenaExec & Space(10) & strSQL
        End If
  
        If ![02-D34] > 0 Then
          strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                 & " values(" & GLOBALES.gInstitucion & "," & iPago _
                 & "," & GLOBALES.glngFechaCR & ",3,'" & Trim(!Cedula) & "'," & ![02-D34] & ",'02-D34')"
          pCadenaExec = pCadenaExec & Space(10) & strSQL
        End If
  
        If ![02-D35] > 0 Then
          strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                 & " values(" & GLOBALES.gInstitucion & "," & iPago _
                 & "," & GLOBALES.glngFechaCR & ",3,'" & Trim(!Cedula) & "'," & ![02-D35] & ",'02-D35')"
          pCadenaExec = pCadenaExec & Space(10) & strSQL
        End If
  
        If ![02-D36] > 0 Then
          strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                 & " values(" & GLOBALES.gInstitucion & "," & iPago _
                 & "," & GLOBALES.glngFechaCR & ",3,'" & Trim(!Cedula) & "'," & ![02-D36] & ",'02-D36')"
          pCadenaExec = pCadenaExec & Space(10) & strSQL
        End If
  
        If ![02-D37] > 0 Then
          strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                 & " values(" & GLOBALES.gInstitucion & "," & iPago _
                 & "," & GLOBALES.glngFechaCR & ",3,'" & Trim(!Cedula) & "'," & ![02-D37] & ",'02-D37')"
          pCadenaExec = pCadenaExec & Space(10) & strSQL
        End If
  
        If ![02-D38] > 0 Then
          strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                 & " values(" & GLOBALES.gInstitucion & "," & iPago _
                 & "," & GLOBALES.glngFechaCR & ",3,'" & Trim(!Cedula) & "'," & ![02-D38] & ",'02-D38')"
          pCadenaExec = pCadenaExec & Space(10) & strSQL
        End If
  
  
  End If
  
  lng = lng + 1
  
    
    If Len(pCadenaExec) > 25000 Then
         lblStatus.Caption = "Subiendo Registros (Espere!)  (" & lng & " / " & .RecordCount + 1 & ")"
         DoEvents
         Call ConectionExecute(pCadenaExec)
         pCadenaExec = ""
    End If
  
  .MoveNext
Loop
.Close

If Len(pCadenaExec) > 0 Then
     lblStatus.Caption = "Subiendo Registros (Espere!) "
     DoEvents
     Call ConectionExecute(pCadenaExec)
     pCadenaExec = ""
End If
   
End With

'Fix de la Cedula en Caso de Utilizar Codigo de Empleado
strSQL = "exec spPrmCargado_Revision_Cedulas " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR & "," & iPago
Call ConectionExecute(strSQL)


lblStatus.Caption = ""

Me.MousePointer = vbDefault

prgProcesoMensual.Value = 1

Call Bitacora("Aplica", "PRM-CREDITO Carga Deducciones Inst:" & GLOBALES.gInstitucion)
Call sbBitacoraPlanilla("03", GLOBALES.gInstitucion, GLOBALES.glngFechaCR, "R", "Pla.Num." & iPago)

strSQL = "update instituciones set pr_carga = 1 where cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)

Call sbEstadoActualProceso

lblStatus.Caption = "Estado..."

Call sbReporteCargado(GLOBALES.glngFechaCR)

MsgBox "Información Cargada ...", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCargaDeduc_Excel_DxC_Costa_Rica()
Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset

Dim vArchivo As String, vPasa As Boolean
Dim vPlanilla As String, vTipoAporte As String, vTipoCredito As String
Dim lng As Long, iPago As Integer, vExisteColPat As Boolean, i As Integer

Dim pCedula     As String, pNombre As String
Dim pCadenaExec As String

On Error GoTo vError


pCadenaExec = ""

strSQL = "select planilla,codigo_aportes,codigo_creditos from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
  vPlanilla = Trim(rs!planilla)
  vTipoAporte = Trim(rs!codigo_aportes & "")
  vTipoCredito = Trim(rs!codigo_creditos & "")
rs.Close


With frmContenedor.CD
    .InitDir = "C:\"
    .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]"
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

    vArchivo = .FileName

End With


MsgBox "Se procederá a cargar los registros del archivo :" & frmContenedor.CD.FileName, vbInformation

'CAPTURA EL NUMERO DE PAGO DE LA PLANILLA
strSQL = ""
Do While Not IsNumeric(strSQL)
  strSQL = InputBox("Digite el Numero de Planilla del Mes: ", "# Pago en el Mes...", 1)
Loop
iPago = strSQL

Me.MousePointer = vbHourglass

'Limpiando Informacion Anterior
lblStatus = "Borrando Información Anterior ..."
DoEvents

strSQL = "delete prm_cargado where fecha_proceso = " & GLOBALES.glngFechaCR _
       & " and pago = " & iPago & " and cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)

Set rsExcel = Excel_Load(vArchivo, "")

With rsExcel


Do While Not .EOF

  lblStatus.Caption = "Procesando Registro : " & lng & " de " & .RecordCount + 1
  DoEvents
  
  If IsNumeric(.Fields(0).Value & "") Then
          
        pCedula = Trim(.Fields(0).Value)
          
        'Aporte Obrero
        If Not IsNull(.Fields(2).Value) Then
          If .Fields(2).Value > 0 Then
            strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                   & " values(" & GLOBALES.gInstitucion & "," & iPago _
                   & "," & GLOBALES.glngFechaCR & ",1,'" & pCedula & "'," & .Fields(2).Value & ",'DE16')"
            pCadenaExec = pCadenaExec & Space(10) & strSQL
          End If
        End If
  
        'Aporte Patronal
        If Not IsNull(.Fields(12).Value) Then
          If .Fields(12).Value > 0 Then
            strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                   & " values(" & GLOBALES.gInstitucion & "," & iPago _
                   & "," & GLOBALES.glngFechaCR & ",1,'" & pCedula & "'," & .Fields(12).Value & ",'DE15')"
            pCadenaExec = pCadenaExec & Space(10) & strSQL
          End If
        End If
  
        'Ahorro Escolar %
        If Not IsNull(.Fields(11).Value) Then
          If .Fields(11).Value > 0 Then
            strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                   & " values(" & GLOBALES.gInstitucion & "," & iPago _
                   & "," & GLOBALES.glngFechaCR & ",3,'" & pCedula & "'," & .Fields(11).Value & ",'DE31')"
            pCadenaExec = pCadenaExec & Space(10) & strSQL
          End If
        End If
  
        'Creditos
        If Not IsNull(.Fields(4).Value) Then
          If .Fields(4).Value + .Fields(5).Value + .Fields(6).Value + .Fields(7).Value + .Fields(8).Value > 0 Then
            strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                   & " values(" & GLOBALES.gInstitucion & "," & iPago _
                   & "," & GLOBALES.glngFechaCR & ",3,'" & pCedula & "'," & .Fields(4).Value + .Fields(5).Value + .Fields(6).Value + .Fields(7).Value + .Fields(8).Value & ",'DE17')"
            pCadenaExec = pCadenaExec & Space(10) & strSQL
          End If
        End If
  
        'Ahorros Extraordinarios
        If Not IsNull(.Fields(3).Value) Then
          If .Fields(3).Value + .Fields(10).Value > 0 Then
            strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                   & " values(" & GLOBALES.gInstitucion & "," & iPago _
                   & "," & GLOBALES.glngFechaCR & ",3,'" & pCedula & "'," & .Fields(3).Value + .Fields(10).Value & ",'DE14')"
            pCadenaExec = pCadenaExec & Space(10) & strSQL
          End If
        End If
  
        'Plan Mutual
        If Not IsNull(.Fields(9).Value) Then
          If .Fields(9).Value > 0 Then
            strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                   & " values(" & GLOBALES.gInstitucion & "," & iPago _
                   & "," & GLOBALES.glngFechaCR & ",3,'" & pCedula & "'," & .Fields(9).Value & ",'DE24')"
            pCadenaExec = pCadenaExec & Space(10) & strSQL
          End If
        End If
  
  
  
  
  End If 'Es Cedula
  
  
  lng = lng + 1
  
    
    If Len(pCadenaExec) > 25000 Then
         lblStatus.Caption = "Subiendo Registros (Espere!)  (" & lng & " / " & .RecordCount + 1 & ")"
         DoEvents
         Call ConectionExecute(pCadenaExec)
         pCadenaExec = ""
    End If
  
  .MoveNext
Loop
.Close

If Len(pCadenaExec) > 0 Then
     lblStatus.Caption = "Subiendo Registros (Espere!) "
     DoEvents
     Call ConectionExecute(pCadenaExec)
     pCadenaExec = ""
End If
   
End With

'Fix de la Cedula en Caso de Utilizar Codigo de Empleado
strSQL = "exec spPrmCargado_Revision_Cedulas " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR & "," & iPago
Call ConectionExecute(strSQL)


lblStatus.Caption = ""

Me.MousePointer = vbDefault

prgProcesoMensual.Value = 1

Call Bitacora("Aplica", "PRM-CREDITO Carga Deducciones Inst:" & GLOBALES.gInstitucion)
Call sbBitacoraPlanilla("03", GLOBALES.gInstitucion, GLOBALES.glngFechaCR, "R", "Pla.Num." & iPago)

strSQL = "update instituciones set pr_carga = 1 where cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)

Call sbEstadoActualProceso

lblStatus.Caption = "Estado..."

Call sbReporteCargado(GLOBALES.glngFechaCR)

MsgBox "Información Cargada ...", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub sbCargaDeduc_Excel_DxC_CentroAmerica()
Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset

Dim vArchivo As String, vPasa As Boolean
Dim vPlanilla As String, vTipoAporte As String, vTipoCredito As String
Dim lng As Long, iPago As Integer, vExisteColPat As Boolean, i As Integer

Dim pCedula     As String, pNombre As String
Dim pCadenaExec As String

On Error GoTo vError


pCadenaExec = ""

strSQL = "select planilla,codigo_aportes,codigo_creditos from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
  vPlanilla = Trim(rs!planilla)
  vTipoAporte = Trim(rs!codigo_aportes & "")
  vTipoCredito = Trim(rs!codigo_creditos & "")
rs.Close


With frmContenedor.CD
    .InitDir = "C:\"
    .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]"
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

    vArchivo = .FileName

End With


MsgBox "Se procederá a cargar los registros del archivo :" & frmContenedor.CD.FileName, vbInformation

'CAPTURA EL NUMERO DE PAGO DE LA PLANILLA
strSQL = ""
Do While Not IsNumeric(strSQL)
  strSQL = InputBox("Digite el Numero de Planilla del Mes: ", "# Pago en el Mes...", 1)
Loop
iPago = strSQL

Me.MousePointer = vbHourglass

'Limpiando Informacion Anterior
lblStatus = "Borrando Información Anterior ..."
DoEvents

strSQL = "delete prm_cargado where fecha_proceso = " & GLOBALES.glngFechaCR _
       & " and pago = " & iPago & " and cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)

Set rsExcel = Excel_Load(vArchivo, "")

With rsExcel


Do While Not .EOF

  lblStatus.Caption = "Procesando Registro : " & lng & " de " & .RecordCount + 1
  DoEvents
  
  If IsNumeric(.Fields(0).Value & "") Then
          
        pCedula = Trim(.Fields(0).Value)
          
        'Aporte Obrero
        If Not IsNull(.Fields(3).Value) Then
          If .Fields(3).Value > 0 Then
            strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                   & " values(" & GLOBALES.gInstitucion & "," & iPago _
                   & "," & GLOBALES.glngFechaCR & ",1,'" & pCedula & "'," & .Fields(3).Value & ",'DE16')"
            pCadenaExec = pCadenaExec & Space(10) & strSQL
          End If
        End If
  
        'Aporte Patronal
        If Not IsNull(.Fields(12).Value) Then
          If .Fields(12).Value > 0 Then
            strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                   & " values(" & GLOBALES.gInstitucion & "," & iPago _
                   & "," & GLOBALES.glngFechaCR & ",1,'" & pCedula & "'," & .Fields(12).Value & ",'DE15')"
            pCadenaExec = pCadenaExec & Space(10) & strSQL
          End If
        End If
  
        'Ahorro Escolar %
        If Not IsNull(.Fields(11).Value) Then
          If .Fields(11).Value > 0 Then
            strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                   & " values(" & GLOBALES.gInstitucion & "," & iPago _
                   & "," & GLOBALES.glngFechaCR & ",3,'" & pCedula & "'," & .Fields(11).Value & ",'DE31')"
            pCadenaExec = pCadenaExec & Space(10) & strSQL
          End If
        End If
  
        'Creditos
        If Not IsNull(.Fields(4).Value) Then
          If .Fields(4).Value + .Fields(5).Value + .Fields(7).Value + .Fields(9).Value > 0 Then
            strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                   & " values(" & GLOBALES.gInstitucion & "," & iPago _
                   & "," & GLOBALES.glngFechaCR & ",3,'" & pCedula & "'," & .Fields(4).Value + .Fields(5).Value + .Fields(7).Value + .Fields(9).Value & ",'DE17')"
            pCadenaExec = pCadenaExec & Space(10) & strSQL
          End If
        End If
  
        'Ahorros Extraordinarios
        If Not IsNull(.Fields(2).Value) Then
          If .Fields(2).Value + .Fields(8).Value > 0 Then
            strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                   & " values(" & GLOBALES.gInstitucion & "," & iPago _
                   & "," & GLOBALES.glngFechaCR & ",3,'" & pCedula & "'," & .Fields(2).Value + .Fields(8).Value & ",'DE14')"
            pCadenaExec = pCadenaExec & Space(10) & strSQL
          End If
        End If
  
        'Plan Mutual
        If Not IsNull(.Fields(6).Value) Then
          If .Fields(6).Value > 0 Then
            strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                   & " values(" & GLOBALES.gInstitucion & "," & iPago _
                   & "," & GLOBALES.glngFechaCR & ",3,'" & pCedula & "'," & .Fields(6).Value & ",'DE24')"
            pCadenaExec = pCadenaExec & Space(10) & strSQL
          End If
        End If
  
  
  
  
  End If 'Es Cedula
  
  
  lng = lng + 1
  
    
    If Len(pCadenaExec) > 25000 Then
         lblStatus.Caption = "Subiendo Registros (Espere!)  (" & lng & " / " & .RecordCount + 1 & ")"
         DoEvents
         Call ConectionExecute(pCadenaExec)
         pCadenaExec = ""
    End If
  
  .MoveNext
Loop
.Close

If Len(pCadenaExec) > 0 Then
     lblStatus.Caption = "Subiendo Registros (Espere!) "
     DoEvents
     Call ConectionExecute(pCadenaExec)
     pCadenaExec = ""
End If
   
End With

'Fix de la Cedula en Caso de Utilizar Codigo de Empleado
strSQL = "exec spPrmCargado_Revision_Cedulas " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR & "," & iPago
Call ConectionExecute(strSQL)


lblStatus.Caption = ""

Me.MousePointer = vbDefault

prgProcesoMensual.Value = 1

Call Bitacora("Aplica", "PRM-CREDITO Carga Deducciones Inst:" & GLOBALES.gInstitucion)
Call sbBitacoraPlanilla("03", GLOBALES.gInstitucion, GLOBALES.glngFechaCR, "R", "Pla.Num." & iPago)

strSQL = "update instituciones set pr_carga = 1 where cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)

Call sbEstadoActualProceso

lblStatus.Caption = "Estado..."

Call sbReporteCargado(GLOBALES.glngFechaCR)

MsgBox "Información Cargada ...", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCargaDeduc_ExcelNew()
Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset

Dim vArchivo As String, vPasa As Boolean
Dim vPlanilla As String, vTipoAporte As String, vTipoCredito As String
Dim lng As Long, iPago As Integer, vTipoPatronal As String


Dim pCadenaExec As String

On Error GoTo vError


pCadenaExec = ""

strSQL = "select I.planilla, I.codigo_creditos" _
       & ",isnull(Ta.COD_DEDUCCION, I.codigo_aportes) AS 'OBRERO'" _
       & ",isnull(TP.COD_DEDUCCION, 'x-PAT-x') AS 'PATRONAL'" _
       & " from instituciones I" _
       & " left join vPrm_Codigos_Patrimonio Ta on I.cod_institucion = Ta.cod_institucion and Ta.Tipo = 'O'" _
       & " left join vPrm_Codigos_Patrimonio Tp on I.cod_institucion = Tp.cod_institucion and Tp.Tipo = 'P'" _
       & " where I.cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
  vPlanilla = Trim(rs!planilla)
  vTipoAporte = Trim(rs!Obrero & "")
  vTipoPatronal = Trim(rs!PATRONAL & "")
  vTipoCredito = Trim(rs!codigo_creditos & "")
rs.Close


With frmContenedor.CD
    .InitDir = "C:\"
    .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]"
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

    vArchivo = .FileName

End With


MsgBox "Se procederá a cargar los registros del archivo :" & frmContenedor.CD.FileName, vbInformation

'CAPTURA EL NUMERO DE PAGO DE LA PLANILLA
strSQL = ""
Do While Not IsNumeric(strSQL)
  strSQL = InputBox("Digite el Numero de Planilla del Mes: ", "# Pago en el Mes...", 1)
Loop
iPago = strSQL

Me.MousePointer = vbHourglass

'Limpiando Informacion Anterior
lblStatus = "Borrando Información Anterior ..."
DoEvents

strSQL = "delete prm_cargado where fecha_proceso = " & GLOBALES.glngFechaCR _
       & " and pago = " & iPago & " and cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)

Set rsExcel = Excel_Load(vArchivo, "Import")

'Verifica Archivo
'TODO: CEDULA, NOMBRE, CODIGO, MONTO

With rsExcel

Do While Not .EOF

  lblStatus.Caption = "Procesando Registro : " & lng & " de " & .RecordCount + 1
  DoEvents
  
  If Trim(!Cedula) <> "" Then
  
        If Trim(!codigo) = vTipoAporte Or Trim(!codigo) = vTipoPatronal Then
            'Patrimonio
            If !Monto > 0 Then
                   strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                          & " values(" & GLOBALES.gInstitucion & "," & iPago _
                          & "," & GLOBALES.glngFechaCR & ",1,'" & Trim(!Cedula) & "'," & CCur(!Monto) & ",'" & Trim(!codigo) & "')"
                   pCadenaExec = pCadenaExec & Space(10) & strSQL
            End If
        
        Else
            'Creditos y Otras deducciones
            If !Monto > 0 Then
                   strSQL = "insert prm_cargado(cod_institucion,pago,fecha_proceso,tipo,cedula,monto,cod_deduccion)" _
                          & " values(" & GLOBALES.gInstitucion & "," & iPago _
                          & "," & GLOBALES.glngFechaCR & ",3,'" & Trim(!Cedula) & "'," & CCur(!Monto) & ",'" & Trim(!codigo) & "')"
                    pCadenaExec = pCadenaExec & Space(10) & strSQL
            End If
        End If
        
  End If
  
  lng = lng + 1
  
    
    If Len(pCadenaExec) > 25000 Then
         lblStatus.Caption = "Subiendo Registros (Espere!)  (" & lng & " / " & .RecordCount + 1 & ")"
         DoEvents
         Call ConectionExecute(pCadenaExec)
         pCadenaExec = ""
    End If
  
  
  .MoveNext
Loop
.Close

If Len(pCadenaExec) > 0 Then
     lblStatus.Caption = "Subiendo Registros (Espere!) "
     DoEvents
     Call ConectionExecute(pCadenaExec)
     pCadenaExec = ""
End If
   
End With

'Fix: Reemplaza el Codigo de Empleado por la Cedula de la Persona
strSQL = "exec spPrmCargado_Revision_Cedulas " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR & "," & iPago
Call ConectionExecute(strSQL)

lblStatus.Caption = ""

Me.MousePointer = vbDefault

prgProcesoMensual.Value = 1

Call Bitacora("Aplica", "PRM-CREDITO Carga Deducciones Inst:" & GLOBALES.gInstitucion)
Call sbBitacoraPlanilla("03", GLOBALES.gInstitucion, GLOBALES.glngFechaCR, "R", "Pla.Num." & iPago)

strSQL = "update instituciones set pr_carga = 1 where cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)

Call sbEstadoActualProceso

lblStatus.Caption = "Estado..."

Call sbReporteCargado(GLOBALES.glngFechaCR)

MsgBox "Información Cargada ...", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCrDesgloce()
Dim strSQL As String, rs As New ADODB.Recordset, iAplInco As Integer
Dim rsTmp As New ADODB.Recordset, rsTmp2 As New ADODB.Recordset
Dim vHistorico As Integer, i As Integer, lngFechaAnterior As Currency

lblStatus = "Limpiando y Actualizando ..."
DoEvents


' Eliminados por Procesos Complementarios
''Call sbCrActualizaSaldos   'Actualiza Estado de los creditos y corrige inconsistencias menores
''Call sbCrCalculaSaldoMes(True, GLOBALES.glngFechaCR)  'Actualiza Saldo del Mes

strSQL = "select isnull(pr_cr_aplica_incon,0) as 'Aplica',historico_cobro_envio, dbo.MyGetdate() as 'FechaServer'" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
  iAplInco = rs!Aplica
  vHistorico = rs!historico_cobro_envio
  mFechaSistema = rs!FechaServer
rs.Close


'Conserva X Meses de detalle (Histórico)
lngFechaAnterior = GLOBALES.glngFechaCR
For i = 1 To vHistorico
 lngFechaAnterior = fxFechaProcesoAnterior(lngFechaAnterior)
Next i

'Borra El Historial en el Proceso de Creditos
strSQL = "delete prm_creditos where Fecha_Proceso <= " & lngFechaAnterior _
       & " or Fecha_Proceso = " & GLOBALES.glngFechaCR _
       & " and cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)

'spPrmCreditoDesgloseNew(@Institucion int, @Proceso int, @FechaApl varchar(30), @AplicaInco smallint = 1,  @Inicializa smallint = 0, @Bloque smallin
Dim vTotal As Long, vPendientes As Long, vProcesados As Long

lblStatus.Caption = "Detallando los Abonos a créditos..."
DoEvents

strSQL = "exec spPrmCreditoDesgloseNew " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR _
       & ",'" & Format(mFechaSistema, "yyyy/mm/dd") & "'," & iAplInco & ",1,50"
Call OpenRecordSet(rs, strSQL)
  vTotal = rs!total + 1
      
  If vTotal = 0 Then
    vPendientes = 0
    vProcesados = 0
  Else
    vPendientes = rs!Pendientes
    vProcesados = rs!Procesados
  End If
rs.Close

prgProcesoMensual.Max = vTotal + 1
prgProcesoMensual.Value = vProcesados + 1

  
lblStatus.Caption = "Detallando..Registro # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
DoEvents

Do While vPendientes > 0
    strSQL = "exec spPrmCreditoDesgloseNew " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR _
           & ",'" & Format(mFechaSistema, "yyyy/mm/dd") & "'," & iAplInco & ",0,150"
    Call OpenRecordSet(rs, strSQL)
    
        vTotal = rs!total
        vPendientes = rs!Pendientes
        vProcesados = rs!Procesados
    
    rs.Close

  prgProcesoMensual.Value = vProcesados
  lblStatus.Caption = "Cargando..Registro # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
  DoEvents

Loop


'''Carga Listado para Procesamiento de Credito (Detallar abonos)
''strSQL = "exec spPrmCreditoListado " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR
''Call OpenRecordSet(rs, strSQL)
''
''
''prgProcesoMensual.Value = 1
''prgProcesoMensual.Max = rs.RecordCount + 2
''
''Do While Not rs.EOF
''  strSQL = "exec spPrmCreditoDetalleAbonos " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR & ",'" & Trim(rs!CEDULA) _
''         & "'," & rs!Monto & ",'" & Format(mFechaSistema, "yyyy/mm/dd") & "','N'," & rs!Mora & "," & rs!Refunde & "," & iAplInco
''  Call ConectionExecute(strSQL)
''  rs.MoveNext
''
''  If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
''  lblStatus.Caption = "Cargando..Registro # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
''  doEvents
''Loop
''rs.Close


'glogon.Conection.Execute "update par_ahcr set cr_des = 1"
'Call Bitacora("Aplica", "PRM-CREDITO Detalla las Deducciones")
'Call sbReporteDetalleDeducciones(GLOBALES.glngFechaCR)
'Call EstadoActualProceso

lblStatus.Caption = ""
Me.MousePointer = vbDefault
prgProcesoMensual.Value = 1

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical



End Sub


Private Sub sbGeneraPlanillaComp(vFechaProceso As Currency, Optional ByVal TablaBase As String = "R")
Dim strSQL As String

Me.MousePointer = vbHourglass

On Error GoTo pErrorCompra

strSQL = "exec spPrm_Planilla_Compra " & GLOBALES.gInstitucion & "," & vFechaProceso
Call ConectionExecute(strSQL)

Exit Sub

pErrorCompra:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbGeneraDeducciones(vFechaProceso As Currency)
Dim strSQL As String, rs As New ADODB.Recordset, rsTmp As New ADODB.Recordset
Dim iRespuesta As Integer, lngFechaAnterior As Currency
Dim vCompara As Integer, vComparaTipo As String
Dim vFecha As Date, i As Integer, vCodigoAporte As String
Dim vProcesaAporte As Boolean, vProcesaCreditos As Boolean, vHistorico As Integer


On Error GoTo vError

vFecha = fxFechaServidor

lblStatus.Visible = True

'Procedimientos Complementarios
Call sbProcesosAdd("02", "PRE", vFechaProceso)

'Carga variables que identifican si se procesan Creditos y Aportes
strSQL = "select codigo_aportes_env,codigo_creditos_env,historico_cobro_envio from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
    vProcesaAporte = IIf((Trim(UCase(rs!Codigo_Aportes_Env)) = "NO"), False, True)
    vProcesaCreditos = IIf((Trim(UCase(rs!codigo_creditos_env)) = "NO"), False, True)
    vCodigoAporte = Trim(rs!Codigo_Aportes_Env & "")
    vHistorico = rs!historico_cobro_envio
rs.Close


''''Actualiza Estado de los creditos y corrige inconsistencias menores
'''lblStatus.Caption = "Corrigiendo Inconsistencias Menores..."
'''doEvents
'''
'''Call sbCrActualizaSaldos

'CALCULA EL CAMPO SALDO MES PARA TODOS LOS REGISTROS ACTIVOS CON PROCESO NORMAL Y SALDO >0
lblStatus = "Calculando Saldo del Mes..."
DoEvents

Call sbCrCalculaSaldoMes(False, 0, False)

lblStatus = "Limpiando Información Anterior..."
DoEvents

'Conserva 6 Meses de detalle
lngFechaAnterior = vFechaProceso
For i = 1 To vHistorico + 1
 lngFechaAnterior = fxFechaProcesoAnterior(lngFechaAnterior)
Next i


If vProcesaCreditos Then

            strSQL = "delete PRM_ENVIADO_DETALLE where fecPro <= " & lngFechaAnterior _
                   & " or fecPro = " & vFechaProceso & " and cod_institucion = " & GLOBALES.gInstitucion
            Call ConectionExecute(strSQL)
        
        If chkPlanillaTransito.Value = vbChecked Then
           'Utilizar Procedimiento de Planilla en Transito
           'Realizar en espejo Abono en Falso
           Call sbCrEnviaConPlanillaTransito(vFechaProceso)

        Else
            
            lblStatus.Caption = "Procesando Cuotas Ordinarias..."
            DoEvents
            
            strSQL = "exec spPrmCreditoEnviaCuotaOrdinaria " & vFechaProceso & "," & GLOBALES.gInstitucion
            Call ConectionExecute(strSQL)
            
            lblStatus.Caption = "Procesando Cuotas Atrasadas..."
            DoEvents
            
            strSQL = "exec spPrmCreditoEnviaMora " & vFechaProceso & "," & GLOBALES.gInstitucion
            Call ConectionExecute(strSQL)
            
            
        End If
        
        'Aplica Deduccion Doble a los casos marcados con esta opción
        strSQL = "insert into PRM_ENVIADO_DETALLE(id_solicitud,codigo,fecpro,cedula,cuota,morosidad" _
               & ",cod_institucion,cargo,poliza,cod_deduccion, Cod_Divisa, IMPORTE)" _
               & " select id_solicitud,codigo,fecpro,cedula,cuota,morosidad,cod_institucion,cargo,poliza" _
               & ",cod_deduccion, Cod_Divisa, IMPORTE" _
               & " from PRM_ENVIADO_DETALLE " _
               & " Where fecpro = " & vFechaProceso & "  and cod_institucion = " & GLOBALES.gInstitucion _
               & " and cedula in(select cedula from socios where ind_doble_deduccion = 1)"
        Call ConectionExecute(strSQL)
        
        
        
''        '***********************************************************************************************************************************
''        'Parche Unicamente para ASECCSS con codigo Nuevo 29/07/2009
''        '***********************************************************************************************************************************
''        If GLOBALES.gInstitucion = 1 And GLOBALES.SysASEVersion Then
''            'Codigo Actual de Separacion
''            'Solo Cuotas Ordinarias
''            strSQL = "insert prm_planilla(cod_institucion,tipo,cedula,proceso,monto_actual,monto_anterior,fecha,movimiento) (" _
''                   & "select cod_institucion,'C',cedula," & vFechaProceso & ",isnull(sum(cuota),0) as Monto" _
''                   & ",0,'" & Format(vFecha, "yyyy/mm/dd") & "','P'" _
''                   & " From PRM_ENVIADO_DETALLE Where fecPro = " & vFechaProceso & " and Morosidad = 0" _
''                   & " and cod_institucion = " & GLOBALES.gInstitucion & " group by cedula,cod_institucion)"
''            Call ConectionExecute(strSQL)
''
''
''            ' Solo Morosidad para el Codigo Nuevo
''            strSQL = "insert prm_planilla(cod_institucion,tipo,cedula,proceso,monto_actual,monto_anterior,fecha,movimiento) (" _
''                   & "select cod_institucion,'M',cedula," & vFechaProceso & ",isnull(sum(cuota),0) as Monto" _
''                   & ",0,'" & Format(vFecha, "yyyy/mm/dd") & "','P'" _
''                   & " From PRM_ENVIADO_DETALLE Where fecPro = " & vFechaProceso & " and Morosidad = 1" _
''                   & " and cod_institucion = " & GLOBALES.gInstitucion & " group by cedula,cod_institucion)"
''            Call ConectionExecute(strSQL)
''        Else
''
''            'Procesa Bloque de Créditos
''            'Codigo Anterior
''            strSQL = "insert prm_planilla(cod_institucion,tipo,cod_deduccion,cedula,proceso,monto_actual,monto_anterior,fecha,movimiento) (" _
''                   & "select cod_institucion,'C',cedula," & vFechaProceso & ",isnull(sum(cuota),0) as Monto" _
''                   & ",0,'" & Format(vFecha, "yyyy/mm/dd") & "','P'" _
''                   & " From PRM_ENVIADO_DETALLE Where fecPro = " & vFechaProceso _
''                   & " and cod_institucion = " & GLOBALES.gInstitucion & " group by cod_institucion,cod_deduccion,cedula)"
''            Call ConectionExecute(strSQL)
''        End If
''        '***********************************************************************************************************************************
''        ' Fin del Parche
''        '***********************************************************************************************************************************
''

End If 'Procesa Creditos


'Nuevo Proceso: 01/07/2013
lblStatus.Caption = "Codificando..."
DoEvents

strSQL = "exec  spPrmProcCodigosSeparacion " & GLOBALES.gInstitucion & "," & vFechaProceso
Call ConectionExecute(strSQL)


lblStatus.Caption = "Consolidadon Envío de Deducciones..."
DoEvents

'Codificacion y Registra Cobros de Aportes de Patrimonio
strSQL = "exec  spPrmDeduccionCodifica_Envio " & GLOBALES.gInstitucion & "," & vFechaProceso & "," & chkRedondeo.Value
Call ConectionExecute(strSQL)


If chkCambioDeducciones.Value = vbChecked Then
  lblStatus.Caption = "Revisando y Aplicando Cambios Manuales (Espere)..."
  DoEvents
  
  strSQL = "exec spPrm_CreditoCambioDeducciones " & GLOBALES.gInstitucion & "," & vFechaProceso _
         & ",'" & glogon.Usuario & "'"
  Call ConectionExecute(strSQL)

End If


lblStatus.Caption = "Excluyendo Casos vía Politica General..."
DoEvents

 strSQL = "exec spPrm_Credito_Excluye_Casos " & GLOBALES.gInstitucion & "," & vFechaProceso
 Call ConectionExecute(strSQL)

lblStatus.Caption = "Generando Archivo de Formatos..."
DoEvents

      
strSQL = "select planilla,Planilla_envio,compara_indicador,compara_valor" _
       & " from instituciones where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)

If rs!Compara_Indicador = 1 Then

    lblStatus.Caption = "Revisando variaciones con planilla anterior..."
    DoEvents

    Call sbGeneraPlanillaComp(vFechaProceso, rs!compara_valor) 'Procedimiento de Comparacion
End If

'Corre Ajustes personalizados por Institución antes del Formato del Archivo
Select Case rs!Planilla_Envio
  Case "08" ' AYA Acueductos y Alcantarillados  / Requiere Comparacion
    
    'Redondedo forzado sin decimales
    strSQL = "update prm_planilla set monto_actual = ROUND(monto_actual,0)" _
           & ",monto_anterior = ROUND(monto_anterior,0)" _
           & " where proceso = " & vFechaProceso _
           & " and cod_institucion = " & GLOBALES.gInstitucion
    Call ConectionExecute(strSQL)
  
  Case "09" '[SPA] Mecanizada Tesoreria Nacional"
  
    'Si la planilla no lleva comparacion aplicar el parche siguiente
    If rs!Compara_Indicador = 0 Then
            lblStatus.Caption = "Redondeando Planilla (Espere)..."
            DoEvents
            
            strSQL = "update prm_planilla set monto_anterior = 0" _
                   & ", monto_actual = ROUND(monto_actual,0), Movimiento = 'I'" _
                   & " where proceso = " & vFechaProceso _
                   & " and cod_institucion = " & GLOBALES.gInstitucion
            Call ConectionExecute(strSQL)
    End If
    
  Case "13" '[INA] Instituto Nacional de Apendizaje
    'Redondedo forzado sin decimales
    strSQL = "update prm_planilla set monto_actual = ROUND(monto_actual,0)" _
           & ",monto_anterior = ROUND(monto_anterior,0)" _
           & " where proceso = " & vFechaProceso _
           & " and cod_institucion = " & GLOBALES.gInstitucion
    Call ConectionExecute(strSQL)
End Select
rs.Close


'Aplica Cambios Manuales
If chkCambioDeducciones.Value = vbChecked Then
    strSQL = "exec spPrm_CreditoCambioDeducciones " & GLOBALES.gInstitucion & "," & vFechaProceso & ",'" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)

End If


strSQL = "update instituciones set pr_genera = 1 where cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)


strSQL = "Transito.: " & IIf((chkPlanillaTransito.Value = vbChecked), "Si", "No")
If chkCambioDeducciones.Value = vbChecked Then
  strSQL = strSQL & "  - Cambios: Sí"
End If

Call sbBitacoraPlanilla("02", GLOBALES.gInstitucion, vFechaProceso, "E", strSQL)
Call Bitacora("Aplica", "Planilla Genera Deducciones Inst:" & GLOBALES.gInstitucion)


'Genera Archivos
Call sbGeneraArchivo_Main(vFechaProceso)

fraGenera.Visible = False

Call sbEstadoActualProceso

'Procedimientos Complementarios
Call sbProcesosAdd("02", "POS", vFechaProceso)


lblStatus.Caption = "Estado..."
 
Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbEstadoActualProceso()
Dim rs As New ADODB.Recordset, strSQL As String
Dim pProcesoClean As Long

lblInstitucion.Caption = GLOBALES.gNombreInstitucion
lblFechaProceso.Caption = Format(GLOBALES.glngFechaCR, "####-##")

pProcesoClean = GLOBALES.glngFechaCR

strSQL = "select *, isnull(Frecuencia,'M') as 'Frecuencia_Id'" _
       & " from instituciones " _
       & " where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)

cboMes.Text = fxConvierteMES(Val(Mid(GLOBALES.glngFechaCR, 5, 2)))
txtAno.Text = Mid(GLOBALES.glngFechaCR, 1, 4)

If rs.EOF And rs.BOF Then
  MsgBox "NO EXISTEN PARAMETROS DEL PROCESO - !! DEBE CREARLOS ANTES DE ENTRAR AQUI !! -"
  rs.Close
  UnLoad Me
End If
  

mFrecuencia = rs!Frecuencia_Id

cboFrecuencia.Clear
Select Case rs!Frecuencia_Id
    Case "M" 'Mensual
        cboFrecuencia.AddItem "Mensual"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "0"
        cboFrecuencia.Text = "Mensual"
    
        'Temporal
        cboFrecuencia.AddItem "1er Quincena"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "1"
        cboFrecuencia.AddItem "2da Quincena"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "2"
    
    Case "Q" 'Quincenal
        cboFrecuencia.AddItem "1er Quincena"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "1"
        cboFrecuencia.AddItem "2da Quincena"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "2"
        
        If (GLOBALES.glngFechaCR - pProcesoClean) = 0.1 Then
            lblFechaProceso.Caption = lblFechaProceso.Caption & "_Q1"
            cboFrecuencia.Text = "1er Quincena"
        Else
            lblFechaProceso.Caption = lblFechaProceso.Caption & "_Q2"
            cboFrecuencia.Text = "2da Quincena"
        End If
End Select

Call sbCbo_Copia(cboFrecuencia, cboFrecuenciaCambia)
  
  
'Aqui el Codigo
imgAplAhAplica.Visible = False
imgAplAhDevolucion.Visible = False
imgAplAhIncon.Visible = False

imgAplCrAplica.Visible = False
imgAplCrIncon.Visible = False
imgAplCrRecalculo.Visible = False

imgAplFecha.Visible = False
imgAplGenera.Visible = False
imgAplCarga.Visible = False
imgAplDesgloce.Visible = False
    
Call btnProceso_Click(0)
    
'Este bloque pasa a ser general
If rs!pr_genera = 1 Then
   optGeneral(2).Value = True 'Carga
   imgAplGenera.Visible = True
   imgAplFecha.Visible = True
   
   Call btnProceso_Click(0)
  
End If



If rs!pr_carga = 1 Then
   optGeneral(3).Value = True
   imgAplCarga.Visible = True

   Call btnProceso_Click(1)
End If


'Hay que Crear un estado nuevo del DEsgloce
If rs!pr_desgloza = 1 Then
  imgAplDesgloce.Visible = True
   Call btnProceso_Click(1)
Else
  imgAplDesgloce.Visible = False
End If



'Ahorros
If rs!pr_apAplica = 1 Then
   optAhorros(1).Value = True
   imgAplAhAplica.Visible = True
   Call btnProceso_Click(2)

End If

If rs!pr_apInco = 1 Then
   optAhorros(2).Value = True
   imgAplAhIncon.Visible = True
   Call btnProceso_Click(2)
End If

If rs!pr_apDev = 1 Then 'Cambio por Devoluciones
   imgAplAhDevolucion.Visible = True
   Call btnProceso_Click(2)
End If

'Credito

If rs!pr_crAplica = 1 Then
   optCreditos(1).Value = True
   imgAplCrAplica.Visible = True
   Call btnProceso_Click(2)
End If

If rs!pr_crInco = 1 Then
   optCreditos(2).Value = True
   imgRepCrIncon.Visible = True
   imgAplCrIncon.Visible = True
   Call btnProceso_Click(2)
End If

If rs!pr_crMora = 1 Then 'Recalculo de moratorios
   optCreditos(3).Value = True
   imgAplCrRecalculo.Visible = True
   Call btnProceso_Click(2)
End If

'imgGeneraArchivo.Visible = imgAplGenera.Visible

imgRepAhAplica.Visible = True
imgRepAhDevolucion.Visible = True
imgRepAhIncon.Visible = True

imgRepGenera.Visible = True
imgRepCarga.Visible = True
imgRepDesgloce.Visible = True

imgRepCrAplica.Visible = True
imgRepCrIncon.Visible = True

rs.Close

Call sbBitacoraConsulta

Call RefrescaTags(Me)

End Sub

Function fxValidaPaso(Optional pTransaccion As String = "08") As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vResultado As Boolean

'Verifica que no se haya aplicado los creditos en la fecha de proceso actual y/o futuras
'Lo cual indicaría un bloqueo general para la fecha de proceso actual por cuanto ya fue ejecutada o debió ser realizada

vResultado = True

If pTransaccion < "09" Then
    strSQL = "select isnull(count(*),0) as Existe from prm_bitacora where cod_institucion = " & GLOBALES.gInstitucion _
            & " and transaccion = '08' and proceso >= " & GLOBALES.glngFechaCR '" & pTransaccion & "
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe > 0 Then
       vResultado = False
       MsgBox "No se puede realizar el movimiento seleccionado ya que se ha aplicado esta planilla y/o otra futura en los auxiliares... verifique.!", vbExclamation
    End If
    rs.Close
End If

If pTransaccion = "05" Then
    strSQL = "select isnull(count(*),0) as Existe from prm_bitacora where cod_institucion = " & GLOBALES.gInstitucion _
            & " and transaccion = '05' and proceso >= " & GLOBALES.glngFechaCR '" & pTransaccion & "
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe > 0 Then
       vResultado = False
       MsgBox "No se puede realizar el movimiento seleccionado ya que se ha aplicado esta planilla y/o otra futura en los auxiliares... verifique.!", vbExclamation
    End If
    rs.Close
End If


If pTransaccion = "08" Then
  strSQL = "select dbo.fxPrmAplicacionValida(" & GLOBALES.glngFechaCR & "," & GLOBALES.gInstitucion & ") as Valida"
  Call OpenRecordSet(rs, strSQL)
  If rs!Valida = 0 Then
       vResultado = False
       MsgBox "La información detallada de los abonos no cuadra con la información cargada... verifique.!", vbExclamation
  End If
  rs.Close
End If


'Valida que para Aplicar Creditos o Patrimonio, de haya realizado el proceso de desgloce
If vResultado And (pTransaccion = "08" Or pTransaccion = "05") Then
        strSQL = "select isnull(count(*),0) as 'Valida' from prm_bitacora where cod_institucion = " & GLOBALES.gInstitucion _
                & " and proceso = " & GLOBALES.glngFechaCR & " and Transaccion = '04'"
  Call OpenRecordSet(rs, strSQL)
  If rs!Valida = 0 Then
       vResultado = False
       MsgBox "No se ha realizado el proceso de detalle de Aportes/Creditos... verifique.!", vbExclamation
  End If
  rs.Close

End If

If vResultado And pTransaccion = "08" Then
    
    Dim pProcesoClean As Long
    
    pProcesoClean = GLOBALES.glngFechaCR
    
    If GLOBALES.glngFechaCR - pProcesoClean = 0.1 Then
        strSQL = "select isnull(count(*),0) as 'Existe' from prm_bitacora where cod_institucion = " & GLOBALES.gInstitucion _
                & " and transaccion = '08' and (proceso = dbo.fxSIFPrmProcesoAnt(" & GLOBALES.glngFechaCR & ") " _
                & " or Proceso = dbo.fxSIFPrmProcesoAnt(" & pProcesoClean & ") )"
    Else
        strSQL = "select isnull(count(*),0) as 'Existe' from prm_bitacora where cod_institucion = " & GLOBALES.gInstitucion _
                & " and transaccion = '08' and proceso = dbo.fxSIFPrmProcesoAnt(" & GLOBALES.glngFechaCR & ")"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe = 0 Then
       vResultado = False
       
       
       
       MsgBox "La Planilla del Mes anterior no ha sido aplicada... verifique.!", vbExclamation
    End If
    rs.Close
End If


fxValidaPaso = vResultado

End Function






Private Sub imgCambiaFecha_Click()
Dim strSQL As String, rs As New ADODB.Recordset

Dim FechaProceso As String, pQuincena As Integer
Dim iMes As Integer, lngAnio As Long, vFecha As Date

On Error GoTo vError

If Not btnEjecucion.Item(0).Enabled Then
  MsgBox "Su usuario no tiene acceso al cambio de fecha!", vbExclamation
  Exit Sub
End If

imgCambiaFecha.BorderStyle = 1



'Cambia la Fecha de Calculo
iMes = fxConvierteMES(cboMes.Text)
lngAnio = txtAno
pQuincena = cboFrecuenciaCambia.ItemData(cboFrecuenciaCambia.ListIndex)

FechaProceso = lngAnio & Format(iMes, "00") & "." & pQuincena

GLOBALES.glngFechaCR = CCur(FechaProceso)

strSQL = "select dbo.fxSIFCorteAFecha(" & FechaProceso & ") as 'Corte'"
Call OpenRecordSet(rs, strSQL)
    vFecha = rs!Corte
rs.Close


strSQL = "update instituciones set pr_genera = 0, pr_carga = 0, pr_desgloza = 0" _
       & ", pr_apAplica = 0, pr_apInco = 0, pr_apDev = 0" _
       & ", pr_crAplica = 0, pr_crInco = 0, pr_crMora = 0 " _
       & ", pr_fecha_corte = '" & Format(vFecha, "yyyy/mm/dd") _
       & "' where cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)

'Cambia Fecha de Formalizaciones
strSQL = "select isnull(IND_CAMBIA_FECPRO,0) as Cambia" _
       & " from instituciones where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
If rs!Cambia = 1 Then
 strSQL = "update par_ahcr set cr_fecha_calculo = '" & Format(vFecha, "yyyy/mm/dd") _
        & "' where cr_fecha_calculo <= '" & Format(vFecha, "yyyy/mm/dd") & "'"
 Call ConectionExecute(strSQL)
End If
rs.Close

Call Bitacora("Aplica", "PRM-CREDITO Cambia Fecha Proceso Inst:" & GLOBALES.gInstitucion)
Call sbBitacoraPlanilla("01", GLOBALES.gInstitucion, GLOBALES.glngFechaCR, "E")

Call sbEstadoActualProceso

fraFechaProceso.Visible = False

MsgBox "La fecha de proceso fue cambiada a : " & GLOBALES.glngFechaCR

imgCambiaFecha.BorderStyle = 0
fraFechaProceso.Enabled = False

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Function fxRellenoArc(strValor As String, i As Integer) As String
Dim vRes As String, x As Integer

vRes = strValor

For x = Len(strValor) To i
  vRes = "0" & vRes
Next

fxRellenoArc = Right(vRes, i)

End Function

Private Function fxRellenoArc2(strValor As String, i As Integer) As String
Dim vRes As String, x As Integer

vRes = strValor

For x = Len(strValor) To i
  vRes = vRes & " "
Next

fxRellenoArc2 = Left(vRes, i)

End Function


Private Function fxLimpiaNombre(vNombre) As String
Dim i As Integer, vRes As String

vRes = ""

For i = 1 To Len(vNombre)
  If Mid(vNombre, i, 1) <> vbTab Then vRes = vRes & Mid(vNombre, i, 1)
Next i

fxLimpiaNombre = vRes

End Function

Private Sub sbGeneraArchivoF01_CCSS(vFechaProceso As Currency)
Dim strSQL As String, rs As New ADODB.Recordset

Dim vRuta As String, vTempo As String, i As Integer
Dim vFile As String, vArchivo As String, vFecha As Date
Dim fnFile

'Dim vMontoAnterior As Currency, vMonto As Currency

Dim vCodigoAportes As String, vCodigoCreditos As String, vCodigoCreditosAlterno As String
Dim vCodApoArc As String, vCodCreArc As String, vCodCreArcAlterno As String
Dim vLinea As String, y As Integer

'********************************************
'* Formato de Planilla de la C.C.S.S.       *
'********************************************

fnFile = FreeFile

vFecha = fxFechaServidor
vArchivo = ""

prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError

'CODIGOS DE ENVIO, y los del archivo
strSQL = "select * from instituciones where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
   vCodigoAportes = Format(Trim(rs!Codigo_Aportes_Env & ""), "0000")
   vCodigoCreditos = Format(Trim(rs!codigo_creditos_env & ""), "0000")
   vCodApoArc = Trim(rs!codigo_aportes & "")
   vCodCreArc = Trim(rs!codigo_creditos & "")

   vCodigoCreditosAlterno = Format(Trim(rs!codigo_creditos_alt_env & ""), "0000")
   vCodCreArcAlterno = Trim(rs!CODIGO_CREDITOS_ALT & "")
rs.Close

'vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
'         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F01].txt"

vArchivo = Format(GLOBALES.gInstitucion, "000") & "[" & Format(vFechaProceso, "####-##") & "]-ARC" & vCodApoArc & ".txt"

         
vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


If vCodigoAportes <> "NO" Then
    Open vTempo For Output As #fnFile  ' Create file name.
      lblStatus = "Creando archivo a enviar"
      DoEvents
    
      strSQL = "select P.*,S.nombre " _
             & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
             & " where P.Proceso = " & vFechaProceso & " and P.cod_institucion = " _
             & GLOBALES.gInstitucion & " and P.tipo = 'A' order by P.cedula"
      Call OpenRecordSet(rs, strSQL)
    
      prgProcesoMensual.Max = rs.RecordCount + 1
      prgProcesoMensual.Value = 1
    
      Do While Not rs.EOF
           vLinea = Trim(rs!Cedula)
           y = Len(vLinea)
           
           If y < 11 Then
              For i = 1 To 11 - y
                vLinea = "0" + vLinea
              Next i
           End If
           
           If Len(vLinea) > 11 Then vLinea = Mid(vLinea, 1, 11)
           
            vLinea = vLinea & "9999" & Mid(vCodigoAportes, 1, 4) & "00000000000000000000099999999    AHORRO00000000000000000000000000100000000"
        
           Print #fnFile, vLinea
         
         If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
         lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
         rs.MoveNext
      Loop
      rs.Close
    
    Close #fnFile

End If 'vCodigoAportes <> "NO"
  


'****************************************
' Crea Archivo para Creditos  Principal *
'****************************************

fnFile = FreeFile


vArchivo = Format(GLOBALES.gInstitucion, "000") & "[" & Format(vFechaProceso, "####-##") & "]-ARC" & vCodCreArc & ".txt"

vTempo = vRuta & "\" & vArchivo
vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Kill vTempo
End If


If vCodigoCreditos <> "NO" Then
    Open vTempo For Output As #fnFile  ' Create file name.
      lblStatus = "Creando archivo a enviar"
      DoEvents
    
      strSQL = "select P.*,S.nombre " _
             & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
             & " where P.Proceso = " & vFechaProceso & " and P.cod_institucion = " _
             & GLOBALES.gInstitucion & " and P.tipo = 'C' and P.cod_Deduccion = '" & vCodigoCreditos _
             & "' order by P.cedula"
      Call OpenRecordSet(rs, strSQL)
    
      prgProcesoMensual.Max = rs.RecordCount + 1
      prgProcesoMensual.Value = 1
    
      Do While Not rs.EOF
           vLinea = Trim(rs!Cedula)
           y = Len(vLinea)
           
           If y < 11 Then
              For i = 1 To 11 - y
                vLinea = "0" + vLinea
              Next i
           End If
           
           If Len(vLinea) > 11 Then vLinea = Mid(vLinea, 1, 11)
           
            vLinea = vLinea & "9999" & Mid(vCodigoCreditos, 1, 4) & Format((rs!Monto_Actual * 100), "0000000000000") & "0000000099999999" _
                   & "   CREDITO00000000000000000000000000100000000"
        
           Print #fnFile, vLinea
         
         If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
         lblStatus.Caption = "Creando Archivo 2 Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
         rs.MoveNext
      Loop
      rs.Close
    
    Close #fnFile
    
  
  
  
'**************************************
' Crea Archivo para Creditos  Alterno *
'**************************************
  
  If vCodigoCreditosAlterno <> vCodigoCreditos And vCodigoCreditosAlterno <> "NO" Then
        
        fnFile = FreeFile
        
        vArchivo = Format(GLOBALES.gInstitucion, "000") & "-ARC" & vCodCreArcAlterno & ".txt"
        
        vTempo = vRuta & "\" & vArchivo
        vFile = Dir(vTempo, vbArchive)
        
        If vFile = vArchivo Then  'El archivo existe
         Kill vTempo
        End If
        
        fnFile = FreeFile
        
            Open vTempo For Output As #fnFile  ' Create file name.
              lblStatus = "Creando archivo a enviar"
              DoEvents
            
              strSQL = "select P.*,S.nombre " _
                     & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
                     & " where P.Proceso = " & vFechaProceso & " and P.cod_institucion = " _
                     & GLOBALES.gInstitucion & " and P.tipo = 'C' and P.cod_deduccion = '" & vCodigoCreditosAlterno _
                     & "' order by P.cedula"
              Call OpenRecordSet(rs, strSQL)
              
              prgProcesoMensual.Value = 1
              prgProcesoMensual.Max = (rs.RecordCount + 2)

            
              Do While Not rs.EOF
                   vLinea = Trim(rs!Cedula)
                   y = Len(vLinea)
                   
                   If y < 11 Then
                      For i = 1 To 11 - y
                        vLinea = "0" + vLinea
                      Next i
                   End If
                   
                   If Len(vLinea) > 11 Then vLinea = Mid(vLinea, 1, 11)
                   
                    vLinea = vLinea & "9999" & Mid(vCodCreArcAlterno, 1, 4) & Format((rs!Monto_Actual * 100), "0000000000000") & "0000000099999999" _
                           & "   CREDITO00000000000000000000000000100000000"
                
                   Print #fnFile, vLinea
                 
                 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
                 lblStatus.Caption = "Creando Archivo 2 Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
                 rs.MoveNext
              Loop
              rs.Close
            
            Close #fnFile
  
  End If 'GLOBALES.SysASEVersion
    
    
End If 'vCodigoCreditos <> "NO"

Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
Call sbReporteGeneracionF02(vFechaProceso, vTempo)
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbGeneraArchivoF09_SPA(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String

Dim vMontoAnterior As Currency, vMonto As Currency
Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency
Dim vMovimiento As String


'
'Formato de Deducciones Sistema SPA. Tesoreria Nacional
'

On Error GoTo vError

fnFile = FreeFile
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vMovimiento = "in('"
  If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
  If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
  If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
  If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
  vMovimiento = vMovimiento & "P')"
rs.Close



'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError

vTempo = vRuta & "\ARC-DED.TXT"
vFile = Dir(vTempo, vbArchive)
If vFile = "ARC-DED.TXT" Then 'El archivo existe
  Close 'Cierra todos los archivos abiertos
  Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.

lblStatus = "Creando archivo a enviar"
DoEvents

strSQL = "select P.*,S.nombre,isnull(S.cod_sector,0) as Sector" _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " and P.movimiento " & vMovimiento _
       & " order by P.cedula,P.tipo,P.movimiento"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1


Do While Not rs.EOF
 
 Select Case rs!Movimiento
   Case "E" 'Exclusion
      i = 1
   Case "I" 'Inclusion
      i = 2
   Case "C" 'Cambio
      i = 3
   Case Else
      i = 4 ' Invalido, Porque se Mantiene o No se Proceso
  End Select
 
 'Ajuste el 19/11/2009 (Se mandan todas como inclusiones)
 i = 2
 
 
 Select Case rs!Tipo
   Case "A" 'Ahorros
     vCadena = i & vTipoAporte ' "463016"
   Case "E" 'Extraordinarios
     vCadena = i & vTipoAporte ' "463017"
   Case "C" 'Creditos
     vCadena = i & vTipoCredito ' "463018"
 End Select
 
 'Redondeo los Montos a Un Decimal
 vMonto = Format(rs!Monto_Actual, "################0.0")
 vMontoAnterior = Format(rs!Monto_Anterior, "################0.0")
 
 vCadena = vCadena & fxRellenoArc2(Mid(Trim(fxLimpiaNombre(rs!Nombre)), 1, 30), 30)
 vCadena = vCadena & Format(Mid(Trim(rs!Cedula), 1, 10), "0000000000")
 vCadena = vCadena & fxRellenoArc(Format((vMontoAnterior * 100), "000"), 8)
 vCadena = vCadena & fxRellenoArc(Format((vMonto * 100), "000"), 8)
 
 
 If i <> 4 Then Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
Call sbReporteGeneracionF02(vFechaProceso, vTempo)
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbGeneraArchivoF02_Integra(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, iRespuesta As Integer, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String

'********************************************
'* Formato INTEGRA Tesoreria Nacional       *
'********************************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  vMovimiento = "in('"
  If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
  If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
  If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
  If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
  vMovimiento = vMovimiento & "P')"
rs.Close



'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)
On Error GoTo vError


'Estandar
'vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
'         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F02].txt"
         
vArchivo = "E-" & vCodInstitucion & "-" & Year(vFecha) & Format(Month(vFecha), "00") _
         & Format(Day(vFecha), "00") & "-01.txt"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents


'*************************************************************
' Nota: En el nuevo procedimiento de planillas de Mecaniza
' se borran las deducciones de las personas en la aplicacion
' de la nueva planilla, por lo tanto se tienen que enviar todas
' las variaciones e inclusiones, pero no las exclusiones ya que
' estas son eliminadas por si solas.
'*************************************************************
strSQL = "select P.cedula,P.Tipo,P.Tipo,P.cod_deduccion,P.Movimiento,P.Monto_Actual,isnull(S.cod_sector,0) as Sector" _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.cod_deduccion, P.movimiento"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1


Do While Not rs.EOF
 
 'Campo 01: Cedula de 10 char, 2-4-4
 'Campo 02: Codigo de Deduccion Asignada
 'Campo 03: Monto o Valor (10 espacios)
 'Campo 04: Tipo de Aplicacion (Defecto 0 = Mensual,  parte la cuota en dos quincenas)
 'Campo 05: Codigo de Institucion (Defecto 0 = Aplica para Todas donde reciba salario)
 If Len(Trim(rs!Cedula)) = 9 Then
   vCadena = Mid(Format(Trim(rs!Cedula), "0000000000"), 1, 10) & vbTab
 Else
   vCadena = Trim(rs!Cedula) & vbTab
 End If
 
 Select Case rs!Tipo
   Case "A" 'Ahorros
'     vCadena = vCadena & vTipoAporte & vbTab & SIFGlobal.fxStringRelleno(Format(vPorcAhorro, "######0.00"), "I", " ", 10) & vbTab
     vCadena = vCadena & Trim(rs!cod_deduccion) & vbTab & Format(vPorcAhorro, "######0.00") & vbTab
   
   Case "E" 'Extraordinarios
'     vCadena = vCadena & vTipoAporte & vbTab & SIFGlobal.fxStringRelleno(Format(rs!monto_actual, "######0.00"), "I", " ", 10) & vbTab
     vCadena = vCadena & Trim(rs!cod_deduccion) & vbTab & Format(rs!Monto_Actual, "############0.00") & vbTab
   
   Case "C" 'Creditos
'     vCadena = vCadena & vTipoCredito & vbTab & SIFGlobal.fxStringRelleno(Format(rs!monto_actual, "######0.00"), "I", " ", 10) & vbTab
     vCadena = vCadena & Trim(rs!cod_deduccion) & vbTab & Format(rs!Monto_Actual, "############0.00") & vbTab
 End Select
 
 If rs!Sector = 2 Then
   'Pensionados (Cuando son pensionados se especifica xx)
   '2018-12-12 [PBN] Para Pensionados para 1 a 2 (indica la quincena en donde se rebajara)
    vCadena = vCadena & "2" & vbTab & "0"
 Else
   'Otros: 0 = Dividir la deduccion en las quincenas, 0 = No indica institucion donde labora.
    vCadena = vCadena & "0" & vbTab & "0"
 End If
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
  
'------------------------------------------------------------------------------------------------------------
'           Formatos Nuevos:
'------------------------------------------------------------------------------------------------------------

lblStatus.Caption = "Formato: Matricula de Operaciones"
DoEvents


vArchivo = "MD-" & vCodInstitucion & "-" & Year(vFecha) & Format(Month(vFecha), "00") _
         & Format(Day(vFecha), "00") & "-01.csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


strSQL = "exec spPrm_Formato_Integra_New_Matricula " & GLOBALES.gInstitucion & "," & vFechaProceso
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1

Do While Not rs.EOF
 
 vCadena = rs!Cadena
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
   
'----------------------------------------------------------------------------------------------------------------------
'Formato Nuevo:
lblStatus.Caption = "Formato: Integra"
DoEvents

vArchivo = "CD-" & vCodInstitucion & "-" & Year(vFecha) & Format(Month(vFecha), "00") _
         & Format(Day(vFecha), "00") & "-01.csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


strSQL = "exec spPrm_Formato_Integra_New " & GLOBALES.gInstitucion & "," & vFechaProceso
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1

Do While Not rs.EOF
 
 vCadena = rs!Cadena
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
  
  
  
  
  
'       Fin de Formatos Nuevos
'------------------------------------------------------------------------------------------------------------
  
Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
Call sbReporteGeneracionF02(vFechaProceso, vTempo)
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbGeneraArchivoF15_PJ(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, iRespuesta As Integer, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String

'***********************************************************
'* Formato Poder Judicial -> Variacion del formato INTEGRA *
'***********************************************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  vMovimiento = "in('"
  If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
  If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
  If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
  If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
  vMovimiento = vMovimiento & "P')"
rs.Close



'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)
On Error GoTo vError

'Estandar
'vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
'         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F15].txt"


vArchivo = "E-" & vCodInstitucion & "-" & Year(vFecha) & Format(Month(vFecha), "00") _
         & Format(Day(vFecha), "00") & "-01.txt"


vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents


'*************************************************************
' Nota: En el nuevo procedimiento de planillas de Mecaniza
' se borran las deducciones de las personas en la aplicacion
' de la nueva planilla, por lo tanto se tienen que enviar todas
' las variaciones e inclusiones, pero no las exclusiones ya que
' estas son eliminadas por si solas.
'*************************************************************
strSQL = "select P.*,S.nombre,isnull(S.cod_sector,0) as Sector " _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.movimiento"
       
       
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1


Do While Not rs.EOF
 
 'Campo 01: Cedula de 10 char, 2-4-4
 'Campo 02: Codigo de Deduccion Asignada
 'Campo 03: Monto o Valor (10 espacios)
 'Campo 04: Tipo de Aplicacion (Defecto 0 = Mensual,  parte la cuota en dos quincenas)
 
' vCadena = Mid(Format(Trim(rs!Cedula), "0000000000"), 1, 10) & vbTab
 vCadena = Format(Trim(rs!Cedula), "0000000000") & vbTab
 
 Select Case rs!Tipo
   Case "A" 'Ahorros
'     vCadena = vCadena & vTipoAporte & vbTab & SIFGlobal.fxStringRelleno(Format(vPorcAhorro, "######0.00"), "I", " ", 10) & vbTab
     vCadena = vCadena & rs!cod_deduccion & vbTab & Format(rs!Monto_Actual, "######0.00") & vbTab
   
'   Case "E" 'Extraordinarios
''     vCadena = vCadena & vTipoAporte & vbTab & SIFGlobal.fxStringRelleno(Format(rs!monto_actual, "######0.00"), "I", " ", 10) & vbTab
'     vCadena = vCadena & rs!Cod_Deduccion & vbTab & Format(rs!Monto_Actual, "############0.00") & vbTab
   Case Else '"C" 'Creditos
'     vCadena = vCadena & vTipoCredito & vbTab & SIFGlobal.fxStringRelleno(Format(rs!monto_actual, "######0.00"), "I", " ", 10) & vbTab
     vCadena = vCadena & rs!cod_deduccion & vbTab & Format(rs!Monto_Actual, "############0.00") & vbTab
 End Select
 
 If rs!Sector = 2 Then
   'Pensionados (Cuando son pensionados se especifica xx)
    vCadena = vCadena & "1"
 Else
   'Otros
    vCadena = vCadena & "0"
 End If
 
 
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 DoEvents
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
'Reporte Generico para Integra y PJ
Call sbReporteGeneracionF02(vFechaProceso, vTempo)
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbGeneraArchivoF18_CONAVI(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, iRespuesta As Integer, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String

'***************************************************
'* Formato CONAVI -> Variacion del formato INTEGRA *
'***************************************************


On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  vMovimiento = "in('"
  If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
  If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
  If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
  If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
  vMovimiento = vMovimiento & "P')"
rs.Close



'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)
On Error GoTo vError


vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F18].txt"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents


'*************************************************************
' Nota: En el nuevo procedimiento de planillas de Mecaniza
' se borran las deducciones de las personas en la aplicacion
' de la nueva planilla, por lo tanto se tienen que enviar todas
' las variaciones e inclusiones, pero no las exclusiones ya que
' estas son eliminadas por si solas.
'*************************************************************
strSQL = "select P.cedula,P.Tipo,P.Tipo,P.cod_deduccion,P.Movimiento,P.Monto_Actual,isnull(S.cod_sector,0) as Sector" _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.cod_deduccion, P.movimiento"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1


Do While Not rs.EOF
 
 'Campo 01: Cedula de 10 char, 2-4-4
 'Campo 02: Codigo de Deduccion Asignada
 'Campo 03: Monto o Valor (10 espacios)
 'Campo 04: Tipo de Aplicacion (Defecto 0 = Mensual,  parte la cuota en dos quincenas)
 'Campo 05: Caracter Identificador (Defecto 0 )
 
 vCadena = Mid(Format(Trim(rs!Cedula), "0000000000"), 1, 10) & vbTab
 
 Select Case rs!Tipo
   Case "A" 'Ahorros
'     vCadena = vCadena & vTipoAporte & vbTab & SIFGlobal.fxStringRelleno(Format(vPorcAhorro, "######0.00"), "I", " ", 10) & vbTab
     vCadena = vCadena & Trim(rs!cod_deduccion) & vbTab & Format(vPorcAhorro, "######0.00") & vbTab
   
   Case "E" 'Extraordinarios
'     vCadena = vCadena & vTipoAporte & vbTab & SIFGlobal.fxStringRelleno(Format(rs!monto_actual, "######0.00"), "I", " ", 10) & vbTab
     vCadena = vCadena & Trim(rs!cod_deduccion) & vbTab & Format(rs!Monto_Actual, "############0.00") & vbTab
   
   Case "C" 'Creditos
'     vCadena = vCadena & vTipoCredito & vbTab & SIFGlobal.fxStringRelleno(Format(rs!monto_actual, "######0.00"), "I", " ", 10) & vbTab
     vCadena = vCadena & Trim(rs!cod_deduccion) & vbTab & Format(rs!Monto_Actual, "############0.00") & vbTab
 End Select
 
 vCadena = vCadena & "0" & vbTab & "0"
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
Call sbReporteGeneracionF02(vFechaProceso, vTempo)
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbGeneraArchivo_Main(vFechaProceso As Currency)
Dim strSQL As String, rs As New ADODB.Recordset


Call sbBitacoraPlanilla("02.1", GLOBALES.gInstitucion, vFechaProceso, "E", "")


strSQL = "select planilla_envio from instituciones where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)

Select Case rs!Planilla_Envio
  Case "00" 'Formato Excel
    Call sbGeneraArchivoF00_Excel(vFechaProceso)
  
  Case "01" '[CCSS] Caja Costarricense Seguro Social
    Call sbGeneraArchivoF01_CCSS(vFechaProceso)
  
  Case "02" '[INTEGRA] Mecanizada Tesoreria Nacional
    Call sbGeneraArchivoF02_Integra(vFechaProceso) 'Formato Nuevo
  
  Case "03" '[ASECCSS] Asociacion Solidarista Emp CCSS
   If gPortal.Empresa_Id = 1 Then
        Call sbGeneraArchivoF03_ASECCSS(vFechaProceso, "01") 'Nuevo (ASE)
        Call sbGeneraArchivoF03_ASECCSS(vFechaProceso, "02") 'Nuevo (ASE)
        Call sbGeneraArchivoF03_ASECCSS(vFechaProceso, "03") 'Nuevo (ASE)
   Else
       Call sbGeneraArchivoF03_SIF(vFechaProceso) 'Anterior (SIF)
   End If
  
  
  Case "04" '[ICE](ACOTEL)Instituto Costarricense Electricidad
    Call sbGeneraArchivoF04_ICE_ACOTEL(vFechaProceso)
  
  Case "05" 'COOPECAJA  / Requiere Comparacion
    Call sbGeneraArchivoF05_CoopeCaja(vFechaProceso)
    Call sbGeneraArchivoF05_CoopeCaja_OLD(vFechaProceso)
  
  Case "06" ' ICE Oficinas Centrales  / Requiere Comparacion
    Call sbGeneraArchivoF06_ICE_Central(vFechaProceso)
  
  Case "07" ' ICE Proyectos  / Requiere Comparacion
    Call sbGeneraArchivoF07_ICE_Proyectos(vFechaProceso)
  
  Case "08" ' AYA Acueductos y Alcantarillados  / Requiere Comparacion
    Call sbGeneraArchivoF08_AyA(vFechaProceso)

  Case "09" '[SPA] Mecanizada Tesoreria Nacional"
    Call sbGeneraArchivoF09_SPA(vFechaProceso)  'Formato Anterior
  
  Case "10" '[SIF] Sistema SIF [F01.Indefinidos]"
    Call sbGeneraArchivoF10_SIF_Indefinidos(vFechaProceso)
  
  Case "11" '[SIF] Sistema SIF [F02.Plazo definido]"
  
  Case "12" '[IMAS]Institucto Mixto de Ayuda de Social"
    Call sbGeneraArchivoF12_IMAS(vFechaProceso)
    
  Case "13" '[INA] Instituto Nacional de Apendizaje"
    Call sbGeneraArchivoF13_INA(vFechaProceso)

  Case "14" '[MSJ] Municipalidad de San José"
    Call sbGeneraArchivoF14_MSJ(vFechaProceso)
  
  Case "15" '[ PJ] Poder Judicial"
    Call sbGeneraArchivoF15_PJ(vFechaProceso)
  
  Case "16" '[StarH] PriceWaterHouseCoopers"
    Call sbGeneraArchivoF16_StarH(vFechaProceso) 'Nuevo

  Case "17" '[UCR] Universidad de Costa Rica
    Call sbGeneraArchivoF17_UCR(vFechaProceso) 'Nuevo

  Case "18" '[CONAVI] Consejo Nacional de Vialidad
    Call sbGeneraArchivoF18_CONAVI(vFechaProceso) 'Nuevo

  Case "19" '[CGR] Contraloría General de la Republica
    Call sbGeneraArchivoF19_CGR(vFechaProceso)

  Case "20" '[CEN-CINAI]
    Call sbGeneraArchivoF20_CEN_CINAI(vFechaProceso)
    
  Case "21" '[UNAFEPROT]
    Call sbGeneraArchivoF21_UNATEPROT(vFechaProceso)
  
  Case "22" '[PANI]
    Call sbGeneraArchivoF22_PANI(vFechaProceso)
    
  Case "23" '[CORREOS]
    Call sbGeneraArchivoF23_CORREOS(vFechaProceso)
  
  Case "24" '[SERVICOOP]
    Call sbGeneraArchivoF24_SERVICOOP(vFechaProceso)
    
  Case "25", "30" '[AGH] Holcim"
    Call sbGeneraArchivoF25_Holcim(vFechaProceso)
  
  Case "26" '[JUPEMA] Junta de Pensiones y Jubilaciones
    Call sbGeneraArchivoF26_JUPEMA(vFechaProceso)
  
  Case "27" '[RECOPE] Refinadora Costarricense de Petrole
    Call sbGeneraArchivoF27_RECOPE(vFechaProceso)
  
  Case "28" '[ASTEK] Tek Experts
    Call sbGeneraArchivoF28_TekExperts(vFechaProceso)
  
  Case "29" '[P&G] Procter & Gamble
    Call sbGeneraArchivoF29_PyG(vFechaProceso)
 
  Case "31" 'Forza Cash Logicstic
    Call sbGeneraArchivoF31_Excel_ForzaCash(vFechaProceso)
  
  
  Case "32", "33" 'DxC Technology
    Call sbGeneraArchivoF32_DxC(vFechaProceso)
  
  Case "34" 'ASOECorr
    Call sbGeneraArchivoF34_ASOECorr(vFechaProceso)
  
  Case "35" 'ProGRX_RRHH
    Call sbGeneraArchivoF35_ProGrX_RRHH(vFechaProceso)
  
  
  Case "36" 'ASOINSVA
    Call sbGeneraArchivoF36_ASOINSVA(vFechaProceso)
  
End Select
rs.Close


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub imgGeneraArchivo_Click()
Dim vFechaProceso As Currency

On Error GoTo vError

vFechaProceso = CCur(InputBox("Especifique la fecha de proceso " & vbCrLf _
        & "La fecha de Proceso Actual es : " & GLOBALES.glngFechaCR, "Generación de la Planilla", GLOBALES.glngFechaCR))

Call sbGeneraArchivo_Main(vFechaProceso)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbGeneraArchivoF10_SIF_Indefinidos(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vMontoAnterior As Currency, vMonto As Currency, vTemp As String
Dim vCodInstitucion As String

vCodInstitucion = ""

'**************************************************
'* Formato Planilla Sistema SIF (F01.Indefinidos) *
'**************************************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError


vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F10].txt"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

strSQL = "select P.*,S.nombre " _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento <> 'M'" _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.movimiento"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1


Do While Not rs.EOF
 
 'Las exclusiones se envian en Cero
 'Campo 01: Cedula de 15 char, 1-4-4
 'Campo 02: Nombre de 50
 'Campo 03: Monto (con punto decimal y dos decimales)
 'Campo 04: Movimiento (I inclusion, E exclusion, C cambio) no se reportan los que se mantienen


 vCadena = SIFGlobal.fxStringRelleno(rs!Cedula, "D", " ", 15) & SIFGlobal.fxStringRelleno(rs!Nombre, "D", " ", 50) _
           & Format(rs!Monto_Actual, "000000000.00") & " " & rs!Movimiento

 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub



Private Sub sbGeneraArchivoF12_IMAS(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vMontoAnterior As Currency, vMonto As Currency, vTemp As String
Dim vCodInstitucion As String

vCodInstitucion = ""

'*********************************
'* Formato Planilla IMAS         *
'*********************************

On Error GoTo vError

'1. Se envia a deducir todos los datos, en cada planilla
'2. Los separadores entre campos es con TABS


fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError


       
vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F12].txt"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If

Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

strSQL = "select P.*,S.nombre " _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento <> 'E'" _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.movimiento"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1


Do While Not rs.EOF
 
 'Campo 01: Cedula de 10 char, 2-4-4
 'Campo 02: Nombre de 30
 'Campo 03: Proceso (Año-Mes)
 'Campo 04: Monto no debe contemplar el punto decimal (9 + 2 decimales)
 '            , y se asume que los últimos dos dígitos corresponden a los decimales.


    vTemp = Format(CCur(rs!Monto_Actual), "000000000.00")
    vTemp = Mid(vTemp, 1, 9) & Right(vTemp, 2) ' Mid(Text2.Text, 10, 2)
    vTemp = Format(CLng(vTemp), "000000000")
    

 vCadena = SIFGlobal.fxStringRelleno(rs!Cedula, "I", "0", 10) & Mid(SIFGlobal.fxStringRelleno(rs!Nombre, "D", " ", 30), 1, 30) & vFechaProceso & vTemp
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbGeneraArchivoF13_INA(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date
Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency


Dim vCodInstitucion As String
Dim vMontoAnterior As Currency, vMonto As Currency, vTemp As String

'************************
'* Formato Planilla INA *
'************************

On Error GoTo vError

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = ""
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
rs.Close


fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
          & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F13].txt"


vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

strSQL = "select P.*,S.nombre " _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.movimiento"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1


Do While Not rs.EOF
 
 'Las exclusiones se envian en Cero
 'Campo 00: Relleno de Ceros de 10 char
 'Campo 01: Cedula de 9 char, 2-4-4 (con cero a la izq)
 'Campo 02: Relleno de Ceros de 11 char
 'Campo 03: Monto no debe contemplar el punto decimal (8 caracteres con ceros a la izq)
 '            , y se asume que los últimos dos dígitos corresponden a los decimales.
 'Campo 04: Codigo de Institucion de 3 digitos
 'Campo 05: Relleno con la cadena siguente -> 0002000
 

    vTemp = Format(rs!Monto_Actual, "00000000.00")
    vTemp = Mid(vTemp, 1, 8) & Mid(vTemp, 10, 2)
    vTemp = Format(CLng(vTemp), "000000000")


 vCadena = SIFGlobal.fxStringRelleno("", "I", "0", 10) & SIFGlobal.fxStringRelleno(rs!Cedula, "I", "0", 10) & SIFGlobal.fxStringRelleno("", "I", "0", 11) & vTemp & vTipoCredito & "0002000"
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub




Private Sub sbGeneraArchivoF14_MSJ(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vCodInstitucion As String
Dim vMontoAnterior As Currency, vMonto As Currency, vTemp As String

'**************************************************
'* Formato Planilla MSJ Municipalidad de San Jose *
'**************************************************

On Error GoTo vError



fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError


vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
          & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F14].txt"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If

Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

strSQL = "select P.*,S.nombre " _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.movimiento"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1


Do While Not rs.EOF
 
 'Las exclusiones se envian en Cero
 'Campo 01: Cedula de 9 char, 1-4-4
 'Campo 02: Nombre de 30
 'Campo 03: Monto no debe contemplar el punto decimal
 '            , y se asume que los últimos dos dígitos corresponden a los decimales.

    vTemp = Format(rs!Monto_Actual, "00000000.00")
    vTemp = Mid(vTemp, 1, 8) & Mid(vTemp, 10, 2)
    vTemp = Format(CLng(vTemp), "000000000")


 vCadena = SIFGlobal.fxStringRelleno(rs!Cedula, "I", "0", 9) & SIFGlobal.fxStringRelleno(rs!Nombre, "D", " ", 30) & vTemp
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub sbGeneraArchivoF03_ASECCSS(vFechaProceso As Currency, pUnidad As String)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, iRespuesta As Integer, strCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date
Dim vMontoAnterior As Currency ', vMonto As Currency
Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency
Dim vMonto As String, vTipoMov As String

'******************************************
'* Formato Planilla ASECCSS  con Star H   *
'******************************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

strSQL = "select planilla,codigo_aportes,codigo_creditos,porc_ahorro from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vTipoAporte = Trim(rs!codigo_aportes & "")
  vTipoCredito = Trim(rs!codigo_creditos & "")
  vPorcAhorro = rs!porc_ahorro
rs.Close


Me.MousePointer = vbHourglass

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)


vArchivo = "F03-[" & Format(vFechaProceso, "####-##") & "] " & Year(vFecha) & Format(Month(vFecha), "00") _
         & Format(Day(vFecha), "00") & "-" & Format(GLOBALES.gInstitucion, "00") & "u" & pUnidad & ".txt"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents


'Solo cargar los casos que Cambian (E,I,C)
'" and P.movimiento <> 'M'"

strSQL = "select P.*,S.nombre " _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula and S.up = '" & pUnidad & "'" _
       & " where P.Proceso = " & vFechaProceso _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.movimiento"
Call OpenRecordSet(rs, strSQL)



prgProcesoMensual.Max = rs.RecordCount + 2
prgProcesoMensual.Value = 1

Do While Not rs.EOF
 
 'Campo 01: Cedula de 15 char, 2-4-4
 'Campo 02: Codigo de Deduccion Asignada
 'Campo 03: Monto o Valor
 'Campo 04: Tipo de Aplicacion (Defecto 0 = Mensual,  parte la cuota en dos quincenas)
 'Campo 05: Codigo de Institucion (Defecto 0 = Aplica para Todas donde reciba salario)
 
 strSQL = Format(rs!Monto_Actual, "00000000.00")
 vMonto = ""
 For i = 1 To Len(strSQL)
  If Mid(strSQL, i, 1) <> "." Then
    vMonto = vMonto & Mid(strSQL, i, 1)
  End If
 Next i
 
 
 Select Case rs!Movimiento
   Case "E" 'Exclusion
      vTipoMov = "B"
   Case "I" 'Inclusion
      vTipoMov = "F"
   Case "C" 'Cambio
      vTipoMov = "F"
   Case Else
      vTipoMov = "F"
  End Select
 
 
 Select Case rs!Tipo
   Case "A" 'Aportes
         strCadena = fxRellenoArc2(Format(Trim(rs!Cedula), "0000000000"), 15) & Trim(vTipoAporte) & " " & vTipoMov & " " & vMonto _
                   & " 0000000000 0000000000 01/" & Mid(CStr(vFechaProceso), 5, 2) & "/" & Mid(CStr(vFechaProceso), 1, 4) _
                   & Space(12) & Format(vPorcAhorro, "00.00") & Space(1) & vFechaProceso & "1"
  
         If rs!Movimiento <> "M" Then
      '       Print #fnFile, strCadena
         End If
   
   Case "C" 'Creditos
         strCadena = fxRellenoArc2(Format(Trim(rs!Cedula), "0000000000"), 15) & Trim(vTipoCredito) & " " & vTipoMov & " " & vMonto _
                   & " 0000000000 0000000000 01/" & Mid(CStr(vFechaProceso), 5, 2) & "/" & Mid(CStr(vFechaProceso), 1, 4) _
                   & Space(12) & "00.00 " & vFechaProceso & "1"
         Print #fnFile, strCadena
   
  End Select
 
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

'Caracter de Cierre
Print #fnFile, "!"

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbGeneraArchivoF16_StarH(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date
Dim vMontoAnterior As Currency ', vMonto As Currency
Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency
Dim vMonto As String, vTipoMov As String, vFechaCorte  As Date
Dim vCodInstitucion As String

vCodInstitucion = ""

'***********************************
'* Formato Star H, Planilla ASECCSS
'***********************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor
vFechaCorte = CDate(Mid(vFechaProceso, 1, 4) & "/" & Mid(vFechaProceso, 5, 2) & "/01")
vFechaCorte = fxPrmUltimoDiaMes(vFechaCorte)

vArchivo = ""
prgProcesoMensual.Min = 1

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
rs.Close


Me.MousePointer = vbHourglass

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError


vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F16].txt"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If

Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents


'Solo cargar los casos que Cambian (E,I,C)
'" and P.movimiento <> 'M'"

strSQL = "select P.*,S.nombre " _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.movimiento"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1


Do While Not rs.EOF
 
 'Campo 01: Cedula de 15 char, 2-4-4
 'Campo 02: Codigo de Deduccion Asignada
 'Campo 03: Tipo de Movimiento (Exclusion, Inclusion, Cambio)
 'Campo 04: Monto a deducir
 'Campo 05: Relleno Estandar + Fechas
 
 strSQL = Format(rs!Monto_Actual, "00000000.00")
 vMonto = ""
 
 For i = 1 To Len(strSQL)
  If Mid(strSQL, i, 1) <> "." Then
    vMonto = vMonto & Mid(strSQL, i, 1)
  End If
 Next i
 
 
 Select Case rs!Movimiento
   Case "E" 'Exclusion
      vTipoMov = "B"
   Case "I" 'Inclusion
      vTipoMov = "F"
   Case "C" 'Cambio
      vTipoMov = "F"
   Case Else
      vTipoMov = "F"
  End Select
 
 
 Select Case rs!Tipo
   Case "A" 'Aportes
         vCadena = fxRellenoArc2(Format(Trim(rs!Cedula), "0000000000"), 15) & Trim(vTipoAporte) & " " & vTipoMov & " " & vMonto _
                   & " 0000000000 0000000000 01/" & Mid(CStr(vFechaProceso), 5, 2) & "/" & Mid(CStr(vFechaProceso), 1, 4) _
                   & Space(12) & Format(vPorcAhorro, "00.00") & Space(1) & vFechaProceso & "1"
  
         If rs!Movimiento <> "M" Then
             Print #fnFile, vCadena
         End If
   
   Case "C" 'Creditos
         vCadena = fxRellenoArc2(Format(Trim(rs!Cedula), "0000000000"), 15) & Trim(vTipoCredito) & " " & vTipoMov & " " & vMonto _
                   & " 0000000000 0000000000 01/" & Mid(CStr(vFechaProceso), 5, 2) & "/" & Mid(CStr(vFechaProceso), 1, 4) _
                   & Space(12) & "00.00 " & vFechaProceso & "1"
         Print #fnFile, vCadena
  End Select
 
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

'Caracter de Cierre
Print #fnFile, "!"

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub sbGeneraArchivoF17_UCR(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date
Dim vMontoAnterior As Currency ', vMonto As Currency
Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency
Dim vMonto As String, vTipoMov As String
Dim vCodInstitucion As String

vCodInstitucion = ""

'*****************************************
'* Formato UCR Universidad de Costa Rica *
'*****************************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
rs.Close


Me.MousePointer = vbHourglass

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError


vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F17].XML"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If

Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents


'Redondear a Cero decimal
strSQL = "update prm_planilla set monto_actual = ROUND(monto_actual,0)" _
       & ",monto_anterior = ROUND(monto_anterior,0)" _
       & " where proceso = " & vFechaProceso _
       & " and cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)


'Se envian todos los casos

strSQL = "select P.*,S.nombre " _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.tipo = 'C'" _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.movimiento"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 2
prgProcesoMensual.Value = 1

'Genera Formato XML con esquemas
'Print #fnFile, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?>"
Print #fnFile, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "utf-8" & Chr(34) & "?>"
Print #fnFile, "<Deducciones_Externas>"
Print #fnFile, "  <xs:schema id=" & Chr(34) & "Deducciones_Externas" & Chr(34) & " xmlns=" & Chr(34) & Chr(34) & " xmlns:xs=" & Chr(34) & "http://www.w3.org/2001/XMLSchema" & Chr(34) & " xmlns:msdata=" & Chr(34) & "urn:schemas-microsoft-com:xml-msdata" & Chr(34) & ">"
Print #fnFile, "    <xs:element name=" & Chr(34) & "Deducciones_Externas" & Chr(34) & " msdata:IsDataSet=" & Chr(34) & "true" & Chr(34) & " msdata:MainDataTable=" & Chr(34) & "Deduccion" & Chr(34) & " msdata:UseCurrentLocale=" & Chr(34) & "true" & Chr(34) & ">"
Print #fnFile, "      <xs:complexType>"
Print #fnFile, "        <xs:choice minOccurs=" & Chr(34); "0" & Chr(34) & " maxOccurs=" & Chr(34); "unbounded" & Chr(34) & ">"
Print #fnFile, "          <xs:element name=" & Chr(34) & "Deduccion" & Chr(34) & ">"
Print #fnFile, "            <xs:complexType>"
Print #fnFile, "              <xs:sequence>"
Print #fnFile, "                <xs:element name=" & Chr(34) & "Identificacion" & Chr(34) & " type=" & Chr(34) & "xs:string" & Chr(34) & " />"
Print #fnFile, "                <xs:element name=" & Chr(34) & "Nombre" & Chr(34) & " type=" & Chr(34) & "xs:string" & Chr(34) & " />"
Print #fnFile, "                <xs:element name=" & Chr(34) & "Valor" & Chr(34) & " type=" & Chr(34) & "xs:double" & Chr(34) & " />"
Print #fnFile, "              </xs:sequence>"
Print #fnFile, "            </xs:complexType>"
Print #fnFile, "          </xs:element>"
Print #fnFile, "        </xs:choice>"
Print #fnFile, "      </xs:complexType>"
Print #fnFile, "    </xs:element>"
Print #fnFile, "  </xs:schema>"

Do While Not rs.EOF
 
    Print #fnFile, "<Deduccion>"
    Print #fnFile, "    <Identificacion>" & Trim(rs!Cedula) & "</Identificacion>"
    Print #fnFile, "    <Nombre>" & Mid(Trim(rs!Nombre), 1, 30) & "</Nombre>"
    Print #fnFile, "    <Valor>" & rs!Monto_Actual & "</Valor>"
    Print #fnFile, " </Deduccion>"
 

 'Formato Antiguo
 'Campo 01: Cedula de 9 char, 1-4-4
 'Campo 02: Nombre de 30 char
 'Campo 03: Fecha (Mes/Año(2 Digitos)) Ej. 1208
 'Campo 04: Monto a deducir (8)
'
' vCadena = SIFGlobal.fxStringRelleno(Trim(rs!Cedula), "I", " ", 9) & SIFGlobal.fxStringRelleno(Trim(rs!Nombre), "D", " ", 30)
' vCadena = vCadena & Format(Month(vFecha), "00") & Mid(Year(vFecha), 3, 2) & Mid(Format(rs!monto_actual * 100, "00000000"), 1, 8)
 
' Print #fnFile, vCadena
 
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close
Print #fnFile, "</Deducciones_Externas>"


Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub sbGeneraArchivoF19_CGR(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, iRespuesta As Integer, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String, strMonto As String

'**********************************************
'* Formato Contraloría General de la Republica*
'**********************************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  vMovimiento = "in('"
  If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
  If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
  If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
  If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
  vMovimiento = vMovimiento & "P')"
rs.Close



'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)
On Error GoTo vError

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F19].txt"


vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents


'*************************************************************
' Nota: En el nuevo procedimiento de planillas de Mecaniza
' se borran las deducciones de las personas en la aplicacion
' de la nueva planilla, por lo tanto se tienen que enviar todas
' las variaciones e inclusiones, pero no las exclusiones ya que
' estas son eliminadas por si solas.
'*************************************************************
strSQL = "select S.Nombre, P.cedula,P.Tipo,P.Tipo,P.cod_deduccion,P.Movimiento,P.Monto_Actual,isnull(S.cod_sector,0) as Sector" _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.cod_deduccion, P.movimiento"
Call OpenRecordSet(rs, strSQL)

If rs.RecordCount > 0 Then
    prgProcesoMensual.Max = rs.RecordCount + 1
    prgProcesoMensual.Value = 1
End If



'Formato:
'
'2533195CORRALES VASQUEZ FABIOLA      0112320145000000000000589357
'2553551CORRALES VASQUEZ FABIOLA      0112320145000000000000500000
'
'Donde el significado es el que se detalla:
'2:                         Código inclusión (Siempre es “2”, si quieren realizar una exclusión deben eliminar del todo la línea del funcionario del archivo texto), (1 espacio)
'533195:                Número de autoridad deductora (Ustedes ya tienen dos números asignados: 535195 Amortización a préstamo y 553551 ahorro). (6 espacios)
'Corrales Vasquez Fabiola: Nombre del funcionario (30 espacios – completar con espacios en blanco hasta agotar los 30 espacios asignados)
'0112320145:         Número de cédula, (10 espacios).
'000000000:           Cuota anterior (Siempre debe estar en cero) (9 espacios).
'000589357:           Cuota por aplicar mensual (Se coloca el monto por aplicar, siempre se dejan los espacios para los decimales aunque sean decimales de “ceros”).
'por ejemplo si envían ese monto del ejemplo se aplica en la planilla 5.893,57 colones. ( para completar los 9 espacios se incluyen ceros adelante).



Do While Not rs.EOF
 
 Select Case rs!Tipo
   Case "A" 'Ahorros
         strMonto = Format(vPorcAhorro, "######0.00")
   Case "E" 'Extraordinarios
         strMonto = Format(rs!Monto_Actual, "############0.00")
   Case "C" 'Creditos
         strMonto = Format(rs!Monto_Actual, "############0.00")
 End Select
 
 strMonto = Replace(strMonto, ".", "")
 
 vCadena = "2" & Trim(rs!cod_deduccion) & " " & SIFGlobal.fxStringRelleno(rs!Nombre, "D", " ", 28) & " " & SIFGlobal.fxStringRelleno(rs!Cedula, "I", "0", 10) _
         & SIFGlobal.fxStringRelleno("", "I", "0", 9) _
         & SIFGlobal.fxStringRelleno(strMonto, "I", "0", 9)
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
Call sbReporteGeneracionF02(vFechaProceso, vTempo)
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbGeneraArchivoF20_CEN_CINAI(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, iRespuesta As Integer, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String
Dim vFechaInicio As Date, vFechaCorte As Date

'***************************
'* Formato CEN_CINAI       *
'***************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte" _
       & ", dbo.fxSIFCorteAFechaInicio(" & vFechaProceso & ") as 'FechaInicio'" _
       & ", dbo.fxSIFCorteAFecha(" & vFechaProceso & ") as 'FechaCorte'" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  vFechaInicio = rs!FechaInicio
  vFechaCorte = rs!FechaCorte
  vMovimiento = "in('"
  If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
  If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
  If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
  If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
  vMovimiento = vMovimiento & "P')"
rs.Close



'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)
On Error GoTo vError


vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F20].csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

'El monto es quincenal hay que dividirlo entre 2
strSQL = "select P.cedula,P.Tipo,P.Tipo,P.cod_deduccion,P.Movimiento,P.Monto_Actual/2 as 'Monto_Actual'" _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.cod_deduccion, P.movimiento"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1

'Socio negocios Codigo (formato XXX-XXXX)    Texto   500-0003
'Tipo de Deducción Codigo de la deduccion (XXXXX)    Texto   00301 para Credito
'cedula/Identificacion formato (XXXXXXXXXXXX) Nacionales no llevan cero delante  Texto   108650795
'Referencia  Texto   201708-00301
'Método de calculo (0) porcentaje (1) valor  Número entero   1
'Valor del procentaje o del monto a rebajar  "Monto
'(decimales separados por punto)"    10000.25
'Monto Préstamo  "Monto
'(decimales separados por punto)"    10000.25
'Fecha de Inicio formato (YYYYMMDD)  Fecha   20170801
'Fecha Fin formato (YYYYMMDD)    Fecha   20170831
'Controla Saldo (0) NO 1 (SI)    Número entero   0
'Norma Int. Monedas (CRC,USD) ver Norma ISO-4217 Formato (XXX)   Texto   CRC


Do While Not rs.EOF
        
 Select Case rs!Tipo
   Case "A" 'Ahorros
        vCadena = vCodInstitucion & "," & vTipoAporte & "," & Trim(rs!Cedula) & "," & Mid(GLOBALES.gstrNombreEmpresa, 1, 30) & ",1," _
               & rs!Monto_Actual & "," & 0 & "," & Format(vFechaInicio, "yyyymmdd") & "," & Format(vFechaCorte, "yyyymmdd") _
               & ",0,CRC"
   
   Case "E" 'Extraordinarios
        vCadena = vCodInstitucion & "," & vTipoCredito & "," & Trim(rs!Cedula) & "," & Mid(GLOBALES.gstrNombreEmpresa, 1, 30) & ",1," _
               & rs!Monto_Actual & "," & 0 & "," & Format(vFechaInicio, "yyyymmdd") & "," & Format(vFechaCorte, "yyyymmdd") _
               & ",0,CRC"
   
   Case "C" 'Creditos
        vCadena = vCodInstitucion & "," & vTipoCredito & "," & Trim(rs!Cedula) & "," & Mid(GLOBALES.gstrNombreEmpresa, 1, 30) & ",1," _
               & rs!Monto_Actual & "," & 0 & "," & Format(vFechaInicio, "yyyymmdd") & "," & Format(vFechaCorte, "yyyymmdd") _
               & ",0,CRC"
 End Select
 
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
  
  

'------------------------------------------------------------------------------------------------------------
'           Formatos Nuevos:
'------------------------------------------------------------------------------------------------------------


lblStatus.Caption = "Formato: CEN-CINAI Nuevo!"
DoEvents

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " -NUEVO-  [" & Format(vFecha, "ddmmyyyy") & "-F20].csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


strSQL = "exec spPrm_Formato_CECINAI_New " & GLOBALES.gInstitucion & "," & vFechaProceso
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1

Do While Not rs.EOF
 
 vCadena = rs!Cadena
 
 If Len(RTrim(vCadena)) > 0 Then
    Print #fnFile, vCadena
 End If
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
  
  
'       Fin de Formatos Nuevos
'------------------------------------------------------------------------------------------------------------
  
  
Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
Call sbReporteGeneracionF02(vFechaProceso, vTempo)
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbGeneraArchivoF21_UNATEPROT(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, iRespuesta As Integer, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String

'********************************************************************************
'* Formato Unión Nacional Técnicos y Profesionales en Tránsito                  *
'********************************************************************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  vMovimiento = "in('"
  If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
  If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
  If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
  If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
  vMovimiento = vMovimiento & "P')"
rs.Close



'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)
On Error GoTo vError


vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F21].csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents


strSQL = "select P.cedula,P.Tipo,P.Tipo,P.cod_deduccion,P.Movimiento,P.Monto_Actual,Dept.Descripcion as 'Departamento', S.Nombre" _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " left join AFDepartamentos Dept on S.cod_Institucion = Dept.Cod_Institucion and S.cod_Departamento = Dept.Cod_Departamento" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.cod_deduccion, P.movimiento"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1

'Cedula, Nombre, Monto, Departamento (Institucion para Ellos)

Do While Not rs.EOF
        
 Select Case rs!Tipo
   Case "A" 'Ahorros
        vCadena = rs!Cedula & "," & rs!Nombre & "," & rs!Monto_Actual & "," & rs!Departamento
   
   Case "E" 'Extraordinarios
        vCadena = rs!Cedula & "," & rs!Nombre & "," & rs!Monto_Actual & "," & rs!Departamento
   
   Case "C" 'Creditos
        vCadena = rs!Cedula & "," & rs!Nombre & "," & rs!Monto_Actual & "," & rs!Departamento
 End Select
 
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
Call sbReporteGeneracionF02(vFechaProceso, vTempo)
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbGeneraArchivoF22_PANI(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String, vMonto As Long
Dim vFile As String, vArchivo As String, vFecha As Date
Dim vCodInstitucion As String

'*********************************
'* Formato Planilla del PANI     *
'*********************************

On Error GoTo vError


fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass


'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F22].txt"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

strSQL = "select P.*,S.nombre " _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento <> 'E'" _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.movimiento"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Value = 1
prgProcesoMensual.Max = rs.RecordCount + 1


Do While Not rs.EOF
 
 'Campo 01: Cedula de 10 char, 2-4-4
 'Campo 02: Nombre de 30
 'Campo 03: YYYYMM
 'Campo 04: Monto (8)
 vMonto = CLng((rs!Monto_Actual * 100))
 
 vCadena = SIFGlobal.fxStringRelleno(rs!Cedula, "I", "0", 10)
 vCadena = vCadena & SIFGlobal.fxStringRelleno(rs!Nombre, "D", " ", 30) & vFechaProceso
 vCadena = vCadena & SIFGlobal.fxStringRelleno(CStr(vMonto), "I", "0", 8)
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub


Private Sub sbGeneraArchivoF23_CORREOS(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date
Dim vCodInstitucion As String

vCodInstitucion = ""

'**********************************
'* Formato Planilla del CORREOS CR*
'**********************************

On Error GoTo vError

'1. Se envia a deducir todos los datos, en cada planilla
'2. Los separadores entre campos ","


fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass


'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F23].csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

strSQL = "select P.*,S.nombre " _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento <> 'E'" _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.movimiento"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Value = 1
prgProcesoMensual.Max = rs.RecordCount + 1


Do While Not rs.EOF
 
 'Campo 01: Cedula de 10 char, 2-4-4
 'Campo 02: Nombre
 'Campo 03: Tipo de Deduccion (C, Credito. A, Aporte)
 'Campo 04: Monto

 vCadena = Trim(rs!Cedula) & "," & Replace(rs!Nombre, ",", " ") & "," & rs!Monto_Actual
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub

Private Sub sbGeneraArchivoF24_SERVICOOP(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date


'**********************************
'* Formato Planilla SERVICOOP     *
'**********************************

On Error GoTo vError

'1. Se envia a deducir todos los datos, en cada planilla
'2. Los separadores entre campos ","


fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass


Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte,compara_indicador" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  If rs!Compara_Indicador = 1 Then
        vMovimiento = "in('"
        If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
        If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
        If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
        If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
        vMovimiento = vMovimiento & "P')"
  Else
        vMovimiento = "in('I','E','M','C','P')"
  End If
rs.Close

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F24].csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

strSQL = "select P.*,S.nombre, isnull(D.descripcion,'') as 'Departamento' " _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " left join AFDepartamentos D on S.cod_institucion = D.cod_institucion and S.cod_departamento = D.cod_departamento" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.tipo,P.movimiento,P.cedula"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Value = 1
prgProcesoMensual.Max = rs.RecordCount + 1


'- Formato ServiCoop (Cedula, nombre, monto, institucion, movimiento) : CSV
'  a) No reporta las que se mantienen

Do While Not rs.EOF
 
 vCadena = Trim(rs!Cedula) & "," & Replace(rs!Nombre, ",", " ") & "," & rs!Monto_Actual & "," & Replace(rs!Departamento, ",", " ") & "," & rs!Movimiento
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub



Private Sub sbGeneraArchivoF25_Holcim(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date


'**********************************
'* Formato Planilla HOLCIM *
'**********************************

On Error GoTo vError

'1. Se envia a deducir todos los datos, en cada planilla
'2. Los separadores entre campos ","


fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass


Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte,compara_indicador" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  If rs!Compara_Indicador = 1 Then
        vMovimiento = "in('"
        If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
        If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
        If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
        If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
        vMovimiento = vMovimiento & "P')"
  Else
        vMovimiento = "in('I','E','M','C','P')"
  End If
rs.Close


'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F25].csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

strSQL = "select P.*, S.CedulaR as 'Cedula_Colilla'" _
       & ", dbo.fxSIFCorteAFechaInicio(P.Proceso) as 'Inicio'" _
       & ", dbo.fxSIFCorteAFecha(P.proceso) as 'Corte'" _
       & ", S.Nombre" _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.tipo,P.movimiento,P.cedula"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Value = 1
prgProcesoMensual.Max = rs.RecordCount + 1


'- Formato ServiCoop (Cedula, nombre, monto, institucion, movimiento) : CSV
'  a) No reporta las que se mantienen

Do While Not rs.EOF
 
 If rs!Tipo_Deduc = "M" Then
    vCadena = Trim(rs!Cedula_Colilla) & ";" & vCodInstitucion & ";" & Format(rs!INICIO, "dd.mm.yyyy") & ";" & Format(rs!Corte, "dd.mm.yyyy") _
            & ";" & rs!cod_deduccion & ";" & rs!Monto_Actual & ";;;" & rs!Cedula & ";" & rs!Nombre
 Else
    vCadena = Trim(rs!Cedula_Colilla) & ";" & vCodInstitucion & ";" & Format(rs!INICIO, "dd.mm.yyyy") & ";31.12.9999" _
            & ";" & rs!cod_deduccion & ";;" & rs!Porc_Deduc & ";;" & rs!Cedula & ";" & rs!Nombre
 End If
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub






Private Sub sbGeneraArchivoF26_JUPEMA(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date


'**********************************
'* Formato Planilla JUPEMA        *
'**********************************

On Error GoTo vError

'1. Se envia a deducir todos los datos, en cada planilla
'2. Los separadores entre campos ","


fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass


Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte,compara_indicador" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  If rs!Compara_Indicador = 1 Then
        vMovimiento = "in('"
        If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
        If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
        If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
        If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
        vMovimiento = vMovimiento & "P')"
  Else
        vMovimiento = "in('I','E','M','C','P')"
  End If
rs.Close


'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F26].txt"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

strSQL = "select P.*, S.CedulaR as 'Cedula_Colilla'" _
       & ", S.Nombre" _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.tipo,P.movimiento,P.cedula"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Value = 1
prgProcesoMensual.Max = rs.RecordCount + 1


'- Formato ServiCoop (Cedula, nombre, monto, institucion, movimiento) : CSV
'  a) No reporta las que se mantienen

Do While Not rs.EOF
 
 If rs!Tipo_Deduc = "M" Then
     vCadena = Trim(rs!Cedula) & "," & rs!Nombre & "," & vFechaProceso & "," & rs!cod_deduccion & ",F," & rs!Monto_Actual
 Else
     vCadena = Trim(rs!Cedula) & "," & rs!Nombre & "," & vFechaProceso & "," & rs!cod_deduccion & ",F," & rs!Porc_Deduc
 End If
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbGeneraArchivoF27_RECOPE(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date


'**********************************
'* Formato Planilla RECOPE        *
'**********************************

On Error GoTo vError

'1. Se envia a deducir todos los datos, en cada planilla
'2. Los separadores entre campos TAB


fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass


Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte,compara_indicador" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  If rs!Compara_Indicador = 1 Then
        vMovimiento = "in('"
        If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
        If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
        If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
        If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
        vMovimiento = vMovimiento & "P')"
  Else
        vMovimiento = "in('I','E','M','C','P')"
  End If
rs.Close


'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F27].txt"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

strSQL = "select P.*, S.CedulaR as 'Cedula_Colilla'" _
       & ", S.Nombre" _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.tipo,P.movimiento,P.cedula"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Value = 1
prgProcesoMensual.Max = rs.RecordCount + 1


vFecha = CDate(Format(vFechaProceso, "####/##") & "/14")

Do While Not rs.EOF
 
 If rs!Tipo_Deduc = "M" Then
     vCadena = Format(Trim(rs!Cedula), "0000000000") & vbTab & Format(vFecha, "dd.mm.yyyy") & vbTab & rs!cod_deduccion & vbTab & rs!Monto_Actual
 Else
     vCadena = Format(Trim(rs!Cedula), "0000000000") & vbTab & Format(vFecha, "dd.mm.yyyy") & vbTab & rs!cod_deduccion & vbTab & vbTab & rs!Monto_Actual
 End If
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub sbGeneraArchivoF28_TekExperts(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date


'**************************************
'* Formato Planilla TekExperts ASETEK *
'**************************************

On Error GoTo vError

'1. Se envia a deducir todos los datos, en cada planilla
'2. Los separadores entre campos ","


fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass


Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte,compara_indicador" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  If rs!Compara_Indicador = 1 Then
        vMovimiento = "in('"
        If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
        If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
        If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
        If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
        vMovimiento = vMovimiento & "P')"
  Else
        vMovimiento = "in('I','E','M','C','P')"
  End If
rs.Close


'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F25].csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

'strSQL = "select P.*, S.CedulaR as 'Cedula_Colilla'" _
'       & ", dbo.fxSIFCorteAFechaInicio(P.Proceso) as 'Inicio'" _
'       & ", dbo.fxSIFCorteAFecha(P.proceso) as 'Corte'" _
'       & ", S.Nombre" _
'       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
'       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
'       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
'       & " order by P.tipo,P.movimiento,P.cedula"

strSQL = "exec spPrm_File_028_ASETEK " & GLOBALES.gInstitucion & "," & vFechaProceso
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Value = 1
prgProcesoMensual.Max = rs.RecordCount + 1

'Titulos
vCadena = "CODIGO;TEAM;COLABORADOR;ENTRY_DATE;LOCATION;TERMINATION_DATE;02-D31;02-D32;02-D33;02-D36;02-D37;02-D35;02-D34;02-D38;02-D30"
Print #fnFile, vCadena

Do While Not rs.EOF
 vCadena = rs!codigo & ";" & rs!Team & ";" & rs!Colaborador & ";" & rs!Entry_Date & ";" & rs!Location & ";" & rs!Termination_Date _
         & ";" & rs![02-D31] & ";" & rs![02-D32] & ";" & rs![02-D33] & ";" & rs![02-D36] & ";" & rs![02-D37] & ";" & rs![02-D35] _
         & ";" & rs![02-D34] & ";" & rs![02-D38] & ";" & rs![02-D30]
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





Private Sub sbGeneraArchivoF29_PyG(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date


'***************************************
'* Formato Planilla Procter and Gamble *
'***************************************

On Error GoTo vError

'1. Se envia a deducir todos los datos, en cada planilla
'2. Los separadores entre campos ","


fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass


Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte,compara_indicador" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  If rs!Compara_Indicador = 1 Then
        vMovimiento = "in('"
        If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
        If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
        If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
        If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
        vMovimiento = vMovimiento & "P')"
  Else
        vMovimiento = "in('I','E','M','C','P')"
  End If
rs.Close


'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F29].csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

strSQL = "select P.*, S.CedulaR as 'Cedula_Colilla', S.nombre" _
       & ", dbo.fxSIFCorteAFechaInicio(P.Proceso) as 'Inicio'" _
       & ", dbo.fxSIFCorteAFecha(P.proceso) as 'Corte'" _
       & ", S.Nombre" _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.tipo,P.movimiento,P.cedula"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Value = 1
prgProcesoMensual.Max = rs.RecordCount + 1

Do While Not rs.EOF
 
 If rs!Tipo_Deduc = "M" Then
    vCadena = "F2;'01;'" & vCodInstitucion & ";" & Trim(rs!Cedula_Colilla) & ";;" _
            & Format(rs!Corte, "yyyymmdd") & ";" _
            & Format(rs!Corte, "yyyymmdd") & ";" & Trim(rs!cod_deduccion) & ";" _
            & rs!Monto_Actual & ";CRC;;" & rs!Cedula & ";" & rs!Nombre
'
'    vCadena = Trim(rs!Cedula_Colilla) & ";" & vCodInstitucion & ";" & Format(rs!Inicio, "dd.mm.yyyy") & ";" & Format(rs!Corte, "dd.mm.yyyy") _
'            & ";" & rs!cod_deduccion & ";" & rs!Monto_Actual & ";;;" & rs!Cedula & ";" & rs!Nombre
' Else
'    vCadena = Trim(rs!Cedula_Colilla) & ";" & vCodInstitucion & ";" & Format(rs!Inicio, "dd.mm.yyyy") & ";31.12.9999" _
'            & ";" & rs!cod_deduccion & ";;" & rs!Porc_Deduc & ";;" & rs!Cedula & ";" & rs!Nombre
     Print #fnFile, vCadena
 End If
 
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub




Private Sub sbGeneraArchivoF00_Excel(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date
Dim vCodInstitucion As String

vCodInstitucion = ""

'*********************************
'* Formato Microsoft Excel      *
'*********************************

On Error GoTo vError

'1. Se envia a deducir todos los datos, en cada planilla
'2. Los separadores entre campos ","


fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass


Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte,compara_indicador" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  If rs!Compara_Indicador = 1 Then
        vMovimiento = "in('"
        If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
        If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
        If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
        If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
        vMovimiento = vMovimiento & "P')"
  Else
        vMovimiento = "in('I','E','M','C','P')"
  End If
rs.Close



'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F00].csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

strSQL = "select P.*,S.nombre, I.Descripcion as 'InstDesc', isnull(S.CedulaR,S.cedula) as 'Id_Alterno'" _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " inner join instituciones I on S.cod_institucion = I.cod_institucion" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.movimiento"
      
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Value = 1
prgProcesoMensual.Max = rs.RecordCount + 1


Do While Not rs.EOF
 
 'Campo 01: Cedula
 'Campo 02: Nombre
 'Campo 03: Tipo de Deduccion (C, Credito. A, Aporte)
 'Campo 04: Monto
 'Campo 05: Movimiento (I,E,C,M)

 vCadena = Trim(rs!Cedula) & ";" & Replace(rs!Nombre, ";", " ") & ";" & rs!Tipo & ";" & rs!Monto_Actual _
         & ";" & rs!Movimiento & ";" & rs!InstDesc & ";" & rs!Id_Alterno
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub


Private Sub sbGeneraArchivoF31_Excel_ForzaCash(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date
Dim vCodInstitucion As String

vCodInstitucion = ""

'****************************************************
'* Formato Microsoft Excel version para ASECASH     *
'****************************************************

On Error GoTo vError

'1. Se envia a deducir todos los datos, en cada planilla
'2. Los separadores entre campos ","


fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass


Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte,compara_indicador" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  If rs!Compara_Indicador = 1 Then
        vMovimiento = "in('"
        If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
        If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
        If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
        If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
        vMovimiento = vMovimiento & "P')"
  Else
        vMovimiento = "in('I','E','M','C','P')"
  End If
rs.Close



'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F31].csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

strSQL = "select P.*,S.nombre, I.Descripcion as 'InstDesc', isnull(S.CedulaR,S.cedula) as 'Id_Alterno' " _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " inner join instituciones I on S.cod_institucion = I.cod_institucion" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.movimiento"
      
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Value = 1
prgProcesoMensual.Max = rs.RecordCount + 1


Do While Not rs.EOF
 

 vCadena = Trim(rs!Cedula) & ";" & Replace(rs!Nombre, ";", " ") _
        & ";" & rs!Tipo & ";" & rs!Monto_Actual & ";" & rs!Movimiento & ";" & rs!InstDesc _
        & ";" & Trim(rs!Id_Alterno)
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max _
        & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub




Private Sub sbGeneraArchivoF32_DxC(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date
Dim vExcelFile As String

'***********************************
'* Formato Planilla DxC Technology *
'***********************************

On Error GoTo vError



fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass


Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte,compara_indicador" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  If rs!Compara_Indicador = 1 Then
        vMovimiento = "in('"
        If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
        If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
        If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
        If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
        vMovimiento = vMovimiento & "P')"
  Else
        vMovimiento = "in('I','E','M','C','P')"
  End If
rs.Close


'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F32].csv"

vExcelFile = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F32]"

vTempo = vRuta & "\" & vArchivo



vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

strSQL = "select P.*, S.CedulaR as 'Cedula_Colilla'" _
       & ", dbo.fxSIFCorteAFechaInicio(P.Proceso) as 'Inicio'" _
       & ", dbo.fxSIFCorteAFecha(P.proceso) as 'Corte'" _
       & ", S.Nombre" _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.tipo,P.movimiento,P.cedula"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Value = 1
prgProcesoMensual.Max = rs.RecordCount + 1

'Linea de titulos
vCadena = "empleado;concepto;cantidad;monto"

Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 4
    vHeaders.Headers(1) = "empleado"
    vHeaders.Headers(2) = "concepto"
    vHeaders.Headers(3) = "cantidad"
    vHeaders.Headers(4) = "Monto"

Print #fnFile, vCadena


With vGrid
    .MaxRows = 0
  

    Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     .Col = 1
     .Text = Trim(rs!Cedula_Colilla)
     .Col = 2
     .Text = Trim(rs!cod_deduccion)
     
     If rs!cod_deduccion = "DE31" Then
        vCadena = Trim(rs!Cedula_Colilla) & ";" & rs!cod_deduccion & ";" & rs!Monto_Actual & ";0"
     
        .Col = 3
        .Text = CStr(rs!Monto_Actual)
        .Col = 4
        .Text = "0"
     
     Else
        vCadena = Trim(rs!Cedula_Colilla) & ";" & rs!cod_deduccion & ";0;" & rs!Monto_Actual
        .Col = 3
        .Text = "0"
        .Col = 4
        .Text = CStr(rs!Monto_Actual)
     End If
      
     Print #fnFile, vCadena
     
             
     
     If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
     lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
     rs.MoveNext
    Loop
    rs.Close

End With

Close #fnFile
  
'Exportar a Excel
vTempo = vRuta & "\" & vExcelFile

Me.MousePointer = vbDefault

Call sbSIFGridExportar(vGrid, vHeaders, vTempo, "Excel", True)
  
Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub







Private Sub sbGeneraArchivoF34_ASOECorr(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date


'***********************************
'* Formato Planilla ASOECorr Correos de Costa Rica *
'***********************************

On Error GoTo vError



fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass


Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte,compara_indicador" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  If rs!Compara_Indicador = 1 Then
        vMovimiento = "in('"
        If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
        If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
        If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
        If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
        vMovimiento = vMovimiento & "P')"
  Else
        vMovimiento = "in('I','E','M','C','P')"
  End If
rs.Close


'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F34].csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

strSQL = "select P.*, S.CedulaR as 'Cedula_Colilla', S.nombre" _
       & ", dbo.fxSIFCorteAFechaInicio(P.Proceso) as 'Inicio'" _
       & ", dbo.fxSIFCorteAFecha(P.proceso) as 'Corte'" _
       & ", S.Nombre" _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.tipo,P.movimiento,P.cedula"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Value = 1
prgProcesoMensual.Max = rs.RecordCount + 1

'Linea de titulos
vCadena = "Identificacion;concepto;valor;nombre"
Print #fnFile, vCadena

Do While Not rs.EOF
 
    vCadena = Trim(rs!Cedula_Colilla) & ";" & rs!cod_deduccion & ";" & rs!Monto_Actual & ";" & rs!Nombre
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub



Private Sub sbGeneraArchivoF35_ProGrX_RRHH(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, iRespuesta As Integer, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String
Dim vFechaInicio As Date, vFechaCorte As Date

'****************************************
'*    Formato ProGrX Recursos Humanos   *
'****************************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte" _
       & ", dbo.fxSIFCorteAFechaInicio(" & vFechaProceso & ") as 'FechaInicio'" _
       & ", dbo.fxSIFCorteAFecha(" & vFechaProceso & ") as 'FechaCorte'" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  vFechaInicio = rs!FechaInicio
  vFechaCorte = rs!FechaCorte
  vMovimiento = "in('"
  If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
  If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
  If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
  If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
  vMovimiento = vMovimiento & "P')"
rs.Close



'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)
On Error GoTo vError


vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F35].txt"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


lblStatus = "Creando archivo a enviar"
DoEvents

Open vTempo For Output As #fnFile  ' Create file name.

strSQL = "exec spPrm_Formato_ProGrX_RRHH " & GLOBALES.gInstitucion & "," & vFechaProceso
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1

Do While Not rs.EOF
 
 vCadena = rs!Cadena
 
 If Len(RTrim(vCadena)) > 0 Then
    Print #fnFile, vCadena
 End If
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
  
Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
Call sbReporteGeneracionF02(vFechaProceso, vTempo)
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbGeneraArchivoF36_ASOINSVA(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, iRespuesta As Integer, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String
Dim vFechaInicio As Date, vFechaCorte As Date

'****************************************
'*    Formato INS VALORES               *
'****************************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte" _
       & ", dbo.fxSIFCorteAFechaInicio(" & vFechaProceso & ") as 'FechaInicio'" _
       & ", dbo.fxSIFCorteAFecha(" & vFechaProceso & ") as 'FechaCorte'" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  vFechaInicio = rs!FechaInicio
  vFechaCorte = rs!FechaCorte
  vMovimiento = "in('"
  If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
  If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
  If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
  If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
  vMovimiento = vMovimiento & "P')"
rs.Close



'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)
On Error GoTo vError


vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F36].txt"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


lblStatus = "Creando archivo a enviar"
DoEvents

Open vTempo For Output As #fnFile  ' Create file name.

strSQL = "exec spPrm_Formato_INSVA " & GLOBALES.gInstitucion & "," & vFechaProceso
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1

Do While Not rs.EOF
 
 vCadena = rs!Cadena
 
 If Len(RTrim(vCadena)) > 0 Then
    Print #fnFile, vCadena
 End If
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
  
Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
Call sbReporteGeneracionF02(vFechaProceso, vTempo)
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub sbGeneraArchivoF03_SIF(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date
Dim vCodInstitucion As String

vCodInstitucion = ""

'*********************************
'* Formato Planilla ASECCSS      *
'*********************************

On Error GoTo vError

'1. Se envia a deducir todos los datos, en cada planilla
'2. Los separadores entre campos es con TABS


fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass



Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte,compara_indicador" _
       & " from instituciones" _
       & " where cod_institucion = " & GLOBALES.gInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  If rs!Compara_Indicador = 1 Then
        vMovimiento = "in('"
        If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
        If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
        If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
        If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
        vMovimiento = vMovimiento & "P')"
  Else
        vMovimiento = "in('I','E','M','C','P')"
  End If
rs.Close


'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F03].txt"
         
vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents

'strSQL = "select P.*,S.nombre " _
'       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
'       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
'       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
'       & " order by P.cedula,P.tipo,P.movimiento"

strSQL = "select P.*,S.nombre, I.Descripcion as 'InstDesc' " _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " inner join instituciones I on S.cod_institucion = I.cod_institucion" _
       & " where P.Proceso = " & vFechaProceso & " and P.movimiento " & vMovimiento _
       & " and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " order by P.cedula,P.tipo,P.movimiento"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Value = 1
prgProcesoMensual.Max = rs.RecordCount + 1


Do While Not rs.EOF
 
 'Campo 01: Cedula de 10 char, 2-4-4
 'Campo 02: Nombre
 'Campo 03: Tipo de Deduccion (C, Credito. A, Aporte)
 'Campo 04: Monto
 'Campo 05: Movimiento

 vCadena = Trim(rs!Cedula) & vbTab & rs!Nombre & vbTab & rs!Tipo & vbTab
 vCadena = vCadena & rs!Monto_Actual & vbTab & rs!Movimiento & vbTab & rs!InstDesc
 
 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub


Private Sub sbGeneraArchivoF04_ICE_ACOTEL(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vCodInstitucion As String

vCodInstitucion = ""

'*********************************
'* Formato Planilla ICE - ACOTEL *
'*********************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError


vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F04].txt"


vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lblStatus = "Creando archivo a enviar"
DoEvents


'Limpia y pone como Exclusion casos infereriores a 100

strSQL = "update prm_planilla set monto_actual = 0, movimiento = 'E'" _
       & " where Proceso = " & vFechaProceso & " and monto_actual <= 100" _
       & " and Tipo = 'C' and cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)


'Las Exclusiones se representan con montos en Cero
strSQL = "select P.cedula,P.movimiento,P.monto_actual,S.nombre,isnull(sum(R.saldo),0) as Saldos " _
       & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
       & " left join reg_creditos R on P.cedula = R.cedula and R.estado = 'A'" _
       & " where P.Proceso = " & vFechaProceso _
       & " and Tipo = 'C' and P.cod_institucion = " & GLOBALES.gInstitucion _
       & " group by P.cedula,P.monto_actual,P.movimiento,S.nombre"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 2
prgProcesoMensual.Value = 1


Do While Not rs.EOF
 
 'Campo 01: Cedula de 9 char, 1-4-4
 'Campo 02: Deuda o Saldo de 11 char (Sin Punto decimal y redondeado)
 'Campo 03: Cuota 10 char (Sin punto Decimal y redondeado)
 
 vCadena = Format(Trim(rs!Cedula), "000000000")
 
 Select Case rs!Movimiento
   Case "E"
     'Exclusiones se envian en CERO
     vCadena = vCadena & Format(0, "00000000000")
     vCadena = vCadena & Format(0, "0000000000")
   Case Else
     vCadena = vCadena & Format((CLng(rs!Saldos) * 100), "00000000000")
     vCadena = vCadena & Format((CLng(rs!Monto_Actual) * 100), "0000000000")
 End Select
 

 Print #fnFile, vCadena
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

Call sbReporteGeneracionF02(vFechaProceso, vTempo)

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Public Function fxgDepuraCadena(vCadena As String) As String
Dim i As Integer, Resultado As String

vCadena = Trim(vCadena)
Resultado = ""

For i = 1 To Len(vCadena)
 If Asc(Mid(vCadena, i, 1)) > 47 And Asc(Mid(vCadena, i, 1)) < 123 Then
   Resultado = Resultado & Mid(vCadena, i, 1)
 Else
   If Asc(Mid(vCadena, i, 1)) = 32 Then Resultado = Resultado & Mid(vCadena, i, 1)
 End If
 
Next i

fxgDepuraCadena = Resultado

End Function



Private Sub sbGeneraArchivoF05_CoopeCaja_OLD(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vCodigoAportes As String, vCodigoCreditos As String
Dim vLinea As String, y As Integer
Dim vApellido1 As String, vApellido2 As String
Dim vNombre1 As String, vNombre2 As String
Dim bPasa As Byte, vMonto As Long

'********************************************
'* Formato de Planilla de COOPECAJA         *
'********************************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)


On Error GoTo vError


strSQL = "select * from instituciones where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
   vCodigoAportes = Trim(rs!codigo_aportes & "")
   vCodigoCreditos = Trim(rs!codigo_creditos & "")
rs.Close

vArchivo = Format(GLOBALES.gInstitucion, "00") & "-" & Format(vFechaProceso, "####-##") & "-ARC-COOPECAJA-A_OLD.txt"

vTempo = vRuta & "\" & vArchivo
vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If

If vCodigoAportes <> "NO" Then
    Open vTempo For Output As #fnFile  ' Create file name.
      lblStatus = "Creando archivo a enviar"
      DoEvents
    
      strSQL = "select P.*,S.nombre,S.direccion" _
             & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
             & " where P.Proceso = " & vFechaProceso & " and P.cod_institucion = " _
             & GLOBALES.gInstitucion & " and P.tipo = 'A' order by P.cedula"
      Call OpenRecordSet(rs, strSQL)
      
    
      prgProcesoMensual.Max = rs.RecordCount + 1
      prgProcesoMensual.Value = 1
    
      Do While Not rs.EOF
           vLinea = Mid(vCodigoAportes, 1, 6) & SIFGlobal.fxStringRelleno(Trim(rs!Cedula), "I", "0", 15)
        
           bPasa = 1
           vApellido1 = ""
           vApellido2 = ""
           vNombre1 = ""
           vNombre2 = ""
           
           For i = 1 To Len(rs!Nombre)
             If Mid(rs!Nombre, i, 1) = " " Then
               bPasa = bPasa + 1
             Else
                Select Case bPasa
                   Case 1
                     vApellido1 = vApellido1 & Mid(rs!Nombre, i, 1)
                   Case 2
                     vApellido2 = vApellido2 & Mid(rs!Nombre, i, 1)
                   Case 3
                     vNombre1 = vNombre1 & Mid(rs!Nombre, i, 1)
                   Case 4
                     vNombre2 = vNombre2 & Mid(rs!Nombre, i, 1)
                End Select
             End If
           Next i
           
           vLinea = vLinea & SIFGlobal.fxStringRelleno(vApellido1, "D", " ", 15) & SIFGlobal.fxStringRelleno(vApellido2, "D", " ", 15)
           vLinea = vLinea & SIFGlobal.fxStringRelleno(vNombre1, "D", " ", 15) & SIFGlobal.fxStringRelleno(vNombre2, "D", " ", 15)
           vLinea = vLinea & SIFGlobal.fxStringRelleno(fxgDepuraCadena(rs!Direccion & ""), "D", " ", 140)
           'Telefonos / Se Envian En Cero
           vLinea = vLinea & SIFGlobal.fxStringRelleno("0", "I", "0", 7) & SIFGlobal.fxStringRelleno("0", "I", "0", 7)
           
           'Monto con Dos Decimales, si Punto Decimal
           vLinea = vLinea & SIFGlobal.fxStringRelleno("0", "I", "0", 11)
           
           Print #fnFile, vLinea
         
         If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
         lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
         rs.MoveNext
      Loop
      rs.Close
    
    Close #fnFile

End If 'vCodigoAportes <> "NO"
  
'*****************************
' Crea Archivo para Creditos *
'*****************************

fnFile = FreeFile


vArchivo = Format(GLOBALES.gInstitucion, "00") & "-" & Format(vFechaProceso, "####-##") & "-ARC-COOPECAJA-C_OLD.txt"
vTempo = vRuta & "\" & vArchivo
vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then Kill vTempo

vFile = ""
vFile = Dir(vRuta, vbDirectory)
'If Not (vFile = CStr(vFechaProceso)) Then MkDir vRuta

If vCodigoCreditos <> "NO" Then
    Open vTempo For Output As #fnFile  ' Create file name.
      lblStatus = "Creando archivo a enviar"
      DoEvents
    
      strSQL = "select P.*,S.nombre,S.direccion" _
             & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
             & " where P.Proceso = " & vFechaProceso & " and P.cod_institucion = " _
             & GLOBALES.gInstitucion & " and P.tipo = 'C' order by P.cedula"
      Call OpenRecordSet(rs, strSQL)
    
      prgProcesoMensual.Max = rs.RecordCount + 1
      prgProcesoMensual.Value = 1
    
      Do While Not rs.EOF
           vLinea = Mid(vCodigoCreditos, 1, 6) & SIFGlobal.fxStringRelleno(Trim(rs!Cedula), "I", "0", 15)
        
           bPasa = 1
           vApellido1 = ""
           vApellido2 = ""
           vNombre1 = ""
           vNombre2 = ""
           
           For i = 1 To Len(rs!Nombre)
             If Mid(rs!Nombre, i, 1) = " " Then
               bPasa = bPasa + 1
             Else
                Select Case bPasa
                   Case 1
                     vApellido1 = vApellido1 & Mid(rs!Nombre, i, 1)
                   Case 2
                     vApellido2 = vApellido2 & Mid(rs!Nombre, i, 1)
                   Case 3
                     vNombre1 = vNombre1 & Mid(rs!Nombre, i, 1)
                   Case 4
                     vNombre2 = vNombre2 & Mid(rs!Nombre, i, 1)
                End Select
             End If
           Next i
        
           vLinea = vLinea & SIFGlobal.fxStringRelleno(vApellido1, "D", " ", 15) & SIFGlobal.fxStringRelleno(vApellido2, "D", " ", 15)
           vLinea = vLinea & SIFGlobal.fxStringRelleno(vNombre1, "D", " ", 15) & SIFGlobal.fxStringRelleno(vNombre2, "D", " ", 15)
           vLinea = vLinea & SIFGlobal.fxStringRelleno(fxgDepuraCadena(rs!Direccion & ""), "D", " ", 140)
           'Telefonos / Se Envian En Cero
           vLinea = vLinea & SIFGlobal.fxStringRelleno("0", "I", "0", 7) & SIFGlobal.fxStringRelleno("0", "I", "0", 7)
           
           
           vMonto = CLng((rs!Monto_Actual * 100))
           'Monto con Dos Decimales, si Punto Decimal
           vLinea = vLinea & SIFGlobal.fxStringRelleno(CStr(vMonto), "I", "0", 11)
        
           'Nuevo 2005/11/14
           Select Case rs!Movimiento
              Case "E" 'Exclusion
                 i = 1
              Case "I" 'Inclusion
                 i = 2
              Case "C" 'Cambio
                 i = 3
              Case Else
                 i = 4 ' Invalido, Porque se Mantiene o No se Proceso
             End Select
               
'           If i <> 4 Then Print #fnFile, vLinea
           If i <> 1 Then Print #fnFile, vLinea

         If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
         lblStatus.Caption = "Creando Archivo 2 Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
         rs.MoveNext
      Loop
      rs.Close
    
    Close #fnFile
    
End If 'vCodigoCreditos <> "NO"

Me.MousePointer = vbDefault

'MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
'
'Call sbReporteGeneracionF02(vFechaProceso, vTempo)
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Resume
  
End Sub

Private Sub sbGeneraArchivoF05_CoopeCaja(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer, vDireccion As String, vTelefono As String
Dim fnFile
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vCodigoAportes As String, vCodigoCreditos As String
Dim vLinea As String, y As Integer
Dim vApellido1 As String, vApellido2 As String
Dim vNombre1 As String, vNombre2 As String
Dim bPasa As Byte, vMonto As Long

Dim vCodInstitucion As String

vCodInstitucion = ""

'****************************************************************************************************************
'Formato de Planilla de COOPECAJA                                                                             *
'Update:2017/04/17: Direccion y Telefono (Empresariales) + Codigo Departamento (Es la Insitutcion en CoopeCaja) *
'****************************************************************************************************************
On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError


strSQL = "select * from instituciones where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
   vCodigoAportes = Trim(rs!codigo_aportes & "")
   vCodigoCreditos = Trim(rs!codigo_creditos & "")
rs.Close

strSQL = "select rtrim(PAG_NOMLARGO) + ', ' + rtrim(PAG_DOMICILIO)  as 'DIRECCION', replace(TELEFONOEMP,'-','') as 'Telefono'" _
       & " From SIF_EMPRESA"
Call OpenRecordSet(rs, strSQL)
    vDireccion = UCase(rs!Direccion)
    vTelefono = rs!Telefono
rs.Close




vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F05] - COOPECAJA-A.txt"


vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


If vCodigoAportes <> "NO" Then
    Open vTempo For Output As #fnFile  ' Create file name.
      lblStatus = "Creando archivo a enviar"
      DoEvents
    
      strSQL = "select P.*,S.nombre,S.cod_departamento" _
             & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
             & " where P.Proceso = " & vFechaProceso & " and P.cod_institucion = " _
             & GLOBALES.gInstitucion & " and P.tipo = 'A' order by P.cedula"
      Call OpenRecordSet(rs, strSQL)
    
      prgProcesoMensual.Max = rs.RecordCount + 1
      prgProcesoMensual.Value = 1
    
      Do While Not rs.EOF
           vLinea = Mid(vCodigoAportes, 1, 6) & SIFGlobal.fxStringRelleno(Trim(rs!Cedula), "I", "0", 15)
        
           bPasa = 1
           vApellido1 = ""
           vApellido2 = ""
           vNombre1 = ""
           vNombre2 = ""
           
           For i = 1 To Len(rs!Nombre)
             If Mid(rs!Nombre, i, 1) = " " Then
               bPasa = bPasa + 1
             Else
                Select Case bPasa
                   Case 1
                     vApellido1 = vApellido1 & Mid(rs!Nombre, i, 1)
                   Case 2
                     vApellido2 = vApellido2 & Mid(rs!Nombre, i, 1)
                   Case 3
                     vNombre1 = vNombre1 & Mid(rs!Nombre, i, 1)
                   Case 4
                     vNombre2 = vNombre2 & Mid(rs!Nombre, i, 1)
                End Select
             End If
           Next i
           
           vLinea = vLinea & SIFGlobal.fxStringRelleno(vApellido1, "D", " ", 15) & SIFGlobal.fxStringRelleno(vApellido2, "D", " ", 15)
           vLinea = vLinea & SIFGlobal.fxStringRelleno(vNombre1, "D", " ", 15) & SIFGlobal.fxStringRelleno(vNombre2, "D", " ", 15)
           vLinea = vLinea & SIFGlobal.fxStringRelleno(fxgDepuraCadena(vDireccion & ""), "D", " ", 140)
           
           'Telefonos / Se Envian En Cero
           vLinea = vLinea & SIFGlobal.fxStringRelleno(vTelefono, "I", "0", 8) & SIFGlobal.fxStringRelleno("0", "I", "0", 8)
           
           'Monto con Dos Decimales, si Punto Decimal
           vLinea = vLinea & SIFGlobal.fxStringRelleno("0", "I", "0", 11)
           
           Print #fnFile, vLinea
         
         If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
         lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
         rs.MoveNext
      Loop
      rs.Close
    
    Close #fnFile

End If 'vCodigoAportes <> "NO"
  
'*****************************
' Crea Archivo para Creditos *
'*****************************

fnFile = FreeFile

vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F05] - COOPECAJA-C.txt"

vTempo = vRuta & "\" & vArchivo
vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then Kill vTempo

'vFile = ""
'vFile = Dir(vRuta, vbDirectory)
'If Not (vFile = CStr(vFechaProceso)) Then MkDir vRuta

If vCodigoCreditos <> "NO" Then
    Open vTempo For Output As #fnFile  ' Create file name.
      lblStatus = "Creando archivo a enviar"
      DoEvents
    
      strSQL = "select P.*,S.nombre,S.direccion,S.cod_Departamento" _
             & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
             & " where P.Proceso = " & vFechaProceso & " and P.cod_institucion = " _
             & GLOBALES.gInstitucion & " and P.tipo = 'C' order by P.cedula"
      Call OpenRecordSet(rs, strSQL)
    
      prgProcesoMensual.Max = rs.RecordCount + 1
      prgProcesoMensual.Value = 1
    
      Do While Not rs.EOF
           vLinea = Mid(vCodigoCreditos, 1, 6) & SIFGlobal.fxStringRelleno(Trim(rs!cod_Departamento), "I", "0", 3) & SIFGlobal.fxStringRelleno(Trim(rs!Cedula), "I", "0", 15)
        
           bPasa = 1
           vApellido1 = ""
           vApellido2 = ""
           vNombre1 = ""
           vNombre2 = ""
           
           For i = 1 To Len(rs!Nombre)
             If Mid(rs!Nombre, i, 1) = " " Then
               bPasa = bPasa + 1
             Else
                Select Case bPasa
                   Case 1
                     vApellido1 = vApellido1 & Mid(rs!Nombre, i, 1)
                   Case 2
                     vApellido2 = vApellido2 & Mid(rs!Nombre, i, 1)
                   Case 3
                     vNombre1 = vNombre1 & Mid(rs!Nombre, i, 1)
                   Case 4
                     vNombre2 = vNombre2 & Mid(rs!Nombre, i, 1)
                End Select
             End If
           Next i
        
           vLinea = vLinea & SIFGlobal.fxStringRelleno(vApellido1, "D", " ", 15) & SIFGlobal.fxStringRelleno(vApellido2, "D", " ", 15)
           vLinea = vLinea & SIFGlobal.fxStringRelleno(vNombre1, "D", " ", 15) & SIFGlobal.fxStringRelleno(vNombre2, "D", " ", 15)
           vLinea = vLinea & SIFGlobal.fxStringRelleno(fxgDepuraCadena(vDireccion & ""), "D", " ", 140)
           'Telefonos / Se Envian En Cero
           vLinea = vLinea & SIFGlobal.fxStringRelleno(vTelefono, "I", "0", 8) & SIFGlobal.fxStringRelleno("0", "I", "0", 8)
           
           
           vMonto = CLng((rs!Monto_Actual * 100))
           'Monto con Dos Decimales, si Punto Decimal
           vLinea = vLinea & SIFGlobal.fxStringRelleno(CStr(vMonto), "I", "0", 11)
        
           'Nuevo 2005/11/14
           Select Case rs!Movimiento
              Case "E" 'Exclusion
                 i = 1
              Case "I" 'Inclusion
                 i = 2
              Case "C" 'Cambio
                 i = 3
              Case Else
                 i = 4 ' Invalido, Porque se Mantiene o No se Proceso
             End Select
               
           'If i <> 4 Then Print #fnFile, vLinea
           
           '2016.10.07: Cambio de Formato. Ahora se reporta todo menos exclusiones
           If i <> 1 Then Print #fnFile, vLinea
         
         If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
         lblStatus.Caption = "Creando Archivo 2 Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
         rs.MoveNext
      Loop
      rs.Close
    
    Close #fnFile
    
End If 'vCodigoCreditos <> "NO"

Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
Call sbReporteGeneracionF02(vFechaProceso, vTempo)
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbGeneraArchivoF06_ICE_Central(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile
Dim vFile As String, vArchivo As String, vFecha As Date

Dim rsX As New ADODB.Recordset

Dim vCodigoAportes As String, vCodigoCreditos As String
Dim vLinea As String, y As Integer
Dim vApellido1 As String, vApellido2 As String
Dim vNombre1 As String, vNombre2 As String
Dim bPasa As Byte, vMonto As Long

Dim vCodInstitucion As String

vCodInstitucion = ""

'********************************************
'* Formato de Planilla de ICE Of.Centrales  *
'********************************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError


strSQL = "select * from instituciones where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
   vCodigoAportes = Trim(rs!codigo_aportes & "")
   vCodigoCreditos = Trim(rs!codigo_creditos & "")
rs.Close

 
'*****************************
' Crea Archivo para Creditos *
'*****************************

fnFile = FreeFile


vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F06].txt"


vTempo = vRuta & "\" & vArchivo
vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


 'Campo 01: Cedula de 9 char, 2-4-4
 'Campo 02: Monto de 11 char  -> Operaciones de la persona
 'Campo 05: Cuota de 10 char
 'Los montos son con decimales (2), sin puntos y sin comas.

If vCodigoCreditos <> "NO" Then
    Open vTempo For Output As #fnFile  ' Create file name.
      lblStatus = "Creando archivo a enviar"
      DoEvents
    
      strSQL = "select P.*,S.nombre,S.direccion" _
             & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
             & " where P.Proceso = " & vFechaProceso & " and P.cod_institucion = " _
             & GLOBALES.gInstitucion & " and P.tipo = 'C' order by P.cedula"
      Call OpenRecordSet(rs, strSQL)
    
      prgProcesoMensual.Max = rs.RecordCount + 1
      prgProcesoMensual.Value = 1
    
      Do While Not rs.EOF
           vLinea = SIFGlobal.fxStringRelleno(Trim(rs!Cedula), "I", "0", 9) & ""
 
            Select Case rs!Movimiento
              Case "E" 'Exclusion
                 i = 1
              Case "I" 'Inclusion
                 i = 2
              Case "C" 'Cambio
                 i = 3
              Case Else
                 i = 4 ' Invalido, Porque se Mantiene o No se Proceso
             End Select
            
               
        
          
          If i = 1 Then 'Exclusiones
            vLinea = SIFGlobal.fxStringRelleno(Trim(rs!Cedula), "I", "0", 9)
            vLinea = SIFGlobal.fxStringRelleno(Trim(vLinea), "D", "0", 30)
          Else
           'Monto con Dos Decimales, si Punto Decimal
           
            strSQL = "select isnull(sum(montoapr),0) as Monto, isnull(sum(Saldo),0) as Saldo" _
                   & " from reg_creditos where prideduc <= " & vFechaProceso & " and estado = 'A'" _
                   & " and cedula = '" & rs!Cedula & "'"
            rsX.Open strSQL, glogon.Conection, adOpenStatic
            If Not rsX.EOF And Not rsX.BOF Then
                vMonto = CLng((rsX!Monto * 100))
            Else
                vMonto = 0
            End If
            rsX.Close
            'Monto Total
            vLinea = vLinea & SIFGlobal.fxStringRelleno(CStr(vMonto), "I", "0", 11)
            
            'Mensualidad
            vMonto = CLng((rs!Monto_Actual * 100))
            vLinea = vLinea & SIFGlobal.fxStringRelleno(CStr(vMonto), "I", "0", 10)
          
          End If
          
          
          If i <> 4 Then Print #fnFile, vLinea
         
         If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
         lblStatus.Caption = "Creando Archivo 2 Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
         rs.MoveNext
      Loop
      rs.Close
    
    Close #fnFile
    
End If 'vCodigoCreditos <> "NO"

Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
Call sbReporteGeneracionF02(vFechaProceso, vTempo)
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub sbGeneraArchivoF07_ICE_Proyectos(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, iRespuesta As Integer, strCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date

Dim rsX As New ADODB.Recordset, x As Integer

Dim vCodigoAportes As String, vCodigoCreditos As String
Dim vLinea As String, y As Integer
Dim vApellido1 As String, vApellido2 As String
Dim vNombre1 As String, vNombre2 As String
Dim bPasa As Byte, vMonto As Long
Dim vCodInstitucion As String

vCodInstitucion = ""

'*********************************************
'* Formato de Planilla de ICE para Proyectos *
'*********************************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError


strSQL = "select * from instituciones where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
   vCodigoAportes = Trim(rs!codigo_aportes & "")
   vCodigoCreditos = Trim(rs!codigo_creditos & "")
rs.Close

 
'*****************************
' Crea Archivo para Creditos *
'*****************************

fnFile = FreeFile


vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F07].txt"


vTempo = vRuta & "\" & vArchivo
vFile = Dir(vTempo, vbArchive)
If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


 'Campo 01: Apellido 1
 'Campo 02: Apellido 2
 'Campo 03: Nombre
 'Campo 04: Monto 'Sin comas
 'Campo 05: Plazo
 'Campo 06: Mensualidad
 
 'Separador con Tabs y solo inclusiones

If vCodigoCreditos <> "NO" Then
    Open vTempo For Output As #fnFile  ' Create file name.
      lblStatus = "Creando archivo a enviar"
      DoEvents
    
      strSQL = "select P.*,S.nombre,S.direccion" _
             & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
             & " where P.Proceso = " & vFechaProceso & " and P.cod_institucion = " _
             & GLOBALES.gInstitucion & " and P.tipo = 'C' order by P.cedula"
      Call OpenRecordSet(rs, strSQL)
    
      prgProcesoMensual.Max = rs.RecordCount + 1
      prgProcesoMensual.Value = 1
    
      Do While Not rs.EOF
 
            Select Case rs!Movimiento
              Case "E" 'Exclusion
                 x = 1
              Case "I" 'Inclusion
                 x = 2
              Case "C" 'Cambio
                 x = 3
              Case Else
                 x = 4 ' Invalido, Porque se Mantiene o No se Proceso
             End Select
            
           bPasa = 1
           vApellido1 = ""
           vApellido2 = ""
           vNombre1 = ""
           vNombre2 = ""
           
           For i = 1 To Len(rs!Nombre)
             If Mid(rs!Nombre, i, 1) = " " Then
               bPasa = bPasa + 1
             Else
                Select Case bPasa
                   Case 1
                     vApellido1 = vApellido1 & Mid(rs!Nombre, i, 1)
                   Case 2
                     vApellido2 = vApellido2 & Mid(rs!Nombre, i, 1)
                   Case 3
                     vNombre1 = vNombre1 & Mid(rs!Nombre, i, 1)
                   Case 4
                     vNombre2 = vNombre2 & Mid(rs!Nombre, i, 1)
                End Select
             End If
           Next i
               
          vLinea = Trim(rs!Cedula) & vbTab & vApellido1 & vbTab & vApellido2 & vbTab & vNombre1 & " " & vNombre2 & vbTab
          
          'Solo se procesan Inclusiones
          If x = 2 Then
            strSQL = "select isnull(sum(montoapr),0) as Monto, isnull(sum(Saldo),0) as Saldo, isnull(avg(plazo),1) as Plazo" _
                   & " from reg_creditos where prideduc <= " & vFechaProceso & " and estado = 'A'" _
                   & " and cedula = '" & rs!Cedula & "'"
            rsX.Open strSQL, glogon.Conection, adOpenStatic
            If Not rsX.EOF And Not rsX.BOF Then
               vLinea = vLinea & rsX!Monto & vbTab & rsX!Plazo & vbTab & rs!Monto_Actual
            Else
               vLinea = vLinea & rs!Monto_Actual & vbTab & 1 & vbTab & rs!Monto_Actual
            End If
            rsX.Close
            
            Print #fnFile, vLinea
          
          End If
          
         If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
         lblStatus.Caption = "Creando Archivo 2 Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
         rs.MoveNext
      Loop
      rs.Close
    
    Close #fnFile
    
End If 'vCodigoCreditos <> "NO"

Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
Call sbReporteGeneracionF02(vFechaProceso, vTempo)
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub



Private Sub imgModificaCtas_Click()
  Call sbFormsCall("frmCC_PlanillaCtaCorreccion", vbModal, , , False, Me)
End Sub

Private Sub imgRepAhAplica_Click()
Dim vFecha As Currency

On Error GoTo vError
vFecha = CCur(InputBox("Especifique la fecha de proceso " & vbCrLf _
        & "La fecha de Proceso Actual es : " & GLOBALES.glngFechaCR, "Aplicación de Ahorros", GLOBALES.glngFechaCR))
Call sbAhAplicaAhorroRep(vFecha)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbGeneraArchivoF08_AyA(vFechaProceso As Currency)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile
Dim vFile As String, vArchivo As String, vFecha As Date

Dim rsX As New ADODB.Recordset

Dim vCodigoAportes As String, vCodigoCreditos As String
Dim vLinea As String, y As Integer
Dim vApellido1 As String, vApellido2 As String
Dim vNombre1 As String, vNombre2 As String
Dim bPasa As Byte, strMonto As String
Dim vFechaCorte As Date
Dim vCodInstitucion As String

vCodInstitucion = ""

'*********************************************
'* Formato de Planilla de A Y A              *
'*********************************************


On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
prgProcesoMensual.Min = 1

Me.MousePointer = vbHourglass

'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & txtInstitucion.Text & "\" & Mid(vFechaProceso, 1, 4)

On Error GoTo vError



strSQL = "select *,dbo.fxSIFCorteAFecha(" & vFechaProceso & ") as 'FechaCorte'" _
       & " from instituciones where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
   vCodigoAportes = Trim(rs!codigo_aportes & "")
   vCodigoCreditos = Trim(rs!codigo_creditos & "")
   vFechaCorte = rs!FechaCorte
rs.Close

 
'*****************************
' Crea Archivo para Creditos *
'*****************************

fnFile = FreeFile


vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F08].txt"


vTempo = vRuta & "\" & vArchivo
vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If

'001 Inclusion
'002 Modificacion
'003 Exclusion
'004 Mantienen no se reportan
 
 'Campo 01: Cedula char(10) relleno cero Izq. + relleno hasta 20 derecho espacios
 'Campo 02: Nombre char(30) relleno espacios derecha
 'Campo 03: Codigo de Deduccion char(2) a como se digita
 'Campo 04: Monto redondeado sin decimales char(8)
 'Campo 05: Codigo de Modificacion (1-4) char(3)
 
'01-30  : CEDULA
'31-34  : CODIGO DE DEDUCCION O INGRESO
'36-36  : TIPO DE MOVIMIENTO (I=INCLUSION/CAMBIO INGRESOS/EGRESOS FIJOS DEL EMPLEADO CON CONTROL AUTOMATICO X LIMITES O FECHA HASTA
'                             P=DIRECTO A PLANILLA EN PROCESO (NO AFECTA INGRESOS/EGRESOS FIJOS) CONTROL TOTAL MANUAL
'                             C=CAMBIOS  INGRESO/EGRESOS FIJOS
'                             E=ELIMINAR DE INGRESOS/EGRESOS FIJOS
'         SI TODOS LOS MESES SE REPORTAN TODAS LAS DEDUCCIONES ENTONCES DEBE LLEVAR P
'38-47  : MONTO DEL INGRESO O DEDUCCION CON DOS DECIMALES SIN PUNTO DECIMAL
'49-58  : LIMITE HASTA EL CUAL DEBE REBAJARSE
'60-69  : ACUMULADO CON EL CUAL DEBE INICIARSE
'71-80  : FECHA A PARTIR DE LA CUAL RIGE EL INGRESO O EGRESO
'82-90  : FECHA HASTA LA CUAL RIGE EL INGRESO O EGRESO
'93-97  : PORCENTAJE DE DEDUCCION
'99-105 : PERIODO DE PLANILLA QUE AFECTA (CUANDO EGRESO/INGRESO VA DIRECTO PLANILLA
'         FORMATO AAAAMMS= AAAA=ANO, MM=MES, S=SEMANA (S=BLANCO PARA PLANILLAS MENSUALES
'107-114  : CANTIDAD
'116-129 : Referencia/Comentario
'!      : INDICADOR DE TERMINACION DE ARCHIVO
 
 
 
If vCodigoCreditos <> "NO" Then
    Open vTempo For Output As #fnFile  ' Create file name.
      lblStatus = "Creando archivo a enviar"
      DoEvents
    
      strSQL = "select P.*,S.nombre,S.direccion" _
             & " from prm_planilla P inner join Socios S on P.cedula = S.cedula" _
             & " where P.Proceso = " & vFechaProceso & " and P.cod_institucion = " _
             & GLOBALES.gInstitucion & " and P.tipo = 'C' order by P.cedula"
      Call OpenRecordSet(rs, strSQL)
    
      prgProcesoMensual.Max = rs.RecordCount + 1
      prgProcesoMensual.Value = 1
    
      Do While Not rs.EOF
           vLinea = SIFGlobal.fxStringRelleno(Trim(Format(rs!Cedula, "0000000000")), "D", " ", 30) & " " & SIFGlobal.fxStringRelleno(Trim(vCodigoCreditos), "D", " ", 3) & " " & "P "
           
    
           'Sin decimales
           strMonto = Format(rs!Monto_Actual, "Standard")
           strMonto = Replace(strMonto, ".", "")
           strMonto = Replace(strMonto, ",", "")
  
           vLinea = vLinea & SIFGlobal.fxStringRelleno(Trim(CStr(strMonto)), "I", "0", 10) & " "
           
           vLinea = vLinea & SIFGlobal.fxStringRelleno("0", "D", "0", 10) & " "
           vLinea = vLinea & SIFGlobal.fxStringRelleno("0", "D", "0", 10) & " "
           
           vLinea = vLinea & "01/" & Mid(CStr(vFechaProceso), 5, 2) & "/" & Mid(CStr(vFechaProceso), 1, 4) & " "
          ' vLinea = vLinea & "01/01/2099 "
           vLinea = vLinea & Format(vFechaCorte, "dd/mm/yyyy") & " "
           vLinea = vLinea & "00.00" & " "
           vLinea = vLinea & vFechaProceso & "1" & " " 'Blanco para Planillas mensuales
           vLinea = vLinea & "00000.00"
           
           vLinea = SIFGlobal.fxStringRelleno(vLinea, "D", " ", 129)
 
            
          Select Case rs!Movimiento
            Case "E" 'Exclusion
                i = 3
            Case "I" 'Inclusion
                i = 1
            Case "C" 'Cambio
                i = 2
            Case Else
                i = 4 ' Invalido, Porque se Mantiene o No se Proceso
           End Select
           
          If i <> 3 Then 'Envia todos menos las exclusiones
            Print #fnFile, vLinea
          End If
          
         If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
         lblStatus.Caption = "Creando Archivo 2 Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
         rs.MoveNext
      Loop
      rs.Close
    
      Print #fnFile, "!"
    
    Close #fnFile
    
End If 'vCodigoCreditos <> "NO"






'------------------------------------------------------------------------------------------------------------
'           Formatos Nuevos:
'------------------------------------------------------------------------------------------------------------
'nuevo
'01-30  : CEDULA
'32-34  : CODIGO DE DEDUCCION O INGRESO
'36-36  : TIPO DE MOVIMIENTO (I=INCLUSION/CAMBIO INGRESOS/EGRESOS FIJOS DEL EMPLEADO CON CONTROL AUTOMATICO X LIMITES O FECHA HASTA
'                             P=DIRECTO A PLANILLA EN PROCESO (NO AFECTA INGRESOS/EGRESOS FIJOS) CONTROL TOTAL MANUAL
'                             C=CAMBIOS  INGRESO/EGRESOS FIJOS
'                             E=ELIMINAR DE INGRESOS/EGRESOS FIJOS
'         SI TODOS LOS MESES SE REPORTAN TODAS LAS DEDUCCIONES ENTONCES DEBE LLEVAR P
'38-47  : MONTO DEL INGRESO O DEDUCCION CON DOS DECIMALES SIN PUNTO DECIMAL
'49-58  : LIMITE HASTA EL CUAL DEBE REBAJARSE
'60-69  : ACUMULADO CON EL CUAL DEBE INICIARSE
'71-80  : FECHA A PARTIR DE LA CUAL RIGE EL INGRESO O EGRESO
'82-91  : FECHA HASTA LA CUAL RIGE EL INGRESO O EGRESO
'93-97  : PORCENTAJE DE DEDUCCION
'99-105 : PERIODO DE PLANILLA QUE AFECTA (CUANDO EGRESO/INGRESO VA DIRECTO PLANILLA
'         FORMATO AAAAMMS= AAAA=ANO, MM=MES, S=SEMANA (S=BLANCO PARA PLANILLAS MENSUALES
'107-114: CANTIDAD
'115-194: Referencia/Comentario
'196-205: Fecha de Formalización.
'!      : INDICADOR DE TERMINACION DE ARCHIVO
'

Dim vCadena As String

lblStatus.Caption = "Formato: AyA Nuevo!"
DoEvents

vArchivo = "CD-" & vCodInstitucion & "-" & Year(vFecha) & Format(Month(vFecha), "00") _
         & Format(Day(vFecha), "00") & "-01.csv"


vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(GLOBALES.gInstitucion, "00"), vCodInstitucion) _
         & "_" & Format(vFechaProceso, "####-##") & " -NUEVO- [" & Format(vFecha, "ddmmyyyy") & "-F08].txt"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


strSQL = "exec spPrm_Formato_AYA_New " & GLOBALES.gInstitucion & "," & vFechaProceso
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 1
prgProcesoMensual.Value = 1

Do While Not rs.EOF
 
 vCadena = rs!Cadena
 
 If Len(RTrim(vCadena)) > 0 Then
    Print #fnFile, vCadena
 End If
 
 If prgProcesoMensual.Max > prgProcesoMensual.Value Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 lblStatus.Caption = "Creando Archivo Reg. # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
  
  
'       Fin de Formatos Nuevos
'------------------------------------------------------------------------------------------------------------


Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
Call sbReporteGeneracionF02(vFechaProceso, vTempo)
 
lblStatus.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub




Private Sub imgRepAhDevolucion_Click()
Dim vFecha As Currency

On Error GoTo vError
vFecha = CCur(InputBox("Especifique la fecha de proceso " & vbCrLf _
        & "La fecha de Proceso Actual es : " & GLOBALES.glngFechaCR, "Devoluciones de Aportes", GLOBALES.glngFechaCR))
Call sbAhDevoluciones(vFecha)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub imgRepAhFondos_Click()
Dim vFecha As Currency

On Error GoTo vError
vFecha = CCur(InputBox("Especifique la fecha de proceso " & vbCrLf _
        & "La fecha de Proceso Actual es : " & GLOBALES.glngFechaCR, "Traslados Al Fondo"))
Call sbRepFondo(vFecha)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub imgRepAhIncon_Click()
Dim vFecha As Currency

On Error GoTo vError
vFecha = CCur(InputBox("Especifique la fecha de proceso " & vbCrLf _
        & "La fecha de Proceso Actual es : " & GLOBALES.glngFechaCR, "Inconsistencias de Patrimonio", GLOBALES.glngFechaCR))
Call sbAhInconsistencias(vFecha)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub sbReporteCargado(vFecha As Currency)

On Error GoTo vError

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "PROCESO MENSUAL - CARGADO DE INFORMACION"
     
    .Connect = glogon.ConectRPT
     
    .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fecha='" & fxFechaProcesoFormat(vFecha) & "'"
    .Formulas(3) = "usuario='" & glogon.Usuario & "'"
    .Formulas(4) = "institucion='" & GLOBALES.gNombreInstitucion & "'"
    .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_Cargada.rpt")
    
    .SelectionFormula = "{PRM_CARGADO.FECHA_PROCESO} = " & vFecha _
              & " AND {PRM_CARGADO.COD_INSTITUCION} = " & GLOBALES.gInstitucion
    

    .PrintReport

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbReporteGeneracionF02(vFecha As Currency, Optional vUbicacion As String = "")

On Error GoTo vError

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Planillas - Información Generada"
    
    .Connect = glogon.ConectRPT
     
    .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fecha='" & fxFechaProcesoFormat(vFecha) & "'"
    .Formulas(3) = "usuario='" & glogon.Usuario & "'"
    .Formulas(4) = "institucion='" & GLOBALES.gNombreInstitucion & "'"
    
    .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_Generada.rpt")
    .SelectionFormula = "{PRM_PLANILLA.PROCESO} = " & vFecha _
                      & " AND {PRM_PLANILLA.COD_INSTITUCION} = " & GLOBALES.gInstitucion
    

    .PrintReport
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub imgRepCarga_Click()
Dim vFecha As Currency

On Error GoTo vError

vFecha = CCur(InputBox("Especifique la fecha de proceso " & vbCrLf _
        & "La fecha de Proceso Actual es : " & GLOBALES.glngFechaCR, "Cargado de Informacion", GLOBALES.glngFechaCR))
Call sbReporteCargado(vFecha)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub imgRepCrAplica_Click()
Dim vFecha As Currency

On Error GoTo vError

vFecha = CCur(InputBox("Especifique la fecha de proceso " & vbCrLf _
        & "La fecha de Proceso Actual es : " & GLOBALES.glngFechaCR, "Aplicación de Abonos", GLOBALES.glngFechaCR))
'Call sbCrReporteAplicado(vFecha)
Call sbReporteDetalleDeducciones(vFecha)
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub imgRepCrIncon_Click()
Dim vFecha As Currency

On Error GoTo vError

vFecha = CCur(InputBox("Especifique la fecha de proceso " & vbCrLf _
        & "La fecha de Proceso Actual es : " & GLOBALES.glngFechaCR, "Inconsistencias", GLOBALES.glngFechaCR))
Call sbCrReporteInconsistencias(vFecha)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbRepFondo(vFecha As Currency)
Dim vRuta As String, strSQL As String


vRuta = SIFGlobal.fxPathReportes("Sys_Planilla_Fondo.rpt")
Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Reportes Módulo de Fondos"
 
 .Connect = glogon.ConectRPT
 
 .Formulas(0) = "Titulo='LISTA DE INCONSISTENCIAS'"
 .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "Usuario = '" & glogon.Usuario & "'"
 .Formulas(3) = "Titulo = 'TRASLADOS AL FONDO DE INVERSIONES POR PLANILLAS'"
 .Formulas(4) = "SubTitulo = 'Fecha Proceso : " & fxFechaProcesoFormat(vFecha) & "'"
 .Formulas(5) = "institucion = '" & GLOBALES.gNombreInstitucion & "'"
 .ReportFileName = vRuta
 .SelectionFormula = "{PRM_FONDO.PROCESO} = " & vFecha & " AND {PRM_FONDO.COD_INSTITUCION} = " & GLOBALES.gInstitucion
 .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Sub sbAhDevoluciones(vFecha As Currency)
Dim vRuta As String, strSQL As String, rs As New ADODB.Recordset
Dim dbPorcAhorro As Double, dbPorcentaje As Double

strSQL = "Select porc_aporte,porc_ahorro from instituciones where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
  dbPorcentaje = IIf(IsNull(rs!PORC_APORTE), 0, rs!PORC_APORTE) / 100
  dbPorcAhorro = IIf(IsNull(rs!porc_ahorro), 0, rs!porc_ahorro) / 100
rs.Close

vRuta = SIFGlobal.fxPathReportes("Sys_Planilla_PatListados.rpt")
Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 
 .Connect = glogon.ConectRPT
 
 .WindowTitle = "Reportes Módulo de Ahorros"
 .Formulas(1) = "Titulo='DEVOLUCIONES (CASOS EX-SOCIOS)'"
 .Formulas(2) = "Usuario = '" & glogon.Usuario & "'"
 .Formulas(3) = "Fecha = '" & fxFechaProcesoFormat(vFecha) & "'"
 .Formulas(4) = "Porcentaje = " & dbPorcentaje
 .Formulas(5) = "PorcAhorro = " & dbPorcAhorro
 .Formulas(6) = "institucion = '" & GLOBALES.gNombreInstitucion & "'"
 .ReportFileName = vRuta
 .SelectionFormula = "{SOCIOSTEMP.EXISTE} = 'D' AND {SOCIOSTEMP.FECHAPROC} = " & vFecha _
                   & " AND {SOCIOSTEMP.COD_INSTITUCION} = " & GLOBALES.gInstitucion
 .PrintReport
End With

Me.MousePointer = vbDefault

End Sub


Private Sub sbAhInconsistencias(vFecha As Currency)
Dim vRuta As String, strSQL As String, rs As New ADODB.Recordset
Dim dbPorcAhorro As Double, dbPorcentaje As Double

strSQL = "Select porc_aporte,porc_ahorro from instituciones where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
  dbPorcentaje = IIf(IsNull(rs!PORC_APORTE), 0, rs!PORC_APORTE) / 100
  dbPorcAhorro = IIf(IsNull(rs!porc_ahorro), 0, rs!porc_ahorro) / 100
rs.Close

vRuta = SIFGlobal.fxPathReportes("Sys_Planilla_PatListados.rpt")
Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 
 .Connect = glogon.ConectRPT
 
 .WindowTitle = "Reportes Módulo de Ahorros"
 .Formulas(1) = "Titulo='INCONSISTENCIAS DE APORTES'"
 .Formulas(2) = "Usuario = '" & glogon.Usuario & "'"
 .Formulas(3) = "Fecha = '" & fxFechaProcesoFormat(vFecha) & "'"
 .Formulas(4) = "Porcentaje = " & dbPorcentaje
 .Formulas(5) = "PorcAhorro = " & dbPorcAhorro
 .Formulas(6) = "institucion = '" & GLOBALES.gNombreInstitucion & "'"
 .ReportFileName = vRuta
 .SelectionFormula = "{SOCIOSTEMP.EXISTE} = 'N' AND {SOCIOSTEMP.FECHAPROC} = " & vFecha _
                   & " AND {SOCIOSTEMP.COD_INSTITUCION} = " & GLOBALES.gInstitucion
 .PrintReport
End With

Me.MousePointer = vbDefault

End Sub



Private Sub sbReporteDetalleDeducciones(vFecha As Currency)

On Error GoTo vError

Me.MousePointer = vbHourglass


With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Planillas - Detalle de Abonos"
     
     .Connect = glogon.ConectRPT
     
    .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fecha='" & Format(fxFechaServidor, "DD/MM/YYYY") & "'"
    .Formulas(3) = "subtitulo='FECHA PROCESO :" & fxFechaProcesoFormat(vFecha) _
                 & " USUARIO : " & glogon.Usuario & "'"
    .Formulas(4) = "institucion = '" & GLOBALES.gNombreInstitucion & "'"
    .Formulas(5) = "Usuario = '" & glogon.Usuario & "'"
    .ReportFileName = SIFGlobal.fxPathReportes("Sys_Planilla_CrdCarga.rpt")
    .SelectionFormula = "{PRM_CREDITOS.FECHA_PROCESO} = " & vFecha _
                      & " AND {PRM_CREDITOS.COD_INSTITUCION} = " & GLOBALES.gInstitucion
    
    .PrintReport
End With

Me.MousePointer = vbDefault


Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub imgRepDesgloce_Click()
Dim vFecha As Currency

On Error GoTo vError

vFecha = CCur(InputBox("Especifique la fecha de proceso " & vbCrLf _
        & "La fecha de Proceso Actual es : " & GLOBALES.glngFechaCR, "Informacion Desglozada", GLOBALES.glngFechaCR))
Call sbReporteDetalleDeducciones(vFecha)

vError:

End Sub

Private Sub imgRepGenera_Click()
Dim vFecha As Currency

On Error GoTo vError

vFecha = CCur(InputBox("Especifique la fecha de proceso " & vbCrLf _
        & "La fecha de Proceso Actual es : " & GLOBALES.glngFechaCR, "Informacion Generada", GLOBALES.glngFechaCR))
Call sbReporteGeneracionF02(vFecha)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub lblCambiaFechaLabel_Click()
Call imgCambiaFecha_Click
End Sub


Private Sub sbCrAplicaAbonos()
Dim strSQL As String, rs As New ADODB.Recordset, vRegistros As Long
Dim rs2 As New ADODB.Recordset, vFecha As Date, vDocumento As String, vCodDoc As String
Dim vTemp As String

'VARIABLE DE MODULO CON LA FECHA DEL SISTEMA, PARA EVITAR CONSTANTES ACCESOS
'A LA BASE DE DATOS POR CONCEPTO DE FECHA

mFechaSistema = Format(fxFechaServidor, "yyyy/mm/dd")
vFecha = fxFechaServidor

On Error GoTo vError

Me.MousePointer = vbHourglass

prgProcesoMensual.Value = 1


strSQL = "select isnull(desc_Corta,convert(varchar(10),cod_institucion)) as 'CodDoc'" _
       & " from  instituciones where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL, 0)
vCodDoc = Trim(rs!CodDoc)
rs.Close

vDocumento = GLOBALES.glngFechaCR & "." & vCodDoc & ".CRD"

lblStatus = "Limpiando Casos Menores a 1 ..."
DoEvents

strSQL = "delete prm_creditos where cod_institucion = " & GLOBALES.gInstitucion & " and fecha_proceso = " & GLOBALES.glngFechaCR _
       & " and abono < 1"
Call ConectionExecute(strSQL)


lblStatus = "Aplicando abonos  "
DoEvents



Dim pPaso As Integer

pPaso = 1

Do While pPaso <= 3
    lblStatus = "Aplicando Abonos Masivo Paso " & pPaso & " / 3 [Espere]"
    DoEvents
    
    strSQL = "exec spPrmCreditoAplicaAbonosMasivo " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR & ",'" & vDocumento & "', " & pPaso
    Call ConectionExecute(strSQL)
    
    pPaso = pPaso + 1
Loop



lblStatus = "Aplicando Abonos Paso 2 [Espere]"
DoEvents

strSQL = "select count(*) + 1 as Total from prm_creditos" _
       & " where fecha_proceso = " & GLOBALES.glngFechaCR _
       & " and id_aplicacion = 1 and ind_paso = 0" _
       & " and cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
    prgProcesoMensual.Max = (rs!total + 1)
    vRegistros = rs!total
rs.Close

Do While vRegistros > 0
   strSQL = "exec spPrmCreditoAplicaAbonos " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR & ",'" & vDocumento & "',150"
   Call OpenRecordSet(rs, strSQL)
     vRegistros = rs.Fields(0).Value
   rs.Close
   
  If (prgProcesoMensual.Value + vRegistros) <= prgProcesoMensual.Max Then prgProcesoMensual.Value = prgProcesoMensual.Value + vRegistros
  lblStatus.Caption = "Aplicando Registro # " & prgProcesoMensual.Value & " de " & prgProcesoMensual.Max & "     " & Format((prgProcesoMensual.Value / prgProcesoMensual.Max) * 100, "##0") & "%"
  DoEvents
Loop

lblStatus = "Generando morosidades para creditos sin abono"
DoEvents

'Genera Morosidad
strSQL = "exec spPrmCreditoMoraGenera " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR
Call ConectionExecute(strSQL)


lblStatus = "Revisando Integridad de la aplicación, Paso 1"
DoEvents

'Integridad de Aplicacion de Fondos
strSQL = "exec spPrm_Deducciones_Porc_Revision " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

lblStatus = "Generando Comprobante de aplicación!"
DoEvents

'Asiento
strSQL = "exec spPrmCreditoAsiento '1','" & vDocumento & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & glogon.Usuario _
       & "'," & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR
Call ConectionExecute(strSQL)


Call sbBitacoraPlanilla("08", GLOBALES.gInstitucion, GLOBALES.glngFechaCR, "R", vDocumento)

''' Este Proceso se traslada al Procedimiento de Detalle (Desgloce) de Abonos
''' >>>>>>>>>>>>>  Eliminado por nuevo proceso de aplicación  <<<<<<<<<<<<<<<<
'''
'''strSQL = "select pr_cr_aplica_incon,cta_inconsistencia from instituciones" _
'''       & " where cod_institucion = " & GLOBALES.gInstitucion
'''Call OpenRecordSet(rs, strSQL)
'''
'''
'''
''''Pregunta si quiere Aplicar Inconsistencias
'''If rs!pr_cr_aplica_incon = 1 Then
'''  'Vuelve a Revisar Saldos del Mes e Inconsistencias Menores
'''  'Para Tener Actualizadas las operaciones (Razon de Calculos) para la Nc
'''  Call sbCrActualizaSaldos
'''  Call sbCrCalculaSaldoMes(False, GLOBALES.glngFechaCR)
'''
'''  Call sbCrAplicaIncon(rs!cta_inconsistencia, GLOBALES.glngFechaCR)
'''End If
'''
''''Cerrar y volver por posible error catastrofico,
''''no se sabe porque da, pero con esto no pasa
''''************************************************
'''rs.Close


   
lblStatus.Caption = "Procesamiento de Sobrantes [Espere!]"
DoEvents

'18/09/2023  Aplicacion de Sobrantes - Masivamente desde el Server
strSQL = "exec spPrm_Sobrantes_Main " & GLOBALES.glngFechaCR & "," & GLOBALES.gInstitucion _
       & ",'" & vDocumento & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

lblStatus.Caption = "Trasladando Retenciones a Fondos de Ahorros!"
DoEvents

'03/08/2010 (Carga Automatica de Retenciones de Fondos al Auxiliar de Fondos)
strSQL = "exec spPrmFndTrasladoRetAFondo " & GLOBALES.glngFechaCR & "," & GLOBALES.gInstitucion _
       & ",'" & vDocumento & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)
 

lblStatus = "Revisando Integridad de la aplicación, Paso 2"
DoEvents

strSQL = "exec spPrm_Deducciones_Fondos_Revision " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)
 
strSQL = "update instituciones set pr_crAplica = 1 where cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "PRM-CREDITO Aplica Abonos Inst:" & GLOBALES.gInstitucion)

lblStatus = "Generando Informes"
DoEvents

'Carga Traslados al Fondo
Call sbRepFondo(GLOBALES.glngFechaCR)


lblStatus.Caption = "Estado..."

prgProcesoMensual.Value = 1
Me.MousePointer = vbDefault


Call sbEstadoActualProceso

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCrRecalculaCuotaEnMora()
Dim strSQL As String

On Error GoTo vError

'Recalculo el interes Moratorio de las cuotas en Mora

Me.MousePointer = vbHourglass

lblStatus.Caption = "Actualizando Int. Moratorio..."
DoEvents

strSQL = "exec spPrmCrdMoraIntCalcula " & GLOBALES.gInstitucion & "," & GLOBALES.glngFechaCR
Call ConectionExecute(strSQL)

Call sbEstadoActualProceso

Me.MousePointer = vbDefault

lblStatus.Caption = "Estatus..."
DoEvents

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCrCalculaSaldoMes(Optional vSoloRetenciones As Boolean = False _
                              , Optional vProceso As Currency = 0, Optional vTodas As Boolean = True)
Dim strSQL As String

On Error GoTo vError


lblStatus.Caption = "Procesando Saldo de Mes (Espere...!)"

strSQL = "exec spPrmSaldoMesCreditos " & GLOBALES.gInstitucion & "," & vProceso
Call ConectionExecute(strSQL)

lblStatus.Caption = ""

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbDesglocePlanilla()
Dim strSQL As String, rs As New ADODB.Recordset
Dim rsTmp As New ADODB.Recordset

'1. Procesar Ahorros
'2. Procesar Extraordinarios
'3. Procesar Creditos

On Error GoTo vError

Me.MousePointer = vbHourglass


lblStatus.Caption = "Detallando Aportes..."
DoEvents

'Detallando Aportes
strSQL = "exec spPrmAporteDetalla " & GLOBALES.glngFechaCR & "," & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)


'Desgloce de Creditos
Call sbCrDesgloce

Me.MousePointer = vbDefault
MsgBox "- Detale de Aportes y Abonos realizado satisfactoriamente..." _
      & vbCrLf & " - Puede Proceder a las Aplicaciones...", vbInformation
      
strSQL = "update instituciones set pr_desgloza = 1 where cod_institucion = " & GLOBALES.gInstitucion
Call ConectionExecute(strSQL)


Call sbBitacoraPlanilla("04", GLOBALES.gInstitucion, GLOBALES.glngFechaCR, "R")

Call sbReporteDetalleDeducciones(GLOBALES.glngFechaCR)

Call sbEstadoActualProceso
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbProcesosAdd(pTransaccion As String, Optional pTipo As String = "PRE", Optional pProceso As Currency = 0)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

pTipo = Trim(UCase(pTipo))
pTransaccion = Trim(pTransaccion)
If pProceso = 0 Then pProceso = GLOBALES.glngFechaCR

lblStatus.Caption = "Aplicado Procesos Complementarios..."
DoEvents

strSQL = "select * from PRM_PROCESOS_ADD" _
       & " where Transaccion = '" & pTransaccion & "' and EJECUCION_TIPO = '" & pTipo & "'" _
       & " order by EJECUCION_ORDEN,PROC_NUM"
Call OpenRecordSet(rs, strSQL)

prgProcesoMensual.Max = rs.RecordCount + 2
prgProcesoMensual.Value = 1

Do While Not rs.EOF
 
 lblStatus.Caption = "Aplicado: " & Trim(rs!Descripcion)
 DoEvents
 
 strSQL = "exec " & Trim(rs!Procedimiento) & " "
 If rs!PARAMETROS_PLANILLAS = 1 Then
    strSQL = strSQL & GLOBALES.gInstitucion & "," & pProceso
    If Len(Trim(rs!PARAMETROS_ADD)) > 0 Then
       strSQL = strSQL & "," & Trim(rs!PARAMETROS_ADD)
    End If
 Else
    'Parametros Adicionales
    strSQL = strSQL & Trim(rs!PARAMETROS_ADD)
 End If
  
 Call ConectionExecute(strSQL)
    
 If prgProcesoMensual.Value < prgProcesoMensual.Max Then prgProcesoMensual.Value = prgProcesoMensual.Value + 1
 DoEvents
 rs.MoveNext
Loop
rs.Close

lblStatus.Caption = ""
prgProcesoMensual.Value = 1


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical, "Procesos Adicionales"

 

End Sub


Private Sub optAhorros_Click(Index As Integer)
'Desmarca cualquier opción de credito
  optCreditos.Item(0).Value = False
  optCreditos.Item(1).Value = False
  optCreditos.Item(2).Value = False
  optCreditos.Item(3).Value = False

End Sub

Private Sub optCreditos_Click(Index As Integer)
  
'Desmarca cualquier opción de Ahorros
  optAhorros.Item(0).Value = False
  optAhorros.Item(1).Value = False
  optAhorros.Item(2).Value = False
  optAhorros.Item(3).Value = False
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
 Call Valida(KeyAscii)
End Sub


