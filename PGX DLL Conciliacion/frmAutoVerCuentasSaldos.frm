VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAutoVerCuentasSaldos 
   Caption         =   "Autoverificación de Saldos de las Cuentas vrs Saldos Cartera"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   12900
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   10200
      Top             =   600
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   375
      Left            =   7920
      TabIndex        =   14
      Top             =   360
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmAutoVerCuentasSaldos.frx":0000
   End
   Begin XtremeSuiteControls.GroupBox gbCuenta 
      Height          =   4452
      Left            =   0
      TabIndex        =   4
      Top             =   4080
      Width           =   12852
      _Version        =   1441793
      _ExtentX        =   22669
      _ExtentY        =   7853
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton bntExportar 
         Height          =   330
         Left            =   9600
         TabIndex        =   25
         Top             =   240
         Width           =   2652
         _Version        =   1441793
         _ExtentX        =   4678
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Exportar:"
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
         Picture         =   "frmAutoVerCuentasSaldos.frx":0700
      End
      Begin XtremeSuiteControls.TabControl tcMain 
         Height          =   3732
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   12612
         _Version        =   1441793
         _ExtentX        =   22246
         _ExtentY        =   6583
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
         ItemCount       =   8
         SelectedItem    =   6
         Item(0).Caption =   "Tendencia"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "vGridT"
         Item(1).Caption =   "Asignación"
         Item(1).ControlCount=   2
         Item(1).Control(0)=   "vGridAsg"
         Item(1).Control(1)=   "lblNotaNew(0)"
         Item(2).Caption =   "Forma de Pago"
         Item(2).ControlCount=   2
         Item(2).Control(0)=   "vGridFP"
         Item(2).Control(1)=   "lblNotaNew(1)"
         Item(3).Caption =   "Revisión de afectación Contable"
         Item(3).ControlCount=   2
         Item(3).Control(0)=   "vGridRc"
         Item(3).Control(1)=   "lblNotaNew(2)"
         Item(4).Caption =   "Movimientos No Contabilizados"
         Item(4).ControlCount=   2
         Item(4).Control(0)=   "vGridMovNoConta"
         Item(4).Control(1)=   "lblNotaNew(3)"
         Item(5).Caption =   "Cambios"
         Item(5).ControlCount=   2
         Item(5).Control(0)=   "vGridCambios"
         Item(5).Control(1)=   "lblNotaNew(4)"
         Item(6).Caption =   "Analítico"
         Item(6).ControlCount=   4
         Item(6).Control(0)=   "vGridAnalitico"
         Item(6).Control(1)=   "lblNotaNew(5)"
         Item(6).Control(2)=   "btnAnalitico(0)"
         Item(6).Control(3)=   "btnAnalitico(1)"
         Item(7).Caption =   "Concilia Mov"
         Item(7).ControlCount=   4
         Item(7).Control(0)=   "lblNotaNew(6)"
         Item(7).Control(1)=   "btnConcilaMov(0)"
         Item(7).Control(2)=   "vGridConcilia"
         Item(7).Control(3)=   "btnConcilaMov(1)"
         Begin XtremeSuiteControls.PushButton btnAnalitico 
            Height          =   312
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   385
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Contabilidad"
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
         End
         Begin FPSpreadADO.fpSpread vGridT 
            Height          =   3012
            Left            =   -69880
            TabIndex        =   6
            Top             =   480
            Visible         =   0   'False
            Width           =   12372
            _Version        =   524288
            _ExtentX        =   21823
            _ExtentY        =   5313
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
            SpreadDesigner  =   "frmAutoVerCuentasSaldos.frx":086A
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGridAsg 
            Height          =   2775
            Left            =   -69880
            TabIndex        =   10
            Top             =   840
            Visible         =   0   'False
            Width           =   12375
            _Version        =   524288
            _ExtentX        =   21828
            _ExtentY        =   4895
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
            MaxCols         =   6
            SpreadDesigner  =   "frmAutoVerCuentasSaldos.frx":10B4
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGridFP 
            Height          =   2772
            Left            =   -69880
            TabIndex        =   11
            Top             =   840
            Visible         =   0   'False
            Width           =   12372
            _Version        =   524288
            _ExtentX        =   21823
            _ExtentY        =   4890
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
            MaxCols         =   13
            SpreadDesigner  =   "frmAutoVerCuentasSaldos.frx":173D
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGridRc 
            Height          =   2772
            Left            =   -69880
            TabIndex        =   12
            Top             =   840
            Visible         =   0   'False
            Width           =   12372
            _Version        =   524288
            _ExtentX        =   21823
            _ExtentY        =   4890
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
            MaxCols         =   4
            SpreadDesigner  =   "frmAutoVerCuentasSaldos.frx":20CD
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGridMovNoConta 
            Height          =   2772
            Left            =   -69880
            TabIndex        =   13
            Top             =   840
            Visible         =   0   'False
            Width           =   12372
            _Version        =   524288
            _ExtentX        =   21823
            _ExtentY        =   4890
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
            MaxCols         =   12
            SpreadDesigner  =   "frmAutoVerCuentasSaldos.frx":275E
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGridCambios 
            Height          =   2772
            Left            =   -69880
            TabIndex        =   17
            Top             =   840
            Visible         =   0   'False
            Width           =   12372
            _Version        =   524288
            _ExtentX        =   21823
            _ExtentY        =   4890
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
            MaxCols         =   15
            SpreadDesigner  =   "frmAutoVerCuentasSaldos.frx":2F8F
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGridAnalitico 
            Height          =   2772
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   12372
            _Version        =   524288
            _ExtentX        =   21823
            _ExtentY        =   4890
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
            MaxCols         =   13
            SpreadDesigner  =   "frmAutoVerCuentasSaldos.frx":3A30
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.PushButton btnAnalitico 
            Height          =   312
            Index           =   1
            Left            =   1560
            TabIndex        =   27
            Top             =   385
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Auxiliar"
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
         End
         Begin XtremeSuiteControls.PushButton btnConcilaMov 
            Height          =   312
            Index           =   1
            Left            =   -68440
            TabIndex        =   30
            Top             =   384
            Visible         =   0   'False
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Créditos"
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
         End
         Begin FPSpreadADO.fpSpread vGridConcilia 
            Height          =   2772
            Left            =   -69880
            TabIndex        =   29
            Top             =   840
            Visible         =   0   'False
            Width           =   12372
            _Version        =   524288
            _ExtentX        =   21823
            _ExtentY        =   4890
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
            SpreadDesigner  =   "frmAutoVerCuentasSaldos.frx":42B6
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.PushButton btnConcilaMov 
            Height          =   312
            Index           =   0
            Left            =   -69880
            TabIndex        =   28
            Top             =   384
            Visible         =   0   'False
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Débitos"
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
         End
         Begin XtremeShortcutBar.ShortcutCaption lblNotaNew 
            Height          =   372
            Index           =   5
            Left            =   0
            TabIndex        =   24
            Top             =   360
            Width           =   12612
            _Version        =   1441793
            _ExtentX        =   22246
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   "Consulta del Analitíco de la Cuenta"
            ForeColor       =   8388608
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
            Alignment       =   1
            ForeColor       =   8388608
         End
         Begin XtremeShortcutBar.ShortcutCaption lblNotaNew 
            Height          =   372
            Index           =   4
            Left            =   -70000
            TabIndex        =   23
            Top             =   360
            Visible         =   0   'False
            Width           =   12612
            _Version        =   1441793
            _ExtentX        =   22246
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   "Rastrea las Operaciones con cambios contables "
            ForeColor       =   8388608
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
            Alignment       =   1
            ForeColor       =   8388608
         End
         Begin XtremeShortcutBar.ShortcutCaption lblNotaNew 
            Height          =   372
            Index           =   3
            Left            =   -70000
            TabIndex        =   22
            Top             =   360
            Visible         =   0   'False
            Width           =   12612
            _Version        =   1441793
            _ExtentX        =   22246
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   "Rastrea los comprobantes del auxiliar para detectar movimietnos no contabilizados"
            ForeColor       =   8388608
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
            Alignment       =   1
            ForeColor       =   8388608
         End
         Begin XtremeShortcutBar.ShortcutCaption lblNotaNew 
            Height          =   372
            Index           =   2
            Left            =   -70000
            TabIndex        =   21
            Top             =   360
            Visible         =   0   'False
            Width           =   12612
            _Version        =   1441793
            _ExtentX        =   22246
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   "Revisa y Compara los débitos y créditos del periodo entre la contabilidad y el auxiliar"
            ForeColor       =   8388608
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
            Alignment       =   1
            ForeColor       =   8388608
         End
         Begin XtremeShortcutBar.ShortcutCaption lblNotaNew 
            Height          =   372
            Index           =   1
            Left            =   -70000
            TabIndex        =   20
            Top             =   360
            Visible         =   0   'False
            Width           =   12612
            _Version        =   1441793
            _ExtentX        =   22246
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   $"frmAutoVerCuentasSaldos.frx":4BAE
            ForeColor       =   8388608
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
            Alignment       =   1
            ForeColor       =   8388608
         End
         Begin XtremeShortcutBar.ShortcutCaption lblNotaNew 
            Height          =   372
            Index           =   0
            Left            =   -70000
            TabIndex        =   19
            Top             =   360
            Visible         =   0   'False
            Width           =   12612
            _Version        =   1441793
            _ExtentX        =   22246
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   "Indica auxiliares tienen asignada la cuenta contable y reciben movimientos"
            ForeColor       =   8388608
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
            Alignment       =   1
            ForeColor       =   8388608
         End
         Begin XtremeShortcutBar.ShortcutCaption lblNotaNew 
            Height          =   372
            Index           =   6
            Left            =   -70000
            TabIndex        =   31
            Top             =   360
            Visible         =   0   'False
            Width           =   12612
            _Version        =   1441793
            _ExtentX        =   22246
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   "Consulta para Conciliación de Movimientos Contabilidad versus Auxiliar"
            ForeColor       =   8388608
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
            Alignment       =   1
            ForeColor       =   8388608
         End
      End
      Begin XtremeSuiteControls.FlatEdit feDescripcion 
         Height          =   330
         Left            =   3480
         TabIndex        =   9
         Top             =   240
         Width           =   6132
         _Version        =   1441793
         _ExtentX        =   10816
         _ExtentY        =   582
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit feCuenta 
         Height          =   330
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   582
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
      Begin VB.Label Label2 
         Caption         =   "Cuenta:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   732
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2892
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   12972
      _Version        =   524288
      _ExtentX        =   22881
      _ExtentY        =   5101
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
      SpreadDesigner  =   "frmAutoVerCuentasSaldos.frx":4C3F
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodos 
      Height          =   312
      Left            =   2160
      TabIndex        =   15
      Top             =   360
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin XtremeSuiteControls.ComboBox cboAuxiliar 
      Height          =   312
      Left            =   4200
      TabIndex        =   16
      Top             =   360
      Width           =   3612
      _Version        =   1441793
      _ExtentX        =   6376
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Auxiliar:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   1
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label lblEstado 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   9240
      TabIndex        =   1
      Top             =   360
      Width           =   4332
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1092
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   12972
   End
End
Attribute VB_Name = "frmAutoVerCuentasSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mTendencia As String
Dim vAnio As Long, vMes As Integer


Private Sub sbResumen(pVista As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim curTotales(8) As Currency, i As Integer


If cboPeriodos.ListCount = 0 Then Exit Sub

Me.MousePointer = vbHourglass

'Visualización Inicial
tcMain.Item(0).Selected = True
vGrid.MaxRows = 0
vGridT.MaxRows = 0

lblEstado.Caption = "Cargando Información (Espere)"
lblEstado.Refresh

strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
Call OpenRecordSet(rs, strSQL)
    vMes = rs!Mes
    vAnio = rs!Anio
rs.Close

strSQL = "select '' as 'Btn', Anio,Mes,Cod_Cuenta_Mask,Descripcion,Saldo,Saldo_Contable,Diferencia, Operaciones, CURRENCY_SIM, DIVISA_DESC" _
       & " From " & pVista _
       & " Where ANIO = " & vAnio & " And MES = " & vMes _
       & " Order by Cod_Cuenta_Mask"

vPaso = True

Call sbCargaGrid(vGrid, 11, strSQL, True)


lblEstado.Caption = "Calculando Saldos!"
lblEstado.Refresh

curTotales(1) = 0
curTotales(2) = 0
curTotales(3) = 0
curTotales(4) = 0

If vGrid.MaxRows >= 2 Then
        For i = 1 To vGrid.MaxRows - 1
         vGrid.Row = i
         vGrid.Col = 6
         curTotales(1) = curTotales(1) + CCur(vGrid.Text)
         vGrid.Col = 7
         curTotales(2) = curTotales(2) + CCur(vGrid.Text)
         vGrid.Col = 8
         curTotales(3) = curTotales(3) + CCur(vGrid.Text)
         vGrid.Col = 9
         curTotales(4) = curTotales(4) + CCur(vGrid.Text)
        Next i
        
        vGrid.Row = vGrid.MaxRows
        vGrid.Col = 5
        vGrid.Text = "TOTALES:"
        vGrid.Col = 6
        vGrid.Text = Format(CCur(curTotales(1)), "Standard")
        vGrid.Col = 7
        vGrid.Text = Format(CCur(curTotales(2)), "Standard")
        vGrid.Col = 8
        vGrid.Text = Format(CCur(curTotales(3)), "Standard")
        vGrid.Col = 9
        vGrid.Text = Format(CCur(curTotales(4)), "Standard")
End If

lblEstado.Caption = ""
vPaso = False

Me.MousePointer = vbDefault


End Sub



Private Sub bntExportar_Click()

Dim vHeaders As vGridHeaders, vTitulo As String

Select Case tcMain.SelectedItem
    Case 0 'Tendencia
        vHeaders.Columnas = vGridT.MaxCols
        vTitulo = "ProGrX_Aux_Tendencia_Cta_" & feCuenta.Text & "_" & cboPeriodos.Text & "_" & cboAuxiliar.Text
    
        vHeaders.Headers(1) = "Año"
        vHeaders.Headers(2) = "Mes"
        vHeaders.Headers(3) = "Cuenta"
        vHeaders.Headers(4) = "Descripción"
        vHeaders.Headers(5) = "Saldo Auxiliar"
        vHeaders.Headers(6) = "Saldo Contabilidad"
        vHeaders.Headers(7) = "Diferencia"
        vHeaders.Headers(8) = "Operaciones"
    
        Call sbSIFGridExportar(vGridT, vHeaders, vTitulo)
    
    Case 1 'Asignacion
        vHeaders.Columnas = vGridAsg.MaxCols
        vTitulo = "ProGrX_Aux_ASG_Cta_" & feCuenta.Text & "_" & cboPeriodos.Text & "_" & cboAuxiliar.Text
        
        vHeaders.Headers(1) = "Año"
        vHeaders.Headers(2) = "Mes"
        vHeaders.Headers(3) = "Módulo"
        vHeaders.Headers(4) = "Localización"
        vHeaders.Headers(5) = "Tipo"
        vHeaders.Headers(6) = "Descripción"
    
        Call sbSIFGridExportar(vGridAsg, vHeaders, vTitulo)
    
    Case 2 'Forma de Pago
        vHeaders.Columnas = vGridFP.MaxCols
        vTitulo = "ProGrX_Aux_FP_Cta_" & feCuenta.Text & "_" & cboPeriodos.Text & "_" & cboAuxiliar.Text
    
        vHeaders.Headers(1) = "Tipo Documento"
        vHeaders.Headers(2) = "Id Transacción"
        vHeaders.Headers(3) = "Fecha"
        vHeaders.Headers(4) = "Usuario"
        vHeaders.Headers(5) = "Monto"
        vHeaders.Headers(6) = "Forma de Pago"
        vHeaders.Headers(7) = "Monto"
        vHeaders.Headers(8) = "Tipo Cambio"
        vHeaders.Headers(9) = "Num. Referencia"
        vHeaders.Headers(10) = "Aux. Afectado"
        vHeaders.Headers(11) = "Cuenta"
        vHeaders.Headers(12) = "Descripción"
        vHeaders.Headers(13) = "Persona Id"
        vHeaders.Headers(14) = "Persona Nombre"
        
        Call sbSIFGridExportar(vGridFP, vHeaders, vTitulo)
        
    
    Case 3 'Revision Contable
        vHeaders.Columnas = vGridRc.MaxCols
        vTitulo = "ProGrX_Aux_RevConta_Cta_" & feCuenta.Text & "_" & cboPeriodos.Text & "_" & cboAuxiliar.Text
    
        vHeaders.Headers(1) = "Fuente"
        vHeaders.Headers(2) = "Débitos"
        vHeaders.Headers(3) = "Créditos"
        vHeaders.Headers(4) = "Neto"
    
        Call sbSIFGridExportar(vGridRc, vHeaders, vTitulo)
    
    
    Case 4 'Mov no Contabilizados
        vHeaders.Columnas = vGridMovNoConta.MaxCols
        vTitulo = "ProGrX_Aux_MovNoConta_Cta_" & feCuenta.Text & "_" & cboPeriodos.Text & "_" & cboAuxiliar.Text
    
        vHeaders.Headers(1) = "Tipo Documento"
        vHeaders.Headers(2) = "Id Transacción"
        vHeaders.Headers(3) = "Desc- Caso"
        vHeaders.Headers(4) = "Módulo"
        vHeaders.Headers(5) = "Cuenta Default"
        vHeaders.Headers(6) = "Ref. No.1"
        vHeaders.Headers(7) = "Ref. No.2"
        vHeaders.Headers(8) = "Ref. No.3"
    
        vHeaders.Headers(9) = "Fecha"
        vHeaders.Headers(10) = "Usuario"
        vHeaders.Headers(11) = "Concepto"
        vHeaders.Headers(12) = "Monto"
    
        Call sbSIFGridExportar(vGridMovNoConta, vHeaders, vTitulo)
    
    Case 5 'Cambios
        vTitulo = "ProGrX_Aux_Cambios_" & feCuenta.Text & "_" & cboPeriodos.Text & "_" & cboAuxiliar.Text
        vHeaders.Columnas = vGridCambios.MaxCols
    
        vHeaders.Headers(1) = "No. Operación"
        vHeaders.Headers(2) = "Linea"
        vHeaders.Headers(3) = "Proceso"
        vHeaders.Headers(4) = "OpEx"
        vHeaders.Headers(5) = "Saldo al Corte"
        vHeaders.Headers(6) = "Cuenta contable"
        vHeaders.Headers(7) = "Descripción"
        vHeaders.Headers(8) = "Linea Ant"
        vHeaders.Headers(9) = "Proceso Ant"
        vHeaders.Headers(10) = "OpEx Ant"
        vHeaders.Headers(11) = "Cuenta Ant"
        vHeaders.Headers(12) = "Desc. Ant"
        vHeaders.Headers(13) = "Saldo Ant"
        vHeaders.Headers(14) = "Cambio Fecha"
        vHeaders.Headers(15) = "Cambio Monto"
    
        Call sbSIFGridExportar(vGridCambios, vHeaders, vTitulo)
    
    Case 6 'Analitico
        
        If btnAnalitico.Item(0).Checked Then
            vTitulo = "ProGrX_Aux_Analitico_Cta_" & feCuenta.Text & "_" & cboPeriodos.Text & "_Conta_" & cboAuxiliar.Text
        Else
            vTitulo = "ProGrX_Aux_Analitico_Cta_" & feCuenta.Text & "_" & cboPeriodos.Text & "_Auxiliar_" & cboAuxiliar.Text
        End If
        
        vHeaders.Columnas = vGridAnalitico.MaxCols
        
        vHeaders.Headers(1) = "Tipo Asiento"
        vHeaders.Headers(2) = "Num. Asiento"
        vHeaders.Headers(3) = "Descripción"
        vHeaders.Headers(4) = "Referencia"
        vHeaders.Headers(5) = "Notas"
        vHeaders.Headers(6) = "Fecha"
        vHeaders.Headers(7) = "Cuenta"
        vHeaders.Headers(8) = "Unidad"
        vHeaders.Headers(9) = "Centro"
        vHeaders.Headers(10) = "Divisa"
        vHeaders.Headers(11) = "T.C."
        vHeaders.Headers(12) = "Débito"
        vHeaders.Headers(13) = "Crébito"
    
        Call sbSIFGridExportar(vGridAnalitico, vHeaders, vTitulo)


    Case 7 'Concilia Movimientos
        
        If btnConcilaMov.Item(0).Checked Then
            vTitulo = "ProGrX_Aux_Concilia_DB_" & feCuenta.Text & "_" & cboPeriodos.Text & "_" & cboAuxiliar.Text
        Else
            vTitulo = "ProGrX_Aux_Concilia_CR_" & feCuenta.Text & "_" & cboPeriodos.Text & "_" & cboAuxiliar.Text
        End If
        
        vHeaders.Columnas = vGridAnalitico.MaxCols
        
        vHeaders.Headers(1) = "Cuenta"
        vHeaders.Headers(2) = "Tipo Mov"
        vHeaders.Headers(3) = "Tipo Asiento"
        vHeaders.Headers(4) = "Num. Asiento"
        vHeaders.Headers(5) = "Fecha Asiento"
        vHeaders.Headers(6) = "Monto Conta"
        vHeaders.Headers(7) = "Monto Aux."
        vHeaders.Headers(8) = "Diferencia"
        vHeaders.Headers(9) = "Aux.TipoDoc"
        vHeaders.Headers(10) = "Aux.NumDoc"
        vHeaders.Headers(11) = "Aux.Fecha"
    
        Call sbSIFGridExportar(vGridConcilia, vHeaders, vTitulo)

End Select


End Sub

Private Sub btnAnalitico_Click(Index As Integer)
      Call sbCuenta_Analitico(Index)
End Sub

Private Sub btnConcilaMov_Click(Index As Integer)
Call sbCuenta_Analitico_Concilia(Index)
End Sub

Private Sub cboAuxiliar_Click()
If vPaso Then Exit Sub
Call cmdBuscar_Click

End Sub

Private Sub cboPeriodos_Click()
If vPaso Then Exit Sub
Call cmdBuscar_Click
End Sub

Private Sub cmdBuscar_Click()

Select Case cboAuxiliar.Text
    Case "Créditos [Cartera Propia]"
        mTendencia = "Creditos"
        Call sbResumen("vSys_Aux_Creditos_Comparativo_Contable")
        
    Case "Créditos [Cartera Administrada]"
        mTendencia = "Creditos_CA"
        Call sbResumen("vSys_Aux_Creditos_CA_Comparativo_Contable")
        
    Case "Recaudos y Pólizas"
        mTendencia = "Creditos_RC"
        Call sbResumen("vSys_Aux_Creditos_RC_Comparativo_Contable")
    
    Case "Producto Acumulado"
        mTendencia = "Producto"
        Call sbResumen("vSys_Aux_Producto_Comparativo_Contable")
    
    Case "Producto En Suspenso"
        mTendencia = "ProductoSuspenso"
        Call sbResumen("vSys_Aux_ProductoSuspenso_Comparativo_Contable")
    
    Case "Interés Cobrado por Adelanto"
'        mTendencia = "Creditos"
'        Call sbResumen("vSys_Aux_Creditos_Comparativo_Contable")
    
    Case "Gastos/Cargos Diferidos"
'        mTendencia = "Creditos"
'        Call sbResumen("vSys_Aux_Creditos_Comparativo_Contable")
    
    Case "Fondos de Ahorros"
        mTendencia = "Fondos"
        Call sbResumen("vSys_Aux_Fondos_Comparativo_Contable")
    
    Case "Patrimonio"
        mTendencia = "Patrimonio"
        Call sbResumen("vSys_Aux_Patrimonio_Comparativo_Contable")
 
    Case "Activos Fijos"
        mTendencia = "Activos"
        Call sbResumen("vSys_Aux_Activos_Comparativo_Contable")
 
    Case "Inversiones"
        mTendencia = "Inversiones"
        Call sbResumen("vSys_Aux_Inversiones_Comparativo_Contable")
 
 
End Select

End Sub


Private Sub Form_Load()
Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

End Sub

Private Sub Form_Resize()

On Error Resume Next

imgBanner.Width = Me.Width

vGrid.Width = Me.Width - 150
vGrid.Height = (Me.Height - (vGrid.Top + 820)) / 2

gbCuenta.Width = vGrid.Width
gbCuenta.Height = vGrid.Height

gbCuenta.Top = vGrid.Top + gbCuenta.Height + 310

tcMain.Width = gbCuenta.Width - 150
tcMain.Height = gbCuenta.Height - 750

vGridT.Width = tcMain.Width - 250
vGridT.Height = tcMain.Height - 550

vGridAsg.Width = vGridT.Width
vGridFP.Width = vGridT.Width
vGridRc.Width = vGridT.Width
vGridMovNoConta.Width = vGridT.Width
vGridCambios.Width = vGridT.Width
vGridAnalitico.Width = vGridT.Width
vGridConcilia.Width = vGridT.Width

vGridAsg.Height = vGridT.Height - 350
vGridFP.Height = vGridAsg.Height
vGridRc.Height = vGridAsg.Height
vGridMovNoConta.Height = vGridAsg.Height
vGridCambios.Height = vGridAsg.Height
vGridAnalitico.Height = vGridAsg.Height
vGridConcilia.Height = vGridAsg.Height

lblNotaNew.Item(0).Width = tcMain.Width
lblNotaNew.Item(1).Width = tcMain.Width
lblNotaNew.Item(2).Width = tcMain.Width
lblNotaNew.Item(3).Width = tcMain.Width
lblNotaNew.Item(4).Width = tcMain.Width
lblNotaNew.Item(5).Width = tcMain.Width
lblNotaNew.Item(6).Width = tcMain.Width

End Sub



Private Sub sbCuenta_Asignacion()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "EXEC spSys_Aux_Cta_Asigna " & vAnio & "," & vMes & ",'" & fxgCntCuentaFormato(False, feCuenta.Text, 0) & "'"

With vGridAsg
   .MaxRows = 0
   Call OpenRecordSet(rs, strSQL)
   Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     .Col = 1
     .Text = CStr(rs!Anio)
     .Col = 2
     .Text = CStr(rs!Mes)
     .Col = 3
     .Text = rs!Modulo
     .Col = 4
     .Text = rs!Localizacion
     .Col = 5
     .Text = rs!Tipo
     .Col = 6
     .Text = rs!Descripcion
     rs.MoveNext
   Loop
   rs.Close
End With

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCuenta_FormaPago()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "EXEC spSys_Aux_Cta_Forma_Pago " & vAnio & "," & vMes

With vGridFP
   .MaxRows = 0
   Call OpenRecordSet(rs, strSQL)
   Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     .Col = 1
     .Text = rs!Tipo_Documento
     .Col = 2
     .Text = rs!Cod_Transaccion
     .Col = 3
     .Text = rs!Registro_Fecha
     .Col = 4
     .Text = rs!Registro_Usuario
     .Col = 5
     .Text = Format(rs!Monto, "Standard")
     .Col = 6
     .Text = rs!Descripcion
     .Col = 7
     .Text = Format(rs!Tipo_Cambio & "", "Standard")
     .Col = 8
     .Text = rs!NUM_REFERENCIA
     .Col = 9
     .Text = rs!AFECTA_AUXILIAR
     .Col = 10
     .Text = rs!CTA_COD
     .Col = 11
     .Text = rs!CTA_DESC
     .Col = 12
     .Text = rs!CLIENTE_IDENTIFICACION & ""
     .Col = 13
     .Text = rs!CLIENTE_NOMBRE & ""
     
     rs.MoveNext
   Loop
   rs.Close
End With

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCuenta_Rev_Mov()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "EXEC spSys_Aux_Cta_Mov_Rev " & vAnio & "," & vMes & ",'" & fxgCntCuentaFormato(False, feCuenta.Text, 0) & "'"

With vGridRc
   .MaxRows = 0
   Call OpenRecordSet(rs, strSQL)
   Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     .Col = 1
     .Text = "Auxiliar"
     .Col = 2
     .Text = Format(rs!AUX_DEBITO, "Standard")
     .Col = 3
     .Text = Format(rs!AUX_CREDITO, "Standard")
     .Col = 4
     .Text = Format(rs!AUX_DEBITO - rs!AUX_CREDITO, "Standard")
     
     
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     .Col = 1
     .Text = "Contabilidad"
     .Col = 2
     .Text = Format(rs!CNT_DEBITO, "Standard")
     .Col = 3
     .Text = Format(rs!CNT_CREDITO, "Standard")
     .Col = 4
     .Text = Format(rs!CNT_DEBITO - rs!CNT_CREDITO, "Standard")
     
     
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     .Col = 1
     .Text = "Diferencias"
     .Col = 2
     .Text = Format(rs!AUX_DEBITO - rs!CNT_DEBITO, "Standard")
     .Col = 3
     .Text = Format(rs!AUX_CREDITO - rs!CNT_CREDITO, "Standard")
     
     
     rs.MoveNext
   Loop
   rs.Close
End With

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub sbCuenta_Mov_No_Contabilizados()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "EXEC spSys_Aux_Cta_Mov_No_Conta " & vAnio & "," & vMes & ",'" & fxgCntCuentaFormato(False, feCuenta.Text, 0) & "'"

With vGridMovNoConta
   .MaxRows = 0
   Call OpenRecordSet(rs, strSQL)
   
   Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     .Col = 1
     .Text = rs!Tipo_Documento
     .Col = 2
     .Text = rs!Cod_Transaccion
     .Col = 3
     .Text = rs!Detalle
     .Col = 4
     .Text = rs!Modulo
     .Col = 5
     .Text = rs!COD_Cuenta
     .Col = 6
     .Text = rs!Ref_01
     .Col = 7
     .Text = rs!Ref_02
     .Col = 8
     .Text = rs!Ref_02
     .Col = 9
     .Text = rs!fecha & ""
     .Col = 10
     .Text = rs!Usuario & ""
     .Col = 11
     .Text = rs!cod_Concepto & ""
     .Col = 12
     .Text = Format(rs!Monto, "Standard")
     
    
     rs.MoveNext
   Loop
   rs.Close
End With

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCuenta_Analitico(Optional pIndex As Integer = 1)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

btnAnalitico.Item(0).Checked = False
btnAnalitico.Item(1).Checked = False

btnAnalitico.Item(pIndex).Checked = True

If pIndex = 0 Then
    lblNotaNew.Item(5).Caption = "Consulta Analitico de la Contabilidad"
    strSQL = "EXEC spSys_Aux_Cta_Analitico " & vAnio & "," & vMes & ",'" & fxgCntCuentaFormato(False, feCuenta.Text, 0) & "'"
Else
    lblNotaNew.Item(5).Caption = "Consulta Analitico del Auxiliar"
    strSQL = "EXEC spSys_Aux_Cta_Analitico_Aux " & vAnio & "," & vMes & ",'" & fxgCntCuentaFormato(False, feCuenta.Text, 0) & "'"
End If


With vGridAnalitico
   .MaxRows = 0
   Call OpenRecordSet(rs, strSQL)
   
   Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     .Col = 1
     .Text = rs!Tipo_Asiento
     .Col = 2
     .Text = rs!Num_Asiento
     .Col = 3
     .Text = rs!Descripcion
     .Col = 4
     .Text = rs!Referencia & ""
     .Col = 5
     .Text = rs!Notas & ""
     .Col = 6
     .Text = Format(rs!Fecha_Asiento, "dd/mm/yyyy")
     .Col = 7
     .Text = rs!cod_Cuenta_Mask & ""
     .Col = 8
     .Text = rs!Cod_Unidad
     .Col = 9
     .Text = rs!Cod_Centro_Costo
     .Col = 10
     .Text = rs!cod_Divisa
     .Col = 11
     .Text = rs!Tipo_Cambio
     .Col = 12
     .Text = Format(rs!Monto_Debito, "Standard")
     .Col = 13
     .Text = Format(rs!Monto_Credito, "Standard")
     
    
     rs.MoveNext
   Loop
   rs.Close
End With

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCuenta_Analitico_Concilia(Optional pIndex As Integer = 1)
Dim strSQL As String, rs As New ADODB.Recordset
Dim pTipo As String, pMonto As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass

 
btnConcilaMov.Item(0).Checked = False
btnConcilaMov.Item(1).Checked = False

btnConcilaMov.Item(pIndex).Checked = True

If pIndex = 0 Then
    pTipo = "D"
Else
    pTipo = "C"
End If




strSQL = "EXEC spSys_Aux_Cta_Concilia_Cnt " & vAnio & "," & vMes _
        & ",'" & fxgCntCuentaFormato(False, feCuenta.Text, 0) & "','" & pTipo & "'"



'(1) = "Cuenta"
'(2) = "Tipo Mov"
'(3) = "Tipo Asiento"
'(4) = "Num. Asiento"
'(5) = "Fecha Asiento"
'(6) = "Monto Conta"
With vGridConcilia
   .MaxRows = 0
   Call OpenRecordSet(rs, strSQL)
   
   Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     .Col = 1
     .Text = rs!cod_Cuenta_Mask
     .Col = 2
     .Text = pTipo
     .Col = 3
     .Text = rs!Tipo_Asiento
     .Col = 4
     .Text = rs!Num_Asiento
     .Col = 5
     .Text = Format(rs!Fecha_Asiento, "dd/mm/yyyy")
     .Col = 6
     .Text = Format(rs!Monto, "Standard")
     rs.MoveNext
   Loop
   rs.Close
End With

'Carga el Auxiliar
Dim pRow As Long

strSQL = "EXEC spSys_Aux_Cta_Concilia_Aux " & vAnio & "," & vMes _
        & ",'" & fxgCntCuentaFormato(False, feCuenta.Text, 0) & "','" & pTipo & "'"

With vGridConcilia
   Call OpenRecordSet(rs, strSQL)
          
   pRow = 0

'(7) = "Monto Aux."
'(8) = "Diferencia"
'(9) = "Aux.TipoDoc"
'(10) = "Aux.NumDoc"
'(11) = "Aux.Fecha"
   
   Do While Not rs.EOF
     pRow = pRow + 1
     If pRow > .MaxRows Then
         .MaxRows = .MaxRows + 1
         pMonto = 0
     Else
        .Row = pRow
        .Col = 6
        pMonto = CCur(.Text)
     End If
     
     .Row = pRow
     .Col = 7
     .Text = Format(rs!Monto, "Standard")
     .Col = 8
     .Text = Format(pMonto - rs!Monto, "Standard")
     .Col = 9
     .Text = rs!Tipo_Asiento
     .Col = 10
     .Text = rs!Num_Asiento
     .Col = 11
     .Text = Format(rs!Fecha_Asiento, "dd/mm/yyyy")
     rs.MoveNext
   Loop
   rs.Close
End With


Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub sbCuenta_Crd_Cambio()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "EXEC spSys_Aux_Creditos_Cambio_Cta " & vAnio & "," & vMes

With vGridCambios
   .MaxRows = 0
   Call OpenRecordSet(rs, strSQL)
   Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     .Col = 1
     .Text = CStr(rs!id_solicitud)
     .Col = 2
     .Text = rs!Codigo
     .Col = 3
     .Text = rs!Proceso_Desc
     .Col = 4
     .Text = rs!Opex_Desc
     
     .Col = 5
     .Text = Format(rs!saldo_final, "Standard")
     .Col = 6
     .Text = rs!CUENTA_CORTE_MASK
     .Col = 7
     .Text = rs!CUENTA_CORTE_DESC
     
     .Col = 8
     .Text = rs!INICIAL_CODIGO
     
     .Col = 9
     .Text = rs!Inicial_Proceso_Desc
     .Col = 10
     .Text = rs!Inicial_Opex_Desc
     
     .Col = 11
     .Text = rs!CUENTA_INICIO_MASK
     .Col = 12
     .Text = rs!CUENTA_INICIO_DESC
     .Col = 13
     .Text = Format(rs!saldo_inicial, "Standard")
     
     .Col = 14
     .Text = rs!CAMBIO_FECHA & ""
     .Col = 15
     .Text = Format(rs!Cambio_Monto, "Standard")
     
     rs.MoveNext
   Loop
   rs.Close
End With

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub fpSpread1_Advance(ByVal AdvanceNext As Boolean)

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If vPaso Then Exit Sub

Select Case Item.Index
    Case 0 'Tendencia
     'Nada
     bntExportar.Caption = "Exportar: Tendencia"
     
    Case 1 'Asignación
      
      bntExportar.Caption = "Exportar: Asignación"
      Call sbCuenta_Asignacion
    
    Case 2 'Forma de Pago
      bntExportar.Caption = "Exportar: Forma de Pago"
      Call sbCuenta_FormaPago
    
    Case 3 'Revision Contable
      bntExportar.Caption = "Exportar: Rev Contable"
      Call sbCuenta_Rev_Mov
    
    Case 4 'No Contabilizados (Movimientos)
      bntExportar.Caption = "Exportar: No Contabilizados"
      Call sbCuenta_Mov_No_Contabilizados
    
    Case 5 'Cambios Contables de la Operacion
      bntExportar.Caption = "Exportar: Cambios"
      Call sbCuenta_Crd_Cambio
    
    Case 6 'Consulta de Analitico Contable
      bntExportar.Caption = "Exportar: Analítico"
      Call sbCuenta_Analitico(0)
    
    Case 7 'Concilia Movimientos
      bntExportar.Caption = "Exportar: Concilia Mov."
      Call sbCuenta_Analitico_Concilia(0)
    
    
End Select


End Sub

Private Sub TimerX_Timer()
Dim strSQL As String

TimerX.Interval = 0
TimerX.Enabled = False

vPaso = True

cboAuxiliar.Clear
cboAuxiliar.AddItem "Créditos [Cartera Propia]"
cboAuxiliar.AddItem "Créditos [Cartera Administrada]"
cboAuxiliar.AddItem "Recaudos y Pólizas"
cboAuxiliar.AddItem "Producto Acumulado"
cboAuxiliar.AddItem "Producto En Suspenso"
'cboAuxiliar.AddItem "Interés Cobrado por Adelanto"
'cboAuxiliar.AddItem "Gastos/Cargos Diferidos"
cboAuxiliar.AddItem "Fondos de Ahorros"
cboAuxiliar.AddItem "Patrimonio"

cboAuxiliar.AddItem "Activos Fijos"
cboAuxiliar.AddItem "Inversiones"

cboAuxiliar.Text = "Créditos"


strSQL = "select Top 36 convert(varchar(10), Anio) + '-' + convert(varchar(10), Mes) as 'ItmX',id_per_historico as 'IdX'" _
       & " from ase_per_historico order by anio desc,mes desc"

Call sbCbo_Llena_New(cboPeriodos, strSQL, False, True)

vPaso = False

Call cmdBuscar_Click
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If vPaso Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

vGridT.MaxRows = 0

lblEstado.Caption = "Cargando Información (Espere)"
lblEstado.Refresh

'strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
'Call OpenRecordSet(rs, strSQL)
'    vMes = rs!Mes
'    vAnio = rs!Anio
'rs.Close


tcMain.Item(0).Selected = True

vGrid.Row = Row

vGrid.Col = 4
feCuenta.Text = vGrid.Text
vGrid.Col = 5
feDescripcion.Text = vGrid.Text

strSQL = "exec spSys_Aux_Tendencia_Contable " & vAnio & "," & vMes & ",'" & feCuenta.Text & "','" & mTendencia & "'"
Call sbCargaGrid(vGridT, 8, strSQL, True)

Me.MousePointer = vbDefault
lblEstado.Caption = ""

Exit Sub

vError:
    Me.MousePointer = vbDefault
    lblEstado.Caption = ""
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

