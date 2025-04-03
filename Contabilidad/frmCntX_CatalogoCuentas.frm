VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCntX_CatalogoCuentas 
   Caption         =   "Catálogo de Cuentas"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   15075
   HelpContextID   =   6
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   15075
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ProgressBar PrgBarX 
      Height          =   135
      Left            =   0
      TabIndex        =   71
      Top             =   7200
      Width           =   6615
      _Version        =   1572864
      _ExtentX        =   11668
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5892
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   14892
      _Version        =   1572864
      _ExtentX        =   26268
      _ExtentY        =   10393
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
      ItemCount       =   4
      Item(0).Caption =   "Catálogo"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Mapeo"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "gbMapeo"
      Item(2).Caption =   "Bajar Nivel"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "gbBajarNivel"
      Item(3).Caption =   "Detalle"
      Item(3).ControlCount=   3
      Item(3).Control(0)=   "tcDetalle"
      Item(3).Control(1)=   "btnDetalle(0)"
      Item(3).Control(2)=   "btnDetalle(1)"
      Begin XtremeSuiteControls.GroupBox gbMapeo 
         Height          =   2892
         Left            =   -68560
         TabIndex        =   22
         Top             =   1080
         Visible         =   0   'False
         Width           =   10452
         _Version        =   1572864
         _ExtentX        =   18436
         _ExtentY        =   5101
         _StockProps     =   79
         Caption         =   "Mapeo de Cuentas"
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
         Begin XtremeSuiteControls.PushButton btnMapeo 
            Height          =   612
            Index           =   0
            Left            =   8040
            TabIndex        =   27
            Top             =   2160
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   1080
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
            Appearance      =   16
            Picture         =   "frmCntX_CatalogoCuentas.frx":0000
         End
         Begin XtremeSuiteControls.FlatEdit txtMapeoCta_Actual 
            Height          =   312
            Left            =   2520
            TabIndex        =   23
            Top             =   840
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3831
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtMapeoCta_Nueva 
            Height          =   312
            Left            =   2520
            TabIndex        =   25
            Top             =   1320
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3831
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.PushButton btnMapeo 
            Height          =   612
            Index           =   1
            Left            =   9600
            TabIndex        =   28
            ToolTipText     =   "Cancelar"
            Top             =   2160
            Width           =   852
            _Version        =   1572864
            _ExtentX        =   1503
            _ExtentY        =   1080
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
            Appearance      =   16
            Picture         =   "frmCntX_CatalogoCuentas.frx":09C3
         End
         Begin XtremeSuiteControls.CheckBox chkMapeo_Transac 
            Height          =   252
            Left            =   2520
            TabIndex        =   31
            Top             =   1800
            Width           =   6252
            _Version        =   1572864
            _ExtentX        =   11028
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cambiar Transacciones (Pendientes)"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtMapeoCta_Actual_Desc 
            Height          =   312
            Left            =   4680
            TabIndex        =   24
            Top             =   840
            Width           =   5772
            _Version        =   1572864
            _ExtentX        =   10181
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtMapeoCta_Nueva_Desc 
            Height          =   312
            Left            =   4680
            TabIndex        =   26
            Top             =   1320
            Width           =   5772
            _Version        =   1572864
            _ExtentX        =   10181
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   5
            Left            =   360
            TabIndex        =   30
            Top             =   1320
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cuenta Nueva"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   4
            Left            =   360
            TabIndex        =   29
            Top             =   840
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cuenta Actual"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5892
         Left            =   0
         TabIndex        =   21
         Top             =   360
         Width           =   14772
         _Version        =   524288
         _ExtentX        =   26056
         _ExtentY        =   10393
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
         MaxCols         =   501
         ScrollBars      =   2
         ScrollBarShowMax=   0   'False
         SpreadDesigner  =   "frmCntX_CatalogoCuentas.frx":1166
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         ScrollBarTrack  =   1
         AppearanceStyle =   1
         ScrollBarStyle  =   2
      End
      Begin XtremeSuiteControls.GroupBox gbBajarNivel 
         Height          =   2892
         Left            =   -68680
         TabIndex        =   32
         Top             =   1200
         Visible         =   0   'False
         Width           =   10452
         _Version        =   1572864
         _ExtentX        =   18436
         _ExtentY        =   5101
         _StockProps     =   79
         Caption         =   "Bajar de Nivel"
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
         Begin XtremeSuiteControls.FlatEdit txtBN_Cta 
            Height          =   312
            Left            =   2520
            TabIndex        =   34
            Top             =   840
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3831
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.PushButton btnBajaNivel 
            Height          =   612
            Index           =   0
            Left            =   8040
            TabIndex        =   33
            Top             =   1800
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   1080
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
            Appearance      =   16
            Picture         =   "frmCntX_CatalogoCuentas.frx":21F1
         End
         Begin XtremeSuiteControls.PushButton btnBajaNivel 
            Height          =   612
            Index           =   1
            Left            =   9600
            TabIndex        =   36
            ToolTipText     =   "Cancelar"
            Top             =   1800
            Width           =   852
            _Version        =   1572864
            _ExtentX        =   1503
            _ExtentY        =   1080
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
            Appearance      =   16
            Picture         =   "frmCntX_CatalogoCuentas.frx":2BB4
         End
         Begin XtremeSuiteControls.FlatEdit txtBN_Cta_Desc 
            Height          =   312
            Left            =   4680
            TabIndex        =   35
            Top             =   840
            Width           =   5772
            _Version        =   1572864
            _ExtentX        =   10181
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   7
            Left            =   360
            TabIndex        =   37
            Top             =   840
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cuenta Actual"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.TabControl tcDetalle 
         Height          =   4812
         Left            =   -69880
         TabIndex        =   38
         Top             =   1080
         Visible         =   0   'False
         Width           =   13812
         _Version        =   1572864
         _ExtentX        =   24363
         _ExtentY        =   8488
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
         Appearance      =   3
         Color           =   4
         ItemCount       =   3
         Item(0).Caption =   "Adicional"
         Item(0).ControlCount=   11
         Item(0).Control(0)=   "gTraduce"
         Item(0).Control(1)=   "txtDT_Cuenta"
         Item(0).Control(2)=   "txtDT_Cuenta_Desc"
         Item(0).Control(3)=   "txtDT_Cuenta_Alterna_Desc"
         Item(0).Control(4)=   "Label2(6)"
         Item(0).Control(5)=   "Label2(8)"
         Item(0).Control(6)=   "chkDT_Exclusiva"
         Item(0).Control(7)=   "cboDT_Exclusiva_Unidad"
         Item(0).Control(8)=   "cboDT_Exclusiva_Centro"
         Item(0).Control(9)=   "Label2(10)"
         Item(0).Control(10)=   "Label2(9)"
         Item(1).Caption =   "Prorrateo"
         Item(1).ControlCount=   8
         Item(1).Control(0)=   "gProrratea"
         Item(1).Control(1)=   "txtDT_Prorrateo_Total"
         Item(1).Control(2)=   "lblProrrateoTotal"
         Item(1).Control(3)=   "chkDT_Prorratea"
         Item(1).Control(4)=   "cboDT_Prorratea_Unidad"
         Item(1).Control(5)=   "cboDT_Prorratea_Centro"
         Item(1).Control(6)=   "Label2(12)"
         Item(1).Control(7)=   "Label2(11)"
         Item(2).Caption =   "Difrencial Cambiario"
         Item(2).ControlCount=   11
         Item(2).Control(0)=   "chkDT_Diferencial"
         Item(2).Control(1)=   "Label2(14)"
         Item(2).Control(2)=   "Label2(13)"
         Item(2).Control(3)=   "Label2(15)"
         Item(2).Control(4)=   "Label2(16)"
         Item(2).Control(5)=   "txtDT_DC_Cuenta_GST_Desc"
         Item(2).Control(6)=   "txtDT_DC_Cuenta_ING"
         Item(2).Control(7)=   "txtDT_DC_Cuenta_ING_Desc"
         Item(2).Control(8)=   "txtDT_DC_Cuenta_GST"
         Item(2).Control(9)=   "cboDT_DC_Unidad"
         Item(2).Control(10)=   "cboDT_DC_Centro"
         Begin FPSpreadADO.fpSpread gTraduce 
            Height          =   2172
            Left            =   480
            TabIndex        =   39
            Top             =   2520
            Width           =   12732
            _Version        =   524288
            _ExtentX        =   22458
            _ExtentY        =   3831
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
            MaxCols         =   2
            ScrollBars      =   2
            ScrollBarShowMax=   0   'False
            SpreadDesigner  =   "frmCntX_CatalogoCuentas.frx":3357
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            ScrollBarTrack  =   1
            AppearanceStyle =   1
            ScrollBarStyle  =   2
         End
         Begin FPSpreadADO.fpSpread gProrratea 
            Height          =   2892
            Left            =   -68560
            TabIndex        =   40
            Top             =   1320
            Visible         =   0   'False
            Width           =   11292
            _Version        =   524288
            _ExtentX        =   19918
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
            MaxCols         =   5
            ScrollBars      =   2
            ScrollBarShowMax=   0   'False
            SpreadDesigner  =   "frmCntX_CatalogoCuentas.frx":38F9
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            ScrollBarTrack  =   1
            AppearanceStyle =   1
            ScrollBarStyle  =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtDT_Prorrateo_Total 
            Height          =   312
            Left            =   -58600
            TabIndex        =   41
            Top             =   4440
            Visible         =   0   'False
            Width           =   972
            _Version        =   1572864
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
            Text            =   "0"
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtDT_Cuenta 
            Height          =   312
            Left            =   2160
            TabIndex        =   42
            Top             =   600
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3831
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtDT_Cuenta_Desc 
            Height          =   312
            Left            =   4320
            TabIndex        =   43
            Top             =   600
            Width           =   7212
            _Version        =   1572864
            _ExtentX        =   12721
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtDT_Cuenta_Alterna_Desc 
            Height          =   312
            Left            =   4320
            TabIndex        =   44
            Top             =   960
            Width           =   7212
            _Version        =   1572864
            _ExtentX        =   12721
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.CheckBox chkDT_Prorratea 
            Height          =   612
            Left            =   -58840
            TabIndex        =   45
            Top             =   600
            Visible         =   0   'False
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Prorratea Cuenta?  "
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.ComboBox cboDT_Prorratea_Unidad 
            Height          =   312
            Left            =   -68080
            TabIndex        =   46
            Top             =   840
            Visible         =   0   'False
            Width           =   4572
            _Version        =   1572864
            _ExtentX        =   8070
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboDT_Prorratea_Centro 
            Height          =   312
            Left            =   -63520
            TabIndex        =   47
            Top             =   840
            Visible         =   0   'False
            Width           =   4572
            _Version        =   1572864
            _ExtentX        =   8070
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.CheckBox chkDT_Exclusiva 
            Height          =   252
            Left            =   5640
            TabIndex        =   48
            Top             =   1320
            Width           =   5892
            _Version        =   1572864
            _ExtentX        =   10393
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Esta Cuenta es Exclusiva a una Unidad Estratégica de Negocios?  "
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.ComboBox cboDT_Exclusiva_Unidad 
            Height          =   312
            Left            =   6960
            TabIndex        =   49
            Top             =   1680
            Width           =   4572
            _Version        =   1572864
            _ExtentX        =   8070
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboDT_Exclusiva_Centro 
            Height          =   312
            Left            =   6960
            TabIndex        =   50
            Top             =   2040
            Width           =   4572
            _Version        =   1572864
            _ExtentX        =   8070
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtDT_DC_Cuenta_GST_Desc 
            Height          =   312
            Left            =   -64000
            TabIndex        =   58
            Top             =   1320
            Visible         =   0   'False
            Width           =   6732
            _Version        =   1572864
            _ExtentX        =   11874
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.CheckBox chkDT_Diferencial 
            Height          =   252
            Left            =   -66160
            TabIndex        =   59
            Top             =   600
            Visible         =   0   'False
            Width           =   5892
            _Version        =   1572864
            _ExtentX        =   10393
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Utiliza Cuentas Especiales para Diferencial Cambiario?  "
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
            TextAlignment   =   4
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtDT_DC_Cuenta_ING 
            Height          =   312
            Left            =   -66160
            TabIndex        =   60
            Top             =   960
            Visible         =   0   'False
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3831
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtDT_DC_Cuenta_ING_Desc 
            Height          =   312
            Left            =   -64000
            TabIndex        =   61
            Top             =   960
            Visible         =   0   'False
            Width           =   6732
            _Version        =   1572864
            _ExtentX        =   11874
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtDT_DC_Cuenta_GST 
            Height          =   312
            Left            =   -66160
            TabIndex        =   62
            Top             =   1320
            Visible         =   0   'False
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3831
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.ComboBox cboDT_DC_Unidad 
            Height          =   312
            Left            =   -61960
            TabIndex        =   65
            Top             =   1800
            Visible         =   0   'False
            Width           =   4572
            _Version        =   1572864
            _ExtentX        =   8070
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboDT_DC_Centro 
            Height          =   312
            Left            =   -61960
            TabIndex        =   66
            Top             =   2160
            Visible         =   0   'False
            Width           =   4572
            _Version        =   1572864
            _ExtentX        =   8070
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   16
            Left            =   -64120
            TabIndex        =   68
            Top             =   2160
            Visible         =   0   'False
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Centro de Costo"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   15
            Left            =   -64120
            TabIndex        =   67
            Top             =   1800
            Visible         =   0   'False
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Unidad"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   13
            Left            =   -68320
            TabIndex        =   64
            Top             =   960
            Visible         =   0   'False
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cuenta DC Ingresos"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   14
            Left            =   -68320
            TabIndex        =   63
            Top             =   1320
            Visible         =   0   'False
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cuenta DC Gastos"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblProrrateoTotal 
            Height          =   252
            Left            =   -60760
            TabIndex        =   57
            Top             =   4440
            Visible         =   0   'False
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Total:"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   6
            Left            =   2160
            TabIndex        =   56
            Top             =   960
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Descripción Alterna"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   8
            Left            =   0
            TabIndex        =   55
            Top             =   600
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cuenta"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   12
            Left            =   -63520
            TabIndex        =   54
            Top             =   600
            Visible         =   0   'False
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Centro de Costo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   11
            Left            =   -68080
            TabIndex        =   53
            Top             =   600
            Visible         =   0   'False
            Width           =   1452
            _Version        =   1572864
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Unidad"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   10
            Left            =   4800
            TabIndex        =   52
            Top             =   2040
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Centro de Costo"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   9
            Left            =   4800
            TabIndex        =   51
            Top             =   1680
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Unidad"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton btnDetalle 
         Height          =   612
         Index           =   0
         Left            =   -58480
         TabIndex        =   69
         Top             =   360
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Guardar"
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
         Picture         =   "frmCntX_CatalogoCuentas.frx":3FD7
      End
      Begin XtremeSuiteControls.PushButton btnDetalle 
         Height          =   612
         Index           =   1
         Left            =   -56920
         TabIndex        =   70
         ToolTipText     =   "Cancelar"
         Top             =   360
         Visible         =   0   'False
         Width           =   852
         _Version        =   1572864
         _ExtentX        =   1503
         _ExtentY        =   1080
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
         Appearance      =   16
         Picture         =   "frmCntX_CatalogoCuentas.frx":46CE
      End
   End
   Begin VB.Timer TimerInicio 
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin XtremeSuiteControls.GroupBox gbFiltros 
      Height          =   705
      Left            =   240
      TabIndex        =   10
      Top             =   495
      Width           =   14055
      _Version        =   1572864
      _ExtentX        =   24791
      _ExtentY        =   1244
      _StockProps     =   79
      Appearance      =   16
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkBalance 
         Height          =   255
         Left            =   11520
         TabIndex        =   11
         Top             =   360
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Mostrar Balance"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBarX 
         Height          =   285
         Left            =   10800
         TabIndex        =   12
         Top             =   360
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   503
         _Version        =   393216
         Arrows          =   65536
         Min             =   1
         Max             =   5
         Orientation     =   1638401
         Value           =   5
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltroCuenta 
         Height          =   312
         Index           =   0
         Left            =   1080
         TabIndex        =   13
         Top             =   360
         Width           =   2172
         _Version        =   1572864
         _ExtentX        =   3831
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltroCuenta 
         Height          =   312
         Index           =   1
         Left            =   3240
         TabIndex        =   14
         Top             =   360
         Width           =   4932
         _Version        =   1572864
         _ExtentX        =   8700
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNivel 
         Height          =   315
         Left            =   9840
         TabIndex        =   15
         Top             =   360
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1503
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
         Text            =   "8"
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboDivisa 
         Height          =   330
         Left            =   8160
         TabIndex        =   73
         Top             =   360
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2990
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
         Index           =   17
         Left            =   8280
         TabIndex        =   72
         Top             =   120
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1503
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Divisa"
         ForeColor       =   4210752
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   19
         Top             =   360
         Width           =   972
         _Version        =   1572864
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Filtros:"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   1
         Left            =   1080
         TabIndex        =   18
         Top             =   120
         Width           =   2172
         _Version        =   1572864
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuenta"
         ForeColor       =   4210752
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   2
         Left            =   3240
         TabIndex        =   17
         Top             =   120
         Width           =   4812
         _Version        =   1572864
         _ExtentX        =   8488
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Descripcion"
         ForeColor       =   4210752
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   3
         Left            =   9840
         TabIndex        =   16
         Top             =   120
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1503
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Nivel"
         ForeColor       =   4210752
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox gbToolAux 
      Height          =   324
      Left            =   4920
      TabIndex        =   1
      Top             =   80
      Width           =   6972
      _Version        =   1572864
      _ExtentX        =   12298
      _ExtentY        =   564
      _StockProps     =   79
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnToolAux 
         Height          =   312
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   12
         Width           =   972
         _Version        =   1572864
         _ExtentX        =   1714
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Histórico"
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
      Begin XtremeSuiteControls.PushButton btnToolAux 
         Height          =   312
         Index           =   1
         Left            =   960
         TabIndex        =   3
         Top             =   12
         Width           =   972
         _Version        =   1572864
         _ExtentX        =   1714
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Formato"
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
      Begin XtremeSuiteControls.PushButton btnToolAux 
         Height          =   312
         Index           =   2
         Left            =   1920
         TabIndex        =   4
         Top             =   12
         Width           =   972
         _Version        =   1572864
         _ExtentX        =   1714
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Revisión"
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
      Begin XtremeSuiteControls.PushButton btnToolAux 
         Height          =   312
         Index           =   3
         Left            =   2880
         TabIndex        =   5
         Top             =   12
         Width           =   1092
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Bajar Nivel"
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
      Begin XtremeSuiteControls.PushButton btnToolAux 
         Height          =   312
         Index           =   4
         Left            =   3960
         TabIndex        =   6
         Top             =   12
         Width           =   972
         _Version        =   1572864
         _ExtentX        =   1714
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Mapeo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   7
      End
      Begin XtremeSuiteControls.PushButton btnToolAux 
         Height          =   312
         Index           =   5
         Left            =   5040
         TabIndex        =   7
         Top             =   12
         Width           =   972
         _Version        =   1572864
         _ExtentX        =   1714
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Exportar"
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
         Appearance      =   7
      End
      Begin XtremeSuiteControls.PushButton btnToolAux 
         Height          =   312
         Index           =   6
         Left            =   6000
         TabIndex        =   9
         Top             =   12
         Width           =   972
         _Version        =   1572864
         _ExtentX        =   1714
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Reporte"
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
         Appearance      =   7
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   270
      Left            =   12360
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   476
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      HelpContextID   =   6
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repResumen"
                  Text            =   "Catálogo - Resumen "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repSep1"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repDetalle"
                  Text            =   "Catálogo - Detalle"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   492
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15132
      _Version        =   1572864
      _ExtentX        =   26691
      _ExtentY        =   868
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   6
   End
End
Attribute VB_Name = "frmCntX_CatalogoCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mTipoCuentaUltimaSel As String, mDivisaUltimaSel As String
Dim mTipoCuentaLista As String, mDivisaLista As String, mCuenta As String
Dim vMensaje As String, blnCuentaCorrecta As Boolean
Dim vPaso As Boolean


Private Sub sbProrratea_Total()
Dim curTotal As Currency, i As Long

On Error GoTo vError

curTotal = 0

With gProrratea

    For i = 1 To .MaxRows
       .Row = i
       .Col = 5
       If IsNumeric(.Text) Then
        curTotal = curTotal + CCur(.Text)
       End If
    Next i
End With

txtDT_Prorrateo_Total.Text = Format(curTotal, "Standard")

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbCuenta_DT_Load(pCuenta As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

pCuenta = fxCntX_CuentaFormato(False, pCuenta, 0)

strSQL = "exec spCntX_Cuenta_Detalle " & gCntX_Parametros.CodigoConta & ",'" & pCuenta & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
    
    txtDT_Cuenta.Text = rs!COD_CUENTA_MASK
    txtDT_Cuenta_Desc.Text = rs!Descripcion
    txtDT_Cuenta_Alterna_Desc.Text = rs!DESCRIPCION_ALTERNA & ""
    
    chkDT_Exclusiva.Value = rs!EXCLUSIVA_INDICA
    Call sbCboAsignaDato(cboDT_Exclusiva_Unidad, rs!EXCLUSIVA_UNIDAD_DESC, True, rs!EXCLUSIVA_UNIDAD)
    Call sbCboAsignaDato(cboDT_Exclusiva_Centro, rs!EXCLUSIVA_CENTRO_COSTO_DESC, True, rs!EXCLUSIVA_CENTRO_COSTO)
    
    
    
    chkDT_Prorratea.Value = rs!PRORRATEA_INDICA
    txtDT_Prorrateo_Total.Text = Format(rs!PRORRATEA_TOTAL, "Standard")
    Call sbCboAsignaDato(cboDT_Prorratea_Unidad, rs!PRORRATEA_UNIDAD_DESC, True, rs!PRORRATEA_UNIDAD)
    Call sbCboAsignaDato(cboDT_Prorratea_Centro, rs!PRORRATEA_CENTRO_COSTO_DESC, True, rs!PRORRATEA_CENTRO_COSTO)
    
    
    chkDT_Diferencial.Value = rs!DC_ESPECIAL_INDICA
    txtDT_DC_Cuenta_ING.Text = rs!DC_CTA_INGRESO_MASK
    txtDT_DC_Cuenta_ING_Desc.Text = rs!DC_CTA_INGRESO_DESC
    
    txtDT_DC_Cuenta_GST.Text = rs!DC_CTA_GASTO_MASK
    txtDT_DC_Cuenta_GST_Desc.Text = rs!DC_CTA_GASTO_DESC
    
    Call sbCboAsignaDato(cboDT_DC_Unidad, rs!DC_UNIDAD_DESC, True, rs!DC_UNIDAD)
    Call sbCboAsignaDato(cboDT_DC_Centro, rs!DC_CENTRO_COSTO_DESC, True, rs!DC_CENTRO_COSTO)
    
End If
Me.MousePointer = vbDefault


'Traducciones
strSQL = "select COD_IDIOMA, DESCRIPCION " _
       & " From CNTX_CUENTAS_TRADUCCION" _
       & " Where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & "  And COD_CUENTA = '" & pCuenta & "'"
Call sbCargaGrid(gTraduce, 2, strSQL)

'Prorrateos
strSQL = "select P.COD_UNIDAD , ISNULL(U.DESCRIPCION,'') AS 'UNIDAD_DESC'" _
       & ", P.COD_CENTRO_COSTO, ISNULL(Cc.DESCRIPCION,'') AS 'CENTRO_DESC'" _
       & ", P.PORCENTAJE" _
       & " from CNTX_CUENTAS_PRORRATA P" _
       & "  LEFT JOIN CNTX_UNIDADES U on P.COD_CONTABILIDAD = U.COD_CONTABILIDAD AND P.COD_UNIDAD = U.COD_UNIDAD" _
       & "  LEFT JOIN CNTX_CENTRO_COSTOS Cc on P.COD_CONTABILIDAD = Cc.COD_CONTABILIDAD AND P.COD_CENTRO_COSTO = Cc.COD_CENTRO_COSTO" _
       & " Where P.COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & "  And P.COD_CUENTA = '" & pCuenta & "'"
Call sbCargaGrid(gProrratea, 5, strSQL)

Call sbProrratea_Total

tcDetalle.Item(0).Selected = True



Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCuenta_DT_Save()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


'spCntX_Cuenta_Detalle_Guardar(@Contabilidad int, @Cuenta varchar(60), @Desc_Alter varchar(255)
'                    , @Ex_Ind smallint , @Ex_Unidad varchar(10), @Ex_Centro varchar(10)
'                    , @Pr_Ind smallint , @Pr_Unidad varchar(10), @Pr_Centro varchar(10), @Pr_Total dec(10,2)
'                    , @Dc_Ind smallint , @Dc_Unidad varchar(10), @Dc_Centro varchar(10)
'                    , @Dc_Cta_Ing varchar(60), @Dc_Cta_Gst varchar(60)
'                    , @Usuario varchar(30))

strSQL = "exec spCntX_Cuenta_Detalle_Guardar " & gCntX_Parametros.CodigoConta & ",'" & fxCntX_CuentaFormato(False, txtDT_Cuenta.Text, 0) _
        & "','" & txtDT_Cuenta_Alterna_Desc.Text _
        & "'," & chkDT_Exclusiva.Value & ",'" _
        & cboDT_Exclusiva_Unidad.ItemData(cboDT_Exclusiva_Unidad.ListIndex) & "','" & cboDT_Exclusiva_Centro.ItemData(cboDT_Exclusiva_Centro.ListIndex) _
        & "'," & chkDT_Prorratea.Value & ",'" _
        & cboDT_Prorratea_Unidad.ItemData(cboDT_Prorratea_Unidad.ListIndex) & "','" & cboDT_Prorratea_Centro.ItemData(cboDT_Prorratea_Centro.ListIndex) _
        & "'," & CCur(txtDT_Prorrateo_Total.Text) _
        & "," & chkDT_Diferencial.Value & ",'" _
        & cboDT_DC_Unidad.ItemData(cboDT_DC_Unidad.ListIndex) & "','" & cboDT_DC_Centro.ItemData(cboDT_DC_Centro.ListIndex) _
        & "','" & fxCntX_CuentaFormato(False, txtDT_DC_Cuenta_ING.Text, 0) & "','" & fxCntX_CuentaFormato(False, txtDT_DC_Cuenta_GST.Text, 0) _
        & "','" & glogon.Usuario & "'"
        
Call ConectionExecute(strSQL)

Call Bitacora("Registra", "Cuentas Info Adicional: " & txtDT_Cuenta.Text)

Me.MousePointer = vbDefault

MsgBox "Información Guardada Satisfactoriamente!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub sbMapeo_Cuentas()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


'spCntX_Mapeo_Cuentas(@CntActual int,  @CtaActual varchar(60), @CntNew int, @CtaNew varchar(60), @Usuario varchar(30)
'        , @CambioCnf smallint = 1, @CambioTrn smallint = 0)

strSQL = "exec spCntX_Mapeo_Cuentas " & gCntX_Parametros.CodigoConta & ",'" & fxCntX_CuentaFormato(False, txtMapeoCta_Actual.Text, 0) _
        & "'," & gCntX_Parametros.CodigoConta & ",'" & fxCntX_CuentaFormato(False, txtMapeoCta_Nueva.Text, 0) _
        & "','" & glogon.Usuario & "', 1," & chkMapeo_Transac.Value
Call ConectionExecute(strSQL)

Call Bitacora("Aplicar", "Mapeo de Cuentas: " & txtMapeoCta_Actual.Text & " -> " & txtMapeoCta_Nueva.Text)

Me.MousePointer = vbDefault

MsgBox "Se realizó el mapeo de cuenta en los auxiliares!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbBaja_Nivel()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spCntX_Cuentas_Baja_Nivel " & gCntX_Parametros.CodigoConta _
       & ",'" & fxCntX_CuentaFormato(False, txtBN_Cta.Text, 0) _
       & "','" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

Call Bitacora("Aplicar", "Baja Nivel: " & txtBN_Cta.Text & " -> " & fxCntX_CuentaFormato(True, rs!Cuenta, 0))

Me.MousePointer = vbDefault

MsgBox "Baja Nivel de " & txtBN_Cta.Text & " -> " & fxCntX_CuentaFormato(True, rs!Cuenta, 0) & ", satisfactoriamente!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnBajaNivel_Click(Index As Integer)

Select Case Index
    Case 0 'Aplicar
      Call sbBaja_Nivel
      Call sbCatalogo_Consulta
    
    Case 1 'Cancelar
      Call sbTab_Visible(0)
End Select

End Sub

Private Sub btnDetalle_Click(Index As Integer)
Select Case Index
    Case 0 'Aplicar
      Call sbCuenta_DT_Save
    Case 1 'Cancelar
      'NAda
End Select

Call sbTab_Visible(0)

End Sub

Private Sub btnMapeo_Click(Index As Integer)

Select Case Index
    Case 0 'Aplicar
      Call sbMapeo_Cuentas
    Case 1 'Cancelar
      'NAda
End Select

Call sbTab_Visible(0)


End Sub

Private Sub btnToolAux_Click(Index As Integer)
Dim strSQL As String, i As Integer, frm As Form

On Error GoTo vError

Select Case Index
    Case 0 'Historico
        GLOBALES.gTag = mCuenta
        
        Call sbFormActivo("frmCntX_CuentaHistorico", frm)
        
        
        If frm Is Nothing Then
            Call sbFormsCall("frmCntX_CuentaHistorico", , , , False, Me)
        Else
            If frm.Name = "frmCntX_CuentaHistorico" Then
               frm.sbConsultaExterna (GLOBALES.gTag)
               frm.SetFocus
            Else
                Call sbFormsCall("frmCntX_CuentaHistorico", , , , False, Me)
            End If
            
        End If
    
    
    Case 1 'Formato
        Call sbFormatoCtasCatalogoActualiza
        
    Case 2 'Revisa
        'Call sbFormsCall("frmCntX_UtilVerificaAsientos", vbModal, , , False, Me)
        i = MsgBox("Esta seguro que desea revisar la Estructura Contable + Balance de comprobación por inconsistencias?", vbYesNo)
        If i = vbYes Then
           Set frm = frmCntX_Procesos
           Call sbCntX_RestructuraMovimientosRSM(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes, frm, , 1)
        End If
    
    Case 3 'Bajar Nivel
      Call sbTab_Visible(2)
      Call Form_Resize
    
    Case 4 'Mapeo
      Call sbTab_Visible(1)
      Call Form_Resize
    
    Case 5 'Exportar
        Dim vHeaders As vGridHeaders
            vHeaders.Columnas = 10
            vHeaders.Headers(1) = "Det."
            vHeaders.Headers(2) = "Hst."
            vHeaders.Headers(3) = "Cuenta"
            vHeaders.Headers(4) = "Descripción"
            vHeaders.Headers(5) = "Divisa"
            vHeaders.Headers(6) = "Tipo Cuenta"
            vHeaders.Headers(7) = "Acepta.Mov?"
            vHeaders.Headers(8) = "Presupuesta?"
            vHeaders.Headers(9) = "Bloqueada?"
            vHeaders.Headers(10) = "Auxiliar?"

        
         Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Catalogo_Cuentas")
    
    Case 6 'Reporte
       strSQL = "{CntX_Tipos_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta
       Call sbCntX_Reportes("CATALOGO", strSQL, "Detalle")
       
End Select

Exit Sub

vError:

End Sub

Private Sub chkBalance_Click()
Call TimerInicio_Timer
End Sub

Private Sub FlatScrollBarX_Change()
txtNivel = FlatScrollBarX.Value
Call sbCatalogo_Consulta
End Sub

Private Sub sbTab_Visible(pTab As Integer)
Dim i As Integer

For i = 0 To tcMain.ItemCount - 1
   tcMain.Item(i).Visible = False
Next i

tcMain.Item(pTab).Visible = True
tcMain.Item(pTab).Selected = True

Me.Refresh

End Sub


Private Sub Form_Activate()
vModulo = 20

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 20

tcMain.Item(0).Visible = False
tcMain.Item(1).Visible = False
tcMain.Item(2).Visible = False


'Carga Listado de Tipos de Cuentas
mTipoCuentaUltimaSel = ""
strSQL = "select rtrim(tipo_Cuenta) + ' - ' + descripcion as 'TipoCuenta' from CntX_Tipos_Cuentas where cod_contabilidad = " _
       & gCntX_Parametros.CodigoConta & " order by tipo_cuenta"
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.EOF And mTipoCuentaUltimaSel = "" Then mTipoCuentaUltimaSel = rs!TipoCuenta

mTipoCuentaLista = ""
Do While Not rs.EOF
  If Len(mTipoCuentaLista) = 0 Then
    mTipoCuentaLista = Chr$(9) & rs!TipoCuenta
  Else
    mTipoCuentaLista = mTipoCuentaLista & Chr$(9) & rs!TipoCuenta
  End If
  rs.MoveNext
Loop
rs.Close

'Carga Listado de Divisas
mDivisaUltimaSel = ""
strSQL = "select rtrim(cod_divisa) as 'Cod_Divisa' from CntX_Divisas where cod_contabilidad = " _
       & gCntX_Parametros.CodigoConta & " order by divisa_local desc"
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.EOF And mDivisaUltimaSel = "" Then mDivisaUltimaSel = rs!cod_Divisa

mDivisaLista = ""

Do While Not rs.EOF
  If Len(mDivisaLista) = 0 Then
    mDivisaLista = Chr$(9) & rs!cod_Divisa
  Else
    mDivisaLista = mDivisaLista & Chr$(9) & rs!cod_Divisa
  End If
  rs.MoveNext
Loop
rs.Close


strSQL = "select rtrim(cod_divisa) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from CntX_Divisas where cod_contabilidad = " _
       & gCntX_Parametros.CodigoConta & " order by divisa_local desc"

Call sbCbo_Llena_New(cboDivisa, strSQL, True, True)

vGrid.AppearanceStyle = fxGridStyle


Call sbToolBarIconos(tlb)

Call Formularios(Me)
Call RefrescaTags(Me)

If tlb.Buttons(1).Enabled = False Then vGrid.Enabled = False

End Sub

Private Sub sbCargaComboTiposCuenta(vCol As Integer, vRow As Long, vGrid As Object)

vGrid.Col = vCol
vGrid.Row = vRow
vGrid.CellType = CellTypeComboBox

vGrid.TypeComboBoxList = mTipoCuentaLista
vGrid.TypeComboBoxEditable = False
vGrid.Text = mTipoCuentaUltimaSel

End Sub

Private Sub sbCargaComboDivisas(vCol As Integer, vRow As Long, vGrid As Object)

vGrid.Col = vCol
vGrid.Row = vRow
vGrid.CellType = CellTypeComboBox

vGrid.TypeComboBoxList = mDivisaLista
vGrid.TypeComboBoxEditable = False
vGrid.Text = mDivisaUltimaSel

End Sub


Private Function fxCntX_TiposCuentas(strTipo As String) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select rtrim(tipo_Cuenta) + ' - ' + descripcion as 'TipoCuenta'from CntX_Tipos_Cuentas where cod_contabilidad = " _
       & gCntX_Parametros.CodigoConta & " and tipo_cuenta = '" & strTipo & "'"
Call OpenRecordSet(rsX, strSQL, 0)
If Not rsX.EOF And Not rsX.BOF Then
    fxCntX_TiposCuentas = rsX!TipoCuenta
Else
    fxCntX_TiposCuentas = ""
End If
rsX.Close

End Function

Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String
Dim strResMoneda As String, vNota As String

Me.MousePointer = vbHourglass

vPaso = True

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1


vGrid.Row = vGrid.MaxRows

PrgBarX.Visible = True

If chkBalance.Value = vbChecked Then
    Call OpenRecordSet(rs, strSQL, 0)
Else
    Call OpenRecordSet(rs, strSQL, 0)
    PrgBarX.Value = 1
    PrgBarX.Max = rs.RecordCount + 1
End If



Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  vGrid.Col = 5 '3
  vGrid.CellType = 8
  vGrid.TypeComboBoxList = mDivisaLista
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = mDivisaUltimaSel
  
  vGrid.Col = 6 '4
  vGrid.CellType = 8
  vGrid.TypeComboBoxList = mTipoCuentaLista
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = mTipoCuentaUltimaSel
    
  
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
     Case 3
        If chkBalance.Value = vbChecked Then
                vNota = "Estado del Periodo:" & vbCrLf _
                      & "___________________" & vbCrLf _
                      & " Saldo Inicial : " & Format(rs!SALDO_INICIAL, "Standard") & vbCrLf _
                      & " Total Debitos : " & Format(Abs(rs!total_debitos), "Standard") & vbCrLf _
                      & " Total Creditos: " & Format(Abs(rs!total_creditos), "Standard") & vbCrLf _
                      & " Mensual       : " & Format(rs!total_debitos + rs!total_creditos, "Standard") & vbCrLf _
                      & " Acumulado     : " & Format((rs!SALDO_INICIAL + rs!total_debitos + rs!total_creditos), "Standard") & vbCrLf _
                      & "___________________"
                
                vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                vGrid.CellNote = vNota
                vGrid.TextTip = TextTipFixed
                vGrid.TextTipDelay = 1000
        End If
        
'        vGrid.Text = fxCntX_CuentaFormato(True, CStr(rs.Fields(I - 1).Value))
        
        vGrid.Text = CStr(rs.Fields(i - 1).Value & "")

     Case Else
        vGrid.Text = CStr(rs.Fields(i - 1).Value & "")
    End Select
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  
   
  rs.MoveNext
    
    If chkBalance.Value = vbUnchecked Then
        If PrgBarX.Value < PrgBarX.Max Then
            PrgBarX.Value = PrgBarX.Value + 1
        End If
    End If


Loop

rs.Close

  vGrid.Row = vGrid.MaxRows
  
  vGrid.Col = 5 '3
  vGrid.CellType = 8
  vGrid.TypeComboBoxList = mDivisaLista
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = mDivisaUltimaSel
    
  vGrid.Col = 6 '4
  vGrid.CellType = 8
  vGrid.TypeComboBoxList = mTipoCuentaLista
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = mTipoCuentaUltimaSel


PrgBarX.Visible = False
Me.MousePointer = vbDefault

vPaso = False

End Sub


Private Function fxCuentaMadreAux1Rellena(TemCuenta As String, iNivel As Integer) As String
Dim i As Integer


With gCntX_Parametros
  Select Case iNivel
    Case 2
       TemCuenta = Mid(TemCuenta, 1, .Nivel1)
       i = Val(TemCuenta)
    Case 3
       TemCuenta = Mid(TemCuenta, 1, .Nivel1 + .Nivel2)
       i = Val(Right(TemCuenta, .Nivel2))
    Case 4
       TemCuenta = Mid(TemCuenta, 1, .Nivel1 + .Nivel2 + .Nivel3)
       i = Val(Right(TemCuenta, .Nivel3))
    Case 5
       TemCuenta = Mid(TemCuenta, 1, .Nivel1 + .Nivel2 + .Nivel3 + .Nivel4)
       i = Val(Right(TemCuenta, .Nivel4))
  End Select
End With

If i > 0 Then
    fxCuentaMadreAux1Rellena = fxCntX_CuentaFormato(False, TemCuenta)
Else
    fxCuentaMadreAux1Rellena = ""
End If


End Function

Private Function fxCuentaMadre(vCuenta As String, vTipoCuenta As String) As String
Dim rsX As New ADODB.Recordset, strSQL As String
'verificar primero el nivel de la cuenta
'si es nivel 1 no tiene cuenta madre
'si no es nivel 1 verificar si el nivel anterior existe una cuenta madre

'Nivel5

fxCuentaMadre = ""
blnCuentaCorrecta = True
With gCntX_Parametros
    
    'VERIFICA NIVEL 5


    If Val(Right(vCuenta, .Nivel5)) > 0 Then
      
      strSQL = "select cod_cuenta from CntX_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
             & " and acepta_movimientos = 0 and cod_cuenta = '" _
             & fxCuentaMadreAux1Rellena(vCuenta, 5) & "' and tipo_cuenta = '" & vTipoCuenta & "'"
      Call OpenRecordSet(rsX, strSQL, 0)
      
      If rsX.EOF And rsX.BOF Then
         MsgBox " No Existe Nivel Superior para esta cuenta ..." _
                & vbCrLf & " MOTIVOS " & vbCrLf & vbCrLf _
                & " - La cuenta Mayor Existe pero acepta movimientos " & vbCrLf _
                & " - No existe cuenta mayor de subnivel anterior " & vbCrLf _
                & " - La cuenta Mayor Existe pero pertenece a otro Tipo de Cuenta " _
                , vbCritical
        blnCuentaCorrecta = False
      Else
         fxCuentaMadre = rsX!cod_cuenta
      End If
      
      rsX.Close
      Exit Function
    
    End If
    
    'VERIFICA NIVEL 4
    
    If Val(Mid(vCuenta, .Nivel1 + .Nivel2 + .Nivel3 + 1, .Nivel4)) > 0 Then
    
      strSQL = "select cod_cuenta from CntX_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
             & " and acepta_movimientos = 0 and cod_cuenta = '" _
             & fxCuentaMadreAux1Rellena(vCuenta, 4) & "' and tipo_cuenta = '" & vTipoCuenta & "'"
      Call OpenRecordSet(rsX, strSQL, 0)
      
      If rsX.EOF And rsX.BOF Then
         MsgBox " No Existe Nivel Superior para esta cuenta ..." _
                & vbCrLf & " MOTIVOS " & vbCrLf & vbCrLf _
                & " - La cuenta Mayor Existe pero acepta movimientos " & vbCrLf _
                & " - No existe cuenta mayor de subnivel anterior " & vbCrLf _
                & " - La cuenta Mayor Existe pero pertenece a otro Tipo de Cuenta " _
                , vbCritical
        blnCuentaCorrecta = False
      Else
         fxCuentaMadre = rsX!cod_cuenta
      End If
      
      rsX.Close
      Exit Function
    
    End If
    
    'VERIFICA NIVEL 3
    
    If Val(Mid(vCuenta, .Nivel1 + .Nivel2 + 1, .Nivel3)) > 0 Then
    
      strSQL = "select cod_cuenta from CntX_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
             & " and acepta_movimientos = 0 and cod_cuenta = '" _
             & fxCuentaMadreAux1Rellena(vCuenta, 3) & "' and tipo_cuenta = '" & vTipoCuenta & "'"
      Call OpenRecordSet(rsX, strSQL, 0)
      
      If rsX.EOF And rsX.BOF Then
         MsgBox " No Existe Nivel Superior para esta cuenta ..." _
                & vbCrLf & " MOTIVOS " & vbCrLf & vbCrLf _
                & " - La cuenta Mayor Existe pero acepta movimientos " & vbCrLf _
                & " - No existe cuenta mayor de subnivel anterior " & vbCrLf _
                & " - La cuenta Mayor Existe pero pertenece a otro Tipo de Cuenta " _
                , vbCritical
        blnCuentaCorrecta = False
      Else
         fxCuentaMadre = rsX!cod_cuenta
      End If
      
      rsX.Close
      Exit Function
    
    End If
    
    'VERIFICA NIVEL 2
    
    If Val(Mid(vCuenta, .Nivel1 + 1, .Nivel2)) > 0 Then
    
      strSQL = "select cod_cuenta from CntX_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
             & " and acepta_movimientos = 0 and cod_cuenta = '" _
             & fxCuentaMadreAux1Rellena(vCuenta, 2) & "' and tipo_cuenta = '" & vTipoCuenta & "'"
      Call OpenRecordSet(rsX, strSQL, 0)
      
      If rsX.EOF And rsX.BOF Then
         MsgBox " No Existe Nivel Superior para esta cuenta ..." _
                & vbCrLf & " MOTIVOS " & vbCrLf & vbCrLf _
                & " - La cuenta Mayor Existe pero acepta movimientos " & vbCrLf _
                & " - No existe cuenta mayor de subnivel anterior " & vbCrLf _
                & " - La cuenta Mayor Existe pero pertenece a otro Tipo de Cuenta " _
                , vbCritical
        blnCuentaCorrecta = False
      Else
         fxCuentaMadre = rsX!cod_cuenta
      End If
      
      rsX.Close
      Exit Function
    
    End If
    
    'NIVEL 1 NO SE VERIFICA
    
End With



End Function

Private Function fxVerificaCambio(vTipo As String, vCuenta As String) As Boolean
Dim strSQL As String, rsX As New ADODB.Recordset

'1. Verificar si tiene el mismo grupo de Cuentas
'2. Verificar si se han mayorizado Asientos que la afecten
'3. Verificar si hay Asientos sin mayorizar que la contengan
'4. Verificar si es una cuenta de mayor y tiene sub Cuentas
'

vMensaje = ""

strSQL = "select * from CntX_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_cuenta = '" & Trim(vCuenta) & "'"
Call OpenRecordSet(rsX, strSQL, 0)
If UCase(rsX!tipo_cuenta) <> UCase(vTipo) Then vMensaje = vMensaje & vbCrLf & "- Esta tipo de cuenta no es válido entre el grupo de pertenencia..."
   
rsX.Close
strSQL = "select isnull(sum(total_debitos),0) as TD,isnull(sum(total_creditos),0) as TC" _
       & " from CntX_Mov_Cuentas_Detallado where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_cuenta = '" & Trim(vCuenta) & "'"
Call OpenRecordSet(rsX, strSQL, 0)
If rsX!TD <> 0 Or rsX!TC <> 0 Then
   vMensaje = vMensaje & vbCrLf & "- Esta cuenta tiene movimientos registrados..."
Else
   rsX.Close
   'Buscar Asientos Sin Mayorizar que tengan esta cuenta incluida...
   strSQL = "select isnull(count(*),0) as Existe from Cntx_Asientos A inner join Cntx_Asientos_detalle D" _
          & " on A.tipo_asiento = D.tipo_asiento and A.num_asiento = D.num_asiento" _
          & " and A.cod_contabilidad = D.cod_contabilidad" _
          & " where A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
          & " and cod_cuenta = '" & Trim(vCuenta) & "' and A.fecha_aplicado is null"
   Call OpenRecordSet(rsX, strSQL, 0)
   If rsX!Existe > 0 Then vMensaje = vMensaje & vbCrLf & "- Esta cuenta tiene Cntx_Asientos registrados sin mayorizar..."

End If

rsX.Close
strSQL = "select isnull(count(*),0) as Existe from CntX_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cuenta_madre = '" & Trim(vCuenta) & "'"
Call OpenRecordSet(rsX, strSQL, 0)
If rsX!Existe > 0 Then vMensaje = vMensaje & vbCrLf & "- Esta cuenta tiene Sub-Cuentas registradas ..."
rsX.Close


If Len(vMensaje) > 0 Then
  vMensaje = "   ****** NO SE PUEDE GUARDAR ******" & vbCrLf & vbCrLf & vMensaje
  fxVerificaCambio = False
Else
  fxVerificaCambio = True
End If

End Function


Private Function fxNivelCuenta(vCuentaMadre As String) As Integer

fxNivelCuenta = 0

If vCuentaMadre = "" Then
  fxNivelCuenta = 1
  Exit Function
End If

With gCntX_Parametros
  If .Nivel1 > 0 And Val(Mid(vCuentaMadre, 1, .Nivel1)) > 0 Then fxNivelCuenta = 2
  If .Nivel2 > 0 And Val(Mid(vCuentaMadre, .Nivel1 + 1, .Nivel1)) > 0 Then fxNivelCuenta = 3
  If .Nivel3 > 0 And Val(Mid(vCuentaMadre, .Nivel1 + .Nivel2 + 1, .Nivel3)) > 0 Then fxNivelCuenta = 4
  If .Nivel4 > 0 And Val(Mid(vCuentaMadre, .Nivel1 + .Nivel2 + .Nivel3 + 1, .Nivel4)) > 0 Then fxNivelCuenta = 5
End With

If fxNivelCuenta = 0 Then fxNivelCuenta = 5

' MsgBox fxNivelCuenta

End Function

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

Dim vCuenta As String, vCuentaMask As String
Dim vDescripcion As String, vDivisa As String, vTipoCuenta As String
Dim vAceptaMovimientos As Byte, vPresupuesto As Byte, vBloqueada As Byte, vAuxiliar As Byte

On Error GoTo vError

vGrid.Col = 5
mDivisaUltimaSel = vGrid.Text

vGrid.Col = 6
mTipoCuentaUltimaSel = vGrid.Text

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow

vGrid.Col = 3
vCuenta = fxCntX_CuentaFormato(False, vGrid.Text)

If Len(Trim(vCuenta)) > gCntX_Parametros.TotalChr Then
   MsgBox "La Cuenta Digitada sobrepasa el total de caracteres permitidos por la Mascara Contable Definida para esta compañía...", vbCritical
   fxGuardar = 0
   Exit Function
End If


vGrid.Col = 4
vDescripcion = Trim(vGrid.Text)

vGrid.Col = 5
vDivisa = Trim(vGrid.Text)

vGrid.Col = 6
vTipoCuenta = SIFGlobal.fxCodText(vGrid.Text)

vGrid.Col = 7
vAceptaMovimientos = vGrid.Value

vGrid.Col = 8
vPresupuesto = vGrid.Value

vGrid.Col = 9
vBloqueada = vGrid.Value

vGrid.Col = 10
vAuxiliar = vGrid.Value

' If fxVerificaCambio(SIFGlobal.fxCodText(vGrid.Text), @Cuenta) Then

strSQL = "exec spCntX_Cuentas_Registro " & gCntX_Parametros.CodigoConta _
       & ", '" & vCuenta & "','" & vDescripcion & "','" & vDivisa & "','" & vTipoCuenta _
       & "', " & vAceptaMovimientos & "," & vPresupuesto & "," & vAuxiliar & "," & vBloqueada _
       & ", 'A','" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
  vCuentaMask = rs!COD_CUENTA_MASK
  Call Bitacora(rs!Movimiento, "Cuenta en el Catalogo: " & vCuentaMask)
  
  vGrid.Col = 3
  vGrid.Text = vCuentaMask
  
  If Len(rs!Mensaje) > 0 Then
      MsgBox rs!Mensaje, vbExclamation
  End If
rs.Close

 
fxGuardar = 1


Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub Form_Resize()
On Error Resume Next

tcMain.Width = Me.Width - 250
tcMain.Height = Me.Height - 1820

vGrid.Width = Me.Width - 250
vGrid.Height = tcMain.Height - 350

scMain.Width = Me.Width

gbMapeo.Left = (tcMain.Width - gbMapeo.Width) / 2
gbMapeo.Top = (tcMain.Height - gbMapeo.Height - tcMain.Top) / 2


gbBajarNivel.Left = (tcMain.Width - gbBajarNivel.Width) / 2
gbBajarNivel.Top = (tcMain.Height - gbBajarNivel.Height - tcMain.Top) / 2

tcDetalle.Left = (tcMain.Width - tcDetalle.Width) / 2
tcDetalle.Height = tcMain.Height - (tcDetalle.Top)

btnDetalle(0).Left = (tcDetalle.Left + tcDetalle.Width) - (btnDetalle(0).Width + btnDetalle(1).Width)
btnDetalle(1).Left = btnDetalle(0).Left + btnDetalle(0).Width - 10

gTraduce.Height = tcDetalle.Height - (gTraduce.Top + 350)
gProrratea.Height = tcDetalle.Height - (gProrratea.Top + 550)

lblProrrateoTotal.Top = gProrratea.Top + gProrratea.Height + 150
txtDT_Prorrateo_Total.Top = lblProrrateoTotal.Top

gbToolAux.Left = (tcMain.Left + tcMain.Width - gbToolAux.Width - 80) / 2
gbFiltros.Left = (tcMain.Left + tcMain.Width - gbFiltros.Width - 80) / 2


PrgBarX.Top = scMain.Top + scMain.Height + 20
PrgBarX.Width = scMain.Width


End Sub

Private Sub sbCatalogo_Consulta()
Dim strSQL As String

Call sbTab_Visible(0)

If chkBalance.Value = vbChecked Then
    strSQL = "select '','',C.cod_Cuenta_Mask,space(nivel*2) + ltrim(C.descripcion) as Descripcion,C.cod_divisa,(rtrim(T.tipo_cuenta) + ' - ' + T.descripcion) as Tipo" _
           & ",acepta_movimientos,Presupuesto,bloqueada,cuenta_auxiliar" _
           & ",isnull(M.saldo_inicial,0) as Saldo_Inicial,isnull(M.total_debitos,0) as Total_debitos" _
           & ",isnull(M.total_creditos,0) as Total_creditos" _
           & " from CntX_Cuentas C inner join CntX_Tipos_Cuentas T on T.tipo_cuenta = C.tipo_cuenta and C.cod_contabilidad = T.cod_contabilidad" _
           & " left join vCntX_Mov_Cuentas_General M on C.cod_cuenta = M.cod_cuenta" _
           & " and C.cod_contabilidad = M.cod_contabilidad and " & gCntX_Parametros.PeriodoAnio & " = M.anio" _
           & " and " & gCntX_Parametros.PeriodoMes & " = M.mes" _
           & " where C.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and C.nivel <= " & txtNivel.Text
Else
    strSQL = "select '','',C.cod_Cuenta_Mask,space(nivel*2) + ltrim(C.descripcion) as Descripcion,C.cod_divisa,(rtrim(T.tipo_cuenta) + ' - ' + T.descripcion) as Tipo" _
           & ",acepta_movimientos,Presupuesto,bloqueada,cuenta_auxiliar" _
           & " from CntX_Cuentas C inner join CntX_Tipos_Cuentas T on T.tipo_cuenta = C.tipo_cuenta and C.cod_contabilidad = T.cod_contabilidad" _
           & " where C.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and C.nivel <= " & txtNivel.Text

End If

If cboDivisa.Text <> "TODOS" Then
    strSQL = strSQL & " and C.cod_Divisa = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
End If

If Trim(txtFiltroCuenta(0).Text) <> "" Then strSQL = strSQL & " and C.cod_cuenta_mask like '" & Trim(txtFiltroCuenta(0).Text) & "%'"
If Trim(txtFiltroCuenta(1).Text) <> "" Then strSQL = strSQL & " and C.descripcion like '%" & Trim(txtFiltroCuenta(1).Text) & "%'"


strSQL = strSQL & " order by C.cod_cuenta"
       
Call sbCargaGridLocal(vGrid, 10, strSQL)


End Sub



Private Function fxgTraduce_Guardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim pCuenta As String

On Error GoTo vError

With gTraduce

pCuenta = fxgCntCuentaFormato(False, txtDT_Cuenta.Text, 0)

fxgTraduce_Guardar = 0

.Row = .ActiveRow
.Col = 1

strSQL = "select isnull(count(*),0) as Total from CNTX_CUENTAS_TRADUCCION" _
        & " where COD_IDIOMA = '" & .Text & "' and COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
        & " AND COD_CUENTA = '" & pCuenta & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Total = 0 Then 'Insertar
  strSQL = "insert into CNTX_CUENTAS_TRADUCCION(COD_IDIOMA,COD_CONTABILIDAD,COD_CUENTA, DESCRIPCION,REGISTRO_USUARIO, REGISTRO_FECHA) values('"
  .Col = 1
  strSQL = strSQL & UCase(.Text) & "'," & gCntX_Parametros.CodigoConta & ",'" & pCuenta & "','"
  .Col = 2
  strSQL = strSQL & .Text & "','" & glogon.Usuario & "',dbo.MyGetdate())"
   
  Call ConectionExecute(strSQL)

  .Col = 1
  
 Call Bitacora("Registra", "Cta. Traducción: " & .Text & " Conta." & gCntX_Parametros.CodigoConta _
        & ", Cta: " & fxgCntCuentaFormato(False, txtDT_Cuenta.Text, 0))
  
  fxgTraduce_Guardar = 1

Else 'Actualizar

 .Col = 2
 strSQL = "update CNTX_CUENTAS_TRADUCCION set descripcion = '" & .Text & "'" _
        & " where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
        & " and cod_cuenta = '" & pCuenta & "' and COD_IDIOMA = '"
 .Col = 1
 strSQL = strSQL & .Text & "'"
 Call ConectionExecute(strSQL)
 
 Call Bitacora("Modifica", "Cta. Traducción: " & .Text & " Conta." & gCntX_Parametros.CodigoConta _
        & ", Cta: " & fxgCntCuentaFormato(False, txtDT_Cuenta.Text, 0))

 fxgTraduce_Guardar = 1
End If

rs.Close

End With

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function




Private Function fxgProrrateo_Guardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim pUnidad As String, pCentro As String, pCuenta As String

On Error GoTo vError

fxgProrrateo_Guardar = 0

With gProrratea

    .Row = .ActiveRow
    .Col = 1
    
    pCuenta = fxgCntCuentaFormato(False, txtDT_Cuenta.Text, 0)
    
    .Row = .ActiveRow
    .Col = 1
    
    pUnidad = .Text
    .Col = 3
    pCentro = .Text
            
    
    strSQL = "select isnull(count(*),0) as Total from CNTX_CUENTAS_PRORRATA" _
            & " where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
            & " AND COD_CUENTA = '" & pCuenta & "' AND COD_UNIDAD = '" & pUnidad & "' AND COD_CENTRO_COSTO = '" & pCentro & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If rs!Total = 0 Then 'Insertar
      
          .Col = 5
          strSQL = "insert into CNTX_CUENTAS_PRORRATA(COD_CONTABILIDAD,COD_CUENTA, COD_UNIDAD, COD_CENTRO_COSTO, PORCENTAJE,REGISTRO_USUARIO, REGISTRO_FECHA)" _
                 & "  values(" & gCntX_Parametros.CodigoConta & ",'" & pCuenta & "','" & pUnidad & "','" & pCentro _
                 & "'," & CCur(.Text) & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
           
          Call ConectionExecute(strSQL)
        
          .Col = 2
          
          Call Bitacora("Registra", "Cta. Prorrateo: Conta." & gCntX_Parametros.CodigoConta _
                    & ", Cta: " & pCuenta & ", Unidad: " & pUnidad & ", Centro: " & pCentro)
          
          fxgProrrateo_Guardar = 1
    
    Else 'Actualizar
    
         .Col = 5
         strSQL = "update CNTX_CUENTAS_PRORRATA set PORCENTAJE = " & CCur(.Text) _
                & " where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
                & " and cod_cuenta = '" & pCuenta _
                & "' and COD_UNIDAD = '" & pUnidad & "' and COD_CENTRO_COSTO = '" & pCentro & "'"
         Call ConectionExecute(strSQL)
         
                
         Call Bitacora("Modifica", "Cta. Prorrateo: Conta." & gCntX_Parametros.CodigoConta _
                    & ", Cta: " & pCuenta & ", Unidad: " & pUnidad & ", Centro: " & pCentro)
        
         fxgProrrateo_Guardar = 1
    
    End If
    
    rs.Close

End With

'Calcula Totales
Call sbProrratea_Total

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function




Private Sub gProrratea_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String
Dim pUnidad As String, pCentro As String, pCuenta As String

On Error GoTo vError

With gProrratea
    
    If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
      i = fxgProrrateo_Guardar
      .Row = .ActiveRow
      .Col = 1
      If .MaxRows <= .ActiveRow Then
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
      End If
    End If
    
    
    'Consulta Unidad
    If KeyCode = vbKeyF4 And .ActiveCol = 1 Then
      gBusquedas.Resultado = ""
      gBusquedas.Resultado2 = ""
      gBusquedas.Columna = "descripcion"
      gBusquedas.Orden = "descripcion"
      gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
      gBusquedas.Consulta = "select cod_unidad,descripcion from CntX_Unidades"
      frmBusquedas.Show vbModal
        
      .Col = .ActiveCol
      .Row = .ActiveRow
      
      .Text = gBusquedas.Resultado
      .Col = 2
      .Text = gBusquedas.Resultado2
    End If
    
    
    'Consulta Centro de Costo
    If KeyCode = vbKeyF4 And .ActiveCol = 3 Then
      .Row = .ActiveRow
      .Col = 1
      pUnidad = .Text
      
      gBusquedas.Resultado = ""
      gBusquedas.Resultado2 = ""
      gBusquedas.Columna = "descripcion"
      gBusquedas.Orden = "descripcion"
      gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta & " and cod_centro_costo in(select cod_centro_costo" _
                        & " from cntX_unidades_cc where cod_unidad = '" & pUnidad & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta & ")"
      gBusquedas.Consulta = "select cod_centro_costo,descripcion from CntX_Centro_Costos"
      frmBusquedas.Show vbModal
        
      .Col = .ActiveCol
      .Row = .ActiveRow
      
      .Text = gBusquedas.Resultado
      .Col = 4
      .Text = gBusquedas.Resultado2
      
    End If
    
    
    'Inserta Linea
    If KeyCode = vbKeyInsert Then
        .MaxRows = .MaxRows + 1
        .InsertRows .ActiveRow, 1
        .Row = .ActiveRow
    End If
    
    
    'Borrar una linea
    If KeyCode = vbKeyDelete Then
         i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
         If i = vbYes Then
            
            pCuenta = fxgCntCuentaFormato(False, txtDT_Cuenta.Text, 0)
            
            .Row = .ActiveRow
            .Col = 1
            
            pUnidad = .Text
            .Col = 3
            pCentro = .Text
            
            strSQL = "delete CNTX_CUENTAS_PRORRATA where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
                   & " and cod_cuenta = '" & pCuenta _
                   & "' and COD_UNIDAD = '" & pUnidad & "' AND COD_CENTRO_COSTO = '" & pCentro & "'"
            Call ConectionExecute(strSQL, 0)
            
            
            
            Call Bitacora("Elimina", "Cta. Prorrateo: Conta." & gCntX_Parametros.CodigoConta _
                    & ", Cta: " & pCuenta & ", Unidad: " & pUnidad & ", Centro: " & pCentro)
            
            
            .DeleteRows .ActiveRow, 1
            .MaxRows = .MaxRows - 1
            If .MaxRows = 0 Then .MaxRows = 1
            
            
            Call sbProrratea_Total
         
         End If
    End If
    

    'Consulta Codigo al Avanzar
    If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
        .Col = .ActiveCol
        .Row = .ActiveRow
        
        Select Case .ActiveCol
          Case 1
            'Buscar la Unidad
            If fxCntx_UnidadVerifica(.Text) Then
              .Col = 2
              .Text = fxCntX_Unidad("D", .Text)
            Else
              MsgBox "La UNIDAD de negocio: " & .Text & ", no es válida : " & vbCrLf & " - No Existe...", vbCritical
            End If
          
          
          Case 3 'Verificar el Centro de Costo
            .Col = 1
            pUnidad = .Text
            .Col = 3
            
            If fxCntX_CentroCostoVerifica(.Text, pUnidad) Then
              .Col = 4
              .Text = fxCntX_CentroCosto("D", .Text)
            Else
              MsgBox "El Centro de Costo no es válido y no puede ser utilizada por esta unidad: " & vbCrLf & " - No Existe...", vbCritical
            End If
          
          
          Case 5 'Porcentaje
            If IsNumeric(.Text) > 0 Then
                Call sbProrratea_Total
            End If
          
        End Select
    
    End If


End With

Exit Sub


vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub gTraduce_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

With gTraduce

If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxgTraduce_Guardar
  .Row = .ActiveRow
  .Col = 1
  If .MaxRows <= .ActiveRow Then
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows
  End If
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    .MaxRows = .MaxRows + 1
    .InsertRows .ActiveRow, 1
    .Row = .ActiveRow
End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        .Row = .ActiveRow
        .Col = 1
        
        .Row = .ActiveRow
        .Col = 1
        strSQL = "delete CNTX_CUENTAS_TRADUCCION where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
               & " and cod_cuenta = '" & fxgCntCuentaFormato(False, txtDT_Cuenta.Text, 0) _
               & "' and cod_idioma = '" & .Text & "'"
        Call ConectionExecute(strSQL, 0)
        
        Call Bitacora("Elimina", "Cta. Traducción: " & .Text & " Conta." & gCntX_Parametros.CodigoConta _
                & ", Cta: " & fxgCntCuentaFormato(False, txtDT_Cuenta.Text, 0))
        
        
        .DeleteRows .ActiveRow, 1
        .MaxRows = .MaxRows - 1
        If .MaxRows = 0 Then .MaxRows = 1
     
     End If
End If


End With

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerInicio_Timer()

TimerInicio.Enabled = False
TimerInicio.Interval = 0

'Inicializa
Dim strSQL As String
strSQL = "select rtrim(cod_unidad) as 'IdX', rtrim(descripcion) as ItmX" _
       & " from CntX_Unidades where cod_contabilidad = " & gCntX_Parametros.CodigoConta

Call sbCbo_Llena_New(cboDT_Exclusiva_Unidad, strSQL, False, True)

Call sbCbo_Copia(cboDT_Exclusiva_Unidad, cboDT_DC_Unidad)
Call sbCbo_Copia(cboDT_Exclusiva_Unidad, cboDT_Prorratea_Unidad)

cboDT_Prorratea_Unidad.AddItem "[CONSOLIDADO]"
cboDT_Prorratea_Unidad.ItemData(cboDT_Prorratea_Unidad.ListCount - 1) = ""

strSQL = "select COD_CENTRO_COSTO AS 'IdX', RTRIM(DESCRIPCION) AS 'ItmX'" _
       & " From CNTX_CENTRO_COSTOS" _
       & " Where Activo = 1 And COD_CONTABILIDAD = 1"
Call sbCbo_Llena_New(cboDT_Exclusiva_Centro, strSQL, False, True)

cboDT_Exclusiva_Centro.AddItem "[No Indica]"
cboDT_Exclusiva_Centro.ItemData(cboDT_Exclusiva_Centro.ListCount - 1) = ""

Call sbCbo_Copia(cboDT_Exclusiva_Centro, cboDT_DC_Centro)
Call sbCbo_Copia(cboDT_Exclusiva_Centro, cboDT_Prorratea_Centro)


cboDT_Exclusiva_Unidad.Text = cboDT_Exclusiva_Unidad.List(0)


Call sbCatalogo_Consulta

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String
     
On Error GoTo vError
     
i = MsgBox("Esta Seguro que desea borrar esta cuenta , puede afectarse todas sus Cuentas", vbYesNo)
If i = vbYes Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 3
   strSQL = "delete CntX_Cuentas where cuenta_madre = '" & fxCntX_CuentaFormato(False, vGrid.Text) _
          & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta
   Call ConectionExecute(strSQL, 0)
           
   strSQL = "delete CntX_Cuentas where cod_cuenta = '" & fxCntX_CuentaFormato(False, vGrid.Text) _
          & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta
   Call ConectionExecute(strSQL, 0)
   strSQL = vGrid.Text
   vGrid.Col = 4
       Call Bitacora("Elimina", "Cuenta : " & vGrid.Text & "- COD : " & strSQL)
   vGrid.Col = 3
   
   vGrid.DeleteRows vGrid.Row, 1
   vGrid.MaxRows = vGrid.MaxRows - 1
   
   If vGrid.MaxRows <= 0 Then
        vGrid.MaxRows = 1
        Call sbCargaComboTiposCuenta(6, vGrid.MaxRows, vGrid)
        Call sbCargaComboDivisas(5, vGrid.MaxRows, vGrid)
   End If

End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer, strSQL As String

On Error Resume Next

Select Case UCase(Button.Key)
  Case "NUEVO"
    vGrid.Row = vGrid.MaxRows
    vGrid.Col = 1
    If vGrid.Text <> "" Then
        vGrid.MaxRows = vGrid.MaxRows + 1
        Call sbCargaComboTiposCuenta(6, vGrid.MaxRows, vGrid)
        Call sbCargaComboDivisas(5, vGrid.MaxRows, vGrid)
    End If
    
  Case "BORRAR"
    Call sbBorrar
    
  Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp


End Select

End Sub



Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim strSQL As String
Select Case ButtonMenu.Key
 Case "repResumen"
   strSQL = "{CntX_Tipos_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
          & " AND {CntX_Cuentas.CUENTA_MADRE} =''"
   Call sbCntX_Reportes("CATALOGO", strSQL, "Resumen")
 Case "repDetalle"
   strSQL = "{CntX_Tipos_Cuentas.cod_contabilidad} = " & gCntX_Parametros.CodigoConta
   Call sbCntX_Reportes("CATALOGO", strSQL, "Detalle")
End Select

End Sub

Private Sub sbFormatoCtasCatalogoActualiza()
Dim strSQL As String

Me.MousePointer = vbHourglass

strSQL = "exec spCntX_Catalogo_Cuenta_Mask " & gCntX_Parametros.CodigoConta
 
'Call ConectionExecute(strSQL)
 
Me.MousePointer = vbDefault
MsgBox "Formato de Cuentas actualizado!", vbInformation


End Sub



Private Sub txtBN_Cta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtBN_Cta_Desc.SetFocus
    
If KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  txtBN_Cta_Desc.Text = fxCntX_Cuenta("D", gCuenta)
  txtBN_Cta.Text = fxCntX_CuentaFormato(True, gCuenta)
End If

End Sub

Private Sub txtBN_Cta_LostFocus()
On Error GoTo vError
  
  gCuenta = fxCntX_CuentaFormato(False, txtBN_Cta, 0)
  
  txtBN_Cta_Desc.Text = fxCntX_Cuenta("D", gCuenta)
  txtBN_Cta.Text = fxCntX_CuentaFormato(True, gCuenta)

Exit Sub
vError:

End Sub


Private Sub txtDT_DC_Cuenta_GST_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDT_DC_Cuenta_GST_Desc.SetFocus
    
If KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  txtDT_DC_Cuenta_GST_Desc.Text = fxCntX_Cuenta("D", gCuenta)
  txtDT_DC_Cuenta_GST.Text = fxCntX_CuentaFormato(True, gCuenta)
End If
End Sub

Private Sub txtDT_DC_Cuenta_GST_LostFocus()
On Error GoTo vError
  
  gCuenta = fxCntX_CuentaFormato(False, txtDT_DC_Cuenta_GST, 0)
  
  txtDT_DC_Cuenta_GST_Desc.Text = fxCntX_Cuenta("D", gCuenta)
  txtDT_DC_Cuenta_GST.Text = fxCntX_CuentaFormato(True, gCuenta)

Exit Sub
vError:
End Sub


Private Sub txtDT_DC_Cuenta_ING_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDT_DC_Cuenta_ING_Desc.SetFocus
    
If KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  txtDT_DC_Cuenta_ING_Desc.Text = fxCntX_Cuenta("D", gCuenta)
  txtDT_DC_Cuenta_ING.Text = fxCntX_CuentaFormato(True, gCuenta)
End If
End Sub

Private Sub txtDT_DC_Cuenta_ING_LostFocus()
On Error GoTo vError
  
  gCuenta = fxCntX_CuentaFormato(False, txtDT_DC_Cuenta_ING, 0)
  
  txtDT_DC_Cuenta_ING_Desc.Text = fxCntX_Cuenta("D", gCuenta)
  txtDT_DC_Cuenta_ING.Text = fxCntX_CuentaFormato(True, gCuenta)

Exit Sub
vError:
End Sub

Private Sub txtFiltroCuenta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then Call sbCatalogo_Consulta

End Sub



Private Sub txtMapeoCta_Actual_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMapeoCta_Actual_Desc.SetFocus
    
If KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  txtMapeoCta_Actual_Desc.Text = fxCntX_Cuenta("D", gCuenta)
  txtMapeoCta_Actual.Text = fxCntX_CuentaFormato(True, gCuenta)
End If

End Sub



Private Sub txtMapeoCta_Actual_LostFocus()
On Error GoTo vError
  
  gCuenta = fxCntX_CuentaFormato(False, txtMapeoCta_Actual, 0)
  
  txtMapeoCta_Actual_Desc.Text = fxCntX_Cuenta("D", gCuenta)
  txtMapeoCta_Actual.Text = fxCntX_CuentaFormato(True, gCuenta)

Exit Sub
vError:

End Sub

Private Sub txtMapeoCta_Nueva_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMapeoCta_Nueva_Desc.SetFocus
    
If KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  txtMapeoCta_Nueva_Desc.Text = fxCntX_Cuenta("D", gCuenta)
  txtMapeoCta_Nueva.Text = fxCntX_CuentaFormato(True, gCuenta)
End If

End Sub

Private Sub txtMapeoCta_Nueva_LostFocus()
On Error GoTo vError
  
  gCuenta = fxCntX_CuentaFormato(False, txtMapeoCta_Nueva, 0)
  
  txtMapeoCta_Nueva_Desc.Text = fxCntX_Cuenta("D", gCuenta)
  txtMapeoCta_Nueva.Text = fxCntX_CuentaFormato(True, gCuenta)

Exit Sub
vError:

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCuenta As String, frm As Form

If vPaso Then Exit Sub

vGrid.Row = Row
vGrid.Col = 3


vCuenta = fxCntX_CuentaFormato(False, vGrid.Text)

If Trim(vCuenta) = "" Then
   Exit Sub
End If

vGrid.Col = Col

strSQL = "select isnull(count(*),0) as Total from CntX_Cuentas where cod_cuenta = '" _
        & vCuenta & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta
        
Call OpenRecordSet(rs, strSQL)
If rs!Total = 0 Then
   rs.Close
   Exit Sub
End If


Select Case Col
  Case 1 'Detalles
    GLOBALES.gTag = vCuenta
    Call sbTab_Visible(3)
    Call Form_Resize
    Call sbCuenta_DT_Load(vCuenta)
    
  Case 2 'Histórico
    GLOBALES.gTag = vCuenta
    
    Call sbFormActivo("frmCntX_CuentaHistorico", frm)
    
    
    If frm Is Nothing Then
        Call sbFormsCall("frmCntX_CuentaHistorico", , , , False, Me)
    Else
        If frm.Name = "frmCntX_CuentaHistorico" Then
           frm.sbConsultaExterna (GLOBALES.gTag)
           frm.SetFocus
        Else
            Call sbFormsCall("frmCntX_CuentaHistorico", , , , False, Me)
        End If
        
    End If

  Case 7 'Acepta Movimientos
    'Valida cuenta para movimientos
    vGrid.Col = 6
    If fxVerificaCambio(SIFGlobal.fxCodText(vGrid.Text), vCuenta) Then
        vGrid.Col = Col
        strSQL = "update CntX_Cuentas set acepta_movimientos = " & vGrid.Value _
               & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
               & " and cod_cuenta = '" & vCuenta & "'"
        Call ConectionExecute(strSQL, 0)
    Else
       MsgBox "No se puede cambiar el estado de movimientos porque ya tiene líneas de Asientos registradas...!", vbExclamation
    End If
    
  Case 8 'Presupuesto
    strSQL = "update CntX_Cuentas set presupuesto = " & vGrid.Value _
           & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and cod_cuenta = '" & vCuenta & "'"
    Call ConectionExecute(strSQL, 0)
  
  Case 9 'Bloqueado
    strSQL = "update CntX_Cuentas set bloqueada = " & vGrid.Value _
           & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and cod_cuenta = '" & vCuenta & "'"
    Call ConectionExecute(strSQL, 0)
    
    
    
  Case 10 'Cuenta de Auxiliar
    strSQL = "update CntX_Cuentas set cuenta_auxiliar = " & vGrid.Value _
           & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and cod_cuenta = '" & vCuenta & "'"
    Call ConectionExecute(strSQL, 0)
End Select


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 3
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    Call sbCargaComboDivisas(5, vGrid.MaxRows, vGrid)
    Call sbCargaComboTiposCuenta(6, vGrid.MaxRows, vGrid)
  End If
End If

If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    Call sbCargaComboDivisas(5, vGrid.ActiveRow, vGrid)
    Call sbCargaComboTiposCuenta(6, vGrid.ActiveRow, vGrid)
End If

If vGrid.ActiveCol = 3 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = fxCntX_CuentaFormato(True, vGrid.Text)
End If

End Sub



Private Sub vGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

If vPaso Then Exit Sub

On Error GoTo vError

vGrid.Row = Row
vGrid.Col = 3

mCuenta = fxCntX_CuentaFormato(False, vGrid.Text)

Exit Sub
vError:


End Sub

Private Sub vGrid_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)

If vPaso Then Exit Sub

On Error GoTo vError

vGrid.Row = Row
vGrid.Col = 3

mCuenta = fxCntX_CuentaFormato(False, vGrid.Text)

Exit Sub
vError:

End Sub
