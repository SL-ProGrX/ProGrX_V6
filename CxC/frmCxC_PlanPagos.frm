VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCxC_PlanPagos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "CxC: Plan de Pagos"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   18000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   1935
      Left            =   240
      TabIndex        =   39
      Top             =   4680
      Width           =   12255
      _Version        =   1572864
      _ExtentX        =   21616
      _ExtentY        =   3413
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
      Appearance      =   17
   End
   Begin VB.Timer TimerX 
      Interval        =   20
      Left            =   6960
      Top             =   360
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6780
      Width           =   18000
      _ExtentX        =   31750
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Text            =   "Cuotas."
            TextSave        =   "Cuotas."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Text            =   "Inicio"
            TextSave        =   "Inicio"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Text            =   "Corte"
            TextSave        =   "Corte"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Text            =   "Dias"
            TextSave        =   "Dias"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Intereses"
            TextSave        =   "Intereses"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Cargos"
            TextSave        =   "Cargos"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Dias Mora"
            TextSave        =   "Dias Mora"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4095
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   12255
      _Version        =   524288
      _ExtentX        =   21616
      _ExtentY        =   7223
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      DisplayRowHeaders=   0   'False
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
      MaxCols         =   22
      SpreadDesigner  =   "frmCxC_PlanPagos.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   960
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2773
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
   Begin XtremeSuiteControls.FlatEdit txtLinea 
      Height          =   315
      Left            =   1680
      TabIndex        =   10
      Top             =   1320
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2773
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   3240
      TabIndex        =   11
      Top             =   960
      Width           =   4335
      _Version        =   1572864
      _ExtentX        =   7641
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
   Begin XtremeSuiteControls.FlatEdit txtLineaDesc 
      Height          =   315
      Left            =   3240
      TabIndex        =   12
      Top             =   1320
      Width           =   4335
      _Version        =   1572864
      _ExtentX        =   7641
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
   Begin XtremeSuiteControls.GroupBox gbTotales 
      Height          =   615
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   17175
      _Version        =   1572864
      _ExtentX        =   30295
      _ExtentY        =   1085
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkCargos 
         Height          =   255
         Left            =   15360
         TabIndex        =   28
         Top             =   240
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cargos"
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
      Begin XtremeSuiteControls.FlatEdit txtPlazo 
         Height          =   315
         Left            =   10680
         TabIndex        =   14
         Top             =   240
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.FlatEdit txtCuota 
         Height          =   312
         Left            =   6480
         TabIndex        =   15
         Top             =   240
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   312
         Left            =   1440
         TabIndex        =   16
         Top             =   240
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
      Begin XtremeSuiteControls.FlatEdit txtSaldo 
         Height          =   312
         Left            =   3960
         TabIndex        =   17
         Top             =   240
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
      Begin XtremeSuiteControls.FlatEdit txtTasa 
         Height          =   315
         Left            =   9000
         TabIndex        =   18
         Top             =   240
         Width           =   735
         _Version        =   1572864
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   315
         Left            =   13320
         TabIndex        =   24
         Top             =   240
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   556
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Documento"
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
         Index           =   6
         Left            =   12120
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tasa"
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
         Index           =   7
         Left            =   8280
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
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
         Left            =   3240
         TabIndex        =   22
         Top             =   240
         Width           =   852
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
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
         Left            =   0
         TabIndex        =   21
         Top             =   240
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuota"
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
         Left            =   5760
         TabIndex        =   20
         Top             =   240
         Width           =   852
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo"
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
         Index           =   9
         Left            =   9960
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtPagador 
      Height          =   315
      Left            =   8760
      TabIndex        =   26
      Top             =   960
      Width           =   3615
      _Version        =   1572864
      _ExtentX        =   6376
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtContrato 
      Height          =   315
      Left            =   8760
      TabIndex        =   27
      Top             =   1320
      Width           =   3615
      _Version        =   1572864
      _ExtentX        =   6376
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.GroupBox gbAcciones 
      Height          =   735
      Left            =   12360
      TabIndex        =   29
      Top             =   960
      Width           =   6615
      _Version        =   1572864
      _ExtentX        =   11668
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Acciones:"
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
      BorderStyle     =   2
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   735
         Left            =   3240
         TabIndex        =   30
         Top             =   0
         Width           =   1935
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   1291
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Begin XtremeSuiteControls.RadioButton rbExport 
            Height          =   252
            Index           =   0
            Left            =   840
            TabIndex        =   31
            Top             =   144
            Width           =   972
            _Version        =   1572864
            _ExtentX        =   1714
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Excel"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
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
         Begin XtremeSuiteControls.PushButton btnExportar 
            Height          =   420
            Left            =   120
            TabIndex        =   32
            Top             =   216
            Width           =   492
            _Version        =   1572864
            _ExtentX        =   868
            _ExtentY        =   741
            _StockProps     =   79
            BackColor       =   16777215
            FlatStyle       =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCxC_PlanPagos.frx":0E1F
         End
         Begin XtremeSuiteControls.RadioButton rbExport 
            Height          =   252
            Index           =   1
            Left            =   840
            TabIndex        =   33
            Top             =   396
            Width           =   972
            _Version        =   1572864
            _ExtentX        =   1714
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "HTML"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
      End
      Begin XtremeSuiteControls.GroupBox gbPrinter 
         Height          =   735
         Left            =   960
         TabIndex        =   34
         Top             =   0
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
         _ExtentY        =   1296
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Begin XtremeSuiteControls.PushButton btnImpresora 
            Height          =   540
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   492
            _Version        =   1572864
            _ExtentX        =   868
            _ExtentY        =   952
            _StockProps     =   79
            BackColor       =   16777215
            FlatStyle       =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCxC_PlanPagos.frx":141E
         End
         Begin XtremeSuiteControls.RadioButton rbPrinter 
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   36
            Top             =   120
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Plan de Pagos"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
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
         Begin XtremeSuiteControls.RadioButton rbPrinter 
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   37
            Top             =   360
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Movimientos"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
      End
      Begin XtremeSuiteControls.PushButton btnRefrescar 
         Height          =   620
         Left            =   240
         TabIndex        =   38
         Top             =   100
         Width           =   612
         _Version        =   1572864
         _ExtentX        =   1080
         _ExtentY        =   1094
         _StockProps     =   79
         BackColor       =   16777215
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCxC_PlanPagos.frx":1BDA
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Plan de Pagos"
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
      Height          =   360
      Index           =   0
      Left            =   1080
      TabIndex        =   8
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblOperacion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7920
      TabIndex        =   7
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
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
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1452
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1452
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pagador"
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
      Index           =   4
      Left            =   7680
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Contrato"
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
      Index           =   5
      Left            =   7680
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblOficina 
      BackStyle       =   0  'Transparent
      Caption         =   "Oficina...."
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
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   525
      Width           =   7095
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17775
   End
End
Attribute VB_Name = "frmCxC_PlanPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders

vHeaders.Columnas = 22
vHeaders.Headers(1) = "Linea"
vHeaders.Headers(2) = "Fec.Inicio"
vHeaders.Headers(3) = "Fec.Corte"
vHeaders.Headers(4) = "Cargos"
vHeaders.Headers(5) = "Int.Cor."
vHeaders.Headers(6) = "Int.Mor."
vHeaders.Headers(7) = "Amortización"
vHeaders.Headers(8) = "Saldo Ant."
vHeaders.Headers(9) = "Saldo Actual"
vHeaders.Headers(10) = "Dias"
vHeaders.Headers(11) = "Estado"
vHeaders.Headers(12) = "Dias Atraso"
vHeaders.Headers(13) = "Mov.Fecha"
vHeaders.Headers(14) = "Mov.Total"
vHeaders.Headers(15) = "Mov.Cargos"
vHeaders.Headers(16) = "Mov.Int.Cor."
vHeaders.Headers(17) = "Mov.Int.Mor."
vHeaders.Headers(18) = "Mov.Amortización"
vHeaders.Headers(19) = "Usr.Caja"
vHeaders.Headers(20) = "Tipo.Doc."
vHeaders.Headers(21) = "# Documento"
vHeaders.Headers(22) = "Concepto"


Select Case True
  Case rbExport.Item(0).Value    'EXCEL
      Call sbSIFGridExportar(vGrid, vHeaders, "CxC_Plan_Pagos_Op" & lblOperacion.Caption)
  Case rbExport.Item(1).Value    'HTML
      Call sbSIFGridExportar(vGrid, vHeaders, "CxC_Plan_Pagos_Op" & lblOperacion.Caption, "HTML")
End Select

End Sub

Private Sub btnImpresora_Click()
Select Case True
  Case rbPrinter.Item(0).Value  'Plan
      Call sbReportes("Plan")
  Case rbPrinter.Item(1).Value  'Movimientos
      Call sbReportes("Movimientos")
'  Case rbPrinter.Item(1).Value  'Estudio
'      Call sbReportes("Estudio")
      
End Select

End Sub

Private Sub btnRefrescar_Click()
       Call sbInicializa
End Sub



Private Sub chkCargos_Click()
If chkCargos.Value = vbChecked Then
   lsw.Visible = True
Else
   lsw.Visible = False
End If

Call Form_Resize

End Sub

Private Sub Form_Load()

Me.Icon = Me.Picture

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
vGrid.AppearanceStyle = fxGridStyle

With lsw.ColumnHeaders
    .Clear
    .Add , , "Línea", 1000
    .Add , , "Seq", 1000
    .Add , , "Detalle", 3500
    .Add , , "Monto", 1500, vbRightJustify
    .Add , , "Saldo", 1500, vbRightJustify
    .Add , , "Unidad", 1000, vbCenter
    .Add , , "C.Costo", 1000, vbCenter
    .Add , , "Cuenta", 2500, vbCenter
    .Add , , "Abono", 1500, vbRightJustify
    .Add , , "Pendiente", 1500, vbRightJustify
End With

lblOperacion.Caption = Operacion.OperacionConsulta

End Sub


Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset


On Error GoTo vError

strSQL = "select R.Operacion,S.cedula,S.nombre,R.cod_Concepto,C.descripcion,R.Monto,R.Saldo,R.cuota,R.Tasa_Corriente as 'TasaO',R.Tipo_Plazo,  R.Dias_Plazo as 'Plazo'" _
       & ",Ofi.descripcion as 'OficinaX',Con.Descripcion as 'Contrato',Pag.Nombre as 'Pagador',R.Num_Documento,R.Fecha_Pago" _
       & " from CxC_Personas S inner join CxC_Cuentas R on S.cedula = R.cedula" _
       & " inner join CxC_Conceptos C on R.cod_Concepto = C.cod_Concepto" _
       & " left join CxC_Contratos Con on R.cod_Contrato = Con.cod_contrato" _
       & " left join CxC_Personas Pag on R.cedula_pagador = Pag.cedula" _
       & " left join SIF_Oficinas Ofi on R.cod_oficina = Ofi.cod_oficina" _
       & " where R.Operacion = " & lblOperacion.Caption
Call OpenRecordSet(rs, strSQL)

txtCedula.Text = rs!Cedula
txtNombre.Text = rs!Nombre

txtDocumento.Text = Trim(rs!Num_Documento & "")
txtDocumento.ToolTipText = "Fecha de Pago..:" & Format(rs!Fecha_Pago & "", "dd/mm/yyyy")

txtPagador.Text = rs!Pagador & ""
txtContrato.Text = rs!Contrato & ""

txtLinea.Text = rs!cod_Concepto
txtLineaDesc.Text = rs!Descripcion

lblOficina.Caption = rs!OficinaX & ""

txtMonto.Text = Format(rs!Monto, "Standard")
txtSaldo.Text = Format(rs!Saldo, "Standard")

txtCuota.Text = Format(rs!Cuota, "Standard")

If rs!Tipo_Plazo = "M" Then
    txtPlazo.Text = CStr(rs!Plazo / 30) & " meses"
Else
    txtPlazo.Text = CStr(rs!Plazo) & " días"
End If

txtTasa.Text = Format(rs!TasaO, "Standard")
txtTasa.ToolTipText = "Tasa Actual: " & Format(rs!TasaO, "Standard")


rs.Close

strSQL = "select Mov.Linea,Mov.Fecha_Inicio,Mov.Fecha_Corte,Mov.Cargos" _
       & ",Mov.Int_Cor,Mov.Int_Mor,Mov.Principal,Mov.Saldo_Inicial,isnull(Mov.Saldo_Final,Mov.Saldo_Inicial-Mov.Principal)" _
       & ",Mov.Dias,case Mov.Estado when 'A' then 'Activa' when 'P' then 'Pendiente'" _
       & " when 'C' then 'Cancelada' when 'N' then 'Anulada' end as 'Estado', Mov.Dias_Mora" _
       & ",Mov.Registro_Fecha,Mov.Mov_Monto,Mov.Mov_Cargos,Mov.Mov_Int_Cor,Mov.Mov_Int_Mor,Mov.Mov_Principal,Mov.Cod_Caja + '/' + Mov.Registro_Usuario" _
       & ",Mov.Tipo_Documento,Mov.Num_Documento,Con.Descripcion as 'Concepto'" _
       & " from CxC_Cuentas_Mov Mov left join SIF_Conceptos Con on Mov.cod_concepto = Con.cod_Concepto" _
       & " where Mov.Operacion = " & lblOperacion.Caption _
       & " order by Mov.Linea"
Call sbCargaGridFps7(vGrid, 22, strSQL, False)

strSQL = "select max(Linea) as Lineas, sum(Int_Cor + Int_Mor) as Intereses, Sum(Cargos) as Cargos" _
       & ", sum(Dias) as Dias, min(Fecha_Corte) as Inicio, max(Fecha_Corte) as Corte, Sum(Dias_Mora) as MoraDias" _
       & " from CxC_Cuentas_Mov" _
       & " where isnull(Linea_Madre,0) = 0 and Operacion = " & lblOperacion.Caption
Call OpenRecordSet(rs, strSQL)
    StatusBarX.Panels.Item(1).Text = "Líneas..: " & rs!Lineas
    StatusBarX.Panels.Item(2).Text = "Inicio..: " & Format(rs!Inicio, "dd/mm/yyyy")
    StatusBarX.Panels.Item(3).Text = "Corte..: " & Format(rs!Corte, "dd/mm/yyyy")
    StatusBarX.Panels.Item(4).Text = "Días..: " & rs!Dias
    StatusBarX.Panels.Item(5).Text = "Intereses..: " & Format(rs!Intereses, "Standard")
    StatusBarX.Panels.Item(6).Text = "Cargos ..:  " & Format(rs!Cargos, "Standard")
    StatusBarX.Panels.Item(7).Text = "Mora días..: " & rs!MoraDias
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

vGrid.Width = Me.Width - 450

If chkCargos.Value = vbChecked Then
    vGrid.Height = Me.Height - vGrid.top - 680 - StatusBarX.Height - lsw.Height
Else
    vGrid.Height = Me.Height - vGrid.top - 580 - StatusBarX.Height
End If

lsw.top = vGrid.top + vGrid.Height + 50
lsw.Left = vGrid.Left
lsw.Width = vGrid.Width

gbTotales.Width = lsw.Width

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
Call sbInicializa
End Sub


Private Sub sbReportes(pTipo As String)
Dim vFecha As Date

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = False
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de CxC"

 .Connect = glogon.ConectRPT
                
    Select Case pTipo
      Case "Plan"
         .ReportFileName = SIFGlobal.fxPathReportes("CxC_PlanPagos.rpt")
      Case "Movimientos"
         .ReportFileName = SIFGlobal.fxPathReportes("CxC_PlanPagosMov.rpt")
      Case "Estudio"
         .ReportFileName = SIFGlobal.fxPathReportes("CxC_PlanPagosEstudio.rpt")
    End Select

 .Formulas(0) = "fxFecha='FECHA: " & Format(vFecha, "dd/mm/yyyy  hh:mm:ss") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxOficina='" & GLOBALES.gOficina & "'"
 
 .SelectionFormula = "{CXC_CUENTAS.OPERACION} = " & lblOperacion.Caption
 
' .SubreportToChange = "sbCorte"
' .StoredProcParam(0) = lblOperacion.Caption
' .StoredProcParam(1) = Format(vFecha, "yyyy/mm/dd")

 .PrintReport

End With

Me.MousePointer = vbDefault


End Sub



Private Sub vGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

If Row = NewRow Then Exit Sub
If NewCol <> 4 Then Exit Sub

vGrid.Row = NewRow
vGrid.Col = NewCol
If CCur(vGrid.Text) = 0 Then Exit Sub

If chkCargos.Value = vbChecked Then
  vGrid.Col = 11
  
  Me.MousePointer = vbHourglass
  
  If Mid(vGrid.Text, 1, 1) = "A" Then
        strSQL = "select *, isnull(Monto - Saldo,0) as 'Abono',0 as 'Linea'" _
               & " From CxC_Cuentas_Cargos" _
               & " where Operacion = " & CLng(lblOperacion.Caption) & " and Saldo > 0"
  Else
        vGrid.Col = 1
        strSQL = "select Car.*, isnull(Mov.Monto,0) as 'Abono',isnull(Mov.Linea,0) as 'Linea'" _
               & " From CxC_Cuentas_Cargos Car inner join CxC_Cuentas_Cargos_Mov Mov on Car.id_Cargo = Mov.id_Cargo" _
               & " and Car.Operacion = Mov.Operacion and Mov.Linea = " & vGrid.Text _
               & " where Car.Operacion = " & CLng(lblOperacion.Caption) _
               & " order by Mov.LINEA"
  End If
  
  lsw.ListItems.Clear
  Call OpenRecordSet(rs, strSQL)
  Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!Id_Cargo)
        itmX.SubItems(1) = rs!Linea
        itmX.SubItems(2) = rs!Notas
        itmX.SubItems(3) = Format(rs!Monto, "Standard")
        itmX.SubItems(4) = Format(rs!Saldo, "Standard")
        itmX.SubItems(5) = rs!Cod_Unidad
        itmX.SubItems(6) = rs!Cod_Centro_Costo
        itmX.SubItems(7) = fxgCntCuentaFormato(True, rs!cod_cuenta)
        itmX.SubItems(8) = Format(rs!Abono, "Standard")
        itmX.SubItems(9) = Format(rs!Saldo, "Standard")
       
    If rs!Abono = 0 Then
       itmX.SmallIcon = 5
    Else
       itmX.SmallIcon = 6
    End If
    rs.MoveNext
  Loop
  rs.Close
  Me.MousePointer = vbDefault
  
End If


Exit Sub

vError:
  Me.MousePointer = vbDefault

End Sub



