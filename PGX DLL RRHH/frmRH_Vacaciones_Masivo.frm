VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmRH_Vacaciones_Masivo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "RRHH: Aplicación Masiva de Periodo de Vacaciones"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   17445
   LinkTopic       =   "Form1"
   ScaleHeight     =   10500
   ScaleWidth      =   17445
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   9615
      Left            =   5760
      TabIndex        =   38
      Top             =   960
      Width           =   11655
      _Version        =   1441792
      _ExtentX        =   20558
      _ExtentY        =   16960
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
      Item(0).Caption =   "Casos"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "chkTodas"
      Item(1).Caption =   "No Procesa!"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "gError"
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   8535
         Left            =   0
         TabIndex        =   39
         Top             =   840
         Width           =   11535
         _Version        =   524288
         _ExtentX        =   20346
         _ExtentY        =   15055
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         SpreadDesigner  =   "frmRH_Vacaciones_Masivo.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.CheckBox chkTodas 
         Height          =   255
         Left            =   600
         TabIndex        =   40
         Top             =   480
         Width           =   2055
         _Version        =   1441792
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Marcar"
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
         Value           =   1
      End
      Begin FPSpreadADO.fpSpread gError 
         Height          =   9015
         Left            =   -69880
         TabIndex        =   41
         Top             =   480
         Visible         =   0   'False
         Width           =   11415
         _Version        =   524288
         _ExtentX        =   20135
         _ExtentY        =   15901
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         SpreadDesigner  =   "frmRH_Vacaciones_Masivo.frx":0787
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.GroupBox gbAplica 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   8040
      Width           =   5535
      _Version        =   1441792
      _ExtentX        =   9763
      _ExtentY        =   4260
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   330
         Left            =   0
         TabIndex        =   29
         Top             =   480
         Width           =   3375
         _Version        =   1441792
         _ExtentX        =   5953
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
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   435
         Left            =   0
         TabIndex        =   30
         Top             =   1080
         Width           =   5535
         _Version        =   1441792
         _ExtentX        =   9763
         _ExtentY        =   767
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnProcesa 
         Height          =   495
         Left            =   2400
         TabIndex        =   33
         Top             =   1680
         Width           =   3135
         _Version        =   1441792
         _ExtentX        =   5530
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Generar Vacaciones Masivas"
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
         Picture         =   "frmRH_Vacaciones_Masivo.frx":0DF8
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   330
         Left            =   3360
         TabIndex        =   42
         Top             =   480
         Width           =   2175
         _Version        =   1441792
         _ExtentX        =   3836
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
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   3360
         TabIndex        =   43
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
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
         Index           =   2
         Left            =   0
         TabIndex        =   32
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   31
         Top             =   840
         Width           =   1095
      End
   End
   Begin XtremeSuiteControls.ComboBox cboNomina 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   5535
      _Version        =   1441792
      _ExtentX        =   9763
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
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
      _Version        =   1441792
      _ExtentX        =   2355
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   315
      Left            =   2880
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
      _Version        =   1441792
      _ExtentX        =   2355
      _ExtentY        =   556
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
   Begin MSComCtl2.FlatScrollBar FlatScroll_Laboral 
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   7
      Top             =   2520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComCtl2.FlatScrollBar FlatScroll_Laboral 
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   8
      Top             =   3240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComCtl2.FlatScrollBar FlatScroll_Laboral 
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   9
      Top             =   3960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComCtl2.FlatScrollBar FlatScroll_Laboral 
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   10
      Top             =   4680
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCentroCod 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   735
      _Version        =   1441792
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCentroDesc 
      Height          =   315
      Left            =   840
      TabIndex        =   12
      Top             =   2760
      Width           =   4815
      _Version        =   1441792
      _ExtentX        =   8488
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
   Begin XtremeSuiteControls.FlatEdit txtDeptCodigo 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   735
      _Version        =   1441792
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDeptDesc 
      Height          =   315
      Left            =   840
      TabIndex        =   14
      Top             =   3480
      Width           =   4815
      _Version        =   1441792
      _ExtentX        =   8488
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
   Begin XtremeSuiteControls.FlatEdit txtSecCodigo 
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   735
      _Version        =   1441792
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtSecDesc 
      Height          =   315
      Left            =   840
      TabIndex        =   16
      Top             =   4200
      Width           =   4815
      _Version        =   1441792
      _ExtentX        =   8488
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
   Begin XtremeSuiteControls.FlatEdit txtPuestoCod 
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Top             =   4920
      Width           =   735
      _Version        =   1441792
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPuestoDesc 
      Height          =   315
      Left            =   840
      TabIndex        =   18
      Top             =   4920
      Width           =   4815
      _Version        =   1441792
      _ExtentX        =   8488
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
   Begin XtremeSuiteControls.ComboBox cboContrato 
      Height          =   315
      Left            =   120
      TabIndex        =   23
      Top             =   5640
      Width           =   5535
      _Version        =   1441792
      _ExtentX        =   9763
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
   End
   Begin XtremeSuiteControls.ComboBox cboJornada 
      Height          =   315
      Left            =   120
      TabIndex        =   24
      Top             =   6240
      Width           =   5535
      _Version        =   1441792
      _ExtentX        =   9763
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
   End
   Begin XtremeSuiteControls.ComboBox cboVacaciones 
      Height          =   315
      Left            =   120
      TabIndex        =   25
      Top             =   6840
      Width           =   5535
      _Version        =   1441792
      _ExtentX        =   9763
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
   End
   Begin XtremeSuiteControls.FlatEdit txtDias 
      Height          =   315
      Left            =   4680
      TabIndex        =   35
      ToolTipText     =   "Dias a Disfrutar"
      Top             =   1800
      Width           =   975
      _Version        =   1441792
      _ExtentX        =   1720
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
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   495
      Left            =   2880
      TabIndex        =   36
      Top             =   7320
      Width           =   1215
      _Version        =   1441792
      _ExtentX        =   2138
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Buscar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmRH_Vacaciones_Masivo.frx":1511
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   495
      Left            =   4080
      TabIndex        =   37
      Top             =   7320
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Exportar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmRH_Vacaciones_Masivo.frx":1C11
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   120
      TabIndex        =   44
      Top             =   7800
      Visible         =   0   'False
      Width           =   5535
      _Version        =   1441792
      _ExtentX        =   9763
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Días:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   34
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Contrato"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   28
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Jornada"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   27
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Régimen de Vacaciones"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   26
      Top             =   6600
      Width           =   2775
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Puesto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label lblSeccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Sección"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblDepartamento 
      BackStyle       =   0  'Transparent
      Caption         =   "Departamento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Centro"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Disfrute:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Nómina"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Procesamiento Masivo de Vacaciones"
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
      Height          =   480
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17655
   End
End
Attribute VB_Name = "frmRH_Vacaciones_Masivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vScroll As Boolean
Dim strSQL As String, rs As New ADODB.Recordset, vPaso As Boolean

Private Sub btnBuscar_Click()

On Error GoTo vError

Dim pContrato As String, pJornada As String, pRegimen As String


Me.MousePointer = vbHourglass

If cboContrato.Text = "TODOS" Then
    pContrato = ""
Else
    pContrato = cboContrato.ItemData(cboContrato.ListIndex)
End If

If cboJornada.Text = "TODOS" Then
    pJornada = ""
Else
    pJornada = cboJornada.ItemData(cboJornada.ListIndex)
End If

If cboVacaciones.Text = "TODOS" Then
    pRegimen = ""
Else
    pRegimen = cboVacaciones.ItemData(cboVacaciones.ListIndex)
End If

tcMain.Item(0).Selected = True

strSQL = "exec spRH_Vacaciones_Masivas_Consulta '" & cboNomina.ItemData(cboNomina.ListIndex) _
        & "','" & Format(dtpInicio.Value, "yyyy-MM-dd") & "','" & Format(dtpCorte.Value, "yyyy-MM-dd") _
        & "','" & txtCentroCod.Text & "','" & txtDeptCodigo.Text & "','" & txtSecCodigo.Text _
        & "','" & txtPuestoCod.Text & "','" & pContrato & "','" & pJornada & "','" & pRegimen _
        & "','A'"

Call OpenRecordSet(rs, strSQL)

With vGrid
    .MaxRows = 0
    
    Do While Not rs.EOF
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .col = 1
        .Value = chkTodas.Value
        .col = 2
        .Text = rs!EMPLEADO_ID
        .col = 3
        .Text = rs!IDENTIFICACION
        .col = 4
        .Text = rs!NOMBRE_COMPLETO
        .col = 5
        .Text = rs!VACA_ACTUALIZA & ""
        .col = 6
        .Text = CStr(IIf(IsNull(rs!VACA_ACUMULADAS), 0, rs!VACA_ACUMULADAS))
        .col = 7
        .Text = CStr(IIf(IsNull(rs!VACA_ACUMULADAS), 0, rs!VACA_ACUMULADAS) - CCur(txtDias.Text))
        
        
        rs.MoveNext
    Loop
    
End With
rs.Close


'Casos con Boletas Previas que chocan con el rango masivo

strSQL = "exec spRH_Vacaciones_Masivas_Consulta '" & cboNomina.ItemData(cboNomina.ListIndex) _
        & "','" & Format(dtpInicio.Value, "yyyy-MM-dd") & "','" & Format(dtpCorte.Value, "yyyy-MM-dd") _
        & "','" & txtCentroCod.Text & "','" & txtDeptCodigo.Text & "','" & txtSecCodigo.Text _
        & "','" & txtPuestoCod.Text & "','" & pContrato & "','" & pJornada & "','" & pRegimen _
        & "','E'"

Call OpenRecordSet(rs, strSQL)

With gError
    .MaxRows = 0
    
    Do While Not rs.EOF
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .col = 1
        .Text = rs!EMPLEADO_ID
        .col = 2
        .Text = rs!IDENTIFICACION
        .col = 3
        .Text = rs!NOMBRE_COMPLETO
        .col = 4
        .Text = rs!VACA_ACTUALIZA & ""
        .col = 5
        .Text = CStr(IIf(IsNull(rs!VACA_ACUMULADAS), 0, rs!VACA_ACUMULADAS))
        rs.MoveNext
    Loop
    
End With
rs.Close




Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnExportar_Click()
 Dim vHeaders As vGridHeaders
 
 If tcMain.Item(0).Selected Then
 
    vHeaders.Columnas = vGrid.MaxCols
    
    vHeaders.Headers(1) = "..."
    vHeaders.Headers(2) = "Empleado Id"
    vHeaders.Headers(3) = "Identificación"
    vHeaders.Headers(4) = "Nombre"
    vHeaders.Headers(5) = "Ult.Corte"
    vHeaders.Headers(6) = "Vaca.Acum."
    vHeaders.Headers(7) = "Vaca.Resultado"
    Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_RRHH_Vaca_Masiva_Detalle")

Else
    'No Procesa
    vHeaders.Columnas = gError.MaxCols
    
    vHeaders.Headers(1) = "Empleado Id"
    vHeaders.Headers(2) = "Identificación"
    vHeaders.Headers(3) = "Nombre"
    vHeaders.Headers(4) = "Ult.Corte"
    vHeaders.Headers(5) = "Vaca.Acum."
    vHeaders.Headers(6) = "Vaca.Resultado"
    Call sbSIFGridExportar(gError, vHeaders, "ProGrX_RRHH_Vaca_Masiva_NoProcesa")

End If

End Sub

Private Sub btnProcesa_Click()

On Error GoTo vError

Dim i As Long
Dim pAutorizador As String

If Mid(cboEstado.Text, 1, 1) = "S" Then
  pAutorizador = "Null"
Else
  pAutorizador = "Null"
End If


Me.MousePointer = vbHourglass

strSQL = ""

With vGrid

ProgressBarX.Visible = True
ProgressBarX.Value = 1
ProgressBarX.Max = .MaxRows


For i = 1 To .MaxRows
    .Row = i
    .col = 1
    If .Value = vbChecked Then
        .col = 2
        
        strSQL = strSQL & Space(10) & "exec spRH_Vacaciones_Registro '" & .Text & "','" & cboTipo.ItemData(cboTipo.ListIndex) _
                & "','" & txtNotas.Text & "','" & glogon.Usuario & "'" _
                & ",'" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00'" _
                & ",'" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
                & "," & CInt(txtDias.Text)
        .col = 6
        strSQL = strSQL & "," & CCur(.Text) & "," & 0 _
                & ",'" & Mid(cboEstado.Text, 1, 1) & "'," & pAutorizador & ",'ProGrX'"


        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If
    
        ProgressBarX.Value = i

    End If
Next i

End With

'Ultimo Lote
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If

ProgressBarX.Visible = False

Call Bitacora("Aplica", "Vacaciones Masivas_I: " & Format(dtpInicio.Value, "yyyy-mm-dd") & "  C: " & Format(dtpCorte.Value, "yyyy-mm-dd"))


Me.MousePointer = vbDefault

MsgBox "Vacaciones registradas satisfactoriamente!", vbInformation

Call btnBuscar_Click

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboContrato_Click()
If vPaso Then Exit Sub
Call sbCleanResultados
End Sub

Private Sub cboJornada_Click()
If vPaso Then Exit Sub
Call sbCleanResultados
End Sub

Private Sub cboNomina_Click()
If vPaso Then Exit Sub
Call sbCleanResultados

End Sub

Private Sub cboTipo_Click()
If vPaso Then Exit Sub
If cboTipo.ListCount = 0 Then Exit Sub

strSQL = "select REQUIERE_AUTORIZACION,PERMITE_LIQUIDACION" _
       & " from RH_VACACIONES_TIPOS " _
       & " WHERE VACA_TIPO = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

cboEstado.Clear
cboEstado.AddItem "Solicitado"
cboEstado.Text = "Solicitado"

If rs!REQUIERE_AUTORIZACION = 0 Then
    cboEstado.AddItem "Autorizado"
End If
rs.Close

End Sub


Private Sub cboVacaciones_Click()
If vPaso Then Exit Sub
Call sbCleanResultados
End Sub

Private Sub chkTodas_Click()
Dim i As Long

For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.col = 1
    vGrid.Value = chkTodas.Value
Next i

End Sub

Private Sub dtpCorte_Change()
If vPaso Then Exit Sub

If cboNomina.ListCount = 0 Then Exit Sub

txtDias.Text = fxRH_Dias_Laborales_Nomina(cboNomina.ItemData(cboNomina.ListIndex), dtpInicio.Value, dtpCorte.Value)
Call sbCleanResultados

End Sub

Private Sub dtpInicio_Change()
    
If vPaso Then Exit Sub

If cboNomina.ListCount = 0 Then Exit Sub

txtDias.Text = fxRH_Dias_Laborales_Nomina(cboNomina.ItemData(cboNomina.ListIndex), dtpInicio.Value, dtpCorte.Value)

Call sbCleanResultados
End Sub

Private Sub FlatScroll_Laboral_Change(Index As Integer)
Dim rs As New ADODB.Recordset
Dim vCodigo As String, vColumna As String, vChar As String, vFiltroAdd As String
Dim txtCodigo As Object, txtDesc As Object

On Error GoTo vError

If Not vScroll Then Exit Sub

vChar = "'"
vFiltroAdd = ""

Call sbCleanResultados

Select Case Index
   Case 0 'Centro
        vCodigo = txtCentroCod.Text
        vColumna = "COD_CENTRO"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from RH_CENTRO_TRABAJO"
        
        Set txtCodigo = txtCentroCod
        Set txtDesc = txtCentroDesc
    
    Case 1 'Departamentos
        vCodigo = txtDeptCodigo.Text
        vColumna = "COD_DEPARTAMENTO"
        vFiltroAdd = " AND COD_CENTRO = '" & txtCentroCod.Text & "'"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from RH_Departamentos"
        
        Set txtCodigo = txtDeptCodigo
        Set txtDesc = txtDeptDesc
        
        
    Case 2 'Secciones
        vCodigo = txtSecCodigo.Text
        vColumna = "COD_SECCION"
        vFiltroAdd = " AND COD_CENTRO = '" & txtCentroCod.Text & "' AND COD_DEPARTAMENTO = '" & txtDeptCodigo.Text & "'"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from RH_SECCIONES"
        
        Set txtCodigo = txtSecCodigo
        Set txtDesc = txtSecDesc

    
    Case 4 'Puesto
        vCodigo = txtPuestoCod.Text
        
        vColumna = "COD_PUESTO"
        vFiltroAdd = " AND ACTIVO = 1"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from RH_PUESTOS"
        
        Set txtCodigo = txtPuestoCod
        Set txtDesc = txtPuestoDesc
    
End Select

If vScroll Then
    
    If FlatScroll_Laboral(Index).Value = 1 Then
       strSQL = strSQL & " where " & vColumna & " > " & vChar & vCodigo & vChar & " " & vFiltroAdd & " order by " & vColumna & " asc"
    Else
       strSQL = strSQL & " where " & vColumna & " < " & vChar & vCodigo & vChar & " " & vFiltroAdd & " order by " & vColumna & " desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!Codigo
      txtDesc.Text = rs!Descripcion

    End If
    rs.Close
End If



vScroll = False
FlatScroll_Laboral(Index).Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Load()

vModulo = 23

 vScroll = False
' FlatScrollBar.Value = 0
 vScroll = True
 
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vGrid.MaxRows = 0

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbCleanResultados()

tcMain.Item(0).Selected = True
vGrid.MaxRows = 0
gError.MaxRows = 0

End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vPaso = True

dtpInicio.Value = fxFechaServidor

dtpInicio.MinDate = dtpInicio.Value

dtpCorte.Value = dtpInicio.Value
dtpCorte.MinDate = dtpInicio.Value

'Tipo de Accion
    strSQL = "select VACA_TIPO as Idx, rtrim(Descripcion) as ItmX" _
           & " from RH_VACACIONES_TIPOS" _
           & " Where Activa = 1"
    Call sbCbo_Llena_New(cboTipo, strSQL, False, True)

'Nomina
    strSQL = "select COD_NOMINA as Idx, rtrim(Descripcion) as ItmX from RH_NOMINAS_CATALOGO"
    Call sbCbo_Llena_New(cboNomina, strSQL, False, True)

''Divisa
'    strSQL = "select COD_DIVISA as Idx, rtrim(Descripcion) as ItmX from vSys_Divisas"
'    Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)

'Jornada
    strSQL = "select JORNADA_TIPO as Idx, rtrim(Descripcion) as ItmX from RH_JORNADAS_TIPOS"
    Call sbCbo_Llena_New(cboJornada, strSQL, True, True)

'Contratos
    strSQL = "Select CONTRATO_TIPO as Idx, rtrim(Descripcion) as ItmX from RH_CONTRATOS_TIPOS"
    Call sbCbo_Llena_New(cboContrato, strSQL, True, True)

'Vacaciones
    strSQL = "Select COD_VACA_REGIMEN as Idx, rtrim(Descripcion) as ItmX from RH_VACACIONES_REGIMEN"
    Call sbCbo_Llena_New(cboVacaciones, strSQL, True, True)


vPaso = False

'Valores Iniciales
Call dtpCorte_Change

Call cboTipo_Click

vGrid.MaxRows = 0
gError.MaxRows = 0

tcMain.Item(0).Selected = True

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Resize()

On Error Resume Next
Dim pH As Long, pW As Long

pH = 10965
pW = 17565


If Me.Height > pH Then
   pH = Me.Height
End If

If Me.Width > pW Then
  pW = Me.Width
End If

Me.Height = pH
Me.Width = pW

tcMain.Height = pH - (tcMain.Top + 250)
tcMain.Width = pW - (tcMain.Left + 100)

gbAplica.Top = pH - (gbAplica.Height + 350)

imgBanner.Width = Me.Width

If gbAplica.Top < 8040 Then
    gbAplica.Top = 8040
End If

vGrid.Width = tcMain.Width - 200
vGrid.Height = tcMain.Height - (vGrid.Top + 250)

gError.Width = tcMain.Width - 200
gError.Height = tcMain.Height - (gError.Top + 250)

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa
End Sub


Private Sub sbConsulta()
Dim strSQL As String, rs As New Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

'strSQL = "select * from vRH_Personas" _
'       & " Where Empleado_Id = '" & txtEmpleadoId.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
'    txtEmpleadoId.Text = rs!EMPLEADO_ID
'    txtIdentificacion.Text = rs!Identificacion
'    txtNombre.Text = rs!Nombre_Completo
    
    txtCentroCod.Text = rs!Cod_Centro
    txtCentroDesc.Text = rs!CentroDesc
    txtDeptCodigo.Text = rs!Cod_Departamento
    txtDeptDesc.Text = rs!DepartamentoDesc
    txtSecCodigo.Text = rs!Cod_Seccion
    txtSecDesc.Text = rs!SeccionDesc
    
    txtPuestoCod.Text = rs!Cod_Puesto
    txtPuestoDesc.Text = rs!PuestoDesc
    
   
   Call sbCboAsignaDato(cboNomina, rs!NominaDesc, True, rs!Cod_Nomina)
   Call sbCboAsignaDato(cboContrato, rs!ContratoDesc, True, rs!Contrato_Tipo)
   Call sbCboAsignaDato(cboJornada, rs!JornadaDesc, True, rs!Jornada_Tipo)
   Call sbCboAsignaDato(cboVacaciones, rs!VacacionesDesc, True, rs!Cod_Vaca_Regimen)
   
  
    
Else
    'Todo
End If
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub txtCentroCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCentroDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select COD_CENTRO,descripcion,desc_Corta from RH_CENTRO_TRABAJO"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtCentroCod.Text = Trim(gBusquedas.Resultado)
    txtCentroDesc.Text = gBusquedas.Resultado2
    Call sbCleanResultados
  End If
End If
End Sub


Private Sub txtCentroCod_LostFocus()
txtCentroDesc.Text = fxgRH_Centro_Trabajo(txtCentroCod.Text)
End Sub

Private Sub txtCentroDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDeptCodigo.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select COD_CENTRO,descripcion,desc_Corta from RH_CENTRO_TRABAJO"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtCentroCod.Text = Trim(gBusquedas.Resultado)
    txtCentroDesc.Text = gBusquedas.Resultado2
    Call sbCleanResultados
  End If
End If
End Sub


Private Sub txtDeptCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDeptDesc.SetFocus

If KeyCode = vbKeyF4 Then

    gBusquedas.Columna = "cod_departamento"
    gBusquedas.Orden = "cod_departamento"
    gBusquedas.Consulta = "select cod_departamento,descripcion from RH_Departamentos"
    gBusquedas.Filtro = " and COD_CENTRO = '" & txtCentroCod.Text & "'"
  
   
  
  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2

    Call sbCleanResultados
End If


End Sub

Private Sub txtDeptCodigo_LostFocus()
 txtDeptDesc.Text = fxgRH_Departamento(txtCentroCod.Text, txtDeptCodigo.Text)
End Sub

Private Sub txtDeptDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecCodigo.SetFocus
If KeyCode = vbKeyF4 Then

    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_departamento,descripcion from RH_Departamentos"
    gBusquedas.Filtro = " and COD_CENTRO = '" & txtCentroCod.Text & "'"

  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
    Call sbCleanResultados
End If


End Sub


Private Sub txtPuestoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboContrato.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "COD_PUESTO"
  gBusquedas.Orden = "COD_PUESTO"
  gBusquedas.Consulta = "select COD_PUESTO,descripcion from Rh_Puestos"
  gBusquedas.Filtro = ""
        
  frmBusquedas.Show vbModal
  txtPuestoCod.Text = gBusquedas.Resultado
  txtPuestoDesc.Text = gBusquedas.Resultado2
    Call sbCleanResultados
End If

End Sub

Private Sub txtPuestoCod_LostFocus()
'Dim strSQL As String, rs As New ADODB.Recordset
'
'strSQL = "select * from PH_Puestos where cod_Puesto = '" & txtPuestoCod.Text & "'"
'Call OpenRecordSet(rs, strSQL)
'If Not glogon.error Then
'   strSQL = "Salario Recomendado: " & Format(rs!Salario_Actual, "Standard") & vbCrLf _
'          & "Salario Máximo     : " & Format(rs!Salario_Maximo, "Standard") & vbCrLf _
'          & "Salario Mínimo     : " & Format(rs!Salario_Minimo, "Standard") & vbCrLf
'   txtSalario.ToolTipText = strSQL
'
'   If gbAccionPersonal.Enabled Then
'    txtSalario.Text = Format(rs!Salario_Actual, "Standard")
'   End If
'
'End If

End Sub


Private Sub txtSecCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecDesc.SetFocus
If KeyCode = vbKeyF4 Then

        gBusquedas.Columna = "cod_seccion"
        gBusquedas.Orden = "cod_seccion"
        gBusquedas.Consulta = "select cod_seccion,descripcion from Rh_Secciones"
        gBusquedas.Filtro = " and COD_CENTRO = '" & txtCentroCod.Text _
                  & "' and cod_departamento = '" & txtDeptCodigo & "'"
  
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
    Call sbCleanResultados
End If

End Sub

Private Sub txtSecCodigo_LostFocus()
 txtSecDesc.Text = fxgRH_Seccion(txtCentroCod.Text, txtDeptCodigo.Text, txtSecCodigo.Text)
End Sub

Private Sub txtSecDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPuestoCod.SetFocus

If KeyCode = vbKeyF4 Then

        gBusquedas.Columna = "descripcion"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Consulta = "select cod_seccion,descripcion from Rh_Secciones"
        gBusquedas.Filtro = " and COD_CENTRO = '" & txtCentroCod.Text _
                  & "' and cod_departamento = '" & txtDeptCodigo & "'"

  
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
    Call sbCleanResultados
End If

End Sub



