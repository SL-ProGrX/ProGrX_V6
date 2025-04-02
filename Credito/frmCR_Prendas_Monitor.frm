VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCR_Prendas_Monitor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Monitor de Garantía Prendarias"
   ClientHeight    =   10665
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17055
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10665
   ScaleWidth      =   17055
   Begin XtremeSuiteControls.ListView lswComercio 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
      _ExtentY        =   3625
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
      Checkboxes      =   -1  'True
      MultiSelect     =   -1  'True
      HideSelection   =   0   'False
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   7920
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
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
   Begin XtremeSuiteControls.ComboBox cboFechas 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   8280
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   8640
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2773
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
      Left            =   1560
      TabIndex        =   4
      Top             =   9000
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2773
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
   Begin XtremeSuiteControls.FlatEdit txtFiltraComercio 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.CheckBox chkComercio 
      Height          =   210
      Left            =   2880
      TabIndex        =   6
      Top             =   2160
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6615
      Left            =   3360
      TabIndex        =   7
      Top             =   2880
      Width           =   10695
      _Version        =   524288
      _ExtentX        =   18865
      _ExtentY        =   11668
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
      MaxCols         =   33
      SpreadDesigner  =   "frmCR_Prendas_Monitor.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.GroupBox gbFiltros 
      Height          =   2535
      Left            =   3360
      TabIndex        =   8
      Top             =   0
      Width           =   16815
      _Version        =   1441793
      _ExtentX        =   29660
      _ExtentY        =   4471
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin VB.Timer TimerX 
         Interval        =   5
         Left            =   12240
         Top             =   960
      End
      Begin XtremeSuiteControls.FlatEdit txtUserActualiza 
         Height          =   330
         Left            =   1920
         TabIndex        =   9
         Top             =   2160
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   4194304
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
      Begin XtremeSuiteControls.FlatEdit txtIdPrincipal 
         Height          =   330
         Left            =   1920
         TabIndex        =   10
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   600
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   4194304
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
      Begin XtremeSuiteControls.FlatEdit txtPersonaId 
         Height          =   330
         Left            =   1920
         TabIndex        =   11
         Top             =   240
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   4194304
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
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   330
         Left            =   5640
         TabIndex        =   12
         Top             =   240
         Width           =   5295
         _Version        =   1441793
         _ExtentX        =   9340
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   4194304
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
      Begin XtremeSuiteControls.ComboBox cboPresentacion 
         Height          =   330
         Left            =   5640
         TabIndex        =   13
         Top             =   1440
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.ComboBox cboCombustible 
         Height          =   330
         Left            =   5640
         TabIndex        =   14
         Top             =   2160
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.FlatEdit txtUserRegistra 
         Height          =   330
         Left            =   1920
         TabIndex        =   15
         Top             =   1800
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   4194304
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
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   615
         Left            =   11040
         TabIndex        =   16
         Top             =   360
         Width           =   1335
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
         Picture         =   "frmCR_Prendas_Monitor.frx":1098
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   615
         Left            =   12360
         TabIndex        =   17
         ToolTipText     =   "Exportar a Excel"
         Top             =   360
         Width           =   615
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
         Picture         =   "frmCR_Prendas_Monitor.frx":1AB6
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBarX 
         Height          =   135
         Left            =   11040
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   233
         _StockProps     =   93
         BackColor       =   -2147483633
         Scrolling       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboEstadoPersona 
         Height          =   330
         Left            =   1920
         TabIndex        =   19
         Top             =   1440
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
      Begin XtremeSuiteControls.FlatEdit txtChasisNo 
         Height          =   330
         Left            =   5640
         TabIndex        =   20
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   600
         Width           =   5295
         _Version        =   1441793
         _ExtentX        =   9340
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   4194304
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
      Begin XtremeSuiteControls.ComboBox cboModelo 
         Height          =   330
         Left            =   5640
         TabIndex        =   48
         Top             =   1800
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.ComboBox cboCilindraje 
         Height          =   330
         Left            =   8520
         TabIndex        =   49
         Top             =   1440
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
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
      Begin XtremeSuiteControls.ComboBox cboPeso 
         Height          =   330
         Left            =   8520
         TabIndex        =   50
         Top             =   2160
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
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
      Begin XtremeSuiteControls.ComboBox cboCapacidad 
         Height          =   330
         Left            =   8520
         TabIndex        =   54
         Top             =   1800
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
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
      Begin XtremeSuiteControls.FlatEdit txtCilindrajeInicio 
         Height          =   330
         Left            =   9480
         TabIndex        =   55
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1440
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1296
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   4194304
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
      Begin XtremeSuiteControls.FlatEdit txtCilindrajeCorte 
         Height          =   330
         Left            =   10200
         TabIndex        =   56
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1440
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1296
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   4194304
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
      Begin XtremeSuiteControls.FlatEdit txtCapacidadInicio 
         Height          =   330
         Left            =   9480
         TabIndex        =   57
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1800
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1296
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   4194304
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
      Begin XtremeSuiteControls.FlatEdit txtCapacidadCorte 
         Height          =   330
         Left            =   10200
         TabIndex        =   58
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1800
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1296
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   4194304
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
      Begin XtremeSuiteControls.FlatEdit txtPesoInicio 
         Height          =   330
         Left            =   9480
         TabIndex        =   59
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2160
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1296
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   4194304
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
      Begin XtremeSuiteControls.FlatEdit txtPesoCorte 
         Height          =   330
         Left            =   10200
         TabIndex        =   60
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2160
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1296
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   4194304
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
      Begin XtremeSuiteControls.FlatEdit txtIdSecundario 
         Height          =   330
         Left            =   1920
         TabIndex        =   61
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   960
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   4194304
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
      Begin XtremeSuiteControls.FlatEdit txtVINMotor 
         Height          =   330
         Left            =   5640
         TabIndex        =   63
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   960
         Width           =   5295
         _Version        =   1441793
         _ExtentX        =   9340
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   4194304
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
      Begin XtremeSuiteControls.ComboBox cboPuertas 
         Height          =   330
         Left            =   12240
         TabIndex        =   65
         Top             =   1440
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
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
      Begin XtremeSuiteControls.FlatEdit txtColor 
         Height          =   330
         Left            =   12240
         TabIndex        =   68
         Top             =   1800
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   4194304
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   17
         Left            =   11040
         TabIndex        =   67
         Top             =   1800
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Color"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   16
         Left            =   11040
         TabIndex        =   66
         Top             =   1440
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Puertas"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   15
         Left            =   3720
         TabIndex        =   64
         Top             =   960
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. VIN Motor"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   14
         Left            =   360
         TabIndex        =   62
         ToolTipText     =   "No Placa Provisional"
         Top             =   960
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Id Secundario"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   13
         Left            =   7200
         TabIndex        =   53
         Top             =   2160
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Peso"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   12
         Left            =   6840
         TabIndex        =   52
         Top             =   1800
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Capacidad"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   11
         Left            =   6840
         TabIndex        =   51
         Top             =   1440
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cilindraje"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   31
         Top             =   240
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Id Persona"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   30
         Top             =   240
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Nombre"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   29
         ToolTipText     =   "No Placa"
         Top             =   600
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Id Principal"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   28
         Top             =   600
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Chasis"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   4
         Left            =   7800
         TabIndex        =   27
         Top             =   720
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Label2"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   5
         Left            =   3960
         TabIndex        =   26
         Top             =   1440
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Presentación"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   6
         Left            =   3960
         TabIndex        =   25
         Top             =   1800
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Modelo"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   7
         Left            =   3960
         TabIndex        =   24
         Top             =   2160
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Combustible"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   23
         Top             =   1800
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario Registra"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   22
         Top             =   2160
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario Actualiza"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   21
         Top             =   1440
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Estado Persona"
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
   Begin XtremeSuiteControls.FlatEdit txtCasos 
      Height          =   315
      Left            =   360
      TabIndex        =   32
      Top             =   9840
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1296
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   1080
      TabIndex        =   33
      Top             =   9840
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.ListView lswMarca 
      Height          =   2055
      Left            =   120
      TabIndex        =   40
      Top             =   5760
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
      _ExtentY        =   3625
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
      Checkboxes      =   -1  'True
      MultiSelect     =   -1  'True
      HideSelection   =   0   'False
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltraMarca 
      Height          =   315
      Left            =   120
      TabIndex        =   41
      Top             =   5400
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.CheckBox chkMarca 
      Height          =   210
      Left            =   2880
      TabIndex        =   42
      Top             =   5040
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.ListView lswPrenda 
      Height          =   1215
      Left            =   120
      TabIndex        =   44
      Top             =   840
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
      _ExtentY        =   2143
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
      Checkboxes      =   -1  'True
      MultiSelect     =   -1  'True
      HideSelection   =   0   'False
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltroPrenda 
      Height          =   315
      Left            =   120
      TabIndex        =   45
      Top             =   480
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.CheckBox chkPrenda 
      Height          =   210
      Left            =   2880
      TabIndex        =   46
      Top             =   120
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Prenda ...:"
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
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   47
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Marca ...:"
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
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   43
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   39
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
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
      Height          =   315
      Index           =   6
      Left            =   120
      TabIndex        =   38
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Comercializa ...:"
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
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   37
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   36
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
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
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   35
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Resumen:"
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
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   34
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Image imgBanner 
      Height          =   10710
      Left            =   0
      Picture         =   "frmCR_Prendas_Monitor.frx":22BB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3285
   End
End
Attribute VB_Name = "frmCR_Prendas_Monitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean, itmX As ListViewItem

Private Sub btnBuscar_Click()
Call sbBuscar
End Sub



Private Sub sbBuscar()
Dim strSQL As String, i As Integer
Dim vCadena As String, iCantidad As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
iCantidad = 0

'...
'Codigo, Consec, Cedula, Nombre, Monto, Estado, Beneficio, Solicita Id, Solicita Nombre
', Fecha Reg, Usuario Reg, Fecha Autoriza, Usuario Autoriza
', Institución, Departamento, Oficina


strSQL = "select 0 as 'Btn', Cod_Beneficio, Consec, Cedula, NOMBRE_BENEFICIARIO, Monto, Estado_Desc, Beneficio_Desc" _
       & ", Solicita, Solicita_Nombre, Registra_Fecha, Registra_User, Autoriza_Fecha, Autoriza_User " _
       & ", Empresa_Desc, Departamento_Desc, Oficina_Desc" _
       & " from vBeneficios_Integral"

Select Case Mid(cboFechas.Text, 1, 1)
    Case "R"
        strSQL = strSQL & " Where Registra_Fecha between '" & Format(dtpInicio.Value, "yyyy-mm-dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'"
    Case "A"
        strSQL = strSQL & " Where Autoriza_Fecha between '" & Format(dtpInicio.Value, "yyyy-mm-dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'"
    Case "P"
        strSQL = strSQL & " Where Pago_Fecha between '" & Format(dtpInicio.Value, "yyyy-mm-dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'"
    Case Else
End Select

If cboEstado.Text <> "Todos" Then
        strSQL = strSQL & " And Estado =  '" & Mid(cboEstado.Text, 1, 1) & "'"
End If

If cboPresentacion.Text <> "TODOS" Then
        strSQL = strSQL & " And ID_PRESENTACION =  " & cboPresentacion.ItemData(cboPresentacion.ListIndex)
End If

If cboCombustible.Text <> "TODOS" Then
        strSQL = strSQL & " And cod_Oficina =  '" & cboCombustible.ItemData(cboCombustible.ListIndex) & "'"
End If

If cboEstadoPersona.Text <> "TODOS" Then
        strSQL = strSQL & " And EstadoActual =  '" & cboEstadoPersona.ItemData(cboEstadoPersona.ListIndex) & "'"
End If



'Lista de Documentos
If lswComercio.ListItems.Count > 0 Then
    vCadena = " and Cod_Beneficio in('"
    For i = 1 To lswComercio.ListItems.Count
      If lswComercio.ListItems.Item(i).Checked Then
        vCadena = vCadena & "','" & lswComercio.ListItems.Item(i).Tag
        iCantidad = iCantidad + 1
      End If
    Next i
    strSQL = strSQL & vCadena & "')"
End If


If Trim(txtUserRegistra.Text) <> "" Then
      strSQL = strSQL & " and Registra_Usuario like '%" & txtUserRegistra.Text & "%'"
End If

If Trim(txtUserRegistra.Text) <> "" Then
      strSQL = strSQL & " and Autoriza_Usuario like '%" & txtUserRegistra.Text & "%'"
End If



If Trim(txtPersonaId.Text) <> "" Then
      strSQL = strSQL & " and Cedula like '%" & txtPersonaId.Text & "%'"
End If


If Trim(txtNombre.Text) <> "" Then
      strSQL = strSQL & " and NOMBRE_BENEFICIARIO like '%" & txtNombre.Text & "%'"
End If

If Trim(txtIdPrincipal.Text) <> "" Then
      strSQL = strSQL & " and Solicita like '%" & txtIdPrincipal.Text & "%'"
End If

If Trim(txtChasisNo.Text) <> "" Then
      strSQL = strSQL & " and Solicita_Nombre like '%" & txtChasisNo.Text & "%'"
End If

strSQL = strSQL & " Order by Registra_fecha desc, Beneficio_Desc, Consec desc"

vPaso = True
    Call sbCargaGridLocal(vGrid, 17, strSQL)
vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim i As Long, curMonto As Currency

On Error GoTo vError

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
For i = 1 To vGrid.MaxCols
 vGrid.col = i
 vGrid.Text = ""
Next i

curMonto = 0

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    vGrid.Text = CStr(rs.Fields(i - 1).Value & "")
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  curMonto = curMonto + rs!Monto
  rs.MoveNext
Loop
rs.Close

vGrid.MaxRows = vGrid.MaxRows - 1

txtCasos.Text = Format(vGrid.MaxRows, "###,###,##0")
txtMonto.Text = Format(curMonto, "Standard")

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 17
    vHeaders.Headers(1) = "..."
    vHeaders.Headers(2) = "Código"
    vHeaders.Headers(3) = "Id Beneficio"
    vHeaders.Headers(4) = "Identificación"
    vHeaders.Headers(5) = "Nombre"
    vHeaders.Headers(6) = "Monto"
    vHeaders.Headers(7) = "Estado"
    vHeaders.Headers(8) = "Beneficio"
    vHeaders.Headers(9) = "Solicita Id"
    vHeaders.Headers(10) = "Solicita Nombre"
    vHeaders.Headers(11) = "Fecha Reg"
    vHeaders.Headers(12) = "Usuario Reg"
    vHeaders.Headers(13) = "Fecha Autoriza"
    vHeaders.Headers(14) = "Usuario Autoriza"
    vHeaders.Headers(15) = "Institución"
    vHeaders.Headers(16) = "Departamento"
    vHeaders.Headers(17) = "Oficina"
    
    
 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Beneficios_Monitor")


End Sub

Private Sub chkComercio_Click()
Dim i As Integer

For i = 1 To lswComercio.ListItems.Count
  lswComercio.ListItems.Item(i).Checked = chkComercio.Value
Next i
End Sub

Private Sub sbInicializa()

Me.MousePointer = vbHourglass

   
    'Estados
    strSQL = "select rtrim(cod_estado) as 'IdX', rtrim(descripcion) as ItmX" _
             & " from  afi_Estados_Persona"
    Call sbCbo_Llena_New(cboEstadoPersona, strSQL, True, True)
    
    
    'Tipos de Presentacion
    strSQL = "select rtrim(descripcion) as Itmx, ID_PRESENTACION as Idx" _
           & " from CRD_PRENDAS_PRESENTACION order by descripcion"
    Call sbCbo_Llena_New(cboPresentacion, strSQL, True, True)
    
    'Modelos
    strSQL = "select rtrim(ID_MODELO) as 'IdX', rtrim(descripcion) as ItmX" _
             & " from  CRD_PRENDAS_MODELOS Where Activo = 1 order by Descripcion"
    Call sbCbo_Llena_New(cboModelo, strSQL, True, True)
    
    'Combustible
    strSQL = "select rtrim(ID_COMBUSTIBLE) as 'IdX', rtrim(descripcion) as ItmX" _
             & " from  CRD_PRENDAS_COMBUSTIBLE order by Descripcion"
    Call sbCbo_Llena_New(cboCombustible, strSQL, True, True)
    
    'Unidades
    strSQL = "select rtrim(ID_Unidad) as 'IdX', rtrim(descripcion) as ItmX" _
             & " from  CRD_PRENDAS_uds Where Peso_Apl = 1 and Activa = 1 order by Descripcion"
    Call sbCbo_Llena_New(cboPeso, strSQL, True, True)
    
    strSQL = "select rtrim(ID_Unidad) as 'IdX', rtrim(descripcion) as ItmX" _
             & " from  CRD_PRENDAS_uds Where Capacidad_Apl = 1 and Activa = 1 order by Descripcion"
    Call sbCbo_Llena_New(cboCapacidad, strSQL, True, True)
    
    strSQL = "select rtrim(ID_Unidad) as 'IdX', rtrim(descripcion) as ItmX" _
             & " from  CRD_PRENDAS_uds Where Cilindraje_Apl = 1 and Activa = 1 order by Descripcion"
    Call sbCbo_Llena_New(cboCilindraje, strSQL, True, True)
    
    txtCilindrajeInicio.Text = 0
    txtCilindrajeCorte.Text = 10000
    
    txtCapacidadInicio.Text = 0
    txtCapacidadInicio.Text = 10000
    txtPesoInicio.Text = 0
    txtPesoCorte.Text = 1000
    
    cboPuertas.Clear
    cboPuertas.AddItem "No Aplica"
    cboPuertas.AddItem "1"
    cboPuertas.AddItem "2"
    cboPuertas.AddItem "3"
    cboPuertas.AddItem "4"
    cboPuertas.AddItem "5"
    cboPuertas.AddItem "6"
    cboPuertas.AddItem "7"
    cboPuertas.AddItem "8"
    cboPuertas.AddItem "9"
    cboPuertas.AddItem "10"
    cboPuertas.Text = "No Aplica"
    
vPaso = False

Me.MousePointer = vbDefault

End Sub


Private Sub chkMarca_Click()
Dim i As Integer

For i = 1 To lswMarca.ListItems.Count
  lswMarca.ListItems.Item(i).Checked = chkMarca.Value
Next i
End Sub

Private Sub chkPrenda_Click()
Dim i As Integer

For i = 1 To lswPrenda.ListItems.Count
  lswPrenda.ListItems.Item(i).Checked = chkPrenda.Value
Next i
End Sub

Private Sub Form_Load()

vModulo = 3

lswComercio.ColumnHeaders.Add , , "", 3150
lswPrenda.ColumnHeaders.Add , , "", 3150
lswMarca.ColumnHeaders.Add , , "", 3150

vGrid.AppearanceStyle = fxGridStyle

cboEstado.AddItem "Todos"
cboEstado.AddItem "Tramite"
cboEstado.AddItem "Formalizada"
cboEstado.Text = "Todos"

cboFechas.AddItem "Registro"
cboFechas.AddItem "Actualiza"
cboFechas.AddItem "Formalización"
cboFechas.Text = "Registro"

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("m", -1, dtpCorte.Value)


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next

gbFiltros.Width = Me.Width - gbFiltros.Left
imgBanner.Height = Me.Height

vGrid.Width = Me.Width - (vGrid.Left + 120)
vGrid.Height = Me.Height - (vGrid.top + 280)
End Sub

Private Sub sbComercios_Load()
On Error GoTo vError

Me.MousePointer = vbHourglass

txtFiltraComercio.Text = fxSysCleanTxtInject(txtFiltraComercio.Text)

lswComercio.ListItems.Clear

strSQL = "select ID_COMERCIO as IdX, rtrim(DESCRIPCION) as ItmX" _
       & " from CRD_PRENDAS_COMERCIA " _
       & " where ACTIVA = 1 and descripcion like '%" & txtFiltraComercio.Text & "%'" _
       & " order by descripcion"
       
      
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswComercio.ListItems.Add(, , rs!itmX)
     itmX.Tag = rs!IdX
     itmX.Checked = chkComercio.Value
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False
Call sbInicializa

End Sub

Private Sub txtFiltraComercio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbComercios_Load
End If
End Sub


Private Sub txtUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF4 Then
'
'    gBusquedas.Columna = "descripcion"
'    gBusquedas.Orden = "descripcion"
'    gBusquedas.Consulta = "select cod_departamento,descripcion from AFDepartamentos"
'    gBusquedas.Filtro = " and ID_PRESENTACION = " & cboPresentacion.ItemData(cboPresentacion.ListIndex)
'
'  frmBusquedas.Show vbModal
'  txtUnidad.Tag = gBusquedas.Resultado
'  txtUnidad.Text = gBusquedas.Resultado2
'End If

End Sub

Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub
If vGrid.MaxRows = 0 Then Exit Sub

vGrid.Row = Row

vGrid.col = 4
If Not IsNumeric(vGrid.Text) Then
  Operacion.Operacion = 0
Else
    Operacion.Operacion = vGrid.Text
End If
Call sbFormsCall("frmCR_Prendas", 1, , , False, Me)


End Sub


