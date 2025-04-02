VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_Beneficios_Monitor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Montor de Beneficios"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20025
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   20025
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lswBeneficios 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
      _ExtentY        =   8916
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
      Top             =   6600
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
      Top             =   6960
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
      Top             =   7320
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
      Top             =   7680
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
   Begin XtremeSuiteControls.FlatEdit txtFiltraBeneficios 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1080
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
   Begin XtremeSuiteControls.CheckBox chkBeneficios 
      Height          =   210
      Left            =   2880
      TabIndex        =   6
      Top             =   840
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
      TabIndex        =   12
      Top             =   2520
      Width           =   15135
      _Version        =   524288
      _ExtentX        =   26696
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
      MaxCols         =   18
      SpreadDesigner  =   "frmAF_Beneficios_Monitor.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.GroupBox gbFiltros 
      Height          =   2415
      Left            =   3360
      TabIndex        =   13
      Top             =   0
      Width           =   16815
      _Version        =   1441793
      _ExtentX        =   29660
      _ExtentY        =   4260
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin VB.Timer TimerX 
         Interval        =   5
         Left            =   12240
         Top             =   960
      End
      Begin XtremeSuiteControls.FlatEdit txtUserAutoriza 
         Height          =   330
         Left            =   1920
         TabIndex        =   14
         Top             =   1920
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
      Begin XtremeSuiteControls.FlatEdit txtSolicitaId 
         Height          =   330
         Left            =   1920
         TabIndex        =   15
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
      Begin XtremeSuiteControls.FlatEdit txtBeneficiarioId 
         Height          =   330
         Left            =   1920
         TabIndex        =   16
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
      Begin XtremeSuiteControls.FlatEdit txtBeneficiarioName 
         Height          =   330
         Left            =   5640
         TabIndex        =   17
         Top             =   240
         Width           =   5175
         _Version        =   1441793
         _ExtentX        =   9128
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
      Begin XtremeSuiteControls.ComboBox cboInstitucion 
         Height          =   330
         Left            =   5640
         TabIndex        =   18
         Top             =   1200
         Width           =   5175
         _Version        =   1441793
         _ExtentX        =   9128
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
      Begin XtremeSuiteControls.ComboBox cboOficina 
         Height          =   330
         Left            =   5640
         TabIndex        =   19
         Top             =   1920
         Width           =   5175
         _Version        =   1441793
         _ExtentX        =   9128
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
         TabIndex        =   20
         Top             =   1560
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
         TabIndex        =   21
         Top             =   1560
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
         Picture         =   "frmAF_Beneficios_Monitor.frx":0AE8
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   615
         Left            =   12360
         TabIndex        =   22
         ToolTipText     =   "Exportar a Excel"
         Top             =   1560
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
         Picture         =   "frmAF_Beneficios_Monitor.frx":1506
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBarX 
         Height          =   135
         Left            =   11040
         TabIndex        =   23
         Top             =   1440
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
         TabIndex        =   24
         Top             =   1200
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
      Begin XtremeSuiteControls.FlatEdit txtSolicitaName 
         Height          =   330
         Left            =   5640
         TabIndex        =   25
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   600
         Width           =   5175
         _Version        =   1441793
         _ExtentX        =   9128
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
      Begin XtremeSuiteControls.FlatEdit txtUnidad 
         Height          =   330
         Left            =   5640
         TabIndex        =   36
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1560
         Width           =   5175
         _Version        =   1441793
         _ExtentX        =   9128
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   37
         Top             =   1200
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   35
         Top             =   1920
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario Autoriza"
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
         Index           =   8
         Left            =   360
         TabIndex        =   34
         Top             =   1560
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
         Index           =   7
         Left            =   3960
         TabIndex        =   33
         Top             =   1920
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Oficina"
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
         TabIndex        =   32
         Top             =   1560
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Unidad"
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
         Index           =   5
         Left            =   3960
         TabIndex        =   31
         Top             =   1200
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Institución"
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
         TabIndex        =   30
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
         Index           =   3
         Left            =   3720
         TabIndex        =   29
         Top             =   600
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Solicita Nombre"
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
         TabIndex        =   28
         Top             =   600
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Solicita Id"
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
         TabIndex        =   27
         Top             =   240
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Beneficiario Nombre"
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
         TabIndex        =   26
         Top             =   240
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Beneficiario Id"
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
      TabIndex        =   38
      Top             =   8880
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
      TabIndex        =   39
      Top             =   8880
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
   Begin XtremeSuiteControls.ComboBox cboTipoBeneficio 
      Height          =   345
      Left            =   120
      TabIndex        =   41
      Top             =   360
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
      _ExtentY        =   609
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Beneficios ...:"
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
      TabIndex        =   42
      Top             =   840
      Width           =   1815
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
      TabIndex        =   40
      Top             =   8640
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
      TabIndex        =   11
      Top             =   7320
      Width           =   1215
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
      TabIndex        =   10
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Categoría ...:"
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
      TabIndex        =   9
      Top             =   120
      Width           =   1815
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
      TabIndex        =   8
      Top             =   6600
      Width           =   1215
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
      TabIndex        =   7
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Image imgBanner 
      Height          =   9390
      Left            =   0
      Picture         =   "frmAF_Beneficios_Monitor.frx":1D0B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3285
   End
End
Attribute VB_Name = "frmAF_Beneficios_Monitor"
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


strSQL = "select 0 as 'Btn', ID_BENEFICIO, Cod_Beneficio, Consec, Cedula, NOMBRE_BENEFICIARIO, Monto, Estado_Desc, Beneficio_Desc" _
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

If cboInstitucion.Text <> "TODOS" Then
        strSQL = strSQL & " And Cod_Institucion =  " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If

If cboOficina.Text <> "TODOS" Then
        strSQL = strSQL & " And cod_Oficina =  '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
End If

If cboEstadoPersona.Text <> "TODOS" Then
        strSQL = strSQL & " And EstadoActual =  '" & cboEstadoPersona.ItemData(cboEstadoPersona.ListIndex) & "'"
End If



'Lista de Documentos
If lswBeneficios.ListItems.Count > 0 Then
    vCadena = " and Cod_Beneficio in('"
    For i = 1 To lswBeneficios.ListItems.Count
      If lswBeneficios.ListItems.Item(i).Checked Then
        vCadena = vCadena & "','" & lswBeneficios.ListItems.Item(i).Tag
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



If Trim(txtBeneficiarioId.Text) <> "" Then
      strSQL = strSQL & " and Cedula like '%" & txtBeneficiarioId.Text & "%'"
End If


If Trim(txtBeneficiarioName.Text) <> "" Then
      strSQL = strSQL & " and NOMBRE_BENEFICIARIO like '%" & txtBeneficiarioName.Text & "%'"
End If

If Trim(txtSolicitaId.Text) <> "" Then
      strSQL = strSQL & " and Solicita like '%" & txtSolicitaId.Text & "%'"
End If

If Trim(txtSolicitaName.Text) <> "" Then
      strSQL = strSQL & " and Solicita_Nombre like '%" & txtSolicitaName.Text & "%'"
End If

strSQL = strSQL & " Order by Registra_fecha desc, Beneficio_Desc, Consec desc"

vPaso = True
    Call sbCargaGridLocal(vGrid, 18, strSQL)
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
 vGrid.Col = i
 vGrid.Text = ""
Next i

curMonto = 0

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    vGrid.Text = CStr(rs.Fields(i - 1).Value & "")
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  curMonto = curMonto + rs!MONTO
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
    vHeaders.Columnas = 18
    vHeaders.Headers(1) = "..."
    vHeaders.Headers(2) = "Id Beneficio"
    vHeaders.Headers(3) = "Código"
    vHeaders.Headers(4) = "Seq.Beneficio"
    vHeaders.Headers(5) = "Identificación"
    vHeaders.Headers(6) = "Nombre"
    vHeaders.Headers(7) = "Monto"
    vHeaders.Headers(8) = "Estado"
    vHeaders.Headers(9) = "Beneficio"
    vHeaders.Headers(10) = "Solicita Id"
    vHeaders.Headers(11) = "Solicita Nombre"
    vHeaders.Headers(12) = "Fecha Reg"
    vHeaders.Headers(13) = "Usuario Reg"
    vHeaders.Headers(14) = "Fecha Autoriza"
    vHeaders.Headers(15) = "Usuario Autoriza"
    vHeaders.Headers(16) = "Institución"
    vHeaders.Headers(17) = "Departamento"
    vHeaders.Headers(18) = "Oficina"
    
    
 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Beneficios_Monitor")


End Sub

Private Sub cboTipoBeneficio_Click()
If vPaso Then Exit Sub
Call sbBeneficiosTipos_Load
End Sub

Private Sub chkBeneficios_Click()
Dim i As Integer

For i = 1 To lswBeneficios.ListItems.Count
  lswBeneficios.ListItems.Item(i).Checked = chkBeneficios.Value
Next i
End Sub

Private Sub sbInicializa()

Me.MousePointer = vbHourglass

    'Instituciones
    strSQL = "select rtrim(descripcion) as Itmx, cod_institucion as Idx" _
           & " from instituciones order by descripcion"
    Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)

   
    'Estados
    strSQL = "select rtrim(cod_estado) as 'IdX', rtrim(descripcion) as ItmX" _
             & " from  afi_Estados_Persona"
    Call sbCbo_Llena_New(cboEstadoPersona, strSQL, True, True)
    
    
    'Oficinas
    strSQL = "select rtrim(cod_Oficina) as 'IdX', rtrim(descripcion) as ItmX" _
             & " from  SIF_Oficinas order by Descripcion"
    Call sbCbo_Llena_New(cboOficina, strSQL, True, True)
    
vPaso = False

Me.MousePointer = vbDefault

End Sub


Private Sub Form_Load()

vModulo = 7

lswBeneficios.ColumnHeaders.Add , , "", 3150

vGrid.AppearanceStyle = fxGridStyle

cboEstado.AddItem "Todos"
cboEstado.AddItem "Ejecutado"
cboEstado.AddItem "Solicitado"
cboEstado.AddItem "Rechazado"
cboEstado.AddItem "Autorizado"
cboEstado.Text = "Todos"

cboFechas.AddItem "Registro"
cboFechas.AddItem "Autorización"
cboFechas.AddItem "Pago"
cboFechas.Text = "Registro"

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("m", -1, dtpCorte.Value)

vPaso = True
    strSQL = "select cod_categoria as 'IdX', descripcion as 'ItmX'" _
           & " From afi_bene_categorias  where Activo = 1 order by descripcion"
    Call sbCbo_Llena_New(cboTipoBeneficio, strSQL, True, True)
vPaso = False

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next

gbFiltros.Width = Me.Width - gbFiltros.Left
imgBanner.Height = Me.Height

vGrid.Width = Me.Width - (vGrid.Left + 120)
vGrid.Height = Me.Height - (vGrid.Top + 280)
End Sub

Private Sub sbBeneficiosTipos_Load()
On Error GoTo vError

Me.MousePointer = vbHourglass

txtFiltraBeneficios.Text = fxSysCleanTxtInject(txtFiltraBeneficios.Text)

lswBeneficios.ListItems.Clear

If cboTipoBeneficio.Text = "TODOS" Then

    strSQL = "select COD_BENEFICIO as IdX, rtrim(DESCRIPCION) as ItmX" _
           & " from AFI_BENEFICIOS " _
           & " where Estado = 'A' and descripcion like '%" & txtFiltraBeneficios.Text & "%'" _
           & " order by descripcion"
Else

    strSQL = "select COD_BENEFICIO as IdX, rtrim(DESCRIPCION) as ItmX" _
           & " from AFI_BENEFICIOS " _
           & " where Estado = 'A' and descripcion like '%" & txtFiltraBeneficios.Text & "%'" _
           & " and cod_Categoria = '" & cboTipoBeneficio.ItemData(cboTipoBeneficio.ListIndex) & "'" _
           & " order by descripcion"
End If
      
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswBeneficios.ListItems.Add(, , rs!itmX)
     itmX.Tag = rs!IdX
     itmX.Checked = chkBeneficios.Value
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

Private Sub txtFiltraBeneficios_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbBeneficiosTipos_Load
End If
End Sub


Private Sub txtUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then

    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_departamento,descripcion from AFDepartamentos"
    gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)

  frmBusquedas.Show vbModal
  txtUnidad.Tag = gBusquedas.Resultado
  txtUnidad.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub
If vGrid.MaxRows = 0 Then Exit Sub

Dim pCedula As String, pCodigo As String, pId As Long

vGrid.Row = Row
vGrid.Col = 5
pCedula = vGrid.Text

vGrid.Col = 4
GLOBALES.gTag = vGrid.Text
vGrid.Col = 3
GLOBALES.gTag2 = vGrid.Text
Call sbFormsCall("frmAF_BeneficioAsg", , , , False, Me, True)

'Dim frm As Form
'
'Call sbFormActivo("frmAF_BeneficioAsg", frm)
'Call frm.sbConsultaX(pCedula)



End Sub
