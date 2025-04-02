VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.ShortcutBar.v22.0.0.ocx"
Begin VB.Form frmCO_CJ_ControlCasos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Control de Casos"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19260
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   19260
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lswJuicios 
      Height          =   1815
      Left            =   120
      TabIndex        =   61
      Top             =   4320
      Width           =   3015
      _Version        =   1441792
      _ExtentX        =   5318
      _ExtentY        =   3201
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
      View            =   2
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   6
   End
   Begin XtremeSuiteControls.GroupBox fraReportes 
      Height          =   2775
      Left            =   8280
      TabIndex        =   52
      Top             =   1320
      Visible         =   0   'False
      Width           =   6735
      _Version        =   1441792
      _ExtentX        =   11880
      _ExtentY        =   4895
      _StockProps     =   79
      Caption         =   "Reportes"
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Begin XtremeSuiteControls.RadioButton rbReportes 
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   54
         Top             =   720
         Width           =   3375
         _Version        =   1441792
         _ExtentX        =   5953
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Listado General"
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
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbReportes 
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   55
         Top             =   1080
         Width           =   3375
         _Version        =   1441792
         _ExtentX        =   5953
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Agrupado por Usuario"
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
      Begin XtremeSuiteControls.RadioButton rbReportes 
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   56
         Top             =   1440
         Width           =   3375
         _Version        =   1441792
         _ExtentX        =   5953
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Agrupado por Oficina"
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
      Begin XtremeSuiteControls.ComboBox cboTipoReporte 
         Height          =   330
         Left            =   720
         TabIndex        =   58
         Top             =   2160
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
      Begin XtremeSuiteControls.PushButton btnInformes 
         Height          =   495
         Index           =   0
         Left            =   4080
         TabIndex        =   59
         Top             =   2160
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Informe"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCO_CJ_ControlCasos.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnInformes 
         Height          =   495
         Index           =   1
         Left            =   5280
         TabIndex        =   60
         Top             =   2160
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Cerrar"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCO_CJ_ControlCasos.frx":0707
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   13
         Left            =   720
         TabIndex        =   57
         Top             =   1920
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo"
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
      Begin XtremeShortcutBar.ShortcutCaption scTitulos 
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Width           =   7215
         _Version        =   1441792
         _ExtentX        =   12726
         _ExtentY        =   873
         _StockProps     =   14
         Caption         =   "Informes"
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
   End
   Begin XtremeSuiteControls.GroupBox fraFilros 
      Height          =   6015
      Left            =   4560
      TabIndex        =   27
      Top             =   1320
      Visible         =   0   'False
      Width           =   7215
      _Version        =   1441792
      _ExtentX        =   12726
      _ExtentY        =   10610
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Begin XtremeSuiteControls.PushButton btnFiltroCerrar 
         Height          =   495
         Left            =   5880
         TabIndex        =   51
         Top             =   5160
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Cerrar"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCO_CJ_ControlCasos.frx":0E1D
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuarioRegistra 
         Height          =   330
         Left            =   4200
         TabIndex        =   38
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1920
         Width           =   2895
         _Version        =   1441792
         _ExtentX        =   5106
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
      Begin XtremeSuiteControls.FlatEdit txtLinea 
         Height          =   330
         Left            =   1320
         TabIndex        =   37
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1920
         Width           =   2895
         _Version        =   1441792
         _ExtentX        =   5106
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
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   330
         Left            =   1320
         TabIndex        =   31
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   600
         Width           =   5775
         _Version        =   1441792
         _ExtentX        =   10186
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
      Begin XtremeSuiteControls.FlatEdit txtTramite 
         Height          =   330
         Left            =   1320
         TabIndex        =   33
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1320
         Width           =   2895
         _Version        =   1441792
         _ExtentX        =   5106
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
      Begin XtremeSuiteControls.FlatEdit txtExpediente 
         Height          =   330
         Left            =   4200
         TabIndex        =   34
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1320
         Width           =   2895
         _Version        =   1441792
         _ExtentX        =   5106
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
      Begin XtremeSuiteControls.ComboBox cboGarantia 
         Height          =   330
         Left            =   1320
         TabIndex        =   41
         Top             =   2640
         Width           =   2895
         _Version        =   1441792
         _ExtentX        =   5106
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
      Begin XtremeSuiteControls.ComboBox cboEstadoPersona 
         Height          =   330
         Left            =   4200
         TabIndex        =   42
         Top             =   2640
         Width           =   2895
         _Version        =   1441792
         _ExtentX        =   5106
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
         Left            =   1320
         TabIndex        =   44
         Top             =   3120
         Width           =   5775
         _Version        =   1441792
         _ExtentX        =   10186
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
      Begin XtremeSuiteControls.ComboBox cboCartera 
         Height          =   330
         Left            =   1320
         TabIndex        =   46
         Top             =   3600
         Width           =   5775
         _Version        =   1441792
         _ExtentX        =   10186
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
      Begin XtremeSuiteControls.ComboBox cboEmbargables 
         Height          =   330
         Left            =   1320
         TabIndex        =   48
         Top             =   4080
         Width           =   5775
         _Version        =   1441792
         _ExtentX        =   10186
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
      Begin XtremeSuiteControls.ComboBox cboGastos 
         Height          =   330
         Left            =   1320
         TabIndex        =   50
         Top             =   4560
         Width           =   5775
         _Version        =   1441792
         _ExtentX        =   10186
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   49
         Top             =   4560
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Gastos Registrados"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   47
         Top             =   4080
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Embargables"
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   45
         Top             =   3600
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cartera"
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   43
         Top             =   3120
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   8
         Left            =   4200
         TabIndex        =   40
         Top             =   2400
         Width           =   2775
         _Version        =   1441792
         _ExtentX        =   4895
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Estado de la Persona"
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
         Left            =   1320
         TabIndex        =   39
         Top             =   2400
         Width           =   1575
         _Version        =   1441792
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Garantía"
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
         Index           =   6
         Left            =   4200
         TabIndex        =   36
         Top             =   1680
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario"
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
         Left            =   1320
         TabIndex        =   35
         Top             =   1680
         Width           =   1575
         _Version        =   1441792
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Línea de Crédito"
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
         Index           =   4
         Left            =   4200
         TabIndex        =   32
         Top             =   1080
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Expediente"
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
         Left            =   1320
         TabIndex        =   30
         Top             =   1080
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Trámite"
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
         Index           =   2
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
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
         Transparent     =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulos 
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   7215
         _Version        =   1441792
         _ExtentX        =   12726
         _ExtentY        =   873
         _StockProps     =   14
         Caption         =   "Filtros adicionales"
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
   End
   Begin XtremeSuiteControls.CheckBox chkAbogados 
      Height          =   255
      Left            =   13080
      TabIndex        =   16
      Top             =   480
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Todos"
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
      Value           =   1
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8085
      Width           =   19260
      _ExtentX        =   33973
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Casos Encontrados..:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Total Registrado..:"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6372
      Left            =   3480
      TabIndex        =   7
      Top             =   1320
      Width           =   12612
      _Version        =   524288
      _ExtentX        =   22246
      _ExtentY        =   11240
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
      MaxCols         =   19
      SpreadDesigner  =   "frmCO_CJ_ControlCasos.frx":1533
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1200
      TabIndex        =   8
      Top             =   6960
      Width           =   1932
      _Version        =   1441792
      _ExtentX        =   3408
      _ExtentY        =   550
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
      Height          =   312
      Left            =   1200
      TabIndex        =   9
      Top             =   7320
      Width           =   1932
      _Version        =   1441792
      _ExtentX        =   3408
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.CheckBox chkJuzgados 
      Height          =   204
      Left            =   2880
      TabIndex        =   10
      Top             =   120
      Width           =   204
      _Version        =   1441792
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "Todos"
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
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkJuicios 
      Height          =   204
      Left            =   2880
      TabIndex        =   11
      Top             =   3960
      Width           =   204
      _Version        =   1441792
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "Todos"
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
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   204
      Left            =   2880
      TabIndex        =   12
      Top             =   7680
      Width           =   204
      _Version        =   1441792
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "Todos"
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
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   375
      Left            =   8400
      TabIndex        =   13
      Top             =   0
      Width           =   1215
      _Version        =   1441792
      _ExtentX        =   2138
      _ExtentY        =   656
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
      Appearance      =   16
      Picture         =   "frmCO_CJ_ControlCasos.frx":1FFB
   End
   Begin XtremeSuiteControls.PushButton btnInforme 
      Height          =   375
      Left            =   9600
      TabIndex        =   14
      Top             =   0
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Informe"
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
      Picture         =   "frmCO_CJ_ControlCasos.frx":26FB
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   375
      Left            =   11160
      TabIndex        =   15
      Top             =   0
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2773
      _ExtentY        =   656
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
      Appearance      =   16
      Picture         =   "frmCO_CJ_ControlCasos.frx":2E02
   End
   Begin XtremeSuiteControls.CheckBox chkBufete 
      Height          =   255
      Left            =   13080
      TabIndex        =   17
      Top             =   840
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Todos"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkFiltros 
      Height          =   255
      Left            =   5280
      TabIndex        =   18
      Top             =   120
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "+ Filtros ?"
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
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtAbogado 
      Height          =   330
      Left            =   8400
      TabIndex        =   21
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   480
      Width           =   4575
      _Version        =   1441792
      _ExtentX        =   8070
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtBufete 
      Height          =   330
      Left            =   8400
      TabIndex        =   22
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   840
      Width           =   4575
      _Version        =   1441792
      _ExtentX        =   8070
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   330
      Left            =   3480
      TabIndex        =   23
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
   Begin XtremeSuiteControls.FlatEdit txtBuscarPor 
      Height          =   330
      Left            =   3480
      TabIndex        =   24
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   840
      Width           =   3375
      _Version        =   1441792
      _ExtentX        =   5953
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
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   330
      Left            =   1200
      TabIndex        =   25
      Top             =   6240
      Width           =   1935
      _Version        =   1441792
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.ComboBox cboProceso 
      Height          =   330
      Left            =   1200
      TabIndex        =   26
      Top             =   6600
      Width           =   1935
      _Version        =   1441792
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.ListView lswJuzgados 
      Height          =   3255
      Left            =   120
      TabIndex        =   62
      Top             =   480
      Width           =   3015
      _Version        =   1441792
      _ExtentX        =   5318
      _ExtentY        =   5741
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
      View            =   2
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   6
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   7080
      TabIndex        =   20
      Top             =   840
      Width           =   1215
      _Version        =   1441792
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Firma/Bufete"
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
      Index           =   0
      Left            =   7080
      TabIndex        =   19
      Top             =   480
      Width           =   1215
      _Version        =   1441792
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Abogado"
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
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
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Juzgados ...:"
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
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Juicio...:"
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
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
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
      Height          =   315
      Index           =   6
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Proceso"
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
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Image imgBanner 
      Height          =   9396
      Left            =   0
      Picture         =   "frmCO_CJ_ControlCasos.frx":36D3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3200
   End
End
Attribute VB_Name = "frmCO_CJ_ControlCasos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String
Dim vTipoDocu As String
Dim vScroll As Boolean
Dim strSQLReporte As String
Dim strDetalle As String


Private Sub btnBuscar_Click()
Call sbBuscar
End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders, i As Integer

    vHeaders.Columnas = 19
    vHeaders.Headers(1) = ""
    vHeaders.Headers(2) = "No.Tramite"
    vHeaders.Headers(3) = "Operación"
    vHeaders.Headers(4) = "Linea"
    vHeaders.Headers(5) = "Identificación"
    vHeaders.Headers(6) = "Nombre"
    vHeaders.Headers(7) = "Plazo Restante"
    vHeaders.Headers(8) = "Total Deuda"
    vHeaders.Headers(9) = "Usuario"
    vHeaders.Headers(10) = "Fecha"
    vHeaders.Headers(11) = "Estado"
    vHeaders.Headers(12) = "Abogado"
    vHeaders.Headers(13) = "Juzgado"
    vHeaders.Headers(14) = "Juicio"
    vHeaders.Headers(15) = "Oficina"
    vHeaders.Headers(16) = "Garantía"
    vHeaders.Headers(17) = "Estado Persona"
    vHeaders.Headers(18) = "Estado Laboral"
    vHeaders.Headers(19) = "Ult. Proceso"
    
    
   Call sbSIFGridExportar(vGrid, vHeaders, "Cobro_CJ_Tramite")

End Sub

Private Sub btnFiltroCerrar_Click()
fraFilros.Visible = False
chkFiltros.Value = vbUnchecked
End Sub

Private Sub btnInforme_Click()
    fraReportes.Top = btnBuscar.Top + 390
    fraReportes.Left = btnBuscar.Left - 1300
    fraReportes.Visible = True
End Sub

Private Sub btnInformes_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFiltros As String, vFiltrosEtiquetas As String

With frmContenedor.Crt
    .Reset
    .WindowTitle = "Reportes del Módulo: Cobro Judicial"
    .WindowState = crptMaximized
    .WindowShowGroupTree = True

    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "Usuario = 'Usuario..:" & glogon.Usuario & "'"
    .Formulas(3) = "Fecha = 'Fecha ...:" & fxFechaServidor & "'"
    .Connect = glogon.ConectRPT


    vFiltros = fxReportesFiltros


Select Case Index
    Case 0 'Imprimir
        Select Case True
            Case rbReportes.Item(0).Value 'Reporte General
'                If Mid(cboTipoReporte, 1, 1) = "D" Then
                   .ReportFileName = SIFGlobal.fxPathReportes("Cobro_Judicial_TramiteDetalle.rpt")
'                Else
'                   .ReportFileName = SIFGlobal.fxPathReportes("SIFDocGeneralRsm.rpt")
'                End If
            Case rbReportes.Item(1).Value 'Agrupado por Usuario
                If Mid(cboTipoReporte, 1, 1) = "D" Then
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_Judicial_TramiteDetallexUsuario.rpt")
                Else
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_Judicial_TramiteResumenxUsuario.rpt")
                End If
            
            Case rbReportes.Item(3).Value 'Agrupado por Oficina
                If Mid(cboTipoReporte, 1, 1) = "D" Then
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_Judicial_TramiteDetallexOficina.rpt")
                Else
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_Judicial_TramiteResumenxOficina.rpt")
                End If

        End Select
'
'        'Si son los reportes Especiales no aplicar filtros default
'        If Not (optReportes.Item(5).Value Or optReportes.Item(6).Value) Then
'            .SelectionFormula = vFiltros
'        End If
        .PrintReport
'.Action = 1

 Case 1 'Cerrar
      fraReportes.Visible = False

End Select

End With

End Sub

Private Sub chkAbogados_Click()
If chkAbogados.Value = vbUnchecked Then
    txtAbogado.Enabled = True
    txtAbogado.SetFocus
Else
    txtAbogado.Enabled = False
    txtAbogado.Text = ""
End If

End Sub

Private Sub chkBufete_Click()
If chkBufete.Value = vbUnchecked Then
    txtBufete.Enabled = True
    txtBufete.SetFocus
Else
    txtBufete.Enabled = False
    txtBufete.Text = ""
End If
End Sub


Private Sub chkFechas_Click()
If chkFechas.Value = vbChecked Then
   dtpCorte.Enabled = False
   dtpInicio.Enabled = False
Else
   dtpCorte.Enabled = True
   dtpInicio.Enabled = True
End If
End Sub

Private Sub chkFiltros_Click()
If chkFiltros.Value = vbChecked Then
    fraFilros.Visible = True
Else
    fraFilros.Visible = False
End If
End Sub

Private Sub chkJuicios_Click()
Dim i As Integer

For i = 1 To lswJuicios.ListItems.Count
  lswJuicios.ListItems.Item(i).Checked = chkJuicios.Value
Next i
End Sub

Private Sub chkJuzgados_Click()
Dim i As Integer

For i = 1 To lswJuzgados.ListItems.Count
  lswJuzgados.ListItems.Item(i).Checked = chkJuzgados.Value
Next i

End Sub

Private Sub sbBuscar()
Dim strSQL As String, i As Integer
Dim vCadena As String, iCantidad As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
iCantidad = 0



If cboEstado.Text <> "Pendiente" Then
    
    strSQL = "select '',COD_TRAMITE,ID_SOLICITUD,Linea" _
            & ",Persona_Id,PERSONA_NOMBRE,dbo.fxCrdPlazoRestante(PLAZO,PRIDEDUC," & GLOBALES.glngFechaCR & ") as 'Plazo'" _
            & ",isnull(TOTAL_DEUDA, Saldo) as 'Monto' ,PROCESO_USUARIO,PROCESO_FECHA" _
            & ",ProcesoDesc,Abogado,Juzgado, TipoJuicio, Oficina,GarantiaDesc" _
            & ",Persona_Estado,Persona_EstadoLaboral,CJ_ProcesoDesc" _
            & " from vCbrCjOperaciones"
    
    ''Lista de juzgados
    If chkJuzgados.Value = vbChecked Then
        vCadena = " Cod_Juzgado in('"
        For i = 1 To lswJuzgados.ListItems.Count
          If lswJuzgados.ListItems.Item(i).Checked Then
            vCadena = vCadena & "','" & lswJuzgados.ListItems.Item(i).Tag
            iCantidad = iCantidad + 1
          End If
        Next i
        
        If iCantidad > 0 Then strSQL = strSQL & " where " & vCadena & "')"
        
    End If
    '
    iCantidad = 0
    ''Lista de Juicios
    If chkJuicios.Value = vbChecked Then
        vCadena = " Tipo_juicio in('"
        For i = 1 To lswJuicios.ListItems.Count
          If lswJuicios.ListItems.Item(i).Checked Then
            vCadena = vCadena & "','" & lswJuicios.ListItems.Item(i).Tag
            iCantidad = iCantidad + 1
          End If
        Next i
        
        If iCantidad > 0 Then
          If InStr(1, strSQL, "where") <= 0 Then
            strSQL = strSQL & " where " & vCadena & "')"
          Else
            strSQL = strSQL & " and " & vCadena & "')"
          End If
        End If
    End If
    
    'Validación De las Fechas
    If chkFechas.Value = vbUnchecked Then
        If InStr(1, strSQL, "where") <= 0 Then
    
            strSQL = strSQL & " where  Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                            & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
        Else
            strSQL = strSQL & " and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                            & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
        End If
    End If
    
    
    Select Case cboEstado.Text
     Case "Reversada"
        If InStr(1, strSQL, "where") <= 0 Then
            strSQL = strSQL & " where Proceso = 'N'"
        Else
            strSQL = strSQL & " and Proceso = 'N'"
        End If
     Case "En Proceso"
        If InStr(1, strSQL, "where") <= 0 Then
            strSQL = strSQL & " where Proceso = 'J'"
        Else
            strSQL = strSQL & " and Proceso = 'J'"
        End If
    End Select
     
    If UCase(cboProceso) <> "TODOS" Then
     If InStr(1, strSQL, "where") <= 0 Then
         strSQL = strSQL & " where COD_PROCESO =  '" & SIFGlobal.fxCodText(cboProceso.Text) & "'"
      Else
         strSQL = strSQL & " and COD_PROCESO=  '" & SIFGlobal.fxCodText(cboProceso.Text) & "' "
      End If
    End If
    
    If chkAbogados.Value = vbUnchecked And Trim(txtAbogado.Text) <> "" Then
       If InStr(1, strSQL, "where") <= 0 Then
          strSQL = strSQL & " where cod_abogado = " & txtAbogado.Tag & " "
       Else
          strSQL = strSQL & " and cod_abogado = " & txtAbogado.Tag & " "
       End If
    End If
    
    If chkBufete.Value = vbUnchecked And Trim(txtBufete.Text) <> "" Then
       If InStr(1, strSQL, "where") <= 0 Then
          strSQL = strSQL & " Where cod_bufete = '" & txtBufete.Tag & "' "
       Else
          strSQL = strSQL & " and cod_bufete = '" & txtBufete.Tag & "' "
       End If
    End If
    

'Filtros Adicionales
    If Trim(txtBuscarPor.Text) <> "" Then
        Select Case Mid(cbo.Text, 1, 2)
            Case "01"
              strSQL = strSQL & " and cedula like '%" & Trim(txtBuscarPor.Text) & "%'"
            Case "02"
              strSQL = strSQL & " and convert(varchar(30),id_solicitud) like '%" & Trim(txtBuscarPor.Text) & "%'"
        End Select
    End If

   
  'Filtros a Nivel General
  If Trim(txtLinea.Text) <> "" Then
      strSQL = strSQL & " and codigo like '%" & Trim(txtLinea.Text) & "%'"
  End If
      
  If Trim(txtNombre.Text) <> "" Then
      strSQL = strSQL & " and Persona_Nombre like '%" & Trim(txtNombre.Text) & "%'"
  End If
    
  If cboGarantia.Text <> "TODOS" Then
      strSQL = strSQL & " and Garantia = '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'"
  End If
  
  If cboEstadoPersona.Text <> "TODOS" Then
      strSQL = strSQL & " and Persona_EstadoId = '" & cboEstadoPersona.ItemData(cboEstadoPersona.ListIndex) & "'"
  End If
  
  If cboOficina.Text <> "TODOS" Then
      strSQL = strSQL & " and COD_OFICINA = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
  End If
  
  
  If cboCartera.Text <> "TODOS" Then
      strSQL = strSQL & " and dbo.fxCBR_CJ_Tramite_ExisteCartera(Codigo,'" & cboCartera.ItemData(cboCartera.ListIndex) & "') >= 1"
  End If
  
  
  
  'Filtros a nivel de Tramite
  If IsNumeric(txtTramite.Text) Then
      strSQL = strSQL & " and convert(varchar(20),Cod_tramite)  like '%" & Trim(txtTramite.Text) & "%'"
  End If
    
  If Trim(txtExpediente.Text) <> "" Then
      strSQL = strSQL & " and expediente_numero like '%" & Trim(txtExpediente.Text) & "%'"
  End If
  
  If Trim(txtUsuarioRegistra.Text) <> "" Then
      strSQL = strSQL & " and Registro_Usuario like '%" & Trim(txtUsuarioRegistra.Text) & "%'"
  End If
  
  If cboGastos.Text <> "TODOS" Then
      strSQL = strSQL & " and dbo.fxCBR_CJ_Tramite_ExisteGasto(Cod_Tramite,'" & cboGastos.ItemData(cboGastos.ListIndex) & "') >= 1"
  End If
  
  If cboEmbargables.Text <> "TODOS" Then
      strSQL = strSQL & " and dbo.fxCBR_CJ_Tramite_ExisteEmbargable(Cod_Tramite,'" & cboEmbargables.ItemData(cboEmbargables.ListIndex) & "') >= 1"
  End If
  

  
    
Else 'Si son los pendientes
    
    strSQL = "select '','',ID_SOLICITUD,CODIGO" _
            & ",CEDULA,Persona_Nombre,dbo.fxCrdPlazoRestante(PLAZO,PRIDEDUC," & GLOBALES.glngFechaCR & " ) as 'Plazo'" _
            & ",Saldo as 'Monto' ,'','',ProcesoDesc,'','','',Oficina,GarantiaDesc" _
            & ",Persona_Estado,Persona_EstadoLaboral,''" _
            & " from vCbrCjOperacionesPendientes" _
            & " where  Proceso = 'J'"
    
    
    'Filtros Adicionales
    If Trim(txtBuscarPor.Text) <> "" Then
        Select Case Mid(cbo.Text, 1, 2)
            Case "01"
              strSQL = strSQL & " and cedula like '%" & Trim(txtBuscarPor.Text) & "%'"
            Case "02"
              strSQL = strSQL & " and convert(varchar(30),id_solicitud) like '%" & Trim(txtBuscarPor.Text) & "%'"
        End Select
    End If
    
    If chkFechas.Value = vbUnchecked Then
        strSQL = strSQL & " and  FECHA_ENVIAPROCESO between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                            & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
    End If

  'Filtros a Nivel General
  If Trim(txtLinea.Text) <> "" Then
      strSQL = strSQL & " and codigo like '%" & Trim(txtLinea.Text) & "%'"
  End If
      
  If Trim(txtNombre.Text) <> "" Then
      strSQL = strSQL & " and Persona_Nombre like '%" & Trim(txtNombre.Text) & "%'"
  End If
    
  If cboGarantia.Text <> "TODOS" Then
      strSQL = strSQL & " and Garantia = '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'"
  End If
  
  If cboEstadoPersona.Text <> "TODOS" Then
      strSQL = strSQL & " and Persona_EstadoId = '" & cboEstadoPersona.ItemData(cboEstadoPersona.ListIndex) & "'"
  End If

  If cboOficina.Text <> "TODOS" Then
      strSQL = strSQL & " and COD_OFICINA = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
  End If
  
  
  If cboCartera.Text <> "TODOS" Then
      strSQL = strSQL & " and dbo.fxCBR_CJ_Tramite_ExisteCartera(Codigo,'" & cboCartera.ItemData(cboCartera.ListIndex) & "') >= 1"
  End If

End If

Call sbCargaGridLocal(vGrid, 19, strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer
Dim curMonto As Currency

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

    If rs.Fields(i - 1).Type = 135 Then
        If Year(rs.Fields(i - 1).Value) > 1900 Then
           vGrid.Text = Format((rs.Fields(i - 1).Value & ""), "dd/mm/yyyy")
        End If
    Else
        vGrid.Text = CStr(rs.Fields(i - 1).Value & "")
    End If

  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  curMonto = curMonto + rs!Monto
  rs.MoveNext
Loop
rs.Close

vGrid.MaxRows = vGrid.MaxRows - 1

StatusBarX.Panels(1).Text = "Total Registros " & vGrid.MaxRows
StatusBarX.Panels(2).Text = "Monto ..: " & Format(curMonto, "Standard")

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxFechaReportes(vTipo As Integer) As String

fxFechaReportes = " in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ")" _
                & " to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"

End Function

Private Function fxUsuarioNombre(vUsuario As String) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select descripcion from usuarios where nombre = '" & vUsuario & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
 fxUsuarioNombre = "[SIN DESCRIPCION]"
Else
 fxUsuarioNombre = "[" & UCase(Trim(rs!Descripcion)) & "]"

End If
rs.Close
End Function

Private Sub Form_Activate()
 vModulo = 6
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

vModulo = 6

Call Formularios(Me)
Call RefrescaTags(Me)

vGrid.AppearanceStyle = fxGridStyle


cbo.Clear
cbo.AddItem "01 - Cédula"
cbo.AddItem "02 - Operación"
cbo.Text = "01 - Cédula"

cboTipoReporte.Clear
cboTipoReporte.AddItem "Detalle"
cboTipoReporte.AddItem "Resumen"
cboTipoReporte.Text = "Detalle"

cboEstado.Clear
cboEstado.AddItem "Pendiente"
cboEstado.AddItem "En Proceso"
cboEstado.AddItem "Reversada"
cboEstado.AddItem "[TODOS]"
cboEstado.Text = "[TODOS]"

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -30, dtpCorte.Value)

vGrid.MaxRows = 0


lswJuzgados.ListItems.Clear
strSQL = "select cod_Juzgado as IdX, rtrim(Nombre) as ItmX from cbr_cj_Juzgados where Activo = 1 order by Nombre"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswJuzgados.ListItems.Add(, , rs!itmX)
     itmX.Tag = rs!IdX
     itmX.Checked = chkJuzgados.Value
 rs.MoveNext
Loop
rs.Close

lswJuicios.ListItems.Clear
strSQL = "select Tipo_Juicio,DESCRIPCION from cbr_cj_Tipos_Juicios where Activo = 1 order by descripcion"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswJuicios.ListItems.Add(, , rs!Descripcion)
     itmX.Tag = rs!TIPO_JUICIO
     itmX.Checked = chkJuicios.Value
 rs.MoveNext
Loop
rs.Close


strSQL = "select cod_proceso as 'IdX', rtrim(Descripcion) as itmx from cbr_cj_Proceso  where activo = 1 order by descripcion"
Call sbCbo_Llena_New(cboProceso, strSQL, True, True)

strSQL = "select RTRIM(TIPO_GASTO) as 'IdX', rtrim(DESCRIPCION) AS 'ItmX' from CBR_CJ_TIPOS_GASTOS where Activo = 1 order by Tipo_Gasto"
Call sbCbo_Llena_New(cboGastos, strSQL, True, True)

strSQL = " select RTRIM(COD_EMBARGABLE ) as 'IdX', rtrim(DESCRIPCION) AS 'ItmX' from CBR_CJ_EMBARGABLES  where Activo = 1 ORDER BY COD_EMBARGABLE "
Call sbCbo_Llena_New(cboEmbargables, strSQL, True, True)

strSQL = " select RTRIM(COD_ESTADO ) as 'IdX', rtrim(DESCRIPCION) AS 'ItmX' from AFI_ESTADOS_PERSONA   where Activo = 1 ORDER BY COD_ESTADO "
Call sbCbo_Llena_New(cboEstadoPersona, strSQL, True, True)

strSQL = " select RTRIM(GARANTIA) as 'IdX', rtrim(DESCRIPCION) AS 'ItmX'  from CRD_GARANTIA_TIPOS  ORDER BY GARANTIA "
Call sbCbo_Llena_New(cboGarantia, strSQL, True, True)

strSQL = " select RTRIM(cod_oficina) as  'IdX',  rtrim(DESCRIPCION) AS 'ItmX' from SIF_OFICINAS" _
       & " Where Estado = 1 ORDER BY COD_OFICINA"
Call sbCbo_Llena_New(cboOficina, strSQL, True, True)

strSQL = " select RTRIM(COD_CLASIFICACION) as  'IdX', rtrim(DESCRIPCION) AS 'ItmX' from CBR_CLASIFICACION_CARTERA" _
       & " where estado = 1    ORDER BY COD_CLASIFICACION"
Call sbCbo_Llena_New(cboCartera, strSQL, True, True)

End Sub

Private Sub Form_Resize()
On Error Resume Next


vGrid.Width = Me.Width - (vGrid.Left + 350)
vGrid.Height = Me.Height - 2100
lswJuicios.Height = Me.Height - 7200


Label1(6).Top = lswJuicios.Top + lswJuicios.Height + 180
cboEstado.Top = Label1(6).Top
Label1(1).Top = Label1(6).Top + Label1(6).Height + 60
cboProceso.Top = Label1(1).Top

Label1(4).Top = cboProceso.Top + cboProceso.Height + 60
dtpInicio.Top = Label1(4).Top
Label1(5).Top = dtpInicio.Top + dtpInicio.Height + 60
dtpCorte.Top = Label1(5).Top
chkFechas.Top = dtpCorte.Top + dtpInicio.Height + 60

imgBanner.Height = Me.Height


End Sub



Private Function fxReportesFiltros() As String
Dim vFiltro As String
Dim vCadena As String, iCantidad As Integer, i As Integer


On Error GoTo vError

iCantidad = 0
vFiltro = ""
strDetalle = ""

If txtBuscarPor <> "" Then
    Select Case Mid(cbo, 1, 2)
        Case "01"
            vFiltro = vFiltro & "{vSIFDocumentos.cliente_Identificacion}  like '*" & txtBuscarPor.Text & "*' "
            strDetalle = "Id. cliente .: " & txtBuscarPor.Text
        Case "02"

            vFiltro = vFiltro & "{vSIFDocumentos.Cliente_Nombre}  like '*" & txtBuscarPor.Text & "*' "
            strDetalle = "Nombre cliente .: " & txtBuscarPor.Text
        Case "03"

            vFiltro = vFiltro & "{vSIFDocumentos.Cod_Transaccion} = '" & txtBuscarPor.Text & "' "
            strDetalle = "Transacción .: " & txtBuscarPor.Text
        Case "04"

             vFiltro = vFiltro & "{vSIFDocumentos.Documento} = '" & txtBuscarPor.Text & "' "
             strDetalle = "Documento .: " & txtBuscarPor.Text
        Case "05"

            vFiltro = vFiltro & "{vSIFDocumentos.Registro_Usuario}  like '" & txtBuscarPor.Text & "' "
            strDetalle = "Usuario .: " & txtBuscarPor.Text
    End Select
End If 'txtBuscarPor

'Lista de Documentos
vCadena = ""
For i = 1 To lswJuzgados.ListItems.Count
  If lswJuzgados.ListItems.Item(i).Checked Then
    If vCadena <> "" Then
        vCadena = vCadena & ","
    End If
    vCadena = vCadena & "'" & lswJuzgados.ListItems.Item(i).Tag & "'"
    iCantidad = iCantidad + 1
  End If
Next i

If iCantidad > 2 Then
  strDetalle = strDetalle & " - Doc .: Filtrados"
ElseIf iCantidad = 0 Then
  strDetalle = strDetalle & " - Doc .: Todos"
Else
   strDetalle = strDetalle & " - Doc.: " & Mid(vCadena, 28, Len(vCadena))
End If

iCantidad = 0
  If vFiltro <> Empty And vFiltro <> " and " Then vFiltro = vFiltro & " and "
  vFiltro = vFiltro & "{vSIFDocumentos.Tipo_Documento} in [" & vCadena & "] "

vCadena = ""
For i = 1 To lswJuicios.ListItems.Count
  If lswJuicios.ListItems.Item(i).Checked Then
    If vCadena <> "" Then
        vCadena = vCadena & ","
    End If
    vCadena = vCadena & "'" & lswJuicios.ListItems.Item(i).Tag & "'"
    iCantidad = iCantidad + 1
  End If
Next i

If iCantidad > 2 Then
  strDetalle = strDetalle & " - Conceptos .: Filtrados"
ElseIf iCantidad = 0 Then
  strDetalle = strDetalle & " - Conceptos .: Todos"
Else
   strDetalle = strDetalle & " - Concepto.:" & Mid(vCadena, 28, Len(vCadena))
End If

If vFiltro <> Empty And vFiltro <> " and " Then vFiltro = vFiltro & " and "
 vFiltro = vFiltro & "{vSIFDocumentos.Cod_Concepto} in [" & vCadena & "] "

If chkFechas.Value = vbUnchecked Then
    Select Case cboProceso.Text
      Case "Registro"
    
        If vFiltro <> Empty And vFiltro <> " and " Then vFiltro = vFiltro & " and "
        vFiltro = vFiltro & "cdate({vSIFDocumentos.registro_fecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
        vFiltro = vFiltro & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    
        strDetalle = strDetalle & " - Fecha Registro.: desde " & Format(dtpInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(dtpCorte.Value, "dd/mm/yyyy")
    
      Case "Anulación"
    
        If vFiltro <> Empty And vFiltro <> " and " Then vFiltro = vFiltro & " and "
        vFiltro = vFiltro & "cdate({vSIFDocumentos.anulacion_fecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
        vFiltro = vFiltro & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    
        strDetalle = strDetalle & " - Fecha Anulación.: desde " & Format(dtpInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(dtpCorte.Value, "dd/mm/yyyy")
    
      Case "Traslado"
        If vFiltro <> Empty And vFiltro <> " and " Then vFiltro = vFiltro & " and "
        vFiltro = vFiltro & "cdate({vSIFDocumentos.traspaso_fecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
        vFiltro = vFiltro & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    
        strDetalle = strDetalle & " - Fecha Traslado.: desde " & Format(dtpInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(dtpCorte.Value, "dd/mm/yyyy")
      Case Else
        strDetalle = strDetalle & " - Todas las Fechas"
    End Select
End If

Select Case cboEstado.Text
  Case "Impreso"
     If vFiltro <> Empty And vFiltro <> " and" Then vFiltro = vFiltro & " and "
     vFiltro = vFiltro & "{vSIFDocumentos.estado}  in ['I','E'] "

  Case "Pendiente"
     If vFiltro <> Empty And vFiltro <> " and" Then vFiltro = vFiltro & " and "
     vFiltro = vFiltro & "{vSIFDocumentos.estado}  = 'P' "
  Case Else

End Select
strDetalle = strDetalle & " - Estado..:" & cboEstado.Text


fxReportesFiltros = vFiltro


Exit Function

vError:
  fxReportesFiltros = ""
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical




End Function




Private Sub txtAbogado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Consulta = "select cod_abogado,nombre from Cbr_Cj_Abogados"
    frmBusquedas.Show vbModal
    txtAbogado = gBusquedas.Resultado2
    txtAbogado.Tag = gBusquedas.Resultado
End If

End Sub



Private Sub txtBufete_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Consulta = "select cod_bufete,nombre from CBR_CJ_BUFETES"
    frmBusquedas.Show vbModal
    txtBufete = gBusquedas.Resultado2
    txtBufete.Tag = gBusquedas.Resultado
End If
End Sub


Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim frm As Form
Dim lOperacion As String

Call sbSIFForms("frmCO_CJ_Tramite")

vGrid.Row = Row

For Each frm In Forms
    If (UCase(frm.Name) = UCase("frmCO_CJ_Tramite")) Then
        vGrid.Col = 3
        lOperacion = vGrid.Text
        vGrid.Col = 2
        Call frm.sbConsultaExterna(vGrid.Text, lOperacion)
        Exit For
    End If
Next frm

End Sub

Private Function fxFiltro(vTipo As String) As String
Dim strSQL As String

If vTipo = "T" Then
    strSQL = "select '',T.COD_TRAMITE,R.ID_SOLICITUD,R.CODIGO" _
            & ",R.CEDULA,S.NOMBRE,dbo.fxCrdPlazoRestante(R.PLAZO,R.PRIDEDUC," & GLOBALES.glngFechaCR & " )as 'Plazo'" _
            & ",isnull(T.TOTAL_DEUDA, R.Saldo) as 'Monto' ,T.PROCESO_USUARIO,T.PROCESO_FECHA" _
            & ",R.ProcesoDesc,A.NOMBRE as 'Abogado'" _
            & ",J.NOMBRE as 'Juzgado',Tj.DESCRIPCION as 'Juicio'" _
            & " from vCbrCjOperaciones R left JOIN CBR_CJ_TRAMITE  T ON R.ID_SOLICITUD = T.ID_SOLICITUD" _
            & " left JOIN SOCIOS S ON R.CEDULA = S.CEDULA" _
            & " left join CBR_CJ_ABOGADOS A On T.COD_ABOGADO = A.COD_ABOGADO" _
            & " left join CBR_CJ_JUZGADOS J on T.COD_JUZGADO = J.COD_JUZGADO" _
            & " left join CBR_CJ_TIPOS_JUICIOS Tj on T.TIPO_JUICIO = Tj.TIPO_JUICIO"
Else

    strSQL = "select '','',R.ID_SOLICITUD,R.CODIGO" _
            & ",R.CEDULA,S.NOMBRE,dbo.fxCrdPlazoRestante(R.PLAZO,R.PRIDEDUC," & GLOBALES.glngFechaCR & " )as 'Plazo'" _
            & ",R.Saldo as 'Monto' ,'','',R.ProcesoDesc,'','',''" _
            & " from vCbrCjOperacionesPendientes R inner JOIN SOCIOS S ON R.CEDULA = S.CEDULA"
            
End If
        
fxFiltro = strSQL

End Function
