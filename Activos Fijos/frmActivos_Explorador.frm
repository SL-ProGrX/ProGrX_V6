VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmActivos_Explorador 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Explorar: Activos Fijos"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   20970
   Icon            =   "frmActivos_Explorador.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   20970
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3015
      Left            =   2880
      TabIndex        =   2
      Top             =   3640
      Width           =   3255
      _Version        =   1441793
      _ExtentX        =   5741
      _ExtentY        =   5318
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
      Sorted          =   -1  'True
      ShowBorder      =   0   'False
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   6960
      Top             =   720
   End
   Begin XtremeSuiteControls.GroupBox gbPeriodo 
      Height          =   615
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   16815
      _Version        =   1441793
      _ExtentX        =   29660
      _ExtentY        =   1085
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnPeriodo 
         Height          =   375
         Left            =   0
         TabIndex        =   29
         Top             =   240
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Periodo"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmActivos_Explorador.frx":000C
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   375
         Left            =   1200
         TabIndex        =   30
         Top             =   240
         Width           =   1320
         _Version        =   1441793
         _ExtentX        =   2328
         _ExtentY        =   661
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
      Begin XtremeSuiteControls.PushButton btnVisualizar 
         Height          =   375
         Index           =   0
         Left            =   5880
         TabIndex        =   32
         Top             =   240
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Listado"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Checked         =   -1  'True
         Picture         =   "frmActivos_Explorador.frx":0714
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnVisualizar 
         Height          =   375
         Index           =   1
         Left            =   7200
         TabIndex        =   33
         Top             =   240
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Actual"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmActivos_Explorador.frx":0E1C
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnVisualizar 
         Height          =   375
         Index           =   2
         Left            =   8640
         TabIndex        =   34
         Top             =   240
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cierre"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmActivos_Explorador.frx":143A
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAcciones 
         Height          =   375
         Index           =   0
         Left            =   11640
         TabIndex        =   44
         ToolTipText     =   "Refrescar"
         Top             =   240
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmActivos_Explorador.frx":1A56
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAcciones 
         Height          =   375
         Index           =   1
         Left            =   12120
         TabIndex        =   45
         ToolTipText     =   "Informes"
         Top             =   240
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmActivos_Explorador.frx":2156
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAcciones 
         Height          =   375
         Index           =   2
         Left            =   12600
         TabIndex        =   46
         Top             =   240
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cierre"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmActivos_Explorador.frx":285D
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAcciones 
         Height          =   375
         Index           =   3
         Left            =   13920
         TabIndex        =   47
         Top             =   240
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Asientos"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmActivos_Explorador.frx":3109
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnAcciones 
         Height          =   375
         Index           =   4
         Left            =   15120
         TabIndex        =   48
         Top             =   240
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Compras"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmActivos_Explorador.frx":3822
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.Label lblTitulo 
         Height          =   375
         Index           =   1
         Left            =   10680
         TabIndex        =   43
         Top             =   240
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Acciones"
         BackColor       =   16761024
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
      Begin XtremeSuiteControls.Label lblTitulo 
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   35
         Top             =   240
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Visualizar"
         BackColor       =   16761024
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
      Begin XtremeSuiteControls.Label lblPeriodo 
         Height          =   375
         Left            =   2640
         TabIndex        =   31
         Top             =   240
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Periodo"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   4
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox gbFiltros 
      Height          =   2415
      Left            =   2880
      TabIndex        =   3
      Top             =   1080
      Width           =   16815
      _Version        =   1441793
      _ExtentX        =   29660
      _ExtentY        =   4260
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkAdquisicion 
         Height          =   255
         Left            =   8640
         TabIndex        =   22
         Top             =   1800
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha de Adquisición"
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
      Begin XtremeSuiteControls.FlatEdit txtMarca 
         Height          =   330
         Left            =   6240
         TabIndex        =   21
         Top             =   2040
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
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
      Begin XtremeSuiteControls.FlatEdit txtSerie 
         Height          =   330
         Left            =   3960
         TabIndex        =   19
         Top             =   2040
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
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
      Begin XtremeSuiteControls.FlatEdit txtResponsable 
         Height          =   330
         Left            =   1680
         TabIndex        =   9
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   960
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
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   330
         Left            =   1680
         TabIndex        =   6
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
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   330
         Left            =   1680
         TabIndex        =   7
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
      Begin XtremeSuiteControls.ComboBox cboDepartamento 
         Height          =   330
         Left            =   8640
         TabIndex        =   13
         Top             =   600
         Width           =   4335
         _Version        =   1441793
         _ExtentX        =   7646
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
      Begin XtremeSuiteControls.ComboBox cboSeccion 
         Height          =   330
         Left            =   8640
         TabIndex        =   14
         Top             =   960
         Width           =   4335
         _Version        =   1441793
         _ExtentX        =   7646
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
      Begin XtremeSuiteControls.ComboBox cboLocaliza 
         Height          =   330
         Left            =   8640
         TabIndex        =   15
         Top             =   1320
         Width           =   4335
         _Version        =   1441793
         _ExtentX        =   7646
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
      Begin XtremeSuiteControls.FlatEdit txtModelo 
         Height          =   330
         Left            =   1680
         TabIndex        =   17
         Top             =   2040
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
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
      Begin XtremeSuiteControls.CheckBox chkInstalacion 
         Height          =   255
         Left            =   11520
         TabIndex        =   23
         Top             =   1800
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4683
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha Instalación"
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
      Begin XtremeSuiteControls.DateTimePicker dtpAdquisicion 
         Height          =   330
         Index           =   0
         Left            =   8640
         TabIndex        =   24
         Top             =   2040
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpAdquisicion 
         Height          =   330
         Index           =   1
         Left            =   9960
         TabIndex        =   25
         Top             =   2040
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInstalacion 
         Height          =   330
         Index           =   0
         Left            =   11520
         TabIndex        =   26
         Top             =   2040
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInstalacion 
         Height          =   330
         Index           =   1
         Left            =   12840
         TabIndex        =   27
         Top             =   2040
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   615
         Left            =   14400
         TabIndex        =   36
         Top             =   1800
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
         Picture         =   "frmActivos_Explorador.frx":3F29
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   615
         Left            =   15720
         TabIndex        =   37
         ToolTipText     =   "Exportar a Excel"
         Top             =   1800
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
         Picture         =   "frmActivos_Explorador.frx":4947
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBarX 
         Height          =   135
         Left            =   14400
         TabIndex        =   38
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   233
         _StockProps     =   93
         BackColor       =   -2147483633
         Scrolling       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   330
         Left            =   14400
         TabIndex        =   39
         Top             =   240
         Width           =   1935
         _Version        =   1441793
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
      Begin XtremeSuiteControls.FlatEdit txtLineas 
         Height          =   330
         Left            =   600
         TabIndex        =   41
         Top             =   2040
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
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
         Text            =   "10000"
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtProveedor 
         Height          =   330
         Left            =   1680
         TabIndex        =   51
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1320
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
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   330
         Left            =   8640
         TabIndex        =   53
         Top             =   240
         Width           =   4335
         _Version        =   1441793
         _ExtentX        =   7646
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
      Begin XtremeSuiteControls.CheckBox chkInfoResponsables 
         Height          =   495
         Left            =   13680
         TabIndex        =   55
         Top             =   1080
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4683
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Incluir Información Adicional (Responsable, Localiza, etc.)"
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
      Begin XtremeSuiteControls.FlatEdit txtPlacaI 
         Height          =   330
         Left            =   14400
         TabIndex        =   56
         Top             =   600
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
      Begin XtremeSuiteControls.FlatEdit txtPlacaC 
         Height          =   330
         Left            =   15360
         TabIndex        =   57
         Top             =   600
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
      Begin XtremeSuiteControls.ComboBox cboPlaca 
         Height          =   330
         Left            =   13200
         TabIndex        =   58
         Top             =   600
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   12
         Left            =   7080
         TabIndex        =   54
         Top             =   240
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo Activo"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   52
         Top             =   1320
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Proveedor"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   10
         Left            =   600
         TabIndex        =   42
         Top             =   1800
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Líneas"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   9
         Left            =   13200
         TabIndex        =   40
         Top             =   240
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Estado"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   8
         Left            =   6240
         TabIndex        =   20
         Top             =   1800
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Marca"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   7
         Left            =   3960
         TabIndex        =   18
         Top             =   1800
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Serie"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   6
         Left            =   1680
         TabIndex        =   16
         Top             =   1800
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
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
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   5
         Left            =   7080
         TabIndex        =   12
         Top             =   1320
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ubicacion"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   4
         Left            =   7080
         TabIndex        =   11
         Top             =   600
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Departamento"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   3
         Left            =   7080
         TabIndex        =   10
         Top             =   960
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Sección"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Responsable"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
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
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Descripción"
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
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   5280
      Left            =   6225
      ScaleHeight     =   5513.901
      ScaleMode       =   0  'User
      ScaleWidth      =   260
      TabIndex        =   0
      Top             =   1035
      Visible         =   0   'False
      Width           =   156
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   9000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":514C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":B9AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":12210
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":18A72
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":1F2D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":25B36
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":2C398
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":32BFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":3945C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":3FCBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":46520
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":4CD82
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgArbol 
      Left            =   9600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":535E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":53EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":54798
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":55072
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":5594C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":56226
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":56B00
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":573DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":57CB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":5858E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":58E68
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":59742
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Explorador.frx":59B94
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvActivos 
      Height          =   5190
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   9155
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImgArbol"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption lblTitle 
      Height          =   330
      Index           =   1
      Left            =   2760
      TabIndex        =   50
      Top             =   720
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4678
      _ExtentY        =   582
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   6
   End
   Begin XtremeShortcutBar.ShortcutCaption lblTitle 
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   49
      Top             =   720
      Width           =   2775
      _Version        =   1441793
      _ExtentX        =   4895
      _ExtentY        =   582
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
   End
   Begin VB.Image imgSplitter 
      Height          =   5235
      Left            =   2730
      MouseIcon       =   "frmActivos_Explorador.frx":59EAE
      MousePointer    =   99  'Custom
      Top             =   1050
      Width           =   150
   End
End
Attribute VB_Name = "frmActivos_Explorador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const sglSplitLimit = 500

Private mbMoving As Boolean
Private Nodo As MSComctlLib.Node

Private itmX As XtremeSuiteControls.ListViewItem
Private ItemSeleccionado As XtremeSuiteControls.ListViewItem

Private vSQL As String
Dim vPaso As Boolean




Public Sub btnAcciones_Click(Index As Integer)
Dim pOpcion As String

On Error GoTo vError
      
pOpcion = "Propiedades"
      
Select Case Index
    Case 0 'Refresh
        Call sbTreeLoad
        Exit Sub

    Case 1 'Reportes
        Call sbClassCall("Activos", 0, "frmActivos_Reportes")
    
    Case 2 'Cierre
        Call sbClassCall("Activos", 0, "frmActivos_CierrePeriodo")
    
    Case 3 'Traslado de Asientos
        Call sbClassCall("Activos", 0, "frmActivos_TrasladoAsientos")
    
    Case 4 'Informe de Compras
        Call sbClassCall("Activos", 0, "frmActivos_ComprasNR")
     
    Case 5
        pOpcion = "Nuevo"
    
    Case 6
        pOpcion = "Propiedades"
    
End Select
    
If Index = 5 Or Index = 6 Then
     Select Case tvActivos.SelectedItem.Tag
        Case "Activo", "Tipo", "Seccion", "Departamentos", "Empresa"
            Call sbActivosToolbar_Clic(pOpcion)

        Case "TipoActivo"
            Call sbTiposActivosToolbar_Clic(pOpcion)

        Case "Justificacion"
            Call sbJustificacionToolbar_Clic(pOpcion)

        Case "Modificacion"
            Call sbAdicionesRetirosToolbar_Clic(pOpcion)

        Case "Departamento"
            Call sbClassCall("Activos", 0, "frmActivos_Departamentos")
    End Select
End If
    
Exit Sub

vError:

End Sub

Private Sub btnBuscar_Click()

Call sbConsultaFiltrada

End Sub

Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lsw, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnPeriodo_Click()
   gActivos.Periodo = dtpFecha.Value
   frmActivos_Periodos.Show vbModal
   dtpFecha.Value = gActivos.Periodo
   Call dtpFecha_Change
End Sub

Private Sub btnVisualizar_Click(Index As Integer)
Dim i As Integer

For i = 0 To btnVisualizar.Count - 1
    btnVisualizar.Item(i).Checked = False
Next i

btnVisualizar.Item(Index).Checked = True



Call tvActivos_NodeClick(tvActivos.SelectedItem)


End Sub


Private Sub cboDepartamento_Click()
Dim strSQL As String

If vPaso Then Exit Sub

If cboDepartamento.ListCount = 0 Then
    cboSeccion.Clear
    Exit Sub
End If

strSQL = "select rtrim(cod_Seccion) as 'Idx', rtrim(descripcion) as 'ItmX' from Activos_Secciones" _
       & " Where cod_departamento = '" & cboDepartamento.ItemData(cboDepartamento.ListIndex) & "' order by cod_Seccion"
Call sbCbo_Llena_New(cboSeccion, strSQL, True, True)


End Sub

Private Sub chkAdquisicion_Click()
If chkAdquisicion.Value = xtpChecked Then
    dtpAdquisicion(0).Enabled = True
Else
    dtpAdquisicion(0).Enabled = False
End If

dtpAdquisicion(1).Enabled = dtpAdquisicion(0).Enabled
End Sub

Private Sub chkInstalacion_Click()
If chkInstalacion.Value = xtpChecked Then
    dtpInstalacion(0).Enabled = True
Else
    dtpInstalacion(0).Enabled = False
End If

dtpInstalacion(1).Enabled = dtpInstalacion(0).Enabled

End Sub

Private Sub dtpFecha_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

gActivos.Anio = Year(dtpFecha.Value)
gActivos.Mes = Month(dtpFecha.Value)
gActivos.Periodo = dtpFecha.Value

strSQL = "select * from Activos_periodos where anio = " & Year(dtpFecha.Value) _
       & " and Mes = " & Month(dtpFecha.Value)
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
 If rs!Estado = "P" Then
  lblPeriodo.Caption = "Período Pendiente"
  lblPeriodo.BackColor = RGB(214, 234, 248)
 
 
 Else
  lblPeriodo.Caption = "Período Cerrado"
  lblPeriodo.BackColor = RGB(227, 233, 243)
 
 End If
Else
  lblPeriodo.Caption = "Período Pendiente"
  lblPeriodo.BackColor = RGB(214, 234, 248)
End If

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub


Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
  
lblTitulo(0).BackColor = RGB(214, 234, 248)
lblTitulo(1).BackColor = RGB(214, 234, 248)
  
  
Me.BackColor = RGB(214, 234, 248)
  
cboPlaca.AddItem "Placa"
cboPlaca.AddItem "Alterna"
cboPlaca.Text = "Placa"
  
Call sbTreeLoad

strSQL = "select dbo.fxActivos_PeriodoActual() as 'Fecha'"
Call OpenRecordSet(rs, strSQL, 0)
    dtpFecha.Value = rs!fecha
rs.Close

Call dtpFecha_Change
    
Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub sbTreeLoad()
On Error GoTo vError

    Call sbTreeCreateRoot(Nodo, "Activos", "Empresa", 1)
    Call sbTreeShowNodes(4, Nodo)
    
    If tvActivos.Nodes.Count > 0 Then
        tvActivos.Nodes(1).Expanded = True
       ' tvActivos.Nodes(1).Selected = True
        Call tvActivos_NodeClick(tvActivos.Nodes(1))
    End If

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Despliega_SubNivel()
On Error GoTo adoErrores
   
   Dim NodoClic As MSComctlLib.Node
   
   Call tvActivos_Expand(tvActivos.Nodes(tvActivos.SelectedItem.Index))
   
   Set NodoClic = sbTreeNodoSearch(ItemSeleccionado.Key)
   
   If Not NodoClic Is Nothing Then
    tvActivos.Nodes(NodoClic.Index).Expanded = True
    tvActivos.Nodes(NodoClic.Index).Selected = True
    Call tvActivos_NodeClick(NodoClic)
   End If
   
Salir:
    Exit Sub
adoErrores:
   If Err.Number = 91 Then Resume Salir
   MsgBox fxSys_Error_Handler(Err.Description)
End Sub

Private Sub Form_Resize()
On Error Resume Next
  If Me.Width < 3000 Then Me.Width = 3000
  SizeControls imgSplitter.Left
  
  
  gbFiltros.Width = Me.Width - gbFiltros.Left
  gbPeriodo.Width = Me.Width
  
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  With imgSplitter
    picSplitter.Move .Left, .top, .Width - 20, .Height - 20
  End With
  picSplitter.Visible = True
  mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim sglPos As Single
  
  If mbMoving Then
    sglPos = x + imgSplitter.Left
    If sglPos < sglSplitLimit Then
      picSplitter.Left = sglSplitLimit
    ElseIf sglPos > Me.Width - sglSplitLimit Then
      picSplitter.Left = Me.Width - sglSplitLimit
    Else
      picSplitter.Left = sglPos
    End If
  End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  SizeControls picSplitter.Left
  picSplitter.Visible = False
  mbMoving = False
End Sub

Sub SizeControls(x As Single)
'  On Error Resume Next
  If x < 1500 Then x = 1500
  If x > (Me.Width - 1500) Then x = Me.Width - 1500
  
  tvActivos.Width = x
  imgSplitter.Left = x
  
  lsw.Left = x + 40
  gbFiltros.Left = lsw.Left
  
  lsw.Width = Me.Width - (tvActivos.Width + 140)

  
   tvActivos.Height = Me.ScaleHeight - (tvActivos.top) ' + 100
  
    lblTitle(0).Width = tvActivos.Width
    lblTitle(1).Left = lblTitle(0).Left + lblTitle(0).Width
    lblTitle(1).Width = Me.Width

  
  lsw.Height = Me.Height - (lsw.top + 450)
  imgSplitter.top = tvActivos.top
  imgSplitter.Height = tvActivos.Height
End Sub

Private Sub tvTreeView_DragDrop(Source As Control, x As Single, y As Single)
  If Source = imgSplitter Then
    SizeControls x
  End If
End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_DblClick()
    Call Despliega_SubNivel
End Sub



Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
    Set ItemSeleccionado = Item
End Sub

Private Sub lsw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim frmX As Form

On Error GoTo vError

    If Button = vbRightButton Then
        
        For Each frmX In Forms
           If Mid(frmX.Name, 1, 3) = "MDI" Then
                Exit For
           End If
        Next
        
       
        Call PopupMenu(frmX.mnuActivosExplorador, , x, y)
    End If
    
vError:

End Sub

Private Sub Tlb_Accion_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim frmX As Form

For Each frmX In Forms
   If Mid(frmX.Name, 1, 3) = "MDI" Then
        Exit For
   End If
Next

' Call PopupMenu(frmX.mnu_Explorador, , 560, 360)

End Sub


Public Sub PropiedadesTiposActivo()
Dim frmX As Form

On Error GoTo vError

     Call sbClassCall("Activos", 0, "frmActivos_TiposActivo")
'     Call sbFormActivo("frmActivos_TiposActivo", frmX)
'
'     Call frmX.sbconsultaExterna(lsw.SelectedItem.SubItems(1))
     
Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description)
End Sub

Public Sub PropiedadesAdicionRetiro()
Dim strSQL As String, rs As New ADODB.Recordset
Dim frmX As Form

On Error GoTo vError

If lsw.ListItems.Count = 0 Then Exit Sub

strSQL = "select Tipo from Activos_retiro_Adicion where num_placa = '" & lsw.SelectedItem.SubItems(1) _
       & "' and ID_ADDRET = " & lsw.SelectedItem.Tag
Call OpenRecordSet(rs, strSQL, 0)
If rs!Tipo = "V" Then
    
     Call sbClassCall("Activos", 0, "frmActivos_Revaluaciones")
'     Call sbFormActivo("frmActivos_Revaluaciones", frmX)
'     frmX.txtCodigo = lsw.SelectedItem.SubItems(1)
'     frmX.lblID.Tag = lsw.SelectedItem.Tag
'     frmX.Show
'     frmX.sbArbolShow
Else
     Call sbClassCall("Activos", 0, "frmActivos_AdicionRetiro")

'     Call sbFormActivo("frmActivos_AdicionRetiro", frmX)
'     frmX.txtCodigo = lsw.SelectedItem.SubItems(1)
'     frmX.lblID.Tag = lsw.SelectedItem.Tag
'     frmX.Show
'     frmX.sbArbolShow
End If
rs.Close

Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Public Sub PropiedadesJustificacion()
Dim frmX As Form

On Error GoTo vError

     Call sbClassCall("Activos", 0, "frmActivos_Justificaciones")
'     Call sbFormActivo("frmActivos_Justificaciones", frmX)
'
'     Call frmX.sbconsultaExterna(lsw.SelectedItem.SubItems(1))
'

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbInicializa()
Dim strSQL As String


dtpAdquisicion(0).Value = fxFechaServidor
dtpAdquisicion(1).Value = dtpAdquisicion(0).Value

dtpInstalacion(0).Value = dtpAdquisicion(0).Value
dtpInstalacion(1).Value = dtpAdquisicion(0).Value

Call chkAdquisicion_Click
Call chkInstalacion_Click

cboEstado.Clear
cboEstado.AddItem "TODOS"
cboEstado.AddItem "Activos"
cboEstado.AddItem "Depreciados"
cboEstado.AddItem "Retirados"
cboEstado.Text = "TODOS"


 
 vPaso = True
  strSQL = "select rtrim(tipo_activo) as 'Idx', rtrim(descripcion) as 'ItmX'" _
       & " from Activos_tipo_activo order by tipo_activo"
  Call sbCbo_Llena_New(cboTipo, strSQL, True, True)
 
  strSQL = "select rtrim(COD_LOCALIZA) as 'Idx', rtrim(descripcion) as 'ItmX'" _
       & " from ACTIVOS_LOCALIZACIONES Where Activa = 1 order by descripcion"
  Call sbCbo_Llena_New(cboLocaliza, strSQL, True, True)
 
  strSQL = "select rtrim(cod_departamento) as 'IdX' , rtrim(descripcion) as 'ItmX' from Activos_departamentos order by cod_departamento"
  Call sbCbo_Llena_New(cboDepartamento, strSQL, True, True)

 vPaso = False
 

Call cboDepartamento_Click


End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub

Private Sub tvActivos_DragDrop(Source As Control, x As Single, y As Single)
If Source = imgSplitter Then
    SizeControls x
  End If
End Sub

Private Sub tvActivos_Expand(ByVal Node As MSComctlLib.Node)

On Error GoTo adoError

If Node.Children = 0 Then Exit Sub

If Left(Node.Child.Key, 3) = "0x0" Then
    tvActivos.Nodes.Remove Node.Child.Index
    
    Select Case Node.Tag
        Case "Departamento"
            Call sbListaDepartamentos(8, Node)
        
        Case "Departamentos"
            
            Call sbListaSecciones(Node)
            
        Case "TipoActivo"
            Call sbTiposActivosArbol(Node)
        
        Case "Asiento"
           ' Call sbListaAsientos(DatePart("yyyy", dtpFecha), DatePart("m", dtpFecha), GlobalTipoAsiento, Node)
            Call sbListaAsientos(DatePart("yyyy", dtpFecha), DatePart("m", dtpFecha), "AS", Node)
            
        Case "Justificacion"
            Call sbJustificacionArbol(Node)
    End Select

End If
   
Salir:
   Exit Sub
adoError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, "Contabilidad"
 
End Sub

Public Sub tvActivos_NodeClick(ByVal Node As MSComctlLib.Node)
Dim vTipo As String, vLlave01 As String, vLlave02 As String, vLlave03 As String
Dim i As Byte

On Error GoTo vError

vReportes.Tipo = Node.Tag
vReportes.Llave1 = ""
vReportes.Llave2 = ""
 
If Not Node.Parent Is Nothing Then
lblTitle(0).Caption = UCase(Node.Parent & "")
lblTitle(1).Caption = UCase(Node.Text & "")
End If

'Saca Tipo de Listado
For i = 0 To btnVisualizar.Count - 1
  If btnVisualizar.Item(i).Checked Then
    Select Case i
      Case 0 'Listado
        vTipo = "L"
      Case 1 'Actual
        vTipo = "A"
      Case 2 'Preliminares / Cierres
        vTipo = "H"
    End Select
  End If
Next i

Select Case Node.Tag
    Case "Empresa"
    '    Call sbListaDepartamentos
    
    Case "Departamento"
        Call sbListaDepartamentosTodos
        
    Case "Departamentos"
'        vReportes.Llave1 = DeCodificaPrimaryKey(Nodo.Key, 5, "(id)")
'        Call Mantenimientos.ListaSecciones(Node)
        vLlave02 = DeCodificaPrimaryKey(Node.Key, 5, "(id)")
        Call sbListadosActivos(vTipo, dtpFecha, vLlave02, "", "")
    
    Case "Seccion"
         vLlave01 = DeCodificaPrimaryKey(Node.Key, 5, "(SC)")
         vLlave02 = DeCodificaPrimaryKey(Node.Key, wPosIni, "(id)")
         Call sbListadosActivos(vTipo, dtpFecha, vLlave01, vLlave02, "")
    
    Case "TipoActivo"
        Call sbTiposActivosLista
    
    Case "Activo"
         Call sbListadosActivos(vTipo, dtpFecha, "", "", "")
    
    Case "Asiento"
        Call sbActivosListaAsientos(Month(dtpFecha.Value), Year(dtpFecha.Value))
        
    
    
    Case "Tipo"
        vLlave03 = DeCodificaPrimaryKey(tvActivos.SelectedItem.Key, 5, "(id)")
        Call sbListadosActivos(vTipo, dtpFecha, "", "", vLlave03)
        
    Case "Num_Asiento"
        Call sbListaAsientosDetalle(Node.Key, DatePart("m", dtpFecha), DatePart("yyyy", dtpFecha))
        
    Case "Justificacion"
        Call sbJustificacionLista
    
    Case "Modificacion"
        Call sbAdicionesRetirosLista
        
End Select


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbConsultaFiltrada()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, vKey As String


Dim vTipo As String, vFecha As Date, vDepartamento As String, vSeccion As String, vTipoActivo As String, vLocaliza As String


Dim curVL As Currency, curVH As Currency, curVR As Currency
Dim curDepAc As Currency, curDepMes As Currency, curDepAnt As Currency


On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ColumnHeaders.Clear

curVL = 0
curVH = 0
curVR = 0
curDepAc = 0
curDepMes = 0
curDepAnt = 0

Select Case True
    Case btnVisualizar.Item(0).Checked
        vTipo = "L"
    Case btnVisualizar.Item(1).Checked
        vTipo = "A"
    Case btnVisualizar.Item(2).Checked
        vTipo = "C"
End Select
'Revisa Inyeccion de Codigo

txtNombre.Text = fxSysCleanTxtInject(txtNombre.Text)
txtDescripcion.Text = fxSysCleanTxtInject(txtDescripcion.Text)
txtModelo.Text = fxSysCleanTxtInject(txtModelo.Text)
txtMarca.Text = fxSysCleanTxtInject(txtMarca.Text)
txtSerie.Text = fxSysCleanTxtInject(txtSerie.Text)
txtLineas.Text = fxSysCleanTxtInject(txtLineas.Text)
txtPlacaI.Text = fxSysCleanTxtInject(txtPlacaI.Text)
txtPlacaC.Text = fxSysCleanTxtInject(txtPlacaC.Text)

vTipoActivo = ""
vDepartamento = ""
vSeccion = ""
vLocaliza = ""


If cboTipo.Text <> "TODOS" Then
   vTipoActivo = cboTipo.ItemData(cboTipo.ListIndex)
End If

If cboLocaliza.Text <> "TODOS" Then
   vLocaliza = cboLocaliza.ItemData(cboLocaliza.ListIndex)
End If

If cboDepartamento.Text <> "TODOS" Then
   vDepartamento = cboDepartamento.ItemData(cboDepartamento.ListIndex)
End If

If cboSeccion.Text <> "TODOS" Then
   vSeccion = cboSeccion.ItemData(cboSeccion.ListIndex)
End If

vFecha = dtpFecha.Value


Select Case UCase(vTipo)
  Case "L"
    lsw.ListItems.Clear
    lsw.ColumnHeaders.Clear
    lsw.ColumnHeaders.Add , , "Placa", 1400
    lsw.ColumnHeaders.Add , , "Id Alterna", 1400, vbCenter
    lsw.ColumnHeaders.Add , , "Nombre", 4000
    lsw.ColumnHeaders.Add , , "Fecha Adq.", 1400
    lsw.ColumnHeaders.Add , , "Fecha Inst.", 1400
    lsw.ColumnHeaders.Add , , "Tipo", 2400
    lsw.ColumnHeaders.Add , , "Vida Util", 1400, 1
    lsw.ColumnHeaders.Add , , "Valor historico", 1400, 1
    lsw.ColumnHeaders.Add , , "Valor Rescate", 1400, 1
    lsw.ColumnHeaders.Add , , "Estado", 1400, vbCenter
    
If chkInfoResponsables.Value = xtpChecked Then
    lsw.ColumnHeaders.Add , , "Id Responsable", 1400, vbCenter
    lsw.ColumnHeaders.Add , , "Responsable", 2500
    lsw.ColumnHeaders.Add , , "Departamento", 2400
    lsw.ColumnHeaders.Add , , "Sección", 2400
    lsw.ColumnHeaders.Add , , "Localización", 2400
    lsw.ColumnHeaders.Add , , "Proveedor", 2400

    lsw.ColumnHeaders.Add , , "Modelo", 2400, vbCenter
    lsw.ColumnHeaders.Add , , "Marca", 2400, vbCenter
    lsw.ColumnHeaders.Add , , "No. Serie", 2400, vbCenter
    lsw.ColumnHeaders.Add , , "Otras Señas", 3000

End If

    
    strSQL = "select * " _
           & " from vActivos_General" _
           & " where Nombre like '%" & txtNombre.Text & "%'"
           
    If chkAdquisicion.Value = xtpChecked Then
        strSQL = strSQL & " and Fecha_Adquisicion between '" & Format(dtpAdquisicion(0).Value, "yyyy-mm-dd") _
               & " 00:00:00' and '" & Format(dtpAdquisicion(1).Value, "yyyy-mm-dd") & " 23:59:59'"
    End If

    If chkInstalacion.Value = xtpChecked Then
        strSQL = strSQL & " and Fecha_Instalacion between '" & Format(dtpInstalacion(0).Value, "yyyy-mm-dd") _
               & " 00:00:00' and '" & Format(dtpInstalacion(1).Value, "yyyy-mm-dd") & " 23:59:59'"
    End If

    Select Case Mid(cboEstado.Text, 1, 1)
        Case "A"
            strSQL = strSQL & " and Estado = 'A'"
        Case "D"
            strSQL = strSQL & " and Estado = 'A' and isnull(Valor_Libros_Periodo,0) = 0"
        Case "R"
            strSQL = strSQL & " and Estado = 'R'"
    End Select

    If Len(txtDescripcion.Text) > 0 Then strSQL = strSQL & " and Descripcion like '%" & txtDescripcion.Text & "%'"
    If Len(txtModelo.Text) > 0 Then strSQL = strSQL & " and Modelo like '%" & txtModelo.Text & "%'"
    If Len(txtMarca.Text) > 0 Then strSQL = strSQL & " and Marca like '%" & txtMarca.Text & "%'"
    If Len(txtSerie.Text) > 0 Then strSQL = strSQL & " and Num_Serie like '%" & txtSerie.Text & "%'"
    
    
    If vTipoActivo <> "" Then strSQL = strSQL & " and Tipo_Activo = '" & vTipoActivo & "'"
    If vDepartamento <> "" Then strSQL = strSQL & " and cod_Departamento = '" & vDepartamento & "'"
    If vSeccion <> "" Then strSQL = strSQL & " and cod_Seccion = '" & vSeccion & "'"
    If vLocaliza <> "" Then strSQL = strSQL & " and cod_Localiza = '" & vLocaliza & "'"
    
    If Len(txtProveedor.Text) > 0 Then strSQL = strSQL & " and cod_Proveedor = '" & txtProveedor.Text & "'"
    If Len(txtResponsable.Text) > 0 Then strSQL = strSQL & " and Identificacion = '" & txtResponsable.Text & "'"
    
    If cboPlaca.Text = "Placa" Then
        If Len(txtPlacaI.Text) > 0 Then strSQL = strSQL & " and Num_Placa between '" & txtPlacaI.Text & "' and '" & txtPlacaC.Text & "'"
        strSQL = strSQL & " order by num_placa"
    Else
        If Len(txtPlacaI.Text) > 0 Then strSQL = strSQL & " and Placa_alterna between '" & txtPlacaI.Text & "' and '" & txtPlacaC.Text & "'"
        strSQL = strSQL & " order by Placa_Alterna, Num_Placa"
    End If
    
    
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
     vKey = "(AF)" & rs!num_placa & "(id)"
     Set itmX = lsw.ListItems.Add(, vKey, rs!num_placa)
         itmX.SubItems(1) = rs!Placa_Alterna & ""
         itmX.SubItems(2) = rs!Nombre
         itmX.SubItems(3) = Format(rs!fecha_adquisicion, "yyyy-mm-dd")
         itmX.SubItems(4) = Format(rs!fecha_instalacion, "yyyy-mm-dd")
         itmX.SubItems(5) = rs!Tipo_Activo_Desc
         itmX.SubItems(6) = rs!vida_util & IIf((rs!VIDA_UTIL_EN = "A"), " Año(s)", " Mes(es)")
         itmX.SubItems(7) = Format(rs!Valor_Historico, "Standard")
         itmX.SubItems(8) = Format(rs!Valor_Desecho, "Standard")
         itmX.SubItems(9) = rs!Estado_Desc
         
         
        If chkInfoResponsables.Value = xtpChecked Then
            itmX.SubItems(10) = rs!Identificacion
            itmX.SubItems(11) = rs!Responsable
            itmX.SubItems(12) = rs!departamento
            itmX.SubItems(13) = rs!seccion
            itmX.SubItems(14) = rs!Localizacion
            itmX.SubItems(15) = rs!Proveedor
        
            itmX.SubItems(16) = rs!Modelo & ""
            itmX.SubItems(17) = rs!Marca & ""
            itmX.SubItems(18) = rs!Num_Serie & ""
            itmX.SubItems(19) = rs!Otras_Senas
        End If
         
         curVH = curVH + rs!Valor_Historico
         curVR = curVR + rs!Valor_Desecho
         
     rs.MoveNext
    Loop
    rs.Close
     Set itmX = lsw.ListItems.Add(, "")
         itmX.SubItems(7) = "____________________"
         itmX.SubItems(8) = "____________________"
    
     Set itmX = lsw.ListItems.Add(, "")
         itmX.SubItems(7) = Format(curVH, "Standard")
         itmX.SubItems(8) = Format(curVR, "Standard")
          
       
       
  Case Else
    lsw.ListItems.Clear
    lsw.ColumnHeaders.Clear
    lsw.ColumnHeaders.Add , , "Placa", 1400
    lsw.ColumnHeaders.Add , , "Id Alterna", 1400, vbCenter
    lsw.ColumnHeaders.Add , , "Nombre", 4000
    lsw.ColumnHeaders.Add , , "Fecha Adq.", 1400
    lsw.ColumnHeaders.Add , , "Tipo", 2400
    lsw.ColumnHeaders.Add , , "Vida Util", 1400, 1
    lsw.ColumnHeaders.Add , , "Valor Historico", 1400, 1
    lsw.ColumnHeaders.Add , , "Dep. Ac. Mes Ant", 1400, 1
    lsw.ColumnHeaders.Add , , "Depreciacion Mes", 1400, 1
    lsw.ColumnHeaders.Add , , "Depreciación Ac", 1400, 1
    lsw.ColumnHeaders.Add , , "Valor Libros", 1400, 1
    lsw.ColumnHeaders.Add , , "Corte", 1400, vbCenter
    
    If chkInfoResponsables.Value = xtpChecked Then
        lsw.ColumnHeaders.Add , , "Id Responsable", 1400, vbCenter
        lsw.ColumnHeaders.Add , , "Responsable", 2500
        lsw.ColumnHeaders.Add , , "Departamento", 2400
        lsw.ColumnHeaders.Add , , "Sección", 2400
        lsw.ColumnHeaders.Add , , "Localización", 2400
        lsw.ColumnHeaders.Add , , "Proveedor", 2400
        lsw.ColumnHeaders.Add , , "Modelo", 2400, vbCenter
        lsw.ColumnHeaders.Add , , "Marca", 2400, vbCenter
        lsw.ColumnHeaders.Add , , "No. Serie", 2400, vbCenter
        lsw.ColumnHeaders.Add , , "Otras Señas", 3000
    End If
    
    If vTipo = "A" Then
            strSQL = "select A.Num_placa, A.Placa_Alterna, A.Nombre, A.Fecha_adquisicion, A.Vida_Util, A.Vida_Util_En, A.TipoActivo, A.DEPRECIACION_PERIODO" _
                   & ", sum(A.Valor_historico) as 'VALOR_HISTORICO',sum(A.depreciacion_acum) as 'DEPRECIACION_AC'" _
                   & ", sum(A.Depreciacion_mes) as 'DEPRECIACION_MES'" _
                   & ", sum(A.VALOR_LIBROS) as 'VALOR_LIBROS'" _
                   & ", A.Identificacion, A.Responsable, A.Departamento, A.Seccion, A.Localizacion, A.Proveedor" _
                   & ", A.Modelo, A.Marca, A.Num_Serie, A.Otras_Senas" _
                   & " from vActivos_depreciacion_actual A " _
                   & " where A.Estado <> 'R'"
            
            If chkAdquisicion.Value = xtpChecked Then
                strSQL = strSQL & " and A.Fecha_Adquisicion between '" & Format(dtpAdquisicion(0).Value, "yyyy-mm-dd") _
                       & " 00:00:00' and '" & Format(dtpAdquisicion(1).Value, "yyyy-mm-dd") & " 23:59:59'"
            End If
        
            If chkInstalacion.Value = xtpChecked Then
                strSQL = strSQL & " and A.Fecha_Instalacion between '" & Format(dtpInstalacion(0).Value, "yyyy-mm-dd") _
                       & " 00:00:00' and '" & Format(dtpInstalacion(1).Value, "yyyy-mm-dd") & " 23:59:59'"
            End If
            
            
            If Len(txtNombre.Text) > 0 Then strSQL = strSQL & " and P.Nombre like '%" & txtNombre.Text & "%'"
            If Len(txtDescripcion.Text) > 0 Then strSQL = strSQL & " and A.Descripcion like '%" & txtDescripcion.Text & "%'"
            If Len(txtModelo.Text) > 0 Then strSQL = strSQL & " and A.Modelo like '%" & txtModelo.Text & "%'"
            If Len(txtMarca.Text) > 0 Then strSQL = strSQL & " and A.Marca like '%" & txtMarca.Text & "%'"
            If Len(txtSerie.Text) > 0 Then strSQL = strSQL & " and A.Num_Serie like '%" & txtSerie.Text & "%'"
            
            If vTipoActivo <> "" Then strSQL = strSQL & " and A.tipo_activo = '" & vTipoActivo & "'"
            If vDepartamento <> "" Then strSQL = strSQL & " and A.cod_departamento = '" & vDepartamento & "'"
            If vSeccion <> "" Then strSQL = strSQL & " and A.cod_seccion = '" & vSeccion & "'"
            If vLocaliza <> "" Then strSQL = strSQL & " and A.cod_Localiza = '" & vSeccion & "'"
                   
                    
            If Len(txtProveedor.Text) > 0 Then strSQL = strSQL & " and A.cod_Proveedor = '" & txtProveedor.Text & "'"
            If Len(txtResponsable.Text) > 0 Then strSQL = strSQL & " and A.Identificacion = '" & txtResponsable.Text & "'"
            
            If cboPlaca.Text = "Placa" Then
                If Len(txtPlacaI.Text) > 0 Then strSQL = strSQL & " and A.Num_Placa between '" & txtPlacaI.Text & "' and '" & txtPlacaC.Text & "'"
            Else
                If Len(txtPlacaI.Text) > 0 Then strSQL = strSQL & " and A.Placa_alterna between '" & txtPlacaI.Text & "' and '" & txtPlacaC.Text & "'"
            End If
            
            
            strSQL = strSQL & " group by A.Num_placa, A.Placa_Alterna, A.Nombre, A.Fecha_adquisicion, A.Vida_Util, A.Vida_Util_En, A.TipoActivo, A.DEPRECIACION_PERIODO" _
                   & " , A.COD_LOCALIZA, A.COD_PROVEEDOR, A.IDENTIFICACION" _
                   & " , A.MODELO, A.MARCA, A.NUM_SERIE, A.FECHA_INSTALACION, A.DESCRIPCION" _
                   & " , A.Modelo, A.Marca, A.Num_Serie, A.Otras_Senas" _
                   & " , A.Responsable, A.Departamento, A.Seccion, A.Localizacion, A.Proveedor"

    
            Call OpenRecordSet(rs, strSQL, 0)
            Do While Not rs.EOF
             vKey = "(AF)" & rs!num_placa & "(id)"
             Set itmX = lsw.ListItems.Add(, vKey, rs!num_placa)
                 itmX.SubItems(1) = rs!Placa_Alterna & ""
                 itmX.SubItems(2) = rs!Nombre
                 itmX.SubItems(3) = Format(rs!fecha_adquisicion, "yyyy-mm-dd")
                 itmX.SubItems(4) = rs!TipoActivo
                 itmX.SubItems(5) = rs!vida_util & IIf((rs!VIDA_UTIL_EN = "A"), " Año(s)", " Mes(es)")
                 itmX.SubItems(6) = Format(rs!Valor_Historico, "Standard")
                 itmX.SubItems(7) = Format(rs!DEPRECIACION_AC - rs!DEPRECIACION_MES, "Standard")
                 itmX.SubItems(8) = Format(rs!DEPRECIACION_MES, "Standard")
                 itmX.SubItems(9) = Format(rs!DEPRECIACION_AC, "Standard")
                 itmX.SubItems(10) = Format(rs!Valor_Libros, "Standard")
                 itmX.SubItems(11) = Format(rs!depreciacion_periodo, "yyyy-mm-dd")
                 
                If chkInfoResponsables.Value = xtpChecked Then
                    itmX.SubItems(12) = rs!Identificacion
                    itmX.SubItems(13) = rs!Responsable
                    itmX.SubItems(14) = rs!departamento
                    itmX.SubItems(15) = rs!seccion
                    itmX.SubItems(16) = rs!Localizacion
                    itmX.SubItems(17) = rs!Proveedor
                
                    itmX.SubItems(18) = rs!Modelo & ""
                    itmX.SubItems(19) = rs!Marca & ""
                    itmX.SubItems(20) = rs!Num_Serie & ""
                    itmX.SubItems(21) = rs!Otras_Senas
                
                End If
                 
                 curVH = curVH + rs!Valor_Historico
                 curDepAnt = curDepAnt + rs!DEPRECIACION_AC - rs!DEPRECIACION_MES
                 curDepMes = curDepMes + rs!DEPRECIACION_MES
                 curDepAc = curDepAc + rs!DEPRECIACION_AC
                 curVL = curVL + rs!Valor_Libros
                 
             rs.MoveNext
            Loop
            rs.Close
            
             Set itmX = lsw.ListItems.Add(, "")
                 itmX.SubItems(6) = "____________________"
                 itmX.SubItems(7) = "____________________"
                 itmX.SubItems(8) = "____________________"
                 itmX.SubItems(9) = "____________________"
                 itmX.SubItems(10) = "____________________"
            
             Set itmX = lsw.ListItems.Add(, "")
                 itmX.SubItems(6) = Format(curVH, "Standard")
                 itmX.SubItems(7) = Format(curDepAnt, "Standard")
                 itmX.SubItems(8) = Format(curDepMes, "Standard")
                 itmX.SubItems(9) = Format(curDepAc, "Standard")
                 itmX.SubItems(10) = Format(curVL, "Standard")
    
    Else
      
            strSQL = "select A.* " _
                    & " from vActivos_AuxiliarConsolidado A inner join Activos_Principal P on A.num_Placa = P.num_Placa" _
                    & " where A.Anio = " & Year(vFecha) & " and A.Mes = " & Month(vFecha)
            
            
            If chkAdquisicion.Value = xtpChecked Then
                strSQL = strSQL & " and P.Fecha_Adquisicion between '" & Format(dtpAdquisicion(0).Value, "yyyy-mm-dd") _
                       & " 00:00:00' and '" & Format(dtpAdquisicion(1).Value, "yyyy-mm-dd") & " 23:59:59'"
            End If
        
            If chkInstalacion.Value = xtpChecked Then
                strSQL = strSQL & " and P.Fecha_Instalacion between '" & Format(dtpInstalacion(0).Value, "yyyy-mm-dd") _
                       & " 00:00:00' and '" & Format(dtpInstalacion(1).Value, "yyyy-mm-dd") & " 23:59:59'"
            End If
            
            
            If Len(txtNombre.Text) > 0 Then strSQL = strSQL & " and P.Nombre like '%" & txtNombre.Text & "%'"
            If Len(txtDescripcion.Text) > 0 Then strSQL = strSQL & " and P.Descripcion like '%" & txtDescripcion.Text & "%'"
            If Len(txtModelo.Text) > 0 Then strSQL = strSQL & " and P.Modelo like '%" & txtModelo.Text & "%'"
            If Len(txtMarca.Text) > 0 Then strSQL = strSQL & " and P.Marca like '%" & txtMarca.Text & "%'"
            If Len(txtSerie.Text) > 0 Then strSQL = strSQL & " and P.Num_Serie like '%" & txtSerie.Text & "%'"
            
            If vTipoActivo <> "" Then strSQL = strSQL & " and A.tipo_activo = '" & vTipoActivo & "'"
            If vDepartamento <> "" Then strSQL = strSQL & " and A.cod_departamento = '" & vDepartamento & "'"
            If vSeccion <> "" Then strSQL = strSQL & " and A.cod_seccion = '" & vSeccion & "'"
            If vLocaliza <> "" Then strSQL = strSQL & " and P.cod_Localiza = '" & vSeccion & "'"
                   
                    
            If Len(txtProveedor.Text) > 0 Then strSQL = strSQL & " and P.cod_Proveedor = '" & txtProveedor.Text & "'"
            If Len(txtResponsable.Text) > 0 Then strSQL = strSQL & " and A.Identificacion = '" & txtResponsable.Text & "'"
            
            If cboPlaca.Text = "Placa" Then
                If Len(txtPlacaI.Text) > 0 Then strSQL = strSQL & " and A.Num_Placa between '" & txtPlacaI.Text & "' and '" & txtPlacaC.Text & "'"
            Else
                If Len(txtPlacaI.Text) > 0 Then strSQL = strSQL & " and A.Placa_alterna between '" & txtPlacaI.Text & "' and '" & txtPlacaC.Text & "'"
            End If
            
            Call OpenRecordSet(rs, strSQL, 0)
            Do While Not rs.EOF
             vKey = "(AF)" & rs!num_placa & "(id)"
             Set itmX = lsw.ListItems.Add(, vKey, rs!num_placa)
                 itmX.SubItems(1) = rs!Placa_Alterna & ""
                 itmX.SubItems(2) = rs!Nombre
                 itmX.SubItems(3) = Format(rs!fecha_adquisicion, "yyyy-mm-dd")
                 itmX.SubItems(4) = rs!TipoActivo
                 itmX.SubItems(5) = rs!vida_util & IIf((rs!VIDA_UTIL_EN = "A"), " Año(s)", " Mes(es)")
                 itmX.SubItems(6) = Format(rs!Valor_Historico, "Standard")
                 itmX.SubItems(7) = Format(rs!DEPRECIACION_AC_CONSOLIDADO - rs!DEPRECIACION_MES_CONSOLIDADO, "Standard")
                 itmX.SubItems(8) = Format(rs!DEPRECIACION_MES_CONSOLIDADO, "Standard")
                 itmX.SubItems(9) = Format(rs!DEPRECIACION_AC_CONSOLIDADO, "Standard")
                 itmX.SubItems(10) = Format(rs!VALOR_LIBROS_CONSOLIDADO, "Standard")
                 
                 
                If chkInfoResponsables.Value = xtpChecked Then
                    itmX.SubItems(11) = rs!Identificacion
                    itmX.SubItems(12) = rs!RESPONSABLE_NOMBRE
                    itmX.SubItems(13) = rs!RESPONSABLE_DEPARTAMENTO
                    itmX.SubItems(14) = rs!RESPONSABLE_SECCION
                    itmX.SubItems(15) = rs!Localizacion
                    itmX.SubItems(16) = rs!Proveedor
                
                    itmX.SubItems(17) = rs!Modelo & ""
                    itmX.SubItems(18) = rs!Marca & ""
                    itmX.SubItems(19) = rs!Num_Serie & ""
                    itmX.SubItems(20) = rs!Otras_Senas
                
                End If
                 
                 curVH = curVH + rs!Valor_Historico
                 curDepAnt = curDepAnt + rs!DEPRECIACION_AC_CONSOLIDADO - rs!DEPRECIACION_MES_CONSOLIDADO
                 curDepMes = curDepMes + rs!DEPRECIACION_MES_CONSOLIDADO
                 curDepAc = curDepAc + rs!DEPRECIACION_AC_CONSOLIDADO
                 curVL = curVL + rs!VALOR_LIBROS_CONSOLIDADO
                 
             rs.MoveNext
            Loop
            rs.Close
    
             Set itmX = lsw.ListItems.Add(, "")
                 itmX.SubItems(6) = "____________________"
                 itmX.SubItems(7) = "____________________"
                 itmX.SubItems(8) = "____________________"
                 itmX.SubItems(9) = "____________________"
                 itmX.SubItems(10) = "____________________"
            
             Set itmX = lsw.ListItems.Add(, "")
                 itmX.SubItems(6) = Format(curVH, "Standard")
                 itmX.SubItems(7) = Format(curDepAnt, "Standard")
                 itmX.SubItems(8) = Format(curDepMes, "Standard")
                 itmX.SubItems(9) = Format(curDepAc, "Standard")
                 itmX.SubItems(10) = Format(curVL, "Standard")
    
    
    End If ' TIPO A

End Select

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub



'----------------------------------------------------------------------------------------------------------------------------
'                   Consultas dentro del Explorador
'----------------------------------------------------------------------------------------------------------------------------




Private Sub sbListadosActivos(vTipo As String, vFecha As Date _
                   , vDepartamento As String, vSeccion As String, vTipoActivo As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, vKey As String

Dim curVL As Currency, curVH As Currency, curVR As Currency
Dim curDepAc As Currency, curDepMes As Currency, curDepAnt As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ColumnHeaders.Clear

curVL = 0
curVH = 0
curVR = 0
curDepAc = 0
curDepMes = 0
curDepAnt = 0

Select Case UCase(vTipo)
  Case "L"
    lsw.ListItems.Clear
    lsw.ColumnHeaders.Clear
    lsw.ColumnHeaders.Add , , "Placa", 1400
    lsw.ColumnHeaders.Add , , "Id Alterna", 1400
    lsw.ColumnHeaders.Add , , "Nombre", 4000
    lsw.ColumnHeaders.Add , , "Fecha Adq.", 1400
    lsw.ColumnHeaders.Add , , "Fecha Inst.", 1400
    lsw.ColumnHeaders.Add , , "Tipo", 2400
    lsw.ColumnHeaders.Add , , "Vida Util", 1400, 1
    lsw.ColumnHeaders.Add , , "Valor historico", 1400, 1
    lsw.ColumnHeaders.Add , , "Valor Rescate", 1400, 1
    strSQL = "select A.*,T.descripcion as TipoActivo" _
           & " from Activos_tipo_Activo T inner join Activos_Principal A on T.tipo_Activo = A.tipo_activo" _
           & " where A.estado <> 'R' and A.fecha_adquisicion <= '" & Format(vFecha, "yyyy/mm/dd") & "'"
    
    If vTipoActivo <> "" Then strSQL = strSQL & " and A.tipo_activo = '" & vTipoActivo & "'"
    If vDepartamento <> "" Then strSQL = strSQL & " and A.cod_departamento = '" & vDepartamento & "'"
    If vSeccion <> "" Then strSQL = strSQL & " and A.cod_seccion = '" & vSeccion & "'"
    
    strSQL = strSQL & " order by A.num_placa"
    
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
     vKey = "(AF)" & rs!num_placa & "(id)"
     Set itmX = lsw.ListItems.Add(, vKey, rs!num_placa)
         itmX.SubItems(1) = rs!Placa_Alterna & ""
         itmX.SubItems(2) = rs!Nombre
         itmX.SubItems(3) = Format(rs!fecha_adquisicion, "dd/mm/yyyy")
         itmX.SubItems(4) = Format(rs!fecha_instalacion, "dd/mm/yyyy")
         itmX.SubItems(5) = rs!TipoActivo
         itmX.SubItems(6) = rs!vida_util & IIf((rs!VIDA_UTIL_EN = "A"), " Año(s)", " Mes(es)")
         itmX.SubItems(7) = Format(rs!Valor_Historico, "Standard")
         itmX.SubItems(8) = Format(rs!Valor_Desecho, "Standard")
         
         curVH = curVH + rs!Valor_Historico
         curVR = curVR + rs!Valor_Desecho
         
     rs.MoveNext
    Loop
    rs.Close
     Set itmX = lsw.ListItems.Add(, "")
         itmX.SubItems(7) = "____________________"
         itmX.SubItems(8) = "____________________"
    
     Set itmX = lsw.ListItems.Add(, "")
         itmX.SubItems(7) = Format(curVH, "Standard")
         itmX.SubItems(8) = Format(curVR, "Standard")
          
       
       
  Case Else
    lsw.ListItems.Clear
    lsw.ColumnHeaders.Clear
    lsw.ColumnHeaders.Add , , "Placa", 1400
    lsw.ColumnHeaders.Add , , "Id Alterna", 1400
    lsw.ColumnHeaders.Add , , "Nombre", 4000
    lsw.ColumnHeaders.Add , , "Fecha Adq.", 1400
    lsw.ColumnHeaders.Add , , "Tipo", 2400
    lsw.ColumnHeaders.Add , , "Vida Util", 1400, 1
    lsw.ColumnHeaders.Add , , "Valor Historico", 1400, 1
    lsw.ColumnHeaders.Add , , "Dep. Ac. Mes Ant", 1400, 1
    lsw.ColumnHeaders.Add , , "Depreciacion Mes", 1400, 1
    lsw.ColumnHeaders.Add , , "Depreciación Ac", 1400, 1
    lsw.ColumnHeaders.Add , , "Valor Libros", 1400, 1
    lsw.ColumnHeaders.Add , , "Corte", 1400, vbCenter
    
    
    If vTipo = "A" Then
            strSQL = "select Num_placa, isnull(Placa_Alterna,'') as 'Placa_Alterna', Nombre, Fecha_adquisicion, Vida_Util, Vida_Util_En, TipoActivo, DEPRECIACION_PERIODO" _
                   & ",sum(Valor_historico) as 'VALOR_HISTORICO',sum(depreciacion_acum) as 'DEPRECIACION_AC'" _
                   & ",sum(Depreciacion_mes) as 'DEPRECIACION_MES'" _
                   & ",sum(VALOR_LIBROS) as 'VALOR_LIBROS'" _
                   & " from vActivos_depreciacion_actual" _
                   & " where Estado <> 'R'"
            
            If vTipoActivo <> "" Then strSQL = strSQL & " and tipo_activo = '" & vTipoActivo & "'"
            If vDepartamento <> "" Then strSQL = strSQL & " and cod_departamento = '" & vDepartamento & "'"
            If vSeccion <> "" Then strSQL = strSQL & " and cod_seccion = '" & vSeccion & "'"
                   
            strSQL = strSQL & " group by Num_placa,Placa_Alterna, Nombre, Fecha_adquisicion,Vida_Util,Vida_Util_En,TipoActivo,DEPRECIACION_PERIODO"
    
            Call OpenRecordSet(rs, strSQL, 0)
            Do While Not rs.EOF
             vKey = "(AF)" & rs!num_placa & "(id)"
             Set itmX = lsw.ListItems.Add(, vKey, rs!num_placa)
                 itmX.SubItems(1) = rs!Placa_Alterna
                 itmX.SubItems(2) = rs!Nombre
                 itmX.SubItems(3) = Format(rs!fecha_adquisicion, "dd/mm/yyyy")
                 itmX.SubItems(4) = rs!TipoActivo
                 itmX.SubItems(5) = rs!vida_util & IIf((rs!VIDA_UTIL_EN = "A"), " Año(s)", " Mes(es)")
                 itmX.SubItems(6) = Format(rs!Valor_Historico, "Standard")
                 itmX.SubItems(7) = Format(rs!DEPRECIACION_AC - rs!DEPRECIACION_MES, "Standard")
                 itmX.SubItems(8) = Format(rs!DEPRECIACION_MES, "Standard")
                 itmX.SubItems(9) = Format(rs!DEPRECIACION_AC, "Standard")
                 itmX.SubItems(10) = Format(rs!Valor_Libros, "Standard")
                 itmX.SubItems(11) = Format(rs!depreciacion_periodo, "dd/mm/yyyy")
                 
                 curVH = curVH + rs!Valor_Historico
                 curDepAnt = curDepAnt + rs!DEPRECIACION_AC - rs!DEPRECIACION_MES
                 curDepMes = curDepMes + rs!DEPRECIACION_MES
                 curDepAc = curDepAc + rs!DEPRECIACION_AC
                 curVL = curVL + rs!Valor_Libros
                 
             rs.MoveNext
            Loop
            rs.Close
            
             Set itmX = lsw.ListItems.Add(, "")
                 itmX.SubItems(6) = "____________________"
                 itmX.SubItems(7) = "____________________"
                 itmX.SubItems(8) = "____________________"
                 itmX.SubItems(9) = "____________________"
                 itmX.SubItems(10) = "____________________"
            
             Set itmX = lsw.ListItems.Add(, "")
                 itmX.SubItems(6) = Format(curVH, "Standard")
                 itmX.SubItems(7) = Format(curDepAnt, "Standard")
                 itmX.SubItems(8) = Format(curDepMes, "Standard")
                 itmX.SubItems(9) = Format(curDepAc, "Standard")
                 itmX.SubItems(10) = Format(curVL, "Standard")
    
    Else
      
            strSQL = "select * from vActivos_AuxiliarConsolidado" _
                    & " where Anio = " & Year(vFecha) & " and Mes = " & Month(vFecha)
            
            If vTipoActivo <> "" Then strSQL = strSQL & " and Tipo_activo = '" & vTipoActivo & "'"
            If vDepartamento <> "" Then strSQL = strSQL & " and cod_departamento = '" & vDepartamento & "'"
            If vSeccion <> "" Then strSQL = strSQL & " and cod_seccion = '" & vSeccion & "'"
            
            Call OpenRecordSet(rs, strSQL, 0)
            Do While Not rs.EOF
             vKey = "(AF)" & rs!num_placa & "(id)"
             Set itmX = lsw.ListItems.Add(, vKey, rs!num_placa)
                 itmX.SubItems(1) = rs!Placa_Alterna & ""
                 itmX.SubItems(2) = rs!Nombre
                 itmX.SubItems(3) = Format(rs!fecha_adquisicion, "dd/mm/yyyy")
                 itmX.SubItems(4) = rs!TipoActivo
                 itmX.SubItems(5) = rs!vida_util & IIf((rs!VIDA_UTIL_EN = "A"), " Año(s)", " Mes(es)")
                 itmX.SubItems(6) = Format(rs!Valor_Historico, "Standard")
                 itmX.SubItems(7) = Format(rs!DEPRECIACION_AC_CONSOLIDADO - rs!DEPRECIACION_MES_CONSOLIDADO, "Standard")
                 itmX.SubItems(8) = Format(rs!DEPRECIACION_MES_CONSOLIDADO, "Standard")
                 itmX.SubItems(9) = Format(rs!DEPRECIACION_AC_CONSOLIDADO, "Standard")
                 itmX.SubItems(10) = Format(rs!VALOR_LIBROS_CONSOLIDADO, "Standard")
                 
                 curVH = curVH + rs!Valor_Historico
                 curDepAnt = curDepAnt + rs!DEPRECIACION_AC_CONSOLIDADO - rs!DEPRECIACION_MES_CONSOLIDADO
                 curDepMes = curDepMes + rs!DEPRECIACION_MES_CONSOLIDADO
                 curDepAc = curDepAc + rs!DEPRECIACION_AC_CONSOLIDADO
                 curVL = curVL + rs!VALOR_LIBROS_CONSOLIDADO
                 
             rs.MoveNext
            Loop
            rs.Close
    
             Set itmX = lsw.ListItems.Add(, "")
                 itmX.SubItems(6) = "____________________"
                 itmX.SubItems(7) = "____________________"
                 itmX.SubItems(8) = "____________________"
                 itmX.SubItems(9) = "____________________"
                 itmX.SubItems(10) = "____________________"
            
             Set itmX = lsw.ListItems.Add(, "")
                 itmX.SubItems(6) = Format(curVH, "Standard")
                 itmX.SubItems(7) = Format(curDepAnt, "Standard")
                 itmX.SubItems(8) = Format(curDepMes, "Standard")
                 itmX.SubItems(9) = Format(curDepAc, "Standard")
                 itmX.SubItems(10) = Format(curVL, "Standard")
    
    
    End If ' TIPO A

End Select

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Public Function ConsultaSecciones(ByVal codigo_departamento As String, Optional ByVal Codigo_Seccion As String) As Boolean

On Error GoTo adoError

vSQL = "select * from dbo.Activos_SECCIONES where COD_DEPARTAMENTO = '" & codigo_departamento & "'"

If Not Codigo_Seccion = Empty Then
    vSQL = vSQL & " and COD_SECCION = '" & Codigo_Seccion & "'"
End If

If dbConsulta(vSQL) Then
    ConsultaSecciones = True
Else
    ConsultaSecciones = False
End If

Salir:
   Exit Function
adoError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
    ConsultaSecciones = False
End Function

Public Sub sbListaSecciones(pNodo As MSComctlLib.Node)
Dim NIndice As Integer

On Error GoTo adoError

wPosIni = 5

If ConsultaSecciones(DeCodificaPrimaryKey(pNodo.Key, wPosIni, "(id)")) Then
    With ADORecordSet
        While Not .EOF
          NIndice = sbTreeNodoCreate("(DP)" & .Fields!cod_departamento & "(SC)" & .Fields!cod_seccion & "(id)", pNodo, .Fields!Descripcion, "Seccion", pNodo.Index, 9)
     '     NIndice = sbTreeNodoCreate("0x0" & .Fields!cod_departamento & "(SC)" & .Fields!COD_SECCION & "(id)", pNodo, "AF0x0", "AF0x0", NIndice, 1)
          .MoveNext
        Wend
        .Close
    End With
End If
Salir:
    Exit Sub
adoError:
    MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Public Function sbListaAsientos(ByVal Anio As Integer, ByVal Mes As Integer, ByVal TipoAsiento As String, ByVal pNodo As MSComctlLib.Node, Optional ByVal pActivo As String) As Boolean
 Dim vKey As String
 Dim NIndice As Integer

 On Error GoTo vError

   vSQL = "select * from dbo.Activos_ASIENTOS where ANIO = " & Anio _
         & " and Mes = " & Mes & " and TIPO_ASIENTO = '" & TipoAsiento & "'"

   If dbConsulta(vSQL) Then
     With ADORecordSet
       While Not .EOF
        vKey = "(TA)" & .Fields!Tipo_Asiento & "(NA)" & .Fields!Num_Asiento & "(id)"
        NIndice = sbTreeNodoCreate(vKey, pNodo, .Fields!Num_Asiento, "Num_Asiento", pNodo.Index, 13)
        .MoveNext
       Wend
       .Close
     End With
   End If

Salir:
 Exit Function
vError:
    MsgBox fxSys_Error_Handler(Err.Description)
End Function


Private Function ConsultaDepartamentos(Optional ByVal Codigo As String) As Boolean

On Error GoTo adoError

vSQL = "select * from dbo.Activos_DEPARTAMENTOS"

If Not Codigo = Empty Then
    vSQL = vSQL & " where COD_DEPARTAMENTO = '" & Codigo & "'"
End If

If dbConsulta(vSQL) Then
    ConsultaDepartamentos = True
Else
    ConsultaDepartamentos = False
End If

Salir:
   Exit Function

adoError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
    ConsultaDepartamentos = False

End Function

Public Sub sbListaDepartamentosTodos()
Dim vKey As String

On Error GoTo vError

    Call CrearColumnas("Departamentos")

    If ConsultaDepartamentos Then

    With ADORecordSet
        While Not .EOF
          vKey = "(DP)" & .Fields!cod_departamento & "(id)"
          Set itmX = lsw.ListItems.Add(, vKey, .Fields!Descripcion)
              itmX.SubItems(1) = .Fields!cod_departamento
          .MoveNext
        Wend
        .Close
    End With
    End If

If tvActivos.Nodes.Count > 0 Then
    tvActivos.Nodes(1).Expanded = True

End If
Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Public Sub sbListaAsientosDetalle(pKey As String, Mes As Integer, Anio As Integer)
Dim Num_Asiento As String, Tipo_Asiento As String
Dim TD As Double, TC As Double
Dim Conta As Double, vKey As String

On Error GoTo vError

Call CrearColumnas("AsientosDetalle")

Tipo_Asiento = DeCodificaPrimaryKey(pKey, 5, "(NA)")
Num_Asiento = DeCodificaPrimaryKey(pKey, wPosIni, "(id)")

  vSQL = "select C.cod_Cuenta_Mask,D.Num_Linea,D.Num_Asiento,D.Tipo_Movimiento,D.Monto,D.Detalle" _
       & ",D.Referencia,D.Num_Documento,D.Cod_cuenta,C.Descripcion,A.Descripcion as DescripcionA" _
       & " From Activos_asiento A inner join Activos_asientos_Detalle D on A.num_asiento = D.num_asiento" _
       & " and A.tipo_asiento = D.tipo_asiento and A.COD_CONTABILIDAD = D.COD_CONTABILIDAD" _
       & " inner join CNTX_CUENTAS C on D.cod_cuenta = C.cod_cuenta and D.COD_CONTABILIDAD = C.COD_CONTABILIDAD" _
       & " Where A.tipo_asiento = '" & Tipo_Asiento & " and A.num_asiento = '" & Num_Asiento _
       & "' and A.anio = " & Anio & " and A.mes = " & Mes _
       & " Order by D.Tipo_Movimiento,D.cod_cuenta"


   If dbConsulta(vSQL) Then


    With ADORecordSet
         While Not .EOF
          vKey = "(LN)" & .Fields!Num_linea & "(id)"
          Set itmX = lsw.ListItems.Add(, vKey, .Fields!Cod_Cuenta_Mask)
             itmX.SubItems(1) = .Fields!Descripcion
             itmX.SubItems(4) = .Fields!Detalle
             If .Fields!TIPO_MOVIMIENTO = "D" Then
               itmX.SubItems(2) = Format(.Fields!Monto, "Standard")
               TD = TD + .Fields!Monto
             Else
               itmX.SubItems(3) = Format(.Fields!Monto, "Standard")
               TC = TC + .Fields!Monto
             End If
             itmX.SubItems(5) = IIf(IsNull(.Fields!Referencia), "", .Fields!Referencia)
             itmX.SubItems(6) = IIf(IsNull(.Fields!Num_Documento), "", .Fields!Num_Documento)
             itmX.SubItems(7) = .Fields!Num_linea
         .MoveNext
         Conta = Conta + 1
      Wend
         .Close
     End With
     
     Set itmX = lsw.ListItems.Add(, "MTOID1", "")
                 itmX.SubItems(2) = "_____________"
                 itmX.SubItems(3) = "_____________"
                 itmX.SubItems(7) = "0x0"
     Set itmX = lsw.ListItems.Add(, "MTOID2", "")
                 itmX.SubItems(1) = "TOTAL:"
                 itmX.SubItems(2) = Format(TD, "Standard")
                 itmX.SubItems(3) = Format(TC, "Standard")
                 itmX.SubItems(7) = "0x0"
   End If

Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Public Sub ListaSecciones(pNodo As MSComctlLib.Node)
Dim vKey As String

On Error GoTo vError

    Call CrearColumnas("Secciones")

    If ConsultaSecciones(DeCodificaPrimaryKey(pNodo.Key, 5, "(id)")) Then

    With ADORecordSet
        While Not .EOF
          vKey = "(DP)" & .Fields!cod_departamento & "(SC)" & .Fields!cod_seccion & "(id)"
          Set itmX = lsw.ListItems.Add(, vKey, .Fields!Descripcion)
              itmX.SubItems(1) = .Fields!cod_seccion
          .MoveNext
        Wend
        .Close
    End With
    End If

If tvActivos.Nodes.Count > 0 Then
    tvActivos.Nodes(1).Expanded = True
End If

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Public Sub sbTreeShowNodes(Indice As Integer, pNodo As MSComctlLib.Node)
Dim TmpIndice As Integer
Dim NIndice As Integer

On Error GoTo vError

    NIndice = sbTreeNodoCreate(1 & "(ET)", pNodo, "Tipos Activos", "TipoActivo", 1, 2)
    NIndice = sbTreeNodoCreate("0x0", pNodo, "AF0x0", "AF0x0", NIndice, 4)

    NIndice = sbTreeNodoCreate(2 & "(ED)", pNodo, "Departamentos", "Departamento", 1, 3)
    NIndice = sbTreeNodoCreate("0x01", pNodo, "AF0x0", "AF0x0", NIndice, 4)

    NIndice = sbTreeNodoCreate(3 & "(A)", pNodo, "Activos", "Activo", 1, 4)
 '   NIndice = sbTreeNodoCreate("0x02", pNodo, "AF0x0", "AF0x0", NIndice, 4)

    NIndice = sbTreeNodoCreate(4 & "(EM)", pNodo, "Modificaciones", "Modificacion", 1, 5)
 '   NIndice = sbTreeNodoCreate("0x03", pNodo, "AF0x0", "AF0x0", NIndice, 4)

    NIndice = sbTreeNodoCreate(5 & "(AS)", pNodo, "Asientos", "Asiento", 1, 6)
'    NIndice = sbTreeNodoCreate("0x04", pNodo, "AF0x0", "AF0x0", NIndice, 4)

    NIndice = sbTreeNodoCreate(5 & "(JR)", pNodo, "Justificaciones", "Justificacion", 1, 7)
 '   NIndice = sbTreeNodoCreate("0x05", pNodo, "AF0x0", "AF0x0", NIndice, 4)





Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description)
End Sub

Public Sub sbListaDepartamentos(Img As Integer, pNodo As MSComctlLib.Node)
Dim NIndice As Integer

On Error GoTo vError
    
    If ConsultaDepartamentos Then
    With ADORecordSet
        While Not .EOF
          NIndice = sbTreeNodoCreate("(DP)" & .Fields!cod_departamento & "(id)", pNodo, .Fields!Descripcion, "Departamentos", pNodo.Index, Img)
          NIndice = sbTreeNodoCreate("0x0" & .Fields!cod_departamento & "(id)", pNodo, "AF0x0", "AF0x0", NIndice, Img)
          .MoveNext
        Wend
        .Close
    End With
    End If

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description)
End Sub

Public Function sbTreeNodoSearch(pKey As String) As MSComctlLib.Node
Dim i As Long
With tvActivos
For i = 1 To .Nodes.Count
  If .Nodes(i).Key = pKey Then
    Set sbTreeNodoSearch = .Nodes(i)
    Exit For
  End If
Next i
End With
End Function

Public Sub CrearColumnas(pcolumna As String)

On Error GoTo vError

With lsw
    .ColumnHeaders.Clear
    .ListItems.Clear
    
    Select Case UCase(pcolumna)
        Case UCase("Departamentos")
            .ColumnHeaders.Add , , "Departamento", 8000
            .ColumnHeaders.Add , , "Código", 1400
        Case UCase("Secciones")
            .ColumnHeaders.Add , , "Sección", 8000
            .ColumnHeaders.Add , , "Código", 1400
        Case UCase("TiposActivos")
            .ColumnHeaders.Add , , "Descripción", 8000
            .ColumnHeaders.Add , , "Código", 1400
        Case UCase("Activos")
            .ColumnHeaders.Add , , "Nombre", 4000
            .ColumnHeaders.Add , , "Placa", 1400
            .ColumnHeaders.Add , , "Fecha Adq.", 1400
            .ColumnHeaders.Add , , "Tipo", 1400
            .ColumnHeaders.Add , , "Vida Util", 1400, 1
            .ColumnHeaders.Add , , "Valor historico", 1400, 1
            .ColumnHeaders.Add , , "Dep. Ac mes ant", 1400, 1
            .ColumnHeaders.Add , , "Depreciacion Mes", 1400, 1
            .ColumnHeaders.Add , , "Depreciación Ac", 1400, 1
            .ColumnHeaders.Add , , "Valor Libros", 1400, 1
        Case UCase("Asientos")
            .ColumnHeaders.Add , , "Nº Asiento", 1200
            .ColumnHeaders.Add , , "Tipo", 900
            .ColumnHeaders.Add , , "Fecha", 1300
            .ColumnHeaders.Add , , "Descripción", 2500
            .ColumnHeaders.Add , , "Debe       ", 1600, 1
            .ColumnHeaders.Add , , "Haber      ", 1600, 1
            .ColumnHeaders.Add , , "Aplicado", 900, 2
            .ColumnHeaders.Add , , "Notas", 4000
            .ColumnHeaders.Add , , "Tipo", 0
        Case UCase("ActivosxTipos")
            .ColumnHeaders.Add , , "Nombre", 4000
            .ColumnHeaders.Add , , "Placa", 1400
            .ColumnHeaders.Add , , "Fecha Adq.", 1400
            .ColumnHeaders.Add , , "Vida Util", 1400, 1
            .ColumnHeaders.Add , , "Valor historico", 1400, 1
            .ColumnHeaders.Add , , "Dep. Ac mes ant", 1400, 1
            .ColumnHeaders.Add , , "Depreciacion Mes", 1400, 1
            .ColumnHeaders.Add , , "Depreciación Ac", 1400, 1
            .ColumnHeaders.Add , , "Valor Libros", 1400, 1
        Case UCase("AsientosDetalle")
            .ColumnHeaders.Add , , "Cuenta", 1700
            .ColumnHeaders.Add , , "Descripción", 2500
            .ColumnHeaders.Add , , "Débito       ", 1600, 1
            .ColumnHeaders.Add , , "Crédito       ", 1600, 1
            .ColumnHeaders.Add , , "Detalle", 2500
            .ColumnHeaders.Add , , "Referencia", 800
            .ColumnHeaders.Add , , "Nº Documento", 800
            .ColumnHeaders.Add , , "Nº Línea", 0
        Case UCase("Justificaciones")
            .ColumnHeaders.Add , , "Justificación", 8000
            .ColumnHeaders.Add , , "Código", 1400
        Case UCase("Adiciones")
            .ColumnHeaders.Add , , "Nombre", 2000
            .ColumnHeaders.Add , , "Placa", 1200
            .ColumnHeaders.Add , , "Tipo", 2000
            .ColumnHeaders.Add , , "Justificacion", 2500
            .ColumnHeaders.Add , , "Fecha", 1400
            .ColumnHeaders.Add , , "Monto", 1400, 1
            .ColumnHeaders.Add , , "Descripción", 4000

    End Select
End With
Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description)
End Sub

Public Sub sbTreeCreateRoot(ByRef pNodo As MSComctlLib.Node, ptexto As String, ptag As String, pImagen As Integer)
On Error GoTo error

    tvActivos.Nodes.Clear
    Set pNodo = tvActivos.Nodes.Add()
        pNodo.Text = ptexto
        pNodo.Tag = ptag
        pNodo.Image = pImagen
Salir:
    Exit Sub
error:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, "SIA"
    Resume Salir
    
End Sub

Public Function SeparaNivelesCuenta(ByVal Cuenta As String, ByVal Niveles As Integer, _
          ByVal N1 As Integer, ByVal N2 As Integer, ByVal N3 As Integer, ByVal N4 As Integer, _
          ByVal N5 As Integer, Optional ByVal ptexto As Object) As String
On Error GoTo adoError

Dim vCuenta As String

If Not ptexto Is Nothing Then
    With ptexto
        If N1 = 0 Then Exit Function
            .Item(0) = Mid(Cuenta, 1, N1)
        If N2 = 0 Then Exit Function
            .Item(1) = Mid(Cuenta, N1 + 1, N2)
        If N3 = 0 Then Exit Function
            .Item(2) = Mid(Cuenta, N1 + N2 + 1, N3)
        If N4 = 0 Then Exit Function
            .Item(3) = Mid(Cuenta, N1 + N2 + N3 + 1, N4)
        If N5 = 0 Then Exit Function
        .Item(4) = Mid(Cuenta, N1 + N2 + N3 + N4 + 1, N5)
    End With
Else
    If N1 = 0 Then GoTo Salir
    vCuenta = Mid(Cuenta, 1, N1)
    If N2 = 0 Then GoTo Salir
    vCuenta = vCuenta & "-" & Mid(Cuenta, N1 + 1, N2)
    If N3 = 0 Then GoTo Salir
    vCuenta = vCuenta & "-" & Mid(Cuenta, N1 + N2 + 1, N3)
    If N4 = 0 Then GoTo Salir
    vCuenta = vCuenta & "-" & Mid(Cuenta, N1 + N2 + N3 + 1, N4)
    If N5 = 0 Then GoTo Salir
    vCuenta = vCuenta & "-" & Mid(Cuenta, N1 + N2 + N3 + N4 + 1, N5)
End If
SeparaNivelesCuenta = vCuenta
Salir:
   SeparaNivelesCuenta = vCuenta
   Exit Function
adoError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, "Activos Fijos"
Resume

End Function

Public Function sbTreeNodoCreate(ByVal pKey As String, ByVal pNodo As MSComctlLib.Node _
                            , ptexto As String, ptag As String, pindice, pImagen As Integer, Optional vExpand As Boolean = True) As Integer

On Error GoTo error
    
    Set pNodo = tvActivos.Nodes.Add(pindice, tvwChild)
        pNodo.Text = ptexto
        pNodo.Tag = ptag
        pNodo.Image = pImagen
        pNodo.Key = pKey
        'pNodo.Expanded = vExpand
        sbTreeNodoCreate = pNodo.Index

Salir:
    Exit Function
error:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
    Resume Salir
    
End Function


'----------------------------------------------------------------------------------------------------------------------------------------------
' clsActivos: Replace
'----------------------------------------------------------------------------------------------------------------------------------------------


Public Function sbActivosConsultas(Optional ByVal fecha As String, Optional ByVal Placa As String, _
               Optional ByVal departamento As Integer, Optional ByVal seccion As Integer, Optional ByVal tipo_activo As String) As Boolean

Dim citem As ComboItem
Dim i As Integer

On Error GoTo vError

    If Not fecha = Empty Then
        vSQL = "exec Activos_calculo_depreciacion_rep 1, '" & Format(fecha, "yyyymmdd") & "'"
    Else
        vSQL = "select * from Activos_Principal "
    End If
     If Not Placa = Empty Then
       vSQL = vSQL & " where Num_Placa = '" & Placa & "'"
     End If
     If Not departamento = Empty Then
        vSQL = vSQL & "," & departamento & "," & seccion
     ElseIf Placa = Empty Then
         vSQL = vSQL & ",Null,Null"
     End If
     If Not tipo_activo = Empty Then
        vSQL = vSQL & ",'" & tipo_activo & "'"
     ElseIf Placa = Empty Then
        vSQL = vSQL & ",Null"
     End If
      
    If dbConsulta(vSQL) Then
        sbActivosConsultas = True
    Else
        sbActivosConsultas = False
    End If

Salir:
    Exit Function
vError:
    MsgBox fxSys_Error_Handler(Err.Description)
    sbActivosConsultas = False
    

End Function


Private Function sbActivosConsultaPorSeccion(departamento As String, seccion As String) As Boolean
On Error GoTo adoError

vSQL = "select T1.Num_Placa,t1.nombre, t1.num_placa, t2.descripcion,0 as DepreciacionAc,t1.valor_historico, " & _
       "t1.valor_desecho,t1.vida_util,vida_util_en,fecha_instalacion " & _
       "from dbo.Activos_Principal t1, dbo.Activos_tipo_activo t2 " & _
       " where t1.cod_departamento = " & departamento & _
       " and t1.cod_seccion = " & seccion & _
       " and t1.tipo_activo = t2.tipo_activo "

If dbConsulta(vSQL) Then
    sbActivosConsultaPorSeccion = True
Else
    sbActivosConsultaPorSeccion = False
End If

Salir:
   Exit Function
adoError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    sbActivosConsultaPorSeccion = False
End Function

'Se puede eliminar
Public Sub sbActivosListaPorTipo()
On Error GoTo vError

Dim VidaUtil As String
Dim vKey As String
Dim vTipo As String
Dim VU As Integer
Dim VH As Double
Dim DAA As Double
Dim DM As Double
Dim DAM As Double
Dim VL As Double
VH = 0
DAA = 0
DM = 0
DAM = 0
VL = 0
    vTipo = DeCodificaPrimaryKey(tvActivos.SelectedItem.Key, 5, "(id)")
    
    Call CrearColumnas("ActivosTipos")

   If sbActivosConsultas(dtpFecha, , , , vTipo) Then
    With ADORecordSet
        While Not .EOF
          vKey = "(AF)" & .Fields!Placa & "(id)"
          Set itmX = lsw.ListItems.Add(, vKey, .Fields!Nombre)
              itmX.SubItems(1) = .Fields!Placa
              itmX.SubItems(2) = Format(.Fields!FECHA_ADQ, "dd/mm/yyyy")
              VidaUtil = .Fields!VIDA_UTIL_EN
              If VidaUtil = "A" Then
                VidaUtil = "Años"
                 VU = .Fields!vida_util / 12
              Else
                VidaUtil = "Meses"
                VU = .Fields!vida_util
              End If
              itmX.SubItems(3) = VU & " " & VidaUtil
              itmX.SubItems(4) = IIf(IsNull(.Fields!Valor_Historico), 0, Format(.Fields!Valor_Historico, "###,###,##0.00"))
              itmX.SubItems(5) = Format(.Fields!Depreciacion_Ac_Ant, "###,###,##0.00")
              itmX.SubItems(6) = Format(.Fields!DEPRECIACION_MES, "###,###,##0.00")
              itmX.SubItems(7) = Format((.Fields!DEPRECIACION_MES + .Fields!Depreciacion_Ac_Ant), "###,###,##0.00")
              itmX.SubItems(8) = Format((.Fields!Valor_Historico - .Fields!Valor_Desecho) - (.Fields!DEPRECIACION_MES + .Fields!Depreciacion_Ac_Ant), "###,###,##0.00")
              VH = VH + IIf(IsNull(.Fields!Valor_Historico), 0, .Fields!Valor_Historico)
              DAA = DAA + .Fields!Depreciacion_Ac_Ant
              DM = DM + .Fields!DEPRECIACION_MES
              DAM = DAM + (.Fields!DEPRECIACION_MES + .Fields!Depreciacion_Ac_Ant)
              VL = VL + (.Fields!Valor_Historico - .Fields!Valor_Desecho) - (.Fields!DEPRECIACION_MES + .Fields!Depreciacion_Ac_Ant)
          .MoveNext
        Wend
        .Close
    End With
    Set itmX = lsw.ListItems.Add(, "MTOID1", "")
                 itmX.SubItems(4) = "_____________"
                 itmX.SubItems(5) = "_____________"
                 itmX.SubItems(6) = "_____________"
                 itmX.SubItems(7) = "_____________"
                 itmX.SubItems(8) = "_____________"
                 
             Set itmX = lsw.ListItems.Add(, "MTOID2", "")
                 itmX.SubItems(2) = "TOTAL:"
                 itmX.SubItems(4) = Format(VH, "###,###,###,##0.00")
                 itmX.SubItems(5) = Format(DAA, "###,###,###,##0.00")
                 itmX.SubItems(6) = Format(DM, "###,###,###,##0.00")
                 itmX.SubItems(7) = Format(DAM, "###,###,###,##0.00")
                 itmX.SubItems(8) = Format(VL, "###,###,###,##0.00")
   End If

Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


'Se puede Eliminar
Public Sub sbActivosListaPorSeccion()
On Error GoTo vError

Dim vDepartamento As String
Dim vSeccion As String
Dim VidaUtil As String
Dim vKey As String
Dim VU As Integer
Dim VH As Double
Dim DAA As Double
Dim DM As Double
Dim DAM As Double
Dim VL As Double
VH = 0
DAA = 0
DM = 0
DAM = 0
VL = 0


    Call CrearColumnas("Activos")
    
    vDepartamento = DeCodificaPrimaryKey(tvActivos.SelectedItem.Key, 5, "(SC)")
    vSeccion = DeCodificaPrimaryKey(tvActivos.SelectedItem.Key, wPosIni, "(id)")
    
    If sbActivosConsultas(dtpFecha, , vDepartamento, vSeccion) Then
    With ADORecordSet
        While Not .EOF
          vKey = "(AF)" & .Fields!Placa & "(id)"
          Set itmX = lsw.ListItems.Add(, vKey, .Fields!Nombre)
              itmX.SubItems(1) = .Fields!Placa
              itmX.SubItems(2) = Format(.Fields!FECHA_ADQ, "dd/mm/yyyy")
              itmX.SubItems(3) = .Fields!Tipo
              VidaUtil = .Fields!VIDA_UTIL_EN
              If VidaUtil = "A" Then
                VidaUtil = "Años"
                VU = .Fields!vida_util / 12
              Else
                VU = .Fields!vida_util
                VidaUtil = "Meses"
              End If
              itmX.SubItems(4) = VU & " " & VidaUtil
              itmX.SubItems(5) = IIf(IsNull(.Fields!Valor_Historico), 0, Format(.Fields!Valor_Historico, "###,###,##0.00"))
              itmX.SubItems(6) = Format(.Fields!Depreciacion_Ac_Ant, "###,###,##0.00")
              itmX.SubItems(7) = Format(.Fields!DEPRECIACION_MES, "###,###,##0.00")
              itmX.SubItems(8) = Format((.Fields!DEPRECIACION_MES + .Fields!Depreciacion_Ac_Ant), "###,###,##0.00")
              itmX.SubItems(9) = Format((.Fields!Valor_Historico - .Fields!Valor_Desecho) - (.Fields!DEPRECIACION_MES + .Fields!Depreciacion_Ac_Ant), "###,###,##0.00")
              VH = VH + IIf(IsNull(.Fields!Valor_Historico), 0, .Fields!Valor_Historico)
              DAA = DAA + .Fields!Depreciacion_Ac_Ant
              DM = DM + .Fields!DEPRECIACION_MES
              DAM = DAM + (.Fields!DEPRECIACION_MES + .Fields!Depreciacion_Ac_Ant)
              VL = VL + (.Fields!Valor_Historico - .Fields!Valor_Desecho) - (.Fields!DEPRECIACION_MES + .Fields!Depreciacion_Ac_Ant)
          .MoveNext
        Wend
        .Close
    End With
    Set itmX = lsw.ListItems.Add(, "MTOID1", "")
                 itmX.SubItems(5) = "_____________"
                 itmX.SubItems(6) = "_____________"
                 itmX.SubItems(7) = "_____________"
                 itmX.SubItems(8) = "_____________"
                 itmX.SubItems(9) = "_____________"
                 
             Set itmX = lsw.ListItems.Add(, "MTOID2", "")
                 itmX.SubItems(2) = "TOTAL:"
                 itmX.SubItems(5) = Format(VH, "###,###,###,##0.00")
                 itmX.SubItems(6) = Format(DAA, "###,###,###,##0.00")
                 itmX.SubItems(7) = Format(DM, "###,###,###,##0.00")
                 itmX.SubItems(8) = Format(DAM, "###,###,###,##0.00")
                 itmX.SubItems(9) = Format(VL, "###,###,###,##0.00")
    End If
Salir:
    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Resume Salir
End Sub


Public Sub sbActivosLista(ByVal fecha As Date)

Dim VidaUtil As String
Dim vKey As String
Dim VU As Integer
Dim VH As Double
Dim DAA As Double
Dim DM As Double
Dim DAM As Double
Dim VL As Double

On Error GoTo vError


VH = 0
DAA = 0
DM = 0
DAM = 0
VL = 0
    
    Call CrearColumnas("Activos")
    
    If sbActivosConsultas(fecha) Then
    
    With ADORecordSet
        While Not .EOF
          vKey = "(AF)" & .Fields!Placa & "(id)"
          Set itmX = lsw.ListItems.Add(, vKey, .Fields!Nombre)
              itmX.SubItems(1) = .Fields!Placa
              itmX.SubItems(2) = Format(.Fields!FECHA_ADQ, "dd/mm/yyyy")
              itmX.SubItems(3) = .Fields!Tipo
              VidaUtil = .Fields!VIDA_UTIL_EN
              If VidaUtil = "A" Then
                VU = .Fields!vida_util / 12
                VidaUtil = "Años"
              Else
                VU = .Fields!vida_util / 12
                VidaUtil = "Meses"
              End If
              itmX.SubItems(4) = VU & " " & VidaUtil
              itmX.SubItems(5) = IIf(IsNull(.Fields!Valor_Historico), 0, Format(.Fields!Valor_Historico, "Standard"))
              itmX.SubItems(6) = Format(.Fields!Depreciacion_Ac_Ant, "Standard")
              itmX.SubItems(7) = Format(.Fields!DEPRECIACION_MES, "Standard")
              itmX.SubItems(8) = Format((.Fields!DEPRECIACION_MES + .Fields!Depreciacion_Ac_Ant), "Standard")
              itmX.SubItems(9) = Format((.Fields!Valor_Historico - .Fields!Valor_Desecho) - (.Fields!DEPRECIACION_MES + .Fields!Depreciacion_Ac_Ant), "Standard")
              VH = VH + IIf(IsNull(.Fields!Valor_Historico), 0, .Fields!Valor_Historico)
              DAA = DAA + .Fields!Depreciacion_Ac_Ant
              DM = DM + .Fields!DEPRECIACION_MES
              DAM = DAM + (.Fields!DEPRECIACION_MES + .Fields!Depreciacion_Ac_Ant)
              VL = VL + (.Fields!Valor_Historico - .Fields!Valor_Desecho) - (.Fields!DEPRECIACION_MES + .Fields!Depreciacion_Ac_Ant)
          .MoveNext
        Wend
        .Close
    End With
    Set itmX = lsw.ListItems.Add(, "MTOID1", "")
                 itmX.SubItems(5) = "_____________"
                 itmX.SubItems(6) = "_____________"
                 itmX.SubItems(7) = "_____________"
                 itmX.SubItems(8) = "_____________"
                 itmX.SubItems(9) = "_____________"
                 
             Set itmX = lsw.ListItems.Add(, "MTOID2", "")
                 itmX.SubItems(2) = "TOTAL:"
                 itmX.SubItems(5) = Format(VH, "Standard")
                 itmX.SubItems(6) = Format(DAA, "Standard")
                 itmX.SubItems(7) = Format(DM, "Standard")
                 itmX.SubItems(8) = Format(DAM, "Standard")
                 itmX.SubItems(9) = Format(VL, "Standard")
                 
    End If

Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Public Sub sbActivosListaAsientos(ByVal pMes As Integer, ByVal pAnio As Integer)
 Dim TmpNAsiento As String
 Dim Linea As Long
 Dim vKey As String
 
 On Error GoTo adoError
   
   
   Linea = 0
   TmpNAsiento = "0x0"
   
   vSQL = "select A.tipo_asiento,A.num_asiento,A.fecha_asiento,A.descripcion,A.notas" _
        & ",A.fecha_traslado,sum(D.monto_debito) as MontoDebito,sum(D.monto_credito) as MontoCredito" _
        & " from Activos_asientos A inner join Activos_asientos_detalle D on A.tipo_Asiento = D.tipo_asiento" _
        & " and A.COD_CONTABILIDAD = D.COD_CONTABILIDAD and A.num_asiento = D.num_asiento" _
        & " where A.anio = " & pAnio & " and A.mes = " & pMes _
        & " group by A.tipo_asiento,A.num_asiento,A.fecha_asiento,A.descripcion,A.notas" _
        & ",A.fecha_traslado,D.monto_debito,D.monto_credito"
   
   Call CrearColumnas("Asientos")

   
   If dbConsulta(vSQL) Then
     With ADORecordSet
        While Not .EOF
            If TmpNAsiento <> .Fields!Num_Asiento Then
                vKey = "(TA)" & .Fields!Tipo_Asiento & "(NA)" & .Fields!Num_Asiento & "(id)"
                Set itmX = lsw.ListItems.Add(, vKey, .Fields!Num_Asiento)
                itmX.SubItems(1) = .Fields!Tipo_Asiento
                itmX.SubItems(2) = .Fields!fecha_asiento
                itmX.SubItems(3) = .Fields!Descripcion
                
                If IsNull(.Fields!fecha_traslado) Then
                    itmX.SubItems(6) = "NO"
                Else
                    itmX.SubItems(6) = "SI"
                End If
                
                If Not IsNull(.Fields!Notas) Then
                    itmX.SubItems(7) = .Fields!Notas
                End If
                
                itmX.SubItems(8) = .Fields!Tipo_Asiento
                TmpNAsiento = .Fields!Num_Asiento
                Linea = Linea + 1
            
            End If
            
                lsw.ListItems(Linea).SubItems(4) = Format(.Fields!montoDebito, "Standard")
                lsw.ListItems(Linea).SubItems(5) = Format(.Fields!montoCredito, "Standard")
            
            .MoveNext
        Wend
     End With
   End If
   
Salir:
   Exit Sub
adoError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Public Sub sbActivosToolbar_Clic(pKey As String)

On Error GoTo vError

Select Case pKey
  Case "Nuevo"
'      frmActivos_Main.Show
      gActivos.Placa = ""
      Call sbClassCall("Activos", 0, "frmActivos_Main")

  Case "Propiedades"
'      Load frmActivos_Main
'      frmActivos_Main.Show
'      Call frmActivos_Main.sbConsultaExterna(frmActivos_Explorador.lsw.SelectedItem.Text)
'      frmActivos_Main.txtCodigo.SetFocus
       
       gActivos.Placa = lsw.SelectedItem.Text
       
       Call sbClassCall("Activos", 0, "frmActivos_Main")
      
  Case "Generar"
'      frmActivos_CierrePeriodo.Show vbModal
      Call sbClassCall("Activos", 0, "frmActivos_CierrePeriodo")
  
  Case "PaseAsientos"
'       frmActivos_TrasladoAsientos.Show vbModal
      Call sbClassCall("Activos", 0, "frmActivos_TrasladoAsientos")
       
       
End Select

Exit Sub

vError:
'    MsgBox fxSys_Error_Handler(Err.Description)
'    Resume
End Sub




'----------------------------------------------------------------------------------------------------------------------------------------------
' clsTiposActivos: Replace
'----------------------------------------------------------------------------------------------------------------------------------------------

Public Sub sbTiposActivosToolbar_Clic(pKey As String)
On Error GoTo vError
    Select Case pKey
        Case "Nuevo"
            Call sbClassCall("Activos", 0, "frmActivos_TiposActivo")

        Case "Propiedades"
            Call PropiedadesTiposActivo
        End Select
Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description)
End Sub

Private Function fxTipoActivosConsulta(Optional Codigo As String) As Boolean
On Error GoTo adoError

vSQL = "SELECT * FROM DBO.Activos_TIPO_ACTIVO"

If Not Codigo = Empty Then
    vSQL = vSQL & " where TIPO_ACTIVO = '" & Codigo & "'"
End If

If dbConsulta(vSQL) Then
    fxTipoActivosConsulta = True
Else
    fxTipoActivosConsulta = False
End If

Salir:
   Exit Function
adoError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    fxTipoActivosConsulta = False
End Function

Public Sub sbTiposActivosArbol(pNodo)
Dim NIndice As Integer

On Error GoTo adoError

wPosIni = 5

If fxTipoActivosConsulta Then
    With ADORecordSet
        While Not .EOF
          NIndice = sbTreeNodoCreate("(TA)" & .Fields!tipo_activo & "(id)", pNodo, .Fields!Descripcion, "Tipo", pNodo.Index, 11)
     '     NIndice = Mantenimientos.CreateNode("0x0" & "(TA)" & .Fields!tipo_activo & "(id)", pNodo, "AF0x0", "AF0x0", NIndice, 11)
          .MoveNext
        Wend
        .Close
    End With
End If


Salir:
    Exit Sub
adoError:
    MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Public Sub sbTiposActivosLista()
Dim vKey As String

On Error GoTo vError
    
    Call CrearColumnas("TiposActivos")

    If fxTipoActivosConsulta Then
    
    With ADORecordSet
        While Not .EOF
          vKey = "(TA)" & .Fields!tipo_activo & "(id)"
          Set itmX = lsw.ListItems.Add(, vKey, .Fields!Descripcion)
              itmX.SubItems(1) = .Fields!tipo_activo
          .MoveNext
        Wend
        .Close
    End With
    End If

Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





'----------------------------------------------------------------------------------------------------------------------------------------------
' clsJustificaciones: Replace
'----------------------------------------------------------------------------------------------------------------------------------------------



Public Function fxJustificacionConsulta(Optional Codigo As String) As Boolean

On Error GoTo adoError

vSQL = "SELECT * FROM DBO.Activos_JUSTIFICACIONES"

If Not Codigo = Empty Then
    vSQL = vSQL & " where CODIGO_JUSTIFICACION = '" & Codigo & "'"
End If

If dbConsulta(vSQL) Then
    fxJustificacionConsulta = True
Else
    fxJustificacionConsulta = False
End If


Salir:
   Exit Function
adoError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    fxJustificacionConsulta = False

End Function

Public Sub sbJustificacionToolbar_Clic(pKey As String)
On Error GoTo vError
    
Select Case pKey
    Case "Nuevo"
        Call sbClassCall("Activos", 0, "frmActivos_Justificaciones")

    Case "Propiedades"
        Call PropiedadesJustificacion
End Select

Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Public Sub sbJustificacionArbol(pNodo)

Dim NIndice As Integer

On Error GoTo adoError

wPosIni = 5

If fxJustificacionConsulta() Then
    With ADORecordSet
        While Not .EOF
          NIndice = sbTreeNodoCreate("(JF)" & .Fields!CODIGO_JUSTIFICACION & "(id)", pNodo, .Fields!Justificacion, "JustificacionDet", pNodo.Index, 10)
          NIndice = sbTreeNodoCreate("0x0" & "(JF)" & .Fields!CODIGO_JUSTIFICACION & "(id)", pNodo, "AF0x0", "AF0x0", NIndice, 10)
          .MoveNext
        Wend
        .Close
    End With
End If


Salir:
    Exit Sub
adoError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Public Sub sbJustificacionLista()
Dim vKey As String

On Error GoTo vError
    
Call CrearColumnas("Justificaciones")

If fxJustificacionConsulta Then
    With ADORecordSet
        While Not .EOF
          vKey = "(JF)" & .Fields!COD_JUSTIFICACION & "(id)"
          Set itmX = lsw.ListItems.Add(, vKey, .Fields!Descripcion)
              itmX.SubItems(1) = .Fields!COD_JUSTIFICACION
          .MoveNext
        Wend
        .Close
    End With
End If

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





'----------------------------------------------------------------------------------------------------------------------------------------------
' clsAdicionesRetiros: Replace
'----------------------------------------------------------------------------------------------------------------------------------------------



Public Function fxAdicionesRetirosConsulta(Optional Placa As String) As Boolean
Dim strSQL As String

On Error GoTo vError
    strSQL = "select A.num_placa,A.nombre,R.id_AddRet,R.tipo,J.descripcion as Justificacion,R.descripcion,R.fecha,R.monto" _
           & " from Activos_retiro_adicion R inner join Activos_justificaciones J on R.cod_justificacion = J.cod_justificacion" _
           & " inner join Activos_Principal A on R.num_placa = A.num_placa" _
           & " Where R.FECHA <= '" & Format(dtpFecha.Value, "yyyy/mm/dd") & "'"
    
     If Placa <> Empty Then
        strSQL = strSQL & " and A.NUM_PLACA = '" & Placa & "'"
     End If

    fxAdicionesRetirosConsulta = dbConsulta(strSQL)
       
Salir:
    Exit Function

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    fxAdicionesRetirosConsulta = False
    
End Function

Public Sub sbAdicionesRetirosToolbar_Clic(pKey As String)
On Error GoTo vError
    Select Case pKey
        Case "Nuevo"
             Call sbClassCall("Activos", 0, "frmActivos_AdicionRetiro")
        Case "Propiedades"
            Call PropiedadesAdicionRetiro
    End Select
Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Public Sub sbAdicionesRetirosLista()

On Error GoTo vError

Call CrearColumnas("Adiciones")


If fxAdicionesRetirosConsulta Then
    With ADORecordSet
        While Not .EOF
            Set itmX = lsw.ListItems.Add()
                itmX.Key = "(AF)" & .Fields!num_placa & "(AR)" & .Fields!Id_AddRet & "(id)"
                itmX.Text = .Fields!Nombre
                itmX.SubItems(1) = .Fields!num_placa
                itmX.SubItems(2) = IIf(.Fields!Tipo = "A", "Adición", "Retiro")
                itmX.SubItems(3) = .Fields!Justificacion
                itmX.SubItems(4) = Format(.Fields!fecha, "dd/mm/yyyy")
                itmX.SubItems(5) = Format(.Fields!Monto, "Standard")
                itmX.SubItems(6) = .Fields!Descripcion
                itmX.Tag = .Fields!Id_AddRet
            .MoveNext
        Wend
    End With
End If

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub






'-----------------------------------------------------------------------------------------------------------------------------
' clsEjectuar: Reemplazo
'-----------------------------------------------------------------------------------------------------------------------------




'Private Function dbModificar(ByVal pSQL As String) As Boolean
'On Error GoTo adoError
'
'Dim Rows As Long
'
''Call glogon.Conection.Execute(pSQL, Rows)
'
'Call ConectionExecute(pSQL, 0, Rows)
'
'If Rows > 0 Then
'    Modificar = True
'Else
'    Modificar = False
'End If
'
'Salir:
'   Exit Function
'adoError:
'    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
'    Modificar = False
'
'End Function

Private Function dbConsulta(ByVal pSQL As String) As Boolean
On Error GoTo adoError

Call OpenRecordSet(ADORecordSet, pSQL, 0)

If ADORecordSet.RecordCount > 0 Then
    dbConsulta = True
Else
    dbConsulta = False
End If

Salir:
   Exit Function
adoError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
    dbConsulta = False

End Function



Private Sub txtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_proveedor,descripcion from Activos_proveedores"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    If Trim(gBusquedas.Resultado) <> Trim(txtProveedor.Tag) Then
       txtProveedor.Tag = gBusquedas.Resultado
       txtProveedor.Text = gBusquedas.Resultado2
    End If
End If

End Sub



Private Sub txtResponsable_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    gBusquedas.Consulta = "select Identificacion, Nombre from Activos_Personas"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    If Trim(gBusquedas.Resultado) <> Trim(txtResponsable.Tag) Then
       txtResponsable.Tag = gBusquedas.Resultado
       txtResponsable.Text = gBusquedas.Resultado2
    End If
End If
End Sub
