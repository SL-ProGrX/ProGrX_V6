VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCO_ControlLista 
   Caption         =   "Listado de Carteras de Cobros a Ejecutivos"
   ClientHeight    =   8775
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   15435
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   15435
   WindowState     =   2  'Maximized
   Begin VB.Frame fraFiltros 
      Caption         =   "Filtros Adicionales"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   7095
      Left            =   3840
      TabIndex        =   27
      Top             =   1440
      Visible         =   0   'False
      Width           =   9735
      Begin VB.CheckBox chkInfoContracto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Incluir Info. de Contracto"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7800
         TabIndex        =   76
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cboInstitucion 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   1320
         Width           =   3015
      End
      Begin VB.ComboBox cboOficina 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   1320
         Width           =   4575
      End
      Begin MSComctlLib.ListView lswConGarantias 
         Height          =   2295
         Left            =   360
         TabIndex        =   70
         Top             =   4560
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4048
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.CheckBox chkAntiguedad 
         Appearance      =   0  'Flat
         Caption         =   "Todas las Antiguedades"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   5040
         TabIndex        =   69
         Top             =   4200
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkGarantias 
         Appearance      =   0  'Flat
         Caption         =   "Todas las Garantías"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   360
         TabIndex        =   68
         Top             =   4200
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.TextBox txtCtaHasta 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   8640
         TabIndex        =   67
         Text            =   "80"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtCtaDesde 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   8040
         TabIndex        =   66
         Text            =   "1"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtArregloDescFiltro 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2520
         TabIndex        =   65
         ToolTipText     =   "Presione (F4) para Consultar"
         Top             =   3600
         Width           =   3735
      End
      Begin VB.TextBox txtArregloFiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   64
         ToolTipText     =   "Presione (F4) para Consultar"
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox txtCausaDescFiltro 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2520
         TabIndex        =   63
         ToolTipText     =   "Presione (F4) para Consultar"
         Top             =   3240
         Width           =   3735
      End
      Begin VB.TextBox txtCausaFiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   62
         ToolTipText     =   "Presione (F4) para Consultar"
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox txtGestionDescFiltro 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2520
         TabIndex        =   61
         ToolTipText     =   "Presione (F4) para Consultar"
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox txtGestionFiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   60
         ToolTipText     =   "Presione (F4) para Consultar"
         Top             =   2880
         Width           =   855
      End
      Begin VB.CheckBox chkFechaPago 
         Appearance      =   0  'Flat
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8400
         TabIndex        =   48
         Top             =   2880
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox txtDiasAtencion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3240
         TabIndex        =   39
         Text            =   "15"
         Top             =   2280
         Width           =   615
      End
      Begin VB.ComboBox cboTipoCasos 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1920
         Width           =   4575
      End
      Begin VB.ComboBox cboOrden 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox cboEstado 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cboOrdenTipo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cboCarteras 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   960
         Width           =   4575
      End
      Begin MSComctlLib.ListView lswConAntiguedad 
         Height          =   2295
         Left            =   5040
         TabIndex        =   71
         Top             =   4560
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4048
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
         EndProperty
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFPInicio 
         Height          =   312
         Left            =   6840
         TabIndex        =   121
         Top             =   3240
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
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
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFPCorte 
         Height          =   312
         Left            =   8040
         TabIndex        =   122
         Top             =   3240
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
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
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "Institución"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   20
         Left            =   6360
         TabIndex        =   75
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "Oficina"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   11
         Left            =   360
         TabIndex        =   73
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   120
         X2              =   9600
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "Fechas de Pago"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   18
         Left            =   6840
         TabIndex        =   47
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "Arreglo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   17
         Left            =   360
         TabIndex        =   46
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "Causa"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   16
         Left            =   360
         TabIndex        =   45
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "Gestión"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   15
         Left            =   360
         TabIndex        =   44
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   120
         X2              =   9600
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label lblAtencion 
         Appearance      =   0  'Flat
         Caption         =   "dias desde la última gestión"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   1
         Left            =   3960
         TabIndex        =   40
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label lblAtencion 
         Appearance      =   0  'Flat
         Caption         =   "Casos con más de "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   0
         Left            =   1680
         TabIndex        =   38
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "Tipos de Casos a visualizar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Index           =   0
         Left            =   360
         TabIndex        =   36
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   240
         X2              =   9600
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "Ordenar por"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   8
         Left            =   360
         TabIndex        =   35
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   9
         Left            =   4080
         TabIndex        =   34
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "# Cuotas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   12
         Left            =   7200
         TabIndex        =   33
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "Carteras"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   13
         Left            =   360
         TabIndex        =   32
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Frame fraGestion 
      Caption         =   "Gestiones"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   600
      TabIndex        =   49
      ToolTipText     =   "Salir"
      Top             =   1440
      Width           =   11295
      Begin XtremeSuiteControls.GroupBox fraLista 
         Height          =   4212
         Left            =   5880
         TabIndex        =   118
         Top             =   1080
         Width           =   5292
         _Version        =   1572864
         _ExtentX        =   9334
         _ExtentY        =   7429
         _StockProps     =   79
         Caption         =   "Gestion"
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
         Appearance      =   21
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lswLista 
            Height          =   3372
            Left            =   120
            TabIndex        =   119
            Top             =   720
            Width           =   4932
            _Version        =   1572864
            _ExtentX        =   8700
            _ExtentY        =   5948
            _StockProps     =   77
            BackColor       =   -2147483643
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
         End
         Begin XtremeSuiteControls.FlatEdit txtListaFiltro 
            Height          =   330
            Left            =   120
            TabIndex        =   130
            Top             =   360
            Width           =   4935
            _Version        =   1572864
            _ExtentX        =   8705
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.PushButton cmdAplica 
         Height          =   612
         Left            =   9120
         TabIndex        =   87
         Top             =   5400
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Picture         =   "frmCO_ControlLista.frx":0000
      End
      Begin VB.ComboBox cboOperacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   3120
         Width           =   1815
      End
      Begin XtremeSuiteControls.DateTimePicker dtpVence 
         Height          =   330
         Left            =   1440
         TabIndex        =   88
         Top             =   2520
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1455
         Left            =   1440
         TabIndex        =   127
         Top             =   3720
         Width           =   4455
         _Version        =   1572864
         _ExtentX        =   7858
         _ExtentY        =   2566
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtPersonaGestion 
         Height          =   330
         Left            =   1440
         TabIndex        =   128
         Top             =   600
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   330
         Left            =   6000
         TabIndex        =   129
         Top             =   600
         Width           =   4935
         _Version        =   1572864
         _ExtentX        =   8705
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtGestion 
         Height          =   330
         Left            =   1440
         TabIndex        =   131
         Top             =   1080
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtGestionDesc 
         Height          =   330
         Left            =   2280
         TabIndex        =   132
         Top             =   1080
         Width           =   3495
         _Version        =   1572864
         _ExtentX        =   6165
         _ExtentY        =   582
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCausa 
         Height          =   330
         Left            =   1440
         TabIndex        =   133
         Top             =   1440
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtCausaDesc 
         Height          =   330
         Left            =   2280
         TabIndex        =   134
         Top             =   1440
         Width           =   3495
         _Version        =   1572864
         _ExtentX        =   6165
         _ExtentY        =   582
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtArreglo 
         Height          =   330
         Left            =   1440
         TabIndex        =   135
         Top             =   1800
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtArregloDesc 
         Height          =   330
         Left            =   2280
         TabIndex        =   136
         Top             =   1800
         Width           =   3495
         _Version        =   1572864
         _ExtentX        =   6165
         _ExtentY        =   582
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtGestionMonto 
         Height          =   330
         Left            =   1440
         TabIndex        =   137
         Top             =   2160
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "Persona"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   19
         Left            =   240
         TabIndex        =   59
         Top             =   600
         Width           =   855
      End
      Begin VB.Image imgSalir 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   10800
         Picture         =   "frmCO_ControlLista.frx":07D8
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Acuerdo"
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
         Index           =   9
         Left            =   240
         TabIndex        =   58
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Causas"
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
         Index           =   8
         Left            =   240
         TabIndex        =   57
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "a la que se le va a registrar el recargo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   3360
         TabIndex        =   56
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Operación"
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
         Index           =   6
         Left            =   240
         TabIndex        =   55
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Monto"
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
         Left            =   240
         TabIndex        =   54
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Index           =   3
         Left            =   240
         TabIndex        =   53
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Pago"
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
         Index           =   2
         Left            =   240
         TabIndex        =   52
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Gestión"
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
         Left            =   240
         TabIndex        =   51
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Frame fraCargaDatos 
      Appearance      =   0  'Flat
      Caption         =   "Cargar Datos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4095
      Left            =   11520
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   10815
      Begin MSComctlLib.Toolbar tlbCargarAnalisis 
         Height          =   810
         Left            =   7800
         TabIndex        =   7
         Top             =   2040
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   1429
         ButtonWidth     =   1111
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cargar"
               Key             =   "Cargar"
               Object.ToolTipText     =   "Cargar"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin VB.Image imgFraCargarDatos 
         Height          =   255
         Left            =   10440
         Picture         =   "frmCO_ControlLista.frx":08FF
         Stretch         =   -1  'True
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Proceso para cargar información de la cartera de cobros por usuario.  (Cubos)"
         Height          =   495
         Left            =   3000
         TabIndex        =   9
         Top             =   960
         Width           =   6615
      End
      Begin VB.Image Image3 
         Height          =   630
         Left            =   1920
         Picture         =   "frmCO_ControlLista.frx":7151
         Top             =   840
         Width           =   585
      End
      Begin VB.Label lblStatus 
         Caption         =   "Este proceso puede tardar varios minutos, espere el mensaje de proceso concluido."
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   1920
         TabIndex        =   8
         Top             =   1800
         Width           =   4695
      End
   End
   Begin VB.Frame FraResumenCartera 
      Appearance      =   0  'Flat
      Caption         =   "Resumen Cartera por Usuario"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4095
      Left            =   11040
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   10815
      Begin VB.Frame fraContenedorTotales 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   12
         Top             =   3000
         Width           =   10335
         Begin VB.Label lblDescripcionCasos 
            Caption         =   "Saldo"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   25
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblSaldoAlDia 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1560
            TabIndex        =   24
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblDescripcionCasos 
            Caption         =   "Operaciones"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   23
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblOperacionesAlDia 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1560
            TabIndex        =   22
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lblSaldoMora 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3840
            TabIndex        =   20
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblOperacionesMora 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3840
            TabIndex        =   19
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lblSaldoCJud 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6120
            TabIndex        =   17
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblOperacionesCJud 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6120
            TabIndex        =   16
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lblOperacionesCartera 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   8400
            TabIndex        =   15
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lblSaldoCartera 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   8400
            TabIndex        =   14
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblDescripcionCasos 
            Caption         =   "Mora"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   3840
            TabIndex        =   21
            Top             =   0
            Width           =   495
         End
         Begin VB.Label lblDescripcionCasos 
            Caption         =   "Al Día"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   26
            Top             =   0
            Width           =   495
         End
         Begin VB.Label lblDescripcionCasos 
            Caption         =   "Cobro Jud"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   6120
            TabIndex        =   18
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblDescripcionCasos 
            Caption         =   "Cartera"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   8400
            TabIndex        =   13
            Top             =   0
            Width           =   855
         End
      End
      Begin FPSpreadADO.fpSpread vGridCartera 
         Height          =   2415
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   10575
         _Version        =   524288
         _ExtentX        =   18653
         _ExtentY        =   4260
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmCO_ControlLista.frx":7620
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Image ImgCerrarFrameCartera 
         Height          =   255
         Left            =   10440
         Picture         =   "frmCO_ControlLista.frx":825B
         Stretch         =   -1  'True
         Top             =   170
         Width           =   255
      End
   End
   Begin VB.CheckBox chkCasosSinEjecutivo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   85
      Top             =   840
      Width           =   200
   End
   Begin VB.CheckBox chkTodos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   80
      Top             =   480
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.CheckBox chkMarcas 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   360
      MaskColor       =   &H00FF8080&
      TabIndex        =   79
      Top             =   1200
      Width           =   200
   End
   Begin VB.CheckBox chkTodosUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      MaskColor       =   &H00FF8080&
      TabIndex        =   78
      Top             =   120
      Width           =   200
   End
   Begin VB.CheckBox chkFiltros 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      MaskColor       =   &H00FF8080&
      TabIndex        =   77
      Top             =   840
      Width           =   200
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      TabIndex        =   42
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox txtCedula 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4440
      TabIndex        =   41
      Top             =   1200
      Width           =   1335
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   570
      Left            =   6480
      TabIndex        =   3
      Top             =   120
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar Casos"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RSM"
                  Text            =   "Resumen"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "DET"
                  Text            =   "Detalle"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "GAR"
                  Text            =   "Listado x Garantia"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CAR"
                  Text            =   "Listado x Cartera"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cartera"
            Object.ToolTipText     =   "Cartera Usuarios"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CarteraUs"
                  Text            =   "Cartera Usuarios"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Analisis"
                  Text            =   "Análisis Cartera Usuarios (Cubos)"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exportar"
            Object.ToolTipText     =   "Exportar a Excel/Html"
            ImageIndex      =   9
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Excel"
                  Text            =   "Microsoft Excel"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "HTML"
                  Text            =   "HTML"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   120
      Width           =   2415
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3372
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   11292
      _Version        =   524288
      _ExtentX        =   19918
      _ExtentY        =   5948
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
      MaxCols         =   25
      SpreadDesigner  =   "frmCO_ControlLista.frx":EAAD
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10080
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":1011B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":1697D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":1D1DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":23A41
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":2A2A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":30B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":30E1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":315DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":31CC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":325E4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   9240
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":33002
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSemaforos 
      Left            =   10680
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":39864
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":39982
            Key             =   "amarillo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":39AA8
            Key             =   "rojo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":39BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":39CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":39DFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":39EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":3A033
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":3A148
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":3A26C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_ControlLista.frx":3A395
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1200
      TabIndex        =   89
      Top             =   480
      Width           =   1212
      _Version        =   1572864
      _ExtentX        =   2138
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
      Enabled         =   0   'False
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   2400
      TabIndex        =   90
      Top             =   480
      Width           =   1212
      _Version        =   1572864
      _ExtentX        =   2138
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
      Enabled         =   0   'False
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   2892
      Left            =   120
      TabIndex        =   91
      Top             =   5640
      Width           =   11412
      _Version        =   1572864
      _ExtentX        =   20129
      _ExtentY        =   5101
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
      ItemCount       =   5
      Item(0).Caption =   "Operaciones"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vgOperaciones"
      Item(1).Caption =   "Datos Personales"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "Frame1"
      Item(2).Caption =   "Gestiones"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "vgCobro"
      Item(2).Control(1)=   "btnNotifica"
      Item(2).Control(2)=   "rbNotificaEmail(0)"
      Item(2).Control(3)=   "rbNotificaEmail(1)"
      Item(3).Caption =   "Fiadores"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "chkFiadoresEstado"
      Item(3).Control(1)=   "vgridFiadores"
      Item(4).Caption =   "Traslados"
      Item(4).ControlCount=   10
      Item(4).Control(0)=   "txtTrasladar"
      Item(4).Control(1)=   "cbo"
      Item(4).Control(2)=   "cboRebajo"
      Item(4).Control(3)=   "cmdMantiene"
      Item(4).Control(4)=   "cmdTrasladar"
      Item(4).Control(5)=   "Label12(4)"
      Item(4).Control(6)=   "Label12(6)"
      Item(4).Control(7)=   "Label12(5)"
      Item(4).Control(8)=   "Label12(7)"
      Item(4).Control(9)=   "Label12(10)"
      Begin XtremeSuiteControls.PushButton btnNotifica 
         Height          =   732
         Left            =   -70000
         TabIndex        =   123
         Top             =   360
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2984
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Notificación Email: Cuotas Atrasadas"
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
      Begin XtremeSuiteControls.CheckBox chkFiadoresEstado 
         Height          =   372
         Left            =   -69280
         TabIndex        =   120
         Top             =   360
         Visible         =   0   'False
         Width           =   6372
         _Version        =   1572864
         _ExtentX        =   11239
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Solo operaciones atrasadas"
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
         Appearance      =   16
      End
      Begin VB.ComboBox cboRebajo 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         ItemData        =   "frmCO_ControlLista.frx":3A77D
         Left            =   -66520
         List            =   "frmCO_ControlLista.frx":3A787
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cbo 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         ItemData        =   "frmCO_ControlLista.frx":3A793
         Left            =   -66520
         List            =   "frmCO_ControlLista.frx":3A79D
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtTrasladar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -62680
         Locked          =   -1  'True
         TabIndex        =   105
         ToolTipText     =   "Presione F4 Para Consultar"
         Top             =   1080
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   -70000
         TabIndex        =   93
         Top             =   360
         Visible         =   0   'False
         Width           =   11412
         Begin XtremeSuiteControls.ListView lswOtros 
            Height          =   1692
            Left            =   7200
            TabIndex        =   117
            Top             =   360
            Width           =   3612
            _Version        =   1572864
            _ExtentX        =   6371
            _ExtentY        =   2984
            _StockProps     =   77
            BackColor       =   -2147483643
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
         End
         Begin XtremeSuiteControls.PushButton btnContacto 
            Height          =   252
            Left            =   10920
            TabIndex        =   116
            Top             =   360
            Width           =   372
            _Version        =   1572864
            _ExtentX        =   656
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "..."
            Appearance      =   16
         End
         Begin VB.Label Label7 
            Caption         =   "Provincia"
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
            TabIndex        =   102
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Cantón"
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
            TabIndex        =   101
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "Distrito"
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
            TabIndex        =   100
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "Email"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   99
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label lblProvincia 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   960
            TabIndex        =   98
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblCanton 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   960
            TabIndex        =   97
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lblDistrito 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   960
            TabIndex        =   96
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label lblEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   960
            TabIndex        =   95
            Top             =   1680
            Width           =   6015
         End
         Begin VB.Label lblDireccion 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   1035
            Left            =   3000
            TabIndex        =   94
            Top             =   360
            Width           =   4095
         End
      End
      Begin FPSpreadADO.fpSpread vgOperaciones 
         Height          =   2532
         Left            =   0
         TabIndex        =   92
         Top             =   360
         Width           =   10812
         _Version        =   524288
         _ExtentX        =   19071
         _ExtentY        =   4466
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
         SpreadDesigner  =   "frmCO_ControlLista.frx":3A7A9
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vgCobro 
         Height          =   2532
         Left            =   -68080
         TabIndex        =   103
         Top             =   360
         Visible         =   0   'False
         Width           =   10932
         _Version        =   524288
         _ExtentX        =   19283
         _ExtentY        =   4466
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
         SpreadDesigner  =   "frmCO_ControlLista.frx":3B718
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vgridFiadores 
         Height          =   2172
         Left            =   -70000
         TabIndex        =   104
         Top             =   720
         Visible         =   0   'False
         Width           =   10812
         _Version        =   524288
         _ExtentX        =   19071
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
         MaxCols         =   7
         SpreadDesigner  =   "frmCO_ControlLista.frx":3C3CA
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton cmdMantiene 
         Height          =   732
         Left            =   -65560
         TabIndex        =   108
         Top             =   720
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   1291
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
         Transparent     =   -1  'True
         Appearance      =   16
         Picture         =   "frmCO_ControlLista.frx":3CACB
      End
      Begin XtremeSuiteControls.PushButton cmdTrasladar 
         Height          =   732
         Left            =   -60400
         TabIndex        =   109
         Top             =   840
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Trasaladar"
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
         Appearance      =   16
         Picture         =   "frmCO_ControlLista.frx":3D2A3
      End
      Begin XtremeSuiteControls.RadioButton rbNotificaEmail 
         Height          =   372
         Index           =   0
         Left            =   -69880
         TabIndex        =   124
         Top             =   1200
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Resumen"
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
      Begin XtremeSuiteControls.RadioButton rbNotificaEmail 
         Height          =   372
         Index           =   1
         Left            =   -69880
         TabIndex        =   125
         Top             =   1560
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Detalle"
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
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Rebajo Doble"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   10
         Left            =   -67960
         TabIndex        =   114
         Top             =   1080
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mantener"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   7
         Left            =   -67960
         TabIndex        =   113
         Top             =   720
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Trasladar casos Marcados al Usuario siguiente:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1212
         Index           =   5
         Left            =   -64120
         TabIndex        =   112
         Top             =   720
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario a Trasladar "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   6
         Left            =   -62680
         TabIndex        =   111
         Top             =   840
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Actualizar los casos marcados de la lista, el mantener y rebajo doble"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   972
         Index           =   4
         Left            =   -69760
         TabIndex        =   110
         Top             =   720
         Visible         =   0   'False
         Width           =   1812
      End
   End
   Begin XtremeSuiteControls.PushButton btnNotificaLista 
      Height          =   312
      Left            =   9840
      TabIndex        =   126
      Top             =   1200
      Width           =   2892
      _Version        =   1572864
      _ExtentX        =   5101
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Notificar a los casos marcados"
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
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   372
      Left            =   120
      TabIndex        =   115
      Top             =   5160
      Width           =   11412
      _Version        =   1572864
      _ExtentX        =   20129
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.44
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Label lblNameCheck 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Casos Sin Asignar"
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
      Left            =   4680
      TabIndex        =   86
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblNameCheck 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Marcar"
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
      Left            =   720
      TabIndex        =   84
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblNameCheck 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Filtros adicionales"
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
      Left            =   720
      TabIndex        =   83
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblNameCheck 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Todos"
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
      Left            =   4680
      TabIndex        =   82
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblNameCheck 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Todos"
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
      Left            =   4680
      TabIndex        =   81
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Persona"
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
      Index           =   14
      Left            =   3600
      TabIndex        =   43
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas"
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
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   10800
      X2              =   120
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario "
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
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgBanner 
      Height          =   1572
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14892
   End
End
Attribute VB_Name = "frmCO_ControlLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean
Dim mCambioCelda As Boolean
Dim vTipoGestion As String
Dim vDesviacionMax As Double, vDesviacionMin As Double

Dim lngCasos As Long, curMora As Currency, curMoraLegal As Currency


Private Sub btnContacto_Click()
   GLOBALES.gCedulaActual = Trim(scTitulo.Tag)
   frmCR_VerificaDatosPersonales.Show vbModal

End Sub

Private Sub btnNotifica_Click()
   Call sbCbr_Notifica_Email(Trim(scTitulo.Tag), IIf(rbNotificaEmail.Item(0).Value, "R", "D"))
End Sub

Private Sub btnNotificaLista_Click()
Dim i As Long, pCount As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

pCount = 0

With vGrid
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .Value = vbChecked Then
           .Col = 5
           Call sbCbr_Notifica_Email(Trim(.Text), "R")
           pCount = pCount + 1
        End If
    
    Next i
End With

Me.MousePointer = vbDefault

If pCount > 0 Then
    MsgBox "Notificaciones enviadas a : " & Format(pCount, "###,##0") & ", personas!", vbInformation
End If

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox "A ocurrido un error en el Proceso!", vbCritical

End Sub

Private Sub chkAntiguedad_Click()
If chkAntiguedad.Value = vbChecked Then
  lswConAntiguedad.Enabled = False
Else
  lswConAntiguedad.Enabled = True
End If

End Sub

Private Sub chkCasosSinEjecutivo_Click()
If chkCasosSinEjecutivo.Value = vbChecked Then
   txtCodigo.Text = ""
   chkTodosUser.Value = vbUnchecked
   Call chkTodosUser_Click
End If
End Sub

Private Sub chkFechaPago_Click()
  If chkFechaPago.Value = vbUnchecked Then
     dtpFPInicio.Enabled = True
     dtpFPCorte.Enabled = True
  Else
     dtpFPInicio.Enabled = False
     dtpFPCorte.Enabled = False
  End If
End Sub

Private Sub chkFiadoresEstado_Click()
Dim strSQL As String
Me.MousePointer = vbHourglass
On Error GoTo vError

    strSQL = "select M.ESTADO as 'EstadoMora','',F.Id_Solicitud, S.cedula,S.nombre,E.descripcion as Estado,I.descripcion as Inst " _
            & " from fiadores F inner join Socios S on F.cedulaf = S.cedula" _
            & " inner join Instituciones I on S.cod_institucion = I.cod_institucion" _
            & " inner join Reg_Creditos R on F.Id_Solicitud = R.Id_Solicitud" _
            & " inner join AFI_ESTADOS_PERSONA E on E.cod_estado = S.estadoActual" _
            & " left join MOROSIDAD M on F.Id_Solicitud = M.Id_Solicitud and M.Estado = 'A'" _
            & "  where F.estado = 'A' and R.cedula = '" & scTitulo.Tag & "' and R.Estado = 'A'"
            
     If chkFiadoresEstado.Value = vbChecked Then
        strSQL = strSQL & " and M.ESTADO = 'A'"
     End If
     
     strSQL = strSQL & " group by F.Id_Solicitud,S.cedula,M.Estado,S.nombre,E.descripcion,I.descripcion"
           
     
    Call sbCargaGridFiadores(vgridFiadores, strSQL)
    
Me.MousePointer = vbDefault

Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub sbCargaGridFiadores(vGrid As Object, strSQL As String)
Dim i As Integer, rs As New ADODB.Recordset

vGrid.MaxCols = 7
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
For i = 1 To vGrid.MaxCols
 vGrid.Col = i
 vGrid.Text = ""
Next i

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)

With vGrid
  
.MaxRows = 0

Do While Not rs.EOF
  .MaxRows = .MaxRows + 1
  .Row = .MaxRows
  
  For i = 1 To .MaxCols
     .Col = i
     Select Case i
        Case 1 'Status
           If rs!EstadoMora = "A" Then
              .TypePictPicture = imgSemaforos.ListImages.Item(3).Picture
           Else
              .TypePictPicture = imgSemaforos.ListImages.Item(1).Picture
           End If
           
              
        Case 3 'Solicitud
           .Text = CStr(rs!ID_SOLICITUD)
           
        Case 4 'Cedula
           .Text = rs!Cedula
           
        Case 5 'Nombre
           .Text = rs!Nombre
           
        Case 6 ' Estado
           .Text = rs!Estado
           
        Case 7 ' Institución
           .Text = rs!Inst
           
     End Select
  Next i
  
  rs.MoveNext
Loop
rs.Close
End With


End Sub

Private Sub chkFiltros_Click()

fraFiltros.top = chkFiltros.top
fraFiltros.Left = dtpInicio.Left

If chkFiltros.Value = vbChecked Then
   fraFiltros.Visible = True
Else
   fraFiltros.Visible = False
End If

End Sub

Private Sub chkGarantias_Click()
If chkGarantias.Value = vbChecked Then
  lswConGarantias.Enabled = False
Else
  lswConGarantias.Enabled = True
End If

End Sub

Private Sub chkMarcas_Click()
Dim lngCasos As Long

For lngCasos = 1 To vGrid.MaxRows
  vGrid.Row = lngCasos
  vGrid.Col = 1
  vGrid.Value = chkMarcas.Value
Next lngCasos

End Sub

Private Sub chkTodos_Click()

If chkTodos.Value = vbChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub

Private Sub chkTodosUser_Click()
If chkTodosUser.Value = vbChecked Then
  txtCodigo.Enabled = False
Else
  txtCodigo.Enabled = True
End If
End Sub



Private Sub cmdMantiene_Click()
Dim strSQL  As String, i As Byte, y As Byte, x As Long

Me.MousePointer = vbHourglass

i = IIf((cbo.Text = "SI"), 1, 0)
y = IIf((cboRebajo.Text = "SI"), 1, 0)

For x = 1 To vGrid.MaxRows
 vGrid.Row = x
 
 vGrid.Col = 1
 
 
 If vGrid.Value = vbChecked Then
    vGrid.Col = 11
    vGrid.Value = i
    vGrid.Col = 12
    vGrid.Value = y
    
    vGrid.Col = 5
    strSQL = "update cbr_asignacion set mantener = " & i & ",rebajo_doble = " & y _
           & " where usuario = '" & txtCodigo & "' and cedula = '" & vGrid.Text & "'"
    Call ConectionExecute(strSQL)
    
 End If
Next x

Me.MousePointer = vbDefault

MsgBox "Estatus de Mantener Actualizado Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cmdTrasladar_Click()
Dim strSQL  As String, x As Long

If txtTrasladar = "" Then Exit Sub

Me.MousePointer = vbHourglass

For x = 1 To vGrid.MaxRows
 vGrid.Row = x
 vGrid.Col = 1
 If vGrid.Value = vbChecked Then
    
    vGrid.Col = 5
    strSQL = "exec spCBRControlAsg '" & vGrid.Text _
           & "','" & txtTrasladar.Text & "',1"
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Aplica", "Traslado Caso CBR Ced:" & vGrid.Text & " de " & txtCodigo _
             & " a " & txtTrasladar)
 End If
Next x

Me.MousePointer = vbDefault

txtTrasladar = ""

MsgBox "Traslado de Expedientes realizado satisfactoriamente...", vbInformation
Call sbBuscar

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub dtpVence_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboOperacion.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

   
If vScroll Then
    strSQL = "select Top 1 usuario from cbr_usuarios"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where estado = 1 and usuario > '" & txtCodigo & "' order by usuario asc"
    Else
       strSQL = strSQL & " where estado = 1 and usuario < '" & txtCodigo & "' order by usuario desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!Usuario
      Call sbBuscar
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
  vModulo = 4
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

tcMain.Item(0).Selected = True

mCambioCelda = False

Me.Width = 14000
Me.Height = 8700

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

dtpFPInicio.Value = dtpInicio.Value
dtpFPCorte.Value = dtpInicio.Value

cbo.Text = "SI"
cboRebajo.Text = "NO"

With lswLista.ColumnHeaders
    .Clear
    .Add , , "ID", 640
    .Add , , "Detalle", 3850
End With

cboOrden.Clear
cboOrden.AddItem "01 - Sin Orden"
cboOrden.AddItem "02 - Cédula"
cboOrden.AddItem "03 - Nombre"
cboOrden.AddItem "04 - Fecha"
cboOrden.AddItem "05 - Fecha + Cedula"
cboOrden.AddItem "06 - Fecha + Nombre"
cboOrden.AddItem "07 - Peso Mora Cuotas"
cboOrden.AddItem "08 - Peso Mora Monto"
cboOrden.Text = "08 - Peso Mora Monto"


cboOrdenTipo.Clear
cboOrdenTipo.AddItem "Asc"
cboOrdenTipo.AddItem "Desc"
cboOrdenTipo.Text = "Desc"


cboEstado.Clear
cboEstado.AddItem "01 - Todos"
cboEstado.AddItem "02 - Morosos"
cboEstado.AddItem "03 - Al Dia"
cboEstado.Text = "02 - Morosos"


cboTipoCasos.Clear
cboTipoCasos.AddItem "TODOS"
cboTipoCasos.AddItem "Casos Sin Atención"
cboTipoCasos.AddItem "Casos Atendidos"
cboTipoCasos.Text = "TODOS"



strSQL = "select garantia as 'Idx',descripcion from crd_garantia_tipos order by descripcion"
rs.Open strSQL, glogon.Conection
Do While Not rs.EOF
  Set itmX = lswConGarantias.ListItems.Add(, , rs!Descripcion)
      itmX.Tag = rs!IdX
  rs.MoveNext
Loop
rs.Close


strSQL = "select cod_Antiguedad as 'Idx',descripcion from CBR_ANTIGUEDAD_TIPOS order by cod_Antiguedad"
rs.Open strSQL, glogon.Conection
Do While Not rs.EOF
  Set itmX = lswConAntiguedad.ListItems.Add(, , rs!Descripcion)
      itmX.Tag = rs!IdX
  rs.MoveNext
Loop
rs.Close

Call chkAntiguedad_Click
Call chkGarantias_Click

strSQL = "select rtrim(cod_clasificacion) + ' - ' + rtrim(descripcion) as Itmx" _
       & " From CBR_CLASIFICACION_CARTERA Where Estado = 1"
Call sbLlenaCbo(cboCarteras, strSQL, True, False)

strSQL = "select rtrim(cod_oficina) + ' - ' + rtrim(descripcion) as Itmx" _
       & " From SIF_OFICINAS Where Estado = 1"
Call sbLlenaCbo(cboOficina, strSQL, True, False)

strSQL = "select rtrim(descripcion) as Itmx,cod_Institucion as Idx" _
       & " From INSTITUCIONES Where Activa = 1"
Call sbLlenaCbo(cboInstitucion, strSQL, True, True)


vGrid.MaxRows = 0
vgOperaciones.MaxRows = 0
vgridFiadores.MaxRows = 0
vgridFiadores.MaxCols = 7

vgridFiadores.Visible = True
Label12(4).Visible = True

fraGestion.Visible = False
fraGestion.Left = vGrid.Left + tcMain.Width + 300
scTitulo.top = vGrid.top + vGrid.Height + 250

txtGestionFiltro.Text = ""
txtGestionDescFiltro.Text = ""
txtCausaFiltro.Text = ""
txtCausaDescFiltro.Text = ""
txtArregloFiltro.Text = ""
txtArregloDescFiltro.Text = ""

Call Formularios(Me)
Call RefrescaTags(Me)


If vGrid.MaxRows = 0 Then
  tcMain.Enabled = False
End If

End Sub

Private Sub Form_Resize()
On Error GoTo vError
    
'    If Me.Width > 10000 Then
       imgBanner.Width = Me.Width
       vGrid.Width = Me.Width - 500
       vGrid.Height = Me.Height - 6000
'        If chkDetalleCuotas.Value = vbUnchecked Then
'            vGrid.Height = Me.Height - 6000
'        Else
'            vGrid.Height = Me.Height - 7000
'        End If
    
        scTitulo.top = vGrid.top + vGrid.Height + 150
        scTitulo.Width = vGrid.Width
        tcMain.top = scTitulo.top + scTitulo.Height + 100
        tcMain.Height = Me.Height - (vGrid.Height + 3000)
        tcMain.Left = vGrid.Left
        tcMain.Width = vGrid.Width
               
                      
        'Ajuste de los Grid
        vgOperaciones.Width = tcMain.Width - 400
                
        vgridFiadores.Width = tcMain.Width - 300
                
        vgCobro.Width = tcMain.Width - 400
        
        If fraCargaDatos.Visible Then
            Call sbMostrarFrameCarga
        End If
        
        If FraResumenCartera.Visible Then
            Call sbMostrarCartera
        End If
        

        
        
        
        fraGestion.Left = vGrid.Left + tcMain.Width + 300
        If fraGestion.Visible Then
           fraGestion.Visible = False
           fraGestion.Visible = True
           
        End If
'    End If
    
    Exit Sub

vError:
End Sub

Private Sub sbMostrarFrameCarga()

    fraCargaDatos.top = vGrid.top
    fraCargaDatos.Left = vGrid.Left
    fraCargaDatos.Width = vGrid.Width
    fraCargaDatos.Height = vGrid.Height
    imgFraCargarDatos.Left = fraCargaDatos.Width - 500
    fraCargaDatos.Visible = True
    
End Sub

Private Sub ImgCerrarFrameCartera_Click()
    FraResumenCartera.Visible = False
End Sub

Private Sub imgFraCargarDatos_Click()
    fraCargaDatos.Visible = False
End Sub

Private Sub imgSalir_Click()
   fraGestion.Visible = False
   fraGestion.Left = vGrid.Height + 5000
   Call sblimpiarFrameGestiones
   txtCedula.Enabled = True
   txtNombre.Enabled = True
End Sub

Private Sub lblNameCheck_Click(Index As Integer)
Select Case Index
  Case 0 'Todos los Usuario
     chkTodosUser.Value = IIf((chkTodosUser.Value = vbChecked), vbUnchecked, vbChecked)
     Call chkTodosUser_Click
  
  Case 1 'Todas las fecha
     chkTodos.Value = IIf((chkTodos.Value = vbChecked), vbUnchecked, vbChecked)
     Call chkTodos_Click
     
  Case 2 'Casos Sin Asignar
     chkCasosSinEjecutivo.Value = IIf((chkCasosSinEjecutivo.Value = vbChecked), vbUnchecked, vbChecked)
     Call sbBuscar
     
  Case 3 'Filtros adicionales
     chkFiltros.Value = IIf((chkFiltros.Value = vbChecked), vbUnchecked, vbChecked)
     Call chkFiltros_Click
     
  Case 4 'Marcar
     chkMarcas.Value = IIf((chkMarcas.Value = vbChecked), vbUnchecked, vbChecked)
     Call chkMarcas_Click
     

End Select
End Sub

'Devuelve el codigo de la causa de mora o la gestion seleccionada
Private Sub lswLista_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
  
  Select Case Mid(vTipoGestion, 1, 1)
    Case "G"
        txtGestion.Text = Item.Text
        txtGestionDesc.Text = Item.SubItems(1)
        
        If Right(vTipoGestion, 1) = "C" Then
           txtGestion.SetFocus
        Else
           txtGestionDesc.SetFocus
        End If
       
    Case "C"
        txtCausa.Text = Item.Text
        txtCausaDesc.Text = Item.SubItems(1)
        
        If Right(vTipoGestion, 1) = "C" Then
           txtCausa.SetFocus
        Else
           txtCausaDesc.SetFocus
        End If
        
    Case "A"
        txtArreglo.Text = Item.Text
        txtArregloDesc.Text = Item.SubItems(1)
        
        If Right(vTipoGestion, 1) = "C" Then
           txtArreglo.SetFocus
        Else
           txtArregloDesc.SetFocus
        End If
        
  End Select
  

  
End Sub


Private Sub sbBuscar()
Dim strSQL As String, iCantidad As Integer, vCadena As String, i As Integer
Dim vFiltroAddGarantias As String, vFiltroAddAntiguedad As String

Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = ""
vCadena = ""
vFiltroAddAntiguedad = ""
vFiltroAddGarantias = ""


lngCasos = 0
curMora = 0
curMoraLegal = 0

chkMarcas.Value = vbUnchecked

strSQL = "select " & chkMarcas.Value & ",Lt.fecha_asignacion,Lt.cedula,Lt.nombre,Lt.cuotaMora,Lt.Mora,Lt.MoraLegal,Lt.Operaciones,Lt.mantener" _
       & ",Lt.rebajo_doble,Lt.DiasUltAtencion,Lt.ULT_GESTION_FECHA,Lt.ULT_GESTION_USUARIO, Lt.GestionDesc,Lt.CausaDesc,Lt.ArregloDesc, Lt.ARREGLO_VENCE" _
       & ",Lt.MoraDias, Isnull(At.Descripcion,'') as 'Antiguedad'"


If chkInfoContracto.Value = vbChecked Then
   strSQL = strSQL & ", dbo.fxAFITelefono(Lt.Cedula,3) as 'Tel_Cel',  dbo.fxAFITelefono(Lt.Cedula,1) as 'Tel_Hab'" _
          & ", dbo.fxAFITelefono(Cedula,2) as 'Tel_Tra', Lt.Email"
Else
   strSQL = strSQL & ", '' as 'Tel_Cel',  '' as 'Tel_Hab'" _
          & ", '' as 'Tel_Tra', '' as 'Email'"
End If

If chkCasosSinEjecutivo.Value = vbChecked Then
    strSQL = strSQL & " From dbo.vCbrControlListadoSinEjecutivo Lt left join CBR_ANTIGUEDAD_TIPOS At on Lt.MoraDias between At.Dias_Desde and At.Dias_Hasta"
Else
    strSQL = strSQL & " From dbo.vCBRControlListado Lt left join CBR_ANTIGUEDAD_TIPOS At on Lt.MoraDias between At.Dias_Desde and At.Dias_Hasta"
End If

If chkTodosUser.Value = vbChecked Then
  strSQL = strSQL & " where Lt.Usuario <> ''"
Else
  strSQL = strSQL & " where Lt.Usuario = '" & txtCodigo.Text & "'"
End If

If txtCedula.Text <> "" Then
  strSQL = strSQL & " and Lt.Cedula like '%" & txtCedula.Text & "%'"
End If

If txtNombre.Text <> "" Then
  strSQL = strSQL & " and Lt.Nombre like '%" & txtNombre.Text & "%'"
End If


Select Case Mid(cboEstado.Text, 1, 2)
  Case "01" 'Todos
  Case "02" 'Morosos
    strSQL = strSQL & " and Lt.Mora > 0"
  Case "03" ' Al dia
    strSQL = strSQL & " and Lt.Mora = 0"
End Select

'Numero de Cuotas
strSQL = strSQL & " and Lt.CuotaMora between " & txtCtaDesde.Text & " and " & txtCtaHasta.Text


If chkTodos.Value = vbUnchecked Then
  strSQL = strSQL & " and Lt.fecha_asignacion between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
         & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End If

If cboCarteras.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
         strSQL = strSQL & "Lt.cedula in(select cedula from vCBRControlListadoCartera where cod_Clasificacion = '" _
                & SIFGlobal.fxCodText(cboCarteras.Text) & "')"
End If


If cboOficina.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
         strSQL = strSQL & " dbo.fxCrdPersonaOficinaExiste(Lt.cedula,'" & SIFGlobal.fxCodText(cboOficina.Text) & "') > 0"
End If

If cboInstitucion.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
         strSQL = strSQL & " Lt.Cod_Institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If


'----------------------Filtros Especiales con Vista de Morosidad -----------------------
'Lista de Garantias
iCantidad = 0
If chkGarantias.Value = vbUnchecked Then
    vFiltroAddGarantias = " Garantia in('"
    For i = 1 To lswConGarantias.ListItems.Count
      If lswConGarantias.ListItems.Item(i).Checked Then
        vFiltroAddGarantias = vFiltroAddGarantias & "','" & lswConGarantias.ListItems.Item(i).Tag
        iCantidad = iCantidad + 1
      End If
    Next i
    
    If iCantidad > 0 Then
        vFiltroAddGarantias = vFiltroAddGarantias & "')"
    End If
End If



'Lista de Antiguedades
iCantidad = 0
If chkAntiguedad.Value = vbUnchecked Then
    vFiltroAddAntiguedad = " Cod_Antiguedad in('"
    For i = 1 To lswConAntiguedad.ListItems.Count
      If lswConAntiguedad.ListItems.Item(i).Checked Then
        vFiltroAddAntiguedad = vFiltroAddAntiguedad & "','" & lswConAntiguedad.ListItems.Item(i).Tag
        iCantidad = iCantidad + 1
      End If
    Next i
    
    If iCantidad > 0 Then
        vFiltroAddAntiguedad = vFiltroAddAntiguedad & "')"
    End If
End If

If Len(vFiltroAddAntiguedad) > 0 Or Len(vFiltroAddGarantias) > 0 Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  
  strSQL = strSQL & "cedula in( select cedula from vista_Morosidad where "
         
  If Len(vFiltroAddAntiguedad) > 0 Then
     strSQL = strSQL & vFiltroAddAntiguedad
  End If
  
  If Len(vFiltroAddGarantias) > 0 Then
     If Len(vFiltroAddAntiguedad) > 0 Then strSQL = strSQL & " and "
     strSQL = strSQL & vFiltroAddGarantias
  End If
  
  strSQL = strSQL & " group by cedula )"

End If
'----------------------FIN.: Filtros Especiales con Vista de Morosidad -----------------------


Select Case cboTipoCasos.Text
  Case "TODOS"
  Case "Casos Sin Atención"
        strSQL = strSQL & " and Lt.DiasUltAtencion >= " & txtDiasAtencion.Text
  Case "Casos Atendidos"
        strSQL = strSQL & " and Lt.DiasUltAtencion <= " & txtDiasAtencion.Text
End Select

If txtGestionFiltro <> "" Then
  strSQL = strSQL & " and Lt.ULT_COD_GESTION = '" & txtGestionFiltro.Text & "'"
End If

If txtCausaFiltro <> "" Then
  strSQL = strSQL & " and Lt.COD_CAUSA = '" & txtCausaFiltro.Text & "'"
End If

If txtArregloFiltro <> "" Then
  strSQL = strSQL & " and Lt.COD_ARREGLO = '" & txtArregloFiltro.Text & "'"
End If

If chkFechaPago.Value = vbUnchecked Then
   strSQL = strSQL & " and Lt.ARREGLO_VENCE between '" & Format(dtpFPInicio.Value, "yyyy/mm/dd") _
         & " 00:00:00' and '" & Format(dtpFPCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End If

Select Case Mid(cboOrden.Text, 1, 2)
  Case "01" 'Sin Orden
  Case "02" 'Cedula
    strSQL = strSQL & " order by Lt.cedula " & cboOrdenTipo.Text
  Case "03" 'Nombre
    strSQL = strSQL & " order by Lt.nombre " & cboOrdenTipo.Text
  Case "04" 'Fecha
    strSQL = strSQL & " order by Lt.fecha_asignacion " & cboOrdenTipo.Text
  Case "05" 'Fecha,cedula
    strSQL = strSQL & " order by Lt.fecha_asignacion,cedula " & cboOrdenTipo.Text
  Case "06" 'Fecha,Nombre
    strSQL = strSQL & " order by Lt.fecha_asignacion,nombre " & cboOrdenTipo.Text
  Case "07" 'Peso en Mora Cuotas
    strSQL = strSQL & " order by Lt.cuotaMora " & cboOrdenTipo.Text
  Case "08" 'Peso en Mora Monto
    strSQL = strSQL & " order by Lt.Mora " & cboOrdenTipo.Text
End Select

vGrid.Sheet = 1

Call sbCargaGridLocal(vGrid, 25, strSQL)
'Elimina Ultima Linea
vGrid.MaxRows = vGrid.MaxRows - 1

For lngCasos = 1 To vGrid.MaxRows
  vGrid.Row = lngCasos
  vGrid.Col = 8
  curMora = curMora + CCur(vGrid.Text)
  vGrid.Col = 9
  curMoraLegal = curMoraLegal + CCur(vGrid.Text)

  vGrid.Col = 1
  vGrid.Value = chkMarcas.Value
Next lngCasos


scTitulo.Caption = scTitulo.Caption & "     [Casos: " & Format(vGrid.MaxRows, "###,###,###,###") _
                & ", Mora Financiera: " & Format(curMora, "Standard") _
                & ", Mora Legal: " & Format(curMoraLegal, "Standard") & "]"

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
  Dim strSQL As String, rs As New ADODB.Recordset
  Dim itmX As ListViewItem
  
  On Error GoTo vError
    
  Select Case Item.Index
     Case 0 'Operaciones
       Call sbCargarDetalleCuotas
       
     Case 1 'Carga la direccion y el correo
        Me.MousePointer = vbHourglass
        
        strSQL = "Select rtrim(Prov.Descripcion) as ProvDesc,rtrim(Cant.Descripcion) as CantonDesc" _
               & ", rtrim(Dist.Descripcion) as DistDesc,s.direccion,s.af_email" _
               & " From socios S" _
               & " left join Provincias Prov on S.Provincia = Prov.Provincia" _
               & " left join Cantones Cant on S.Provincia = Cant.Provincia and S.Canton = Cant.Canton" _
               & " left join Distritos Dist on S.Provincia = Dist.Provincia and S.Canton = Dist.Canton and S.distrito = Dist.distrito" _
               & " where cedula='" & scTitulo.Tag & "'"
               
        Call OpenRecordSet(rs, strSQL)
     
        lblProvincia.Caption = IIf(IsNull(rs!ProvDesc), "", rs!ProvDesc)
        lblCanton.Caption = IIf(IsNull(rs!CantonDesc), "", rs!CantonDesc)
        lblDistrito.Caption = IIf(IsNull(rs!DistDesc), "", rs!DistDesc)
      
        
        lblDireccion.Caption = IIf(IsNull(rs!direccion), "", rs!direccion)
        lblEmail.Caption = IIf(IsNull(rs!AF_Email), "", rs!AF_Email)
        
        rs.Close
        
        lswOtros.ListItems.Clear
        lswOtros.ColumnHeaders.Clear
        
        'Carga los telefonos
        lswOtros.ColumnHeaders.Add 1, , "Numero", 1500
        lswOtros.ColumnHeaders.Add 2, , "Tipo", 1500
        lswOtros.ColumnHeaders.Add 3, , "Extension", 1500
        lswOtros.ColumnHeaders.Add 4, , "Contacto", 2500
        
        
        strSQL = "Select * From Telefonos where " _
               & "Cedula='" & Trim(scTitulo.Tag) & "'"
        Call OpenRecordSet(rs, strSQL)
        
        Do While Not rs.EOF
           Set itmX = lswOtros.ListItems.Add(, , Trim(rs!Numero))
               itmX.SubItems(1) = fxTipoTelefono(rs!Tipo)
               itmX.SubItems(2) = Trim(rs!Ext) & ""
               itmX.SubItems(3) = Trim(rs!contacto) & ""
           rs.MoveNext
        Loop
        rs.Close
  
        Me.MousePointer = vbDefault
      
      Case 2 'Trae las gestiones realizadas
       
        Call vgCobro_SheetChanged(vgCobro.ActiveSheet, vgCobro.ActiveSheet)
  
      
      Case 3 'Carga el listado de los fiadores
      
        Call chkFiadoresEstado_Click
  
  End Select

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 


End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Call sbBuscar

End Sub

Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim strSQL As String, vSubTitulo As String
Dim i As Integer, iCantidad As Integer, vCadena As String

On Error GoTo vError

Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 25
    vHeaders.Headers(1) = "Caso Seleccionado"
    vHeaders.Headers(2) = " ... "
    vHeaders.Headers(3) = " ... "
    vHeaders.Headers(4) = "Fec.Asignación"
    vHeaders.Headers(5) = "Cedula"
    vHeaders.Headers(6) = "Nombre"
    vHeaders.Headers(7) = "Mora Cuotas"
    vHeaders.Headers(8) = "Mora Financiera"
    vHeaders.Headers(9) = "Mora Legal"
    vHeaders.Headers(10) = "No.Operaciones"
    vHeaders.Headers(11) = "Mantiene Asig."
    vHeaders.Headers(12) = "Rebajo Doble?"
    vHeaders.Headers(13) = "Dias Ult. Atención"
    vHeaders.Headers(14) = "Ult.Gestión Fecha"
    vHeaders.Headers(15) = "Ult.Gestión Usuario"
    vHeaders.Headers(16) = "Ult.Gestión"
    vHeaders.Headers(17) = "Ult.Causa"
    vHeaders.Headers(18) = "Ult.Arreglo"
    vHeaders.Headers(19) = "Fec.Arreglo Pago"
    
    vHeaders.Headers(20) = "Dias Mora"
    vHeaders.Headers(21) = "Antiguedad"
    vHeaders.Headers(22) = "Tel.Celular"
    vHeaders.Headers(23) = "Tel.Habitación"
    vHeaders.Headers(24) = "Tel.Trabajo"
    vHeaders.Headers(25) = "Email"

    Me.MousePointer = vbHourglass

    fraCargaDatos.Visible = False
    FraResumenCartera.Visible = False
    
    Select Case UCase(ButtonMenu.Key)
      Case "CARTERAUS"
        FraResumenCartera.Visible = False
        Call sbConsultaCartera
        Call sbMostrarCartera
      Case "ANALISIS"
        Call sbMostrarFrameCarga
        
      'Exportar
      Case "EXCEL"
        Call sbSIFGridExportar(vGrid, vHeaders, "Cobros_ListadoCasos_" & Format(fxFechaServidor, "yyyy.mm.dd_hh.mm"))
      Case "HTML"
        Call sbSIFGridExportar(vGrid, vHeaders, "Cobros_ListadoCasos_" & Format(fxFechaServidor, "yyyy.mm.dd_hh.mm"), "HTML")
      
      
      'Reportes
      Case "RSM", "DET", "GAR", "CAR"
            With frmContenedor.Crt
                .Reset
                .WindowShowGroupTree = True
                .WindowShowPrintSetupBtn = True
                .WindowShowRefreshBtn = True
                .WindowShowSearchBtn = True
                .WindowState = crptMaximized
                .WindowTitle = "Reportes del Módulo de Cobro"
                
                .Connect = glogon.ConectRPT
                
                .Formulas(0) = "Empresa = '" & GLOBALES.gstrNombreEmpresa & "'"
                .Formulas(1) = "Fecha = '" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
                .Formulas(2) = "Titulo = 'LISTADO DE CONTROL DE ASIGNACION DE CASOS'"
                
                strSQL = ""
                
                If chkTodosUser.Value = vbChecked Then
                       vSubTitulo = "US: TODOS"
                Else
                    If Trim(txtCodigo.Text) <> "" Then
                       strSQL = "{vCBRControlListado.Usuario} = '" & txtCodigo.Text & "'"
                       vSubTitulo = "US: " & UCase(txtCodigo.Text)
                    Else
                       vSubTitulo = "US: TODOS"
                    End If
                End If
                
                
                Select Case Mid(cboEstado.Text, 1, 2)
                  Case "01" 'Todos
                    vSubTitulo = vSubTitulo & Space(3) & "TIPO: TODOS"
                  Case "02" 'Morosos
                    vSubTitulo = vSubTitulo & Space(3) & "TIPO: MOROSOS"
                    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
                    strSQL = strSQL & "{vCBRControlListado.Mora} > 0"
                  Case "03" ' Al dia
                    vSubTitulo = vSubTitulo & Space(3) & "TIPO: AL DIA"
                    strSQL = strSQL & "{vCBRControlListado.Mora} = 0"
                End Select
            
            
                 vSubTitulo = vSubTitulo & Space(3) & "NUM.CTA.:" & txtCtaDesde.Text & "-" & txtCtaHasta.Text
                    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
                    strSQL = strSQL & "{vCBRControlListado.CuotaMora} >= " & txtCtaDesde.Text & " AND {vCBRControlListado.CuotaMora} <= " & txtCtaHasta.Text
                
            
                If chkTodos.Value = vbUnchecked Then
                    vSubTitulo = vSubTitulo & Space(3) & "FECHA: " & Format(dtpInicio.Value, "dd/mm/yyyy") & " al " & Format(dtpCorte.Value, "dd/mm/yyyy")
                    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
                     strSQL = strSQL & "cdate({vCBRControlListado.fecha_asignacion}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                         & ") to date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
                Else
                    vSubTitulo = vSubTitulo & Space(3) & "FECHA: TODAS"
                End If
            
            
                vSubTitulo = vSubTitulo & Space(3) & "ORDEN: " & UCase(Mid(cboOrden.Text, 5, 30)) & " " & cboOrdenTipo.Text
                
                vSubTitulo = vSubTitulo & Space(3) & "GARANTIAS Todas.:" & chkGarantias.Value
            
                .Formulas(3) = "SubTitulo = '" & vSubTitulo & "'"
                    
'                'Lista de Antiguedades
'                iCantidad = 0
'                If chkAntiguedad.Value = vbUnchecked Then
'                    vCadena = " and At.Cod_Antiguedad in('"
'                    For i = 1 To lswConAntiguedad.ListItems.Count
'                      If lswConAntiguedad.ListItems.Item(i).Checked Then
'                        vCadena = vCadena & "','" & lswConAntiguedad.ListItems.Item(i).Tag
'                        iCantidad = iCantidad + 1
'                      End If
'                    Next i
'
'                    If i > 0 Then
'                        strSQL = strSQL & vCadena & "')"
'                    End If
'                End If
                
                'Lista de Garantias
                iCantidad = 0
                If chkGarantias.Value = vbUnchecked Then
                    vCadena = " AND ( {vCBRControlListadoGarantia.Garantia} in ['"
                    For i = 1 To lswConGarantias.ListItems.Count
                      If lswConGarantias.ListItems.Item(i).Checked Then
                        vCadena = vCadena & "','" & lswConGarantias.ListItems.Item(i).Tag
                        iCantidad = iCantidad + 1
                      End If
                    Next i
                
                    If i > 0 Then
                        strSQL = strSQL & vCadena & "'])"
                    End If
                End If
                                
                                
                If chkGarantias.Value = vbUnchecked Then
                   vCadena = "[CBRControlListadoGarantiaDetalle"
                Else
                   vCadena = "vCBRControlListado"
                End If
                
                Select Case Mid(cboOrden.Text, 1, 2)
                  Case "01" 'Sin Orden
                  Case "02" 'Cedula
                     .SortFields(0) = IIf((cboOrdenTipo.Text = "Asc"), "+", "-") & "{" & vCadena & ".cedula}"
                  Case "03" 'Nombre
                     .SortFields(0) = IIf((cboOrdenTipo.Text = "Asc"), "+", "-") & "{" & vCadena & ".nombre}"
                  Case "04" 'Fecha
                     .SortFields(0) = IIf((cboOrdenTipo.Text = "Asc"), "+", "-") & "{" & vCadena & ".fecha_asignacion}"
                  Case "05" 'Fecha,cedula
                     .SortFields(0) = IIf((cboOrdenTipo.Text = "Asc"), "+", "-") & "{" & vCadena & ".fecha_asignacion}"
                     .SortFields(1) = IIf((cboOrdenTipo.Text = "Asc"), "+", "-") & "{" & vCadena & ".cedula}"
                  Case "06" 'Fecha,Nombre
                     .SortFields(0) = IIf((cboOrdenTipo.Text = "Asc"), "+", "-") & "{" & vCadena & ".fecha_asignacion}"
                     .SortFields(1) = IIf((cboOrdenTipo.Text = "Asc"), "+", "-") & "{" & vCadena & ".nombre}"
                  Case "07" 'Peso en Mora Cuotas
                     .SortFields(0) = IIf((cboOrdenTipo.Text = "Asc"), "+", "-") & "{" & vCadena & ".cuotaMora}"
                  Case "08" 'Peso en Mora Monto
                     .SortFields(0) = IIf((cboOrdenTipo.Text = "Asc"), "+", "-") & "{" & vCadena & ".Mora}"
                End Select
                                
                
                Select Case ButtonMenu.Key
                  Case "RSM"
                        If chkGarantias.Value = vbUnchecked Then
                            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ControlListadoGarantia.rpt")
                        Else
                            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ControlListadoRsm.rpt")
                        End If
                        
                  Case "DET"
                        If chkGarantias.Value = vbUnchecked Then
                            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ControlListadoGarantia.rpt")
                        Else
                            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ControlListado.rpt")
                        End If
                  Case "GAR"
                            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ControlListadoGarantia.rpt")
                  Case "CAR"
                
                
                End Select
                    
                
                .SelectionFormula = strSQL
                
                .PrintReport
                
                Me.MousePointer = vbDefault
            
            End With
        End Select

    Me.MousePointer = vbDefault

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub

Private Sub tlbCargarAnalisis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
      Case "CARGAR"
        Call sbProcesa
 End Select
End Sub

Private Sub sbProcesa()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String
On Error GoTo vError

    Me.MousePointer = vbHourglass

    strSQL = "exec spCbrGestionAnalisis"
    vMensaje = "CobroCarteraUsuarios"
    Call ConectionExecute(strSQL)
    
    lblStatus.Caption = "Proceso Concluido con éxito, la información puede ser utilizada desde la base de datos de análisis: " & vMensaje
    Me.MousePointer = vbDefault

    Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbMostrarCartera()
    Dim strSQL As String
    FraResumenCartera.top = vGrid.top
    FraResumenCartera.Left = vGrid.Left
    FraResumenCartera.Width = vGrid.Width
    FraResumenCartera.Height = vGrid.Height
    vGridCartera.Width = FraResumenCartera.Width - 400
    ImgCerrarFrameCartera.Left = FraResumenCartera.Width - 300
    
    vGridCartera.Height = (FraResumenCartera.Height) - 1600
    vGridCartera.top = 300
    
    fraContenedorTotales.top = vGridCartera.Height + 400
    
    FraResumenCartera.Visible = True
    
End Sub

Private Sub sbConsultaCartera()
    Dim strSQL As String

On Error GoTo vError

     Me.MousePointer = vbHourglass
   
    strSQL = "exec spCbrListaMoraGarantia '" & Trim(txtCodigo.Text) & "'"
    
    vGridCartera.Sheet = 2
    
    Call sbCargaGridCartera(vGridCartera, 11, strSQL)
    vGridCartera.MaxRows = vGridCartera.MaxRows - 1

    
    strSQL = "exec spCbrPersonaAlDiaGarantia '" & Trim(txtCodigo.Text) & "'"
    
    vGridCartera.Sheet = 1
    
    Call sbCargaGridCartera(vGridCartera, 5, strSQL)
    vGridCartera.MaxRows = vGridCartera.MaxRows - 1
    
    If IsNumeric(lblSaldoAlDia.Caption) Then
        lblSaldoCartera = Format(CDbl(lblSaldoAlDia) + CDbl(lblSaldoCJud) + CDbl(lblSaldoMora), "Standard")
        lblOperacionesCartera = CDbl(lblOperacionesAlDia) + CDbl(lblOperacionesCJud) + CDbl(lblOperacionesMora)
    End If
    Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

    
End Sub

Private Sub txtArregloDesc_GotFocus()

If vTipoGestion <> "AD" Then
 vTipoGestion = "AD"
 Call sbCargaLista(vTipoGestion)
End If
  
End Sub

Private Sub txtArregloDescFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtGestionFiltro.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select COD_ARREGLO,DESCRIPCION from CBR_TIPOS_ARREGLOS"
    gBusquedas.Columna = "COD_ARREGLO"
    gBusquedas.Orden = "COD_ARREGLO"
    gBusquedas.Filtro = " and ACTIVO = 1 "
    frmBusquedas.Show vbModal
    txtArregloFiltro = Trim(gBusquedas.Resultado)
    txtArregloDescFiltro = Trim(gBusquedas.Resultado2)
End If
End Sub


Private Sub txtArregloDescFiltro_LostFocus()
  If vTipoGestion <> "AD" Then
     vTipoGestion = "AD"
     Call sbConsultaGestionFlt(vTipoGestion)
  End If
End Sub

Private Sub txtArregloFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtArregloDescFiltro.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select COD_ARREGLO,DESCRIPCION from CBR_TIPOS_ARREGLOS"
    gBusquedas.Columna = "COD_ARREGLO"
    gBusquedas.Orden = "COD_ARREGLO"
    gBusquedas.Filtro = " and ACTIVO = 1 "
    frmBusquedas.Show vbModal
    txtArregloFiltro = Trim(gBusquedas.Resultado)
    txtArregloDescFiltro = Trim(gBusquedas.Resultado2)
End If
End Sub


Private Sub txtArregloFiltro_LostFocus()
  If vTipoGestion <> "AC" Then
     vTipoGestion = "AC"
     Call sbConsultaGestionFlt(vTipoGestion)
  End If
End Sub

Private Sub txtCausaDesc_GotFocus()
If vTipoGestion <> "CD" Then
 vTipoGestion = "CD"
 Call sbCargaLista(vTipoGestion)
End If
    
End Sub

Private Sub txtCausaDescFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtArregloFiltro.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select COD_CAUSA,DESCRIPCION from CBR_CAUSAS_MOROSIDAD"
    gBusquedas.Columna = "COD_CAUSA"
    gBusquedas.Orden = "COD_CAUSA"
    gBusquedas.Filtro = " and ACTIVA = 1  "
    frmBusquedas.Show vbModal
    txtCausaFiltro = Trim(gBusquedas.Resultado)
    txtCausaDescFiltro = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtCausaDescFiltro_LostFocus()
  If vTipoGestion <> "CD" Then
     vTipoGestion = "CD"
     Call sbConsultaGestionFlt(vTipoGestion)
  End If
End Sub

Private Sub txtCausaFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
       txtCausaDescFiltro.SetFocus
    End If
    
    If KeyCode = vbKeyF4 Then
        gBusquedas.Convertir = "N"
        gBusquedas.Consulta = "Select COD_CAUSA,DESCRIPCION from CBR_CAUSAS_MOROSIDAD"
        gBusquedas.Columna = "COD_CAUSA"
        gBusquedas.Orden = "COD_CAUSA"
        gBusquedas.Filtro = " and ACTIVA = 1  "
        frmBusquedas.Show vbModal
        txtCausaFiltro = Trim(gBusquedas.Resultado)
        txtCausaDescFiltro = Trim(gBusquedas.Resultado2)
    End If
End Sub

Private Sub txtCausaFiltro_LostFocus()
  If vTipoGestion <> "CC" Then
     vTipoGestion = "CC"
     Call sbConsultaGestionFlt(vTipoGestion)
  End If
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  Call sbBuscar
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select Cedula,nombre from vCBRControlListado"
    gBusquedas.Columna = "Cedula"
    gBusquedas.Orden = "Cedula"
    gBusquedas.Filtro = " and Usuario = '" & txtCodigo.Text & "'"
    frmBusquedas.Show vbModal
    txtCedula.Text = Trim(gBusquedas.Resultado)
    txtNombre.Text = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub sbCargaNombreAsociado(xCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset
On Error GoTo vError

  strSQL = "Select Cedula,nombre from vCBRControlListado" _
         & " where Cedula = '" & txtCedula.Text & "' and Usuario = '" & txtCodigo.Text & "'"
  Call OpenRecordSet(rs, strSQL)
            
  If Not rs.EOF Then
    txtCedula.Text = rs!Cedula
    txtNombre.Text = rs!Nombre
  End If
  
  rs.Close

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCodigo_Change()
 vGrid.MaxRows = 0
 chkMarcas.Value = vbUnchecked
 txtTrasladar.Text = ""
 txtCedula.Text = ""
 txtNombre.Text = ""
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then Call sbBuscar

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select usuario,nombre from cbr_usuarios"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    txtCodigo = Trim(gBusquedas.Resultado)
End If

End Sub

Private Sub txtGestion_LostFocus()
   Call sbCBRControlGestion
End Sub

Private Sub txtGestionDescFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
       txtCausaFiltro.SetFocus
    End If
    
    If KeyCode = vbKeyF4 Then
        gBusquedas.Convertir = "N"
        gBusquedas.Consulta = "Select cod_gestion,descripcion from cbr_gestiones"
        gBusquedas.Columna = "descripcion"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Filtro = " and estado = 1 and nivel_gestion = 'U' "
        frmBusquedas.Show vbModal
        txtGestionFiltro = Trim(gBusquedas.Resultado)
        txtGestionDescFiltro = Trim(gBusquedas.Resultado2)
    End If
End Sub


Private Sub txtGestionDescFiltro_LostFocus()
  If vTipoGestion <> "GD" Then
     vTipoGestion = "GD"
     Call sbConsultaGestionFlt(vTipoGestion)
  End If
End Sub

Private Sub txtGestionFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
       txtGestionDescFiltro.SetFocus
    End If
    
    If KeyCode = vbKeyF4 Then
        gBusquedas.Convertir = "N"
        gBusquedas.Consulta = "Select cod_gestion,descripcion from cbr_gestiones"
        gBusquedas.Columna = "cod_gestion"
        gBusquedas.Orden = "cod_gestion"
        gBusquedas.Filtro = " and estado = 1 and nivel_gestion = 'U' "
        frmBusquedas.Show vbModal
        txtGestionFiltro = Trim(gBusquedas.Resultado)
        txtGestionDescFiltro = Trim(gBusquedas.Resultado2)
    End If
End Sub

Private Sub txtGestionFiltro_LostFocus()
  If vTipoGestion <> "GC" Then
     vTipoGestion = "GC"
     Call sbConsultaGestionFlt(vTipoGestion)
  End If
End Sub

Private Sub txtGestionMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpVence.SetFocus
End Sub

Private Sub txtListaFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    Call sbCargaLista(vTipoGestion, txtListaFiltro.Text)
  End If
End Sub

Private Sub txtTrasladar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado = ""
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select usuario,nombre from cbr_usuarios"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Filtro = " and estado = 1 and usuario <> '" & txtCodigo & "'"
    frmBusquedas.Show vbModal
    txtTrasladar = Trim(gBusquedas.Resultado)
End If
End Sub


Public Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String, Optional vBorra As Boolean = True)
Dim rs As New ADODB.Recordset, i As Integer

If vBorra Then
    vGrid.MaxCols = vGridMaxCol
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
    For i = 1 To vGrid.MaxCols
     vGrid.Col = i
     vGrid.Text = ""
    Next i
End If

Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    If i <> 2 And i <> 3 Then
        If i = 1 Then
            vGrid.Text = CStr(IIf(IsNull(rs.Fields(i - 1).Value), "", rs.Fields(i - 1)))
        Else
            vGrid.Text = CStr(IIf(IsNull(rs.Fields(i - 3).Value), "", rs.Fields(i - 3)))
        End If
    End If
    
    If i = 19 Or i = 14 Then
        vGrid.Text = Format(IIf(IsNull(rs.Fields(i - 3).Value), "", rs.Fields(i - 3)), "yyyy-mm-dd")
    End If
    
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

End Sub

Private Sub vgCobro_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
        
With vgCobro
    .Sheet = NewSheet
    .MaxRows = 0
    
    Select Case NewSheet
      Case 1 'Gestiones
       strSQL = "select S.*, isnull(G.descripcion,'') as 'Gestion'" _
              & "   , isnull(C.DESCRIPCION,'') as 'Causa'" _
              & "   , isnull(A.descripcion,'') as 'Arreglo'" _
              & " from CBR_Seguimiento S  left join cbr_gestiones G on S.cod_gestion = G.cod_gestion" _
              & "  left join CBR_CAUSAS_MOROSIDAD C on S.COD_CAUSA = C.COD_CAUSA" _
              & "  left join CBR_TIPOS_ARREGLOS A on S.COD_ARREGLO = A.COD_ARREGLO" _
              & " where cedula = '" & scTitulo.Tag & "' order by S.cod_seg desc"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          
          For i = 1 To 11
            .Col = i
            Select Case i
              Case 1 'ID
                .Text = CStr(rs!Cod_Seg)
              Case 2 'Fecha
                .Text = Format(rs!fecha, "dd/mm/yyyy")
              Case 3 'vencimiento
                .Text = Format(DateAdd("d", rs!tiempo_resolucion, rs!fecha), "dd/mm/yyyy")
              Case 4 'Gestión
                .Text = rs!Gestion
              Case 5 ' Detalle
                .Text = rs!Notas
                .RowHeight(.Row) = .MaxTextRowHeight(.Row)

              Case 6 ' Ejecutivo
                .Text = rs!Usuario
              Case 7 ' Monto
                .Text = Format(rs!Monto, "Standard")
              Case 8 ' Dias
                .Text = CStr(rs!tiempo_resolucion)
              Case 9  'Arrelgo de Pago
                .Text = rs!Arreglo
              Case 10 'Promesa de Pago
                .Text = Format(rs!Arreglo_Vence & "", "dd/mm/yyyy")
              Case 11 'Causa de Morosidad
                .Text = rs!Causa
                
            End Select
          Next i
          rs.MoveNext
        Loop
        rs.Close
      
      Case 2 'Oficiales
      
        strSQL = "select * from cbr_asignacion_h where cedula = '" & scTitulo.Tag _
               & "' order by fecha_asignacion desc"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          
          For i = 1 To 5
            .Col = i
            Select Case i
              Case 1 'Fecha
                .Text = Format(rs!fecha_asignacion, "dd/mm/yyyy")
              Case 2 'Oficial
                .Text = UCase(rs!Usuario)
              Case 3 'Mantiene
                .Value = rs!mantener
              Case 4 ' Rebajo 2x
                .Value = rs!rebajo_doble
              Case 5 ' Mora
                .Value = rs!aplica_mora
            End Select
          Next i
          rs.MoveNext
        Loop
        rs.Close
      
    End Select
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub vGrid_Click(ByVal Col As Long, ByVal Row As Long)
Dim frm As Form


On Error GoTo vError
    
    tcMain.Enabled = True
    chkFiadoresEstado.Value = vbChecked

    If vGrid.MaxRows = 0 Then
       Exit Sub
    Else
       tcMain.Enabled = True
    End If
    
    If vGrid.Sheet = 1 Then
        
        vGrid.Row = Row
        Select Case Col
    
        Case 2
           
           If vGrid.MaxRows = 0 Then Exit Sub
           Call sbFormsCall("frmCR_ConsultaCreditos")
       
          For Each frm In Forms
            If UCase(frm.Name) = UCase("frmCR_ConsultaCreditos") Then
              vGrid.Row = Row
              vGrid.Col = 5
              Call frm.sbXConsultaAsistida(vGrid.Text)
              Exit For
            End If
          Next frm
             
             
        Case 3
          

             fraGestion.Visible = True
             fraGestion.top = vGrid.top
             fraGestion.Left = vGrid.Left
          
          
          vGrid.Col = 5
          scTitulo.Tag = vGrid.Text
          txtPersonaGestion.Tag = vGrid.Text
          
          Call sbTraeUltimaGestiones(vGrid.Text, txtCodigo.Text)
          Call sbEstadoTxtCbo(vGrid.Text, txtEstado, cboOperacion)
          
          vGrid.Col = 6
          scTitulo.Caption = vGrid.Text
          txtPersonaGestion.Text = vGrid.Text
          
          txtCedula.Enabled = False
          txtNombre.Enabled = False
          

        End Select
        
    End If
     
    vGrid.Col = 5
    If scTitulo.Tag <> vGrid.Text Then
        scTitulo.Tag = vGrid.Text
        vGrid.Col = 6
        scTitulo.Caption = vGrid.Text
        
        Call sbCargarDetalleCuotas
        
    End If


scTitulo.Caption = scTitulo.Caption & "     [Casos: " & Format(vGrid.MaxRows, "###,###,###,###") _
                & ", Mora Financiera: " & Format(curMora, "Standard") _
                & ", Mora Legal: " & Format(curMoraLegal, "Standard") & "]"
Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargarDetalleCuotas()
 Dim strSQL As String
    
tcMain.Item(0).Selected = True
    
 On Error GoTo vError
    If scTitulo.Tag = "" Then Exit Sub
    
    Me.MousePointer = vbHourglass

    vGrid.Col = 5
   
    strSQL = "exec spCbrPersonaMoraGarantia '" & scTitulo.Tag & "','V'"
    vgOperaciones.Sheet = 1
    Call sbCargaGrid(vgOperaciones, 12, strSQL)
    vgOperaciones.MaxRows = vgOperaciones.MaxRows - 1

    
    strSQL = "exec spCbrPersonaMoraDetallada '" & scTitulo.Tag & "'"
    vgOperaciones.Sheet = 2
    Call sbCargaGrid(vgOperaciones, 13, strSQL)
    
    vgOperaciones.MaxRows = vgOperaciones.MaxRows - 1
    
    vgOperaciones.ActiveSheet = 2
    
  
    Me.MousePointer = vbDefault
     
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaGridCartera(vGrid As Object, vGridMaxCol As Integer, strSQL As String, Optional vBorra As Boolean = True)
Dim rs As New ADODB.Recordset, i As Integer
Dim OperacionesAlDia As Long, OperacionesMora As Long, OperacionesCbrJud As Long
Dim SaldoAlDia As Currency, SaldoMora As Currency, SaldoCbrJud As Currency

OperacionesAlDia = 0
OperacionesMora = 0
OperacionesCbrJud = 0
SaldoAlDia = 0
SaldoMora = 0
SaldoCbrJud = 0

If vBorra Then
    vGrid.MaxCols = vGridMaxCol
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
    For i = 1 To vGrid.MaxCols
     vGrid.Col = i
     vGrid.Text = ""
    Next i
End If

Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    vGrid.Text = CStr(IIf(IsNull(rs.Fields(i - 1).Value), "", rs.Fields(i - 1)))
    
    If vGrid.Sheet = 1 Then 'Cartera Al dia y Judicial
        
        ' Totales Saldo Al Dia y Judicial
        If vGrid.Col = 3 Then
            If rs.Fields!Proceso = "Normal" Then
                SaldoAlDia = SaldoAlDia + vGrid.Text
            Else
                SaldoCbrJud = SaldoCbrJud + vGrid.Text
            End If
        End If
        
        ' Totales Operaciones Al Dia y Judicial
        If vGrid.Col = 5 Then
            If rs.Fields!Proceso = "Normal" Then
                OperacionesAlDia = OperacionesAlDia + vGrid.Text
            Else
                OperacionesCbrJud = OperacionesCbrJud + vGrid.Text
            End If
        End If
        
    Else 'Cartera En Mora
    
        ' Totales Saldo Mora
        If vGrid.Col = 3 Then
            SaldoMora = SaldoMora + vGrid.Text
        End If
        
        ' Totales Operaciones Mora
        If vGrid.Col = 4 Then
            OperacionesMora = OperacionesMora + vGrid.Text

        End If
    End If
    
  Next i
  
  ' Asigna totales a etiquetas en pantalla
  If vGrid.Sheet = 1 Then
    lblOperacionesAlDia = OperacionesAlDia
    lblSaldoAlDia = Format(SaldoAlDia, "Standard")
    lblOperacionesCJud = OperacionesCbrJud
    lblSaldoCJud = Format(SaldoCbrJud, "Standard")
  Else
    lblOperacionesMora = OperacionesMora
    lblSaldoMora = Format(SaldoMora, "Standard")
  End If
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        mCambioCelda = True
    Else
        mCambioCelda = False
    End If
End Sub

Private Sub vGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If mCambioCelda Then
        mCambioCelda = False
        
        If Row = NewRow Then Exit Sub
        
        vGrid.Row = NewRow
        vGrid.Col = 5
        scTitulo.Tag = vGrid.Text
        vGrid.Col = 6
        scTitulo.Caption = vGrid.Text
        
        
        scTitulo.Caption = scTitulo.Caption & "     [Casos: " & Format(vGrid.MaxRows, "###,###,###,###") _
                        & ", Mora Financiera: " & Format(curMora, "Standard") _
                        & ", Mora Legal: " & Format(curMoraLegal, "Standard") & "]"
        Call sbCargarDetalleCuotas
        
    Else
        mCambioCelda = False
    End If
End Sub


Private Sub vgridFiadores_Click(ByVal Col As Long, ByVal Row As Long)

  With vgridFiadores
    .Row = Row
    
    Select Case Col
      Case 2
        .Col = 4
        If .Text = "" Then Exit Sub
        GLOBALES.gCedulaActual = .Text
        Call sbFormsCall("frmCR_VerificaDatosPersonales", 1, , , False, Me)
    
    End Select
  End With
  
End Sub


'-----------------------------------------------------------------------------------------
'                          Seguimiento de Gestiones (FRAME)
'-----------------------------------------------------------------------------------------
Private Sub txtArreglo_GotFocus()

If vTipoGestion <> "AC" Then
 vTipoGestion = "AC"
 Call sbCargaLista(vTipoGestion)
End If
 
End Sub

Private Sub txtArreglo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtArregloDesc.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select COD_ARREGLO,DESCRIPCION from CBR_TIPOS_ARREGLOS"
    gBusquedas.Columna = "COD_ARREGLO"
    gBusquedas.Orden = "COD_ARREGLO"
    gBusquedas.Filtro = " and ACTIVO = 1 "
    frmBusquedas.Show vbModal
    txtArreglo = Trim(gBusquedas.Resultado)
    txtArregloDesc = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtCausa_GotFocus()
If vTipoGestion <> "CC" Then
 vTipoGestion = "CC"
 Call sbCargaLista(vTipoGestion)
End If
End Sub

Private Sub txtCausa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtCausaDesc.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select COD_CAUSA,DESCRIPCION from CBR_CAUSAS_MOROSIDAD"
    gBusquedas.Columna = "COD_CAUSA"
    gBusquedas.Orden = "COD_CAUSA"
    gBusquedas.Filtro = " and ACTIVA = 1  "
    frmBusquedas.Show vbModal
    txtCausa = Trim(gBusquedas.Resultado)
    txtCausaDesc = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtArregloDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtGestionMonto.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select COD_ARREGLO,DESCRIPCION from CBR_TIPOS_ARREGLOS"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "COD_ARREGLO"
    gBusquedas.Filtro = " and ACTIVO = 1 "
    frmBusquedas.Show vbModal
    txtArreglo = Trim(gBusquedas.Resultado)
    txtArregloDesc = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtCausaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtArreglo.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select COD_CAUSA,DESCRIPCION from CBR_CAUSAS_MOROSIDAD"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "COD_CAUSA"
    gBusquedas.Filtro = " and ACTIVA = 1  "
    frmBusquedas.Show vbModal
    txtCausa = Trim(gBusquedas.Resultado)
    txtCausaDesc = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtGestion_GotFocus()
If vTipoGestion <> "GC" Then
 vTipoGestion = "GC"
 Call sbCargaLista(vTipoGestion)
End If
End Sub

Private Sub txtGestion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtGestionDesc.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select cod_gestion,descripcion from cbr_gestiones"
    gBusquedas.Columna = "cod_gestion"
    gBusquedas.Orden = "cod_gestion"
    gBusquedas.Filtro = " and estado = 1 and nivel_gestion = 'U' "
    frmBusquedas.Show vbModal
    txtGestion.Text = Trim(gBusquedas.Resultado)
    txtGestionDesc.Text = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtGestionDesc_GotFocus()
If vTipoGestion <> "GD" Then
 vTipoGestion = "GD"
 Call sbCargaLista(vTipoGestion)
End If
End Sub

Private Sub txtGestionDesc_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtCausa.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select cod_gestion,descripcion from cbr_gestiones"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Filtro = " and estado = 1 and nivel_gestion = 'U' "
    frmBusquedas.Show vbModal
    txtGestion = Trim(gBusquedas.Resultado)
    txtGestionDesc = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtGestionDesc_LostFocus()
    Call sbCBRControlGestion
End Sub

Private Sub txtGestionMonto_LostFocus()

    If txtGestionMonto.Text = Empty Then
        txtGestionMonto.Text = Format(0, "Standard")
    End If
    
    If Not IsNumeric(txtGestionMonto) Then
        txtGestionMonto.Text = Format(0, "Standard")
    End If
    
    txtGestionMonto = Format(txtGestionMonto, "Standard")
    
    If txtGestionMonto < vDesviacionMin Then
        MsgBox "El monto es menor que la desviación mínima"
        txtGestionMonto = Format(vDesviacionMin, "Standard")
        txtGestionMonto.SetFocus
    End If

    If txtGestionMonto > vDesviacionMax Then
        MsgBox "El monto es mayor que la desviación máxima"
        txtGestionMonto = Format(vDesviacionMax, "Standard")
        txtGestionMonto.SetFocus
    End If
    
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Call sbBuscar
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select Cedula,nombre from vCBRControlListado"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Filtro = " and Usuario = '" & txtCodigo.Text & "'"
    frmBusquedas.Show vbModal
    txtCedula.Text = Trim(gBusquedas.Resultado)
    txtNombre.Text = Trim(gBusquedas.Resultado2)
End If

End Sub

'Carga la lista con : Gestiones, Causas de morosidad o tipos de areglos
Private Sub sbCargaLista(vTipoGestion As String, Optional vFiltro As String = "")
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

On Error GoTo vError

If vTipoGestion = "" Then Exit Sub

txtListaFiltro.Text = vFiltro

Select Case Mid(vTipoGestion, 1, 1)
   
   Case "G" 'Consulta de gestiones
     fraLista.Caption = "Gestiones"
     strSQL = "Select COD_GESTION as 'Codigo',DESCRIPCION from CBR_GESTIONES" _
            & " where ESTADO = 1 and  NIVEL_GESTION = 'U'"
    
        If vFiltro <> "" Then
            If Right(vTipoGestion, 1) = "C" Then
               strSQL = strSQL & " and COD_GESTION like '%" & txtListaFiltro.Text & "%' order by COD_GESTION"
            Else
               strSQL = strSQL & " and DESCRIPCION like '%" & txtListaFiltro.Text & "%' order by DESCRIPCION"
            End If
        End If
            
    Case "C"  'Consulta de Causas de Mora
      fraLista.Caption = "Causas de Mora"
      strSQL = "Select COD_CAUSA as 'Codigo',DESCRIPCION from CBR_CAUSAS_MOROSIDAD" _
             & " where ACTIVA = 1"
      If vFiltro <> "" Then
         If Right(vTipoGestion, 1) = "C" Then
            strSQL = strSQL & " and COD_CAUSA like '%" & txtListaFiltro.Text & "%' order by COD_CAUSA"
         Else
            strSQL = strSQL & " and DESCRIPCION like '%" & txtListaFiltro.Text & "%' order by DESCRIPCION"
         End If
      End If
            
    Case "A" 'Consulta de Tipos de Arreglos
      fraLista.Caption = "Arreglos"
      strSQL = "Select COD_ARREGLO as 'Codigo',DESCRIPCION from CBR_TIPOS_ARREGLOS" _
             & " where ACTIVO = 1"
      
      If vFiltro <> "" Then
         If Right(vTipoGestion, 1) = "C" Then
          strSQL = strSQL & " and COD_ARREGLO like '%" & txtListaFiltro.Text & "%' order by COD_CAUSA"
         Else
          strSQL = strSQL & " and DESCRIPCION like '%" & txtListaFiltro.Text & "%' order by DESCRIPCION"
         End If
      End If
      
End Select

If Right(vTipoGestion, 1) = "C" Then
    fraLista.Caption = fraLista.Caption & " [Código]"
Else
    fraLista.Caption = fraLista.Caption & " [Descripción]"
End If

Call OpenRecordSet(rs, strSQL)
     
lswLista.ListItems.Clear
     
Do While Not rs.EOF
  Set itmX = lswLista.ListItems.Add(, , Trim(rs!Codigo))
      itmX.SubItems(1) = rs!Descripcion
  rs.MoveNext
Loop

rs.Close
Me.MousePointer = vbDefault


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

'Consulta el codigo y la descripcion de las gestiones
'la causa de mora y el tipo de arreglo
Private Sub sbConsultaGestion(vTipoGestion As String)
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

On Error GoTo vError

Select Case Mid(vTipoGestion, 1, 1)
   
   Case "G" 'Consulta de gestiones
      If Mid(vTipoGestion, 2, 1) = "C" Then
        strSQL = "Select COD_GESTION as 'Codigo',DESCRIPCION, MONTO from CBR_GESTIONES" _
               & " where COD_GESTION like '%" & txtGestion.Text & "%' and ESTADO = 1 and  NIVEL_GESTION = 'U'"
      Else
        strSQL = "Select COD_GESTION as 'Codigo',DESCRIPCION, MONTO from CBR_GESTIONES" _
               & " where DESCRIPCION like  '%" & Trim(txtGestionDesc.Text) & "%' and ESTADO = 1 and  NIVEL_GESTION = 'U'"
      End If
      
      Call OpenRecordSet(rs, strSQL)
      strSQL = ""

      If Not rs.EOF Then
        txtGestionDesc = Trim(rs!Descripcion)
        txtGestionMonto = Format(rs!Monto, "Standard")
      Else
        txtGestion.Text = Empty
        txtGestionDesc.Text = Empty
        txtGestionMonto = Format(0, "Standard")
      End If
   
   
   Case "C"
      If Mid(vTipoGestion, 2, 1) = "C" Then
        strSQL = "Select COD_CAUSA as 'Codigo',DESCRIPCION from CBR_CAUSAS_MOROSIDAD" _
               & " where COD_CAUSA like  '%" & Trim(txtCausa.Text) & "%' and ACTIVA = 1"
      Else
        strSQL = "Select COD_CAUSA as 'Codigo',DESCRIPCION from CBR_CAUSAS_MOROSIDAD" _
               & " where DESCRIPCION like '%" & Trim(txtCausaDesc.Text) & "%' and ACTIVA = 1"
      End If
      
      Call OpenRecordSet(rs, strSQL)
      strSQL = ""

      If Not rs.EOF Then
        txtCausaDesc = Trim(rs!Descripcion)
      Else
        txtCausa.Text = Empty
        txtCausaDesc.Text = Empty
      End If
    
    
    Case "A"
      If Mid(vTipoGestion, 2, 1) = "C" Then
        strSQL = "Select COD_ARREGLO as 'Codigo',DESCRIPCION from CBR_TIPOS_ARREGLOS" _
               & " where COD_ARREGLO like '%" & Trim(txtArreglo.Text) & "%' and ACTIVO = 1"
      Else
        strSQL = "Select COD_ARREGLO as 'Codigo',DESCRIPCION from CBR_TIPOS_ARREGLOS" _
               & " where DESCRIPCION like '%" & Trim(txtArregloDesc.Text) & "%' and ACTIVO = 1"
      End If
      
      Call OpenRecordSet(rs, strSQL)
      strSQL = ""

      If Not rs.EOF Then
        txtArregloDesc = Trim(rs!Descripcion)
      Else
        txtArreglo.Text = Empty
        txtArregloDesc.Text = Empty
      End If

End Select

rs.Close

Me.MousePointer = vbDefault

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdAplica_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String, mAccesoRestringido As Boolean

On Error GoTo vError

Me.MousePointer = vbHourglass

'Verifica datos
vMensaje = ""

If txtEstado.Tag = "N" Then vMensaje = vMensaje & " - La persona no se encuentra morosa verifique..." & vbCrLf
If txtNotas.Text = "" Then vMensaje = vMensaje & " - No se especificó ninguna observación..." & vbCrLf

strSQL = "select isnull(count(*),0) as Existe from cbr_usuarios where usuario = '" _
       & glogon.Usuario & "' and estado = 1"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then vMensaje = vMensaje & " - El usuario actual no se encuentra activo..." & vbCrLf
rs.Close

strSQL = "select isnull(count(*),0) as Existe from cbr_gestiones where cod_gestion = '" _
       & txtGestion & "' and estado = 1 and NIVEL_GESTION = 'U'"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then vMensaje = vMensaje & " - La gestion actual no se encuentra activa..." & vbCrLf
rs.Close

'Preguntar si existe el parametro de sgt sin asignacion previa, de lo contrario buscar asignacion
strSQL = "select valor from cbr_parametros where cod_parametro = '05'"
Call OpenRecordSet(rs, strSQL)
If Mid(rs!Valor, 1, 1) <> "S" Then
  rs.Close
  strSQL = "select isnull(count(*),0) as Existe from cbr_asignacion where usuario = '" _
       & glogon.Usuario & "' and cedula = '" & txtPersonaGestion.Tag & "'"
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe = 0 Then vMensaje = vMensaje & " - Este expediente no se encuentra asignado al usuario actual, verifique..." & vbCrLf
End If
rs.Close

If vMensaje <> "" Then
  Me.MousePointer = vbDefault
  MsgBox vMensaje, vbExclamation
  Exit Sub
End If

strSQL = "exec spCBRControlSGT '" & txtPersonaGestion.Tag & "','" & glogon.Usuario & "','" & txtGestion.Text _
       & "','" & Format(dtpVence.Value, "yyyy/mm/dd") & "','" & txtNotas & "','" & GLOBALES.gOficinaTitular _
       & "'," & CCur(txtGestionMonto.Text) & ""
       
If cboOperacion.Text = "+ Antigua" Then
   strSQL = strSQL & ",0"
Else
   strSQL = strSQL & "," & cboOperacion.Text
End If

strSQL = strSQL & "," & "'" & txtCausa.Text & "','" & txtArreglo.Text & "'"
       
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

MsgBox "Seguimiento Registrado Satisfactoriamente...", vbInformation

fraGestion.Visible = False
fraGestion.Left = vGrid.Height + 5000

Call sblimpiarFrameGestiones

txtCedula.Enabled = True
txtNombre.Enabled = True

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCBRControlGestion()
Dim strSQL As String, rs As New ADODB.Recordset

    If txtGestion.Text = Empty Then
        Exit Sub
    End If

    strSQL = "select descripcion, isnull(monto,0) as Monto, MODIFICA_USUARIO, isnull(MODIFICA_DESVIACION,0) as MODIFICA_DESVIACION " _
        & " from cbr_gestiones where estado = 1 " _
        & " and nivel_gestion = 'U' and cod_gestion = '" & Trim(txtGestion.Text) & "'"
    Call OpenRecordSet(rs, strSQL)

    strSQL = ""
    If Not rs.EOF Then
    
        txtGestionDesc.Text = Trim(rs!Descripcion)
        txtGestionMonto.Text = Format(rs!Monto, "Standard")
        vDesviacionMax = rs!Monto + rs!MODIFICA_DESVIACION
        vDesviacionMin = rs!Monto - rs!MODIFICA_DESVIACION
        txtGestionMonto.ToolTipText = "Min: " & Format(vDesviacionMin, "Standard") & " Max: " & Format(vDesviacionMax, "Standard")
        If rs!MODIFICA_USUARIO = 1 Then
           txtGestionMonto.Locked = False
        Else
           txtGestionMonto.Locked = True
        End If
        
    Else
    
        txtGestion.Text = Empty
        txtGestionDesc.Text = Empty
        txtGestionMonto = Format(0, "Standard")
        txtGestionMonto.Locked = True
        vDesviacionMax = 0
        vDesviacionMin = 0
        
    End If
    rs.Close
    
    strSQL = "select dbo.fxCBRGestionUsuario('" & Trim(txtGestion.Text) & "','" & glogon.Usuario & "') as acceso"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
         If rs!Acceso = 0 Then
            MsgBox "El usuario no tiene acceso a esta gestión"
            
            txtGestion.Text = Empty
            txtGestionDesc.Text = Empty
            txtGestionMonto = Format(0, "Standard")
            txtGestionMonto.Locked = True
            vDesviacionMax = 0
            vDesviacionMin = 0
            
            txtGestion.SetFocus
            Exit Sub
         End If
    End If
    rs.Close

End Sub


Public Sub sbEstadoTxtCbo(xCedula As String, txtX As TextBox, cboX As ComboBox)
Dim strSQL As String, rs As New ADODB.Recordset
   
   cboX.Clear
   cboX.AddItem "+ Antigua"
   cboX.Text = "+ Antigua"

'Busca Datos de la Persona
strSQL = "select R.id_solicitud,R.codigo,R.saldo,R.plazo,R.interesv,R.Garantia,V.cuota,V.intc,V.intm,V.amortiza" _
       & " from reg_creditos R inner join vista_morosidad V on R.id_solicitud = V.id_solicitud" _
       & " inner join catalogo C on R.codigo = C.codigo and C.Linea_Interna = 1" _
       & " where R.cedula = '" & xCedula & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
   txtX.Tag = "N"
   txtX.Text = "Operaciones Activas al ** Dia **"
Else
   txtX.Tag = "S"
   txtX.Text = "Presenta operaciones con morosidad" & vbCrLf & vbCrLf
   Do While Not rs.EOF
     cboX.AddItem rs!ID_SOLICITUD
     rs.MoveNext
   Loop
End If
rs.Close

End Sub

Private Sub sblimpiarFrameGestiones()

  txtGestion.Text = ""
  txtGestionDesc.Text = ""
  txtCausa.Text = ""
  txtCausaDesc.Text = ""
  txtArreglo.Text = ""
  txtArregloDesc.Text = ""
  txtPersonaGestion.Text = ""
  txtPersonaGestion.Tag = ""
  txtGestionMonto.Text = Format("0.00", "standard")
  txtListaFiltro.Text = ""
  lswLista.ListItems.Clear
  cboOperacion.Clear
  txtNotas.Text = ""
  
End Sub


'Consulta el codigo y la descripcion de las gestiones
'la causa de mora y el tipo de arreglo
Private Sub sbConsultaGestionFlt(vTipoGestion As String)
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

On Error GoTo vError

Select Case Mid(vTipoGestion, 1, 1)
   
   Case "G" 'Consulta de gestiones
      If Mid(vTipoGestion, 2, 1) = "C" Then
        strSQL = "Select COD_GESTION as 'Codigo',DESCRIPCION, MONTO from CBR_GESTIONES" _
               & " where COD_GESTION like '%" & txtGestionFiltro.Text & "%' and ESTADO = 1 and  NIVEL_GESTION = 'U'"
      Else
        strSQL = "Select COD_GESTION as 'Codigo',DESCRIPCION, MONTO from CBR_GESTIONES" _
               & " where DESCRIPCION like  '%" & Trim(txtGestionDescFiltro.Text) & "%' and ESTADO = 1 and  NIVEL_GESTION = 'U'"
      End If
      
      Call OpenRecordSet(rs, strSQL)
      strSQL = ""

      If Not rs.EOF Then
        txtGestionDescFiltro.Text = Trim(rs!Descripcion)
      Else
        txtGestionFiltro.Text = Empty
        txtGestionDescFiltro.Text = Empty
      End If
   
   
   Case "C"
      If Mid(vTipoGestion, 2, 1) = "C" Then
        strSQL = "Select COD_CAUSA as 'Codigo',DESCRIPCION from CBR_CAUSAS_MOROSIDAD" _
               & " where COD_CAUSA like  '%" & Trim(txtCausaFiltro.Text) & "%' and ACTIVA = 1"
      Else
        strSQL = "Select COD_CAUSA as 'Codigo',DESCRIPCION from CBR_CAUSAS_MOROSIDAD" _
               & " where DESCRIPCION like '%" & Trim(txtCausaDescFiltro.Text) & "%' and ACTIVA = 1"
      End If
      
      Call OpenRecordSet(rs, strSQL)
      strSQL = ""

      If Not rs.EOF Then
        txtCausaDescFiltro.Text = Trim(rs!Descripcion)
      Else
        txtCausaFiltro.Text = Empty
        txtCausaDescFiltro.Text = Empty
      End If
    
    
    Case "A"
      If Mid(vTipoGestion, 2, 1) = "C" Then
        strSQL = "Select COD_ARREGLO as 'Codigo',DESCRIPCION from CBR_TIPOS_ARREGLOS" _
               & " where COD_ARREGLO like '%" & Trim(txtArregloFiltro.Text) & "%' and ACTIVO = 1"
      Else
        strSQL = "Select COD_ARREGLO as 'Codigo',DESCRIPCION from CBR_TIPOS_ARREGLOS" _
               & " where DESCRIPCION like '%" & Trim(txtArregloDescFiltro.Text) & "%' and ACTIVO = 1"
      End If
      
      Call OpenRecordSet(rs, strSQL)
      strSQL = ""

      If Not rs.EOF Then
        txtArregloDescFiltro = Trim(rs!Descripcion)
      Else
        txtArregloFiltro.Text = Empty
        txtArregloDescFiltro.Text = Empty
      End If

End Select

rs.Close

Me.MousePointer = vbDefault

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbTraeUltimaGestiones(vCedula As String, vUsuario As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


strSQL = "Select ULT_COD_GESTION, GestionDesc , COD_CAUSA, CausaDesc, COD_ARREGLO, ArregloDesc, isnull(Arreglo_Vence, dbo.MyGetdate()) as 'Arreglo_Vence'" _
       & " From dbo.vCBRControlListado" _
       & " where Cedula = '" & vCedula & "'" ' and usuario = '" & vUsuario & "'"

Call OpenRecordSet(rs, strSQL, 0)

If Not rs.EOF Then
  txtGestion.Text = rs!ULT_COD_GESTION & ""
  txtGestionDesc.Text = rs!GestionDesc & ""
  txtCausa.Text = rs!COD_CAUSA & ""
  txtCausaDesc.Text = rs!CausaDesc & ""
  txtArreglo.Text = rs!COD_ARREGLO & ""
  txtArregloDesc.Text = rs!ArregloDesc & ""
  dtpVence.Value = rs!Arreglo_Vence
  Call sbCBRControlGestion
End If

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


