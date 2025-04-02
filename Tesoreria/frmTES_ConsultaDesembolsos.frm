VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmTES_ConsultaDesembolsos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Consulta de Desembolsos (Documentos)"
   ClientHeight    =   11190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17685
   Icon            =   "frmTES_ConsultaDesembolsos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11190
   ScaleWidth      =   17685
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   3600
      Top             =   1920
   End
   Begin XtremeSuiteControls.GroupBox gbFiltros 
      Height          =   2295
      Left            =   3240
      TabIndex        =   14
      Top             =   -120
      Width           =   13575
      _Version        =   1441793
      _ExtentX        =   23945
      _ExtentY        =   4048
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtBeneficiario 
         Height          =   315
         Left            =   1440
         TabIndex        =   25
         Top             =   840
         Width           =   7695
         _Version        =   1441793
         _ExtentX        =   13573
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   315
         Left            =   1440
         TabIndex        =   26
         Top             =   1200
         Width           =   7695
         _Version        =   1441793
         _ExtentX        =   13573
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   495
         Left            =   6840
         TabIndex        =   33
         Top             =   1680
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Buscar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTES_ConsultaDesembolsos.frx":6852
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnInforme 
         Height          =   495
         Left            =   8040
         TabIndex        =   34
         Top             =   1680
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Informe"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTES_ConsultaDesembolsos.frx":6F52
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   495
         Left            =   9360
         TabIndex        =   35
         Top             =   1680
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Exportar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTES_ConsultaDesembolsos.frx":7659
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   330
         Left            =   1440
         TabIndex        =   36
         Top             =   1560
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   315
         Left            =   1440
         TabIndex        =   18
         Top             =   480
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   315
         Left            =   3360
         TabIndex        =   20
         Top             =   480
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin XtremeSuiteControls.FlatEdit txtTF 
         Height          =   315
         Left            =   5280
         TabIndex        =   22
         Top             =   480
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin XtremeSuiteControls.FlatEdit txtAppCod 
         Height          =   315
         Left            =   7200
         TabIndex        =   24
         Top             =   480
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin XtremeSuiteControls.FlatEdit txtRef01 
         Height          =   315
         Left            =   9960
         TabIndex        =   28
         Top             =   480
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
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
      Begin XtremeSuiteControls.FlatEdit txtRef02 
         Height          =   315
         Left            =   9960
         TabIndex        =   30
         Top             =   840
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
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
      Begin XtremeSuiteControls.FlatEdit txtRef03 
         Height          =   315
         Left            =   9960
         TabIndex        =   32
         Top             =   1200
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
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
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   330
         Left            =   3360
         TabIndex        =   37
         Top             =   1560
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.CheckBox chkModoProtegido 
         Height          =   375
         Left            =   2400
         TabIndex        =   38
         Top             =   1920
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Visualizar TF en Modo Protegido"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ListView lswDocumentos 
         Height          =   1815
         Left            =   11280
         TabIndex        =   42
         Top             =   480
         Width           =   3015
         _Version        =   1441793
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
         MultiSelect     =   -1  'True
         HideSelection   =   0   'False
         View            =   3
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkDocumentos 
         Height          =   210
         Left            =   13200
         TabIndex        =   43
         Top             =   240
         Width           =   210
         _Version        =   1441793
         _ExtentX        =   370
         _ExtentY        =   370
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   9
         Left            =   11280
         TabIndex        =   44
         Top             =   240
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo Documento"
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
         Left            =   9240
         TabIndex        =   31
         Top             =   1200
         Width           =   615
         _Version        =   1441793
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ref.03"
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
         Left            =   9240
         TabIndex        =   29
         Top             =   840
         Width           =   615
         _Version        =   1441793
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ref.02"
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
         Left            =   9240
         TabIndex        =   27
         Top             =   480
         Width           =   615
         _Version        =   1441793
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ref.01"
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
         Left            =   7200
         TabIndex        =   23
         Top             =   240
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Id Aplicación"
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
         Left            =   5280
         TabIndex        =   21
         Top             =   240
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Transferencia"
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
         Left            =   3360
         TabIndex        =   19
         Top             =   240
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Documento"
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
         TabIndex        =   17
         Top             =   1200
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Beneficiario"
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
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Código"
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
   Begin XtremeSuiteControls.CheckBox chkBancos 
      Height          =   210
      Left            =   2880
      TabIndex        =   13
      Top             =   600
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
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
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.ListView lswBancos 
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
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
      Checkboxes      =   -1  'True
      MultiSelect     =   -1  'True
      HideSelection   =   0   'False
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   16
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   10935
      Width           =   17685
      _ExtentX        =   31194
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
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6495
      Left            =   3360
      TabIndex        =   3
      Top             =   2280
      Width           =   10695
      _Version        =   524288
      _ExtentX        =   18865
      _ExtentY        =   11456
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
      MaxCols         =   28
      SpreadDesigner  =   "frmTES_ConsultaDesembolsos.frx":7F2A
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   315
      Left            =   1560
      TabIndex        =   8
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
      TabIndex        =   9
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
      TabIndex        =   10
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
      TabIndex        =   11
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
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   330
      Left            =   120
      TabIndex        =   39
      Top             =   120
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
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
   Begin XtremeSuiteControls.FlatEdit txtFiltroCta 
      Height          =   330
      Left            =   120
      TabIndex        =   40
      Top             =   840
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtFiltroConceptos 
      Height          =   330
      Left            =   120
      TabIndex        =   41
      Top             =   4680
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.ListView lswConceptos 
      Height          =   2775
      Left            =   120
      TabIndex        =   45
      Top             =   5040
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
      _ExtentY        =   4895
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
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkConceptos 
      Height          =   210
      Left            =   2880
      TabIndex        =   46
      Top             =   4440
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
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
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
      Alignment       =   1
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
      Left            =   240
      TabIndex        =   7
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Conceptos ...:"
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
      TabIndex        =   6
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta Bancaria...:"
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
      TabIndex        =   5
      Top             =   600
      Width           =   1815
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
      Left            =   240
      TabIndex        =   2
      Top             =   9000
      Width           =   1215
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
      Left            =   240
      TabIndex        =   1
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Left            =   240
      TabIndex        =   0
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Image imgBanner 
      Height          =   9396
      Left            =   0
      Picture         =   "frmTES_ConsultaDesembolsos.frx":8DAB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3240
   End
End
Attribute VB_Name = "frmTES_ConsultaDesembolsos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub btnBuscar_Click()
    Call sbBuscar
End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 28
    vHeaders.Headers(2) = "No. Solicitud"
    vHeaders.Headers(3) = "No. Documento"
    vHeaders.Headers(4) = "Tipo Doc."
    vHeaders.Headers(5) = "Monto"
    vHeaders.Headers(6) = "Estado"
    vHeaders.Headers(7) = "Fec.Emisión"
    vHeaders.Headers(8) = "Fec.Anulación"
    vHeaders.Headers(9) = "Beneficiario"
    vHeaders.Headers(10) = "No. Cuenta"
    vHeaders.Headers(11) = "Banco"
    vHeaders.Headers(12) = "Código"
    vHeaders.Headers(13) = "Detalle"
    vHeaders.Headers(14) = "Unidad"
    vHeaders.Headers(15) = "Concepto"
    vHeaders.Headers(16) = "Tipo Cliente"
    vHeaders.Headers(17) = "Us.Solicita"
    vHeaders.Headers(18) = "Us.Emite"
    vHeaders.Headers(19) = "Us.Anula"
    vHeaders.Headers(20) = "Divisa"
    vHeaders.Headers(21) = "Tipo Cambio"
    vHeaders.Headers(22) = "Grupo Bancario"
    vHeaders.Headers(23) = "Periodo"
    vHeaders.Headers(24) = "Ref.No.1"
    vHeaders.Headers(25) = "Ref.No.2"
    vHeaders.Headers(26) = "Ref.No.3"
    vHeaders.Headers(27) = "No.Doc.Bancario"
    vHeaders.Headers(28) = "No.Desembolso"
    
      Call sbSIFGridExportar(vGrid, vHeaders, "Bancos_ConsultaDesembolsos")

    
'Select Case ButtonMenu.Key
'  Case "Excel"
'      Call sbSIFGridExportar(vGrid, vHeaders, "Bancos_ConsultaDesembolsos")
'  Case "HTML"
'      Call sbSIFGridExportar(vGrid, vHeaders, "Bancos_ConsultaDesembolsos", "HTML")
'End Select

End Sub

Private Sub btnInforme_Click()
        vGrid.PrintHeader = "Bancos: Consulta Desembolsos, Fecha : " & fxFechaServidor & " Usuario : " & glogon.Usuario
        vGrid.PrintFooter = "Estado:" & cboEstado.Text & "..." & cboFechas.Text & "...I:" & Format(dtpInicio.Value, "dd/mm/yyyy") & " C.:" & Format(dtpCorte.Value, "dd/mm/yyyy")
        vGrid.PrintOrientation = PrintOrientationLandscape
        vGrid.PrintSheet
End Sub

Private Sub cboBanco_Click()

If vPaso Then Exit Sub

Call sbCuentas_Load

End Sub

Private Sub cboFechas_Click()

If cboFechas.Text = "[Todas]" Then
  dtpInicio.Enabled = False
  dtpCorte.Enabled = False
Else
  dtpInicio.Enabled = True
  dtpCorte.Enabled = True
End If

End Sub

Private Sub chkBancos_Click()
Dim i As Integer

For i = 1 To lswBancos.ListItems.Count
  lswBancos.ListItems.Item(i).Checked = chkBancos.Value
Next i

End Sub

Private Sub chkConceptos_Click()
Dim i As Integer

For i = 1 To lswConceptos.ListItems.Count
  lswConceptos.ListItems.Item(i).Checked = chkConceptos.Value
Next i
End Sub

Private Sub chkDocumentos_Click()
Dim i As Integer

For i = 1 To lswDocumentos.ListItems.Count
  lswDocumentos.ListItems.Item(i).Checked = chkDocumentos.Value
Next i

End Sub


Private Function fxValida_Textos() As Boolean
Dim i As Integer

i = 0
If Not fxSIFValidaCadena(txtCodigo.Text) Then
   i = i + 1
End If
If Not fxSIFValidaCadena(txtBeneficiario.Text) Then
   i = i + 1
End If
If Not fxSIFValidaCadena(txtDetalle.Text) Then
   i = i + 1
End If
If Not fxSIFValidaCadena(txtDocumento.Text) Then
   i = i + 1
End If
If Not fxSIFValidaCadena(txtTF.Text) Then
   i = i + 1
End If
If Not fxSIFValidaCadena(txtRef01.Text) Then
   i = i + 1
End If
If Not fxSIFValidaCadena(txtRef02.Text) Then
   i = i + 1
End If
If Not fxSIFValidaCadena(txtRef03.Text) Then
   i = i + 1
End If
If Not fxSIFValidaCadena(txtUsuario.Text) Then
   i = i + 1
End If
If Not fxSIFValidaCadena(txtAppCod.Text) Then
   i = i + 1
End If

If i = 0 Then
  fxValida_Textos = True
Else
  fxValida_Textos = False
End If

End Function

Private Sub sbBuscar()
Dim strSQL As String, i As Integer
Dim vCadena As String

On Error GoTo vError

If Not fxValida_Textos Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "select 0,C.nsolicitud,isnull(C.ndocumento,0),C.Tipo,C.monto,case when C.estado = 'I' or C.estado = 'E' or C.estado = 'T' then 'Emitido'" _
       & " when C.estado = 'A' then 'Anulado' when C.estado = 'P' then 'Pendiente' end as Estado" _
       & ",isnull(C.fecha_emision,'') as Fecha_Emision, isnull(C.fecha_Anula,'') as Fecha_Anula,C.Beneficiario,C.Cta_Ahorros" _
       & ",B.descripcion as Banco,C.codigo, (isnull(C.Detalle1,'') + ' '  + isnull(C.Detalle2, '') + ' ' + isnull(C.Detalle3,'') + ' ' + isnull(C.Detalle4,'') + ' ' + isnull(C.Detalle5,'')) as 'Detalle'" _
       & ",U.descripcion as 'Unidad',Con.descripcion as 'Concepto', case C.Tipo_Beneficiario when 1 then 'Personas'" _
       & " when 2 then 'Bancos' when 3 then 'Proveedores' when 4 then 'Acreedores' end as TipoBeneficio" _
       & ",C.User_Solicita,C.User_Genera,C.User_Anula,C.cod_divisa,C.Tipo_Cambio,Grp.Descripcion as 'GrupoBancario'" _
       & ", case month(C.fecha_emision) when 1 then convert(varchar(4),year(C.fecha_emision)) + ' - 01 Enero'" _
       & "       when 2 then convert(varchar(4),year(C.fecha_emision)) + ' - 02 Febero' when 3 then convert(varchar(4),year(C.fecha_emision)) + ' - 03 Marzo'" _
       & "       when 4 then convert(varchar(4),year(C.fecha_emision)) + ' - 04 Abril'  when 5 then convert(varchar(4),year(C.fecha_emision)) + ' - 05 Mayo'" _
       & "       when 6 then convert(varchar(4),year(C.fecha_emision)) + ' - 06 Junio'  when 7 then convert(varchar(4),year(C.fecha_emision)) + ' - 07 Julio'" _
       & "       when 8 then convert(varchar(4),year(C.fecha_emision)) + ' - 08 Agosto' when 9 then convert(varchar(4),year(C.fecha_emision)) + ' - 09 Setiembre'" _
       & "       when 10 then convert(varchar(4),year(C.fecha_emision)) + ' - 10 Octubre' when 11 then convert(varchar(4),year(C.fecha_emision)) + ' - 11 Noviembre'" _
       & "       when 12 then convert(varchar(4),year(C.fecha_emision)) + ' - 12 Diciembre' else '' end as 'Periodo'" _
       & ",C.REF_01,C.REF_02,C.REF_03, '' as 'NoDocBancario', Null as 'IdDesembolso'" _
       & " from Tes_Transacciones C inner join Tes_Bancos B on C.id_banco = B.id_Banco" _
       & " left join tes_bancos_grupos Grp on B.cod_grupo = Grp.Cod_Grupo " _
       & " left join CntX_Unidades U on C.cod_unidad = U.cod_unidad and U.cod_contabilidad = " & GLOBALES.gEnlace _
       & " left join Tes_Conceptos Con on C.cod_concepto = Con.cod_concepto"


Select Case cboEstado.Text
  Case "Emitido"
     strSQL = strSQL & " Where C.estado in('I','T','E')"
  Case "Anulado"
     strSQL = strSQL & " Where C.estado = 'A'"
  Case "Solicitado"
     strSQL = strSQL & " Where C.estado = 'P'"
  Case Else
     strSQL = strSQL & " Where C.estado = C.Estado"
End Select

If chkModoProtegido.Value = xtpUnchecked Then
     strSQL = strSQL & " AND ISNULL(C.MODO_PROTEGIDO,0) = 0"
End If

If Len(Trim(txtUsuario.Text)) > 0 Then
    Select Case cbo.Text
        Case "Solicita"
          strSQL = strSQL & " and C.user_solicita like '%" & txtUsuario.Text & "%'"
        Case "Autoriza"
          strSQL = strSQL & " and C.user_autoriza like '%" & txtUsuario.Text & "%'"
        Case "Emite"
          strSQL = strSQL & " and C.user_genera like '%" & txtUsuario.Text & "%'"
        Case "Anula"
          strSQL = strSQL & " and C.user_anula like '%" & txtUsuario.Text & "%'"
    End Select
End If

If Len(Trim(txtCodigo.Text)) > 0 Then
      strSQL = strSQL & " and C.codigo like '%" & txtCodigo.Text & "%'"
End If

If Len(Trim(txtBeneficiario.Text)) > 0 Then
      strSQL = strSQL & " and C.beneficiario like '%" & txtBeneficiario.Text & "%'"
End If


If Len(Trim(txtDetalle.Text)) > 0 Then
      strSQL = strSQL & " and (C.Detalle1 + C.Detalle2 + isnull(C.Detalle3,'') ) like '%" & txtDetalle.Text & "%'"
End If

If Len(Trim(txtDocumento.Text)) > 0 Then
      strSQL = strSQL & " and C.ndocumento like '%" & txtDocumento.Text & "%'"
End If

If Len(Trim(txtAppCod.Text)) > 0 Then
      strSQL = strSQL & " and isnull(C.Cod_App,'') like '%" & txtAppCod.Text & "%'"
End If


If Len(Trim(txtRef01.Text)) > 0 Then
      strSQL = strSQL & " and isnull(C.Ref_01,'') like '%" & txtRef01.Text & "%'"
End If

If Len(Trim(txtRef02.Text)) > 0 Then
      strSQL = strSQL & " and isnull(C.Ref_02,'') like '%" & txtRef02.Text & "%'"
End If

If Len(Trim(txtRef03.Text)) > 0 Then
      strSQL = strSQL & " and isnull(C.Ref_03,'') like '%" & txtRef03.Text & "%'"
End If

If Len(Trim(txtTF.Text)) > 0 Then
      strSQL = strSQL & " and C.Documento_Base like '%" & txtTF.Text & "%'"
End If



'Lista de Tes_Bancos
vCadena = " and C.id_banco in(0"
For i = 1 To lswBancos.ListItems.Count
  If lswBancos.ListItems.Item(i).Checked Then
    vCadena = vCadena & "," & lswBancos.ListItems.Item(i).Tag
  End If
Next i
strSQL = strSQL & vCadena & ")"


'Lista de Tipos de Documentos
vCadena = " and C.Tipo in('"
For i = 1 To lswDocumentos.ListItems.Count
  If lswDocumentos.ListItems.Item(i).Checked Then
    vCadena = vCadena & "','" & lswDocumentos.ListItems.Item(i).Tag
  End If
Next i
strSQL = strSQL & vCadena & "')"


'Lista de Conceptos
vCadena = " and C.Cod_Concepto in('"
For i = 1 To lswConceptos.ListItems.Count
  If lswConceptos.ListItems.Item(i).Checked Then
    vCadena = vCadena & "','" & lswConceptos.ListItems.Item(i).Tag
  End If
Next i
strSQL = strSQL & vCadena & "')"



Select Case cboFechas.Text
  Case "Emisión"
    strSQL = strSQL & " and C.fecha_emision between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
  Case "Anulación"
    strSQL = strSQL & " and C.fecha_anula between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
  Case "Solicitud"
    strSQL = strSQL & " and C.fecha_solicitud between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End Select

strSQL = strSQL & " order by C.fecha_emision asc "

vPaso = True
    Call sbCargaGridLocal(vGrid, 28, strSQL)
vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbReporte()
Dim strSQL As String, vRango As String

On Error GoTo vError

Me.MousePointer = vbHourglass

If Len(txtCodigo.Text) > 0 Then
  vRango = "Código : " & txtCodigo.Text
  strSQL = "MID({CHEQUES.CODIGO},1," & Len(Trim(txtCodigo)) & ") = '" & txtCodigo & "'"
End If

If Len(txtBeneficiario.Text) > 0 Then
  vRango = "Beneficiario : " & txtBeneficiario.Text
  strSQL = "MID({CHEQUES.BENEFICIARIO},1," & Len(Trim(txtBeneficiario)) & ") = '" & txtBeneficiario & "'"
End If


If cboFechas.Text <> "Todas" Then
    If Mid(cboFechas, 1, 2) = "01" Then
      vRango = vRango & " Emision entre " & dtpInicio.Value & " y " & dtpCorte.Value
      strSQL = strSQL & " AND {CHEQUES.FECHA_EMISION} in date(" & Year(dtpInicio.Value) & "," & Month(dtpInicio.Value) _
             & "," & Day(dtpInicio.Value) & ") to Date(" & Year(dtpCorte.Value) & "," & Month(dtpCorte.Value) _
             & "," & Day(dtpCorte.Value) & ")"
    Else
      vRango = vRango & " Anulación entre " & dtpInicio.Value & " y " & dtpCorte.Value
      strSQL = strSQL & " AND {CHEQUES.FECHA_ANULA} in date(" & Year(dtpInicio.Value) & "," & Month(dtpInicio.Value) _
             & "," & Day(dtpInicio.Value) & ") to Date(" & Year(dtpCorte.Value) & "," & Month(dtpCorte.Value) _
             & "," & Day(dtpCorte.Value) & ")"
    End If
End If


With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Banking"
    
    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "Fecha='" & Format(fxFechaServidor, ("dd/mm/yyyy")) & "'"
    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "rango='" & vRango & "'"
    
    .ReportFileName = SIFGlobal.fxPathReportes("Banking_Desembolsos.rpt")
    .SelectionFormula = strSQL
    
    
    .PrintReport
End With

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
 vGrid.col = i
 vGrid.Text = ""
Next i

curMonto = 0

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i

    If rs.Fields(i - 1).Type = 135 Then
        If Year(rs.Fields(i - 1).Value & "") > 1900 Then
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

StatusBarX.Panels(1).Text = "Casos ..: " & Format(rs.RecordCount, "###,###,##0")
StatusBarX.Panels(2).Text = "Monto ..: " & Format(curMonto, "Standard")

rs.Close

vGrid.MaxRows = vGrid.MaxRows - 1

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()

vModulo = 9
vGrid.AppearanceStyle = fxGridStyle

cbo.AddItem "Solicita"
cbo.AddItem "Autoriza"
cbo.AddItem "Emite"
cbo.AddItem "Anula"
cbo.Text = "Solicita"

cboFechas.AddItem "Emisión"
cboFechas.AddItem "Anulación"
cboFechas.AddItem "Solicitud"
cboFechas.AddItem "[Todas]"

cboFechas.Text = "Emisión"

cboEstado.Clear
cboEstado.AddItem "[Todos]"
cboEstado.AddItem "Solicitado"
cboEstado.AddItem "Emitido"
cboEstado.AddItem "Anulado"
cboEstado.Text = "[Todos]"

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value


lswBancos.ColumnHeaders.Clear
lswBancos.ColumnHeaders.Add , , "", lswBancos.Width - 250

lswConceptos.ColumnHeaders.Clear
lswConceptos.ColumnHeaders.Add , , "", lswConceptos.Width - 250


lswDocumentos.ColumnHeaders.Clear
lswDocumentos.ColumnHeaders.Add , , "", lswDocumentos.Width - 250

strSQL = "select TIPO,DESCRIPCION from TES_TIPOS_DOC"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswDocumentos.ListItems.Add(, , rs!DESCRIPCION)
     itmX.Tag = rs!Tipo
     itmX.Checked = chkDocumentos.Value
 rs.MoveNext
Loop
rs.Close


 strSQL = "select count(*) as 'Existe' from TES_AUTORIZACIONES" _
        & " Where ESTADO = 'A' and NOMBRE = '" & glogon.Usuario & "'"
 Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 1 Then
    chkModoProtegido.Enabled = True
 Else
    chkModoProtegido.Enabled = False
 End If


End Sub

Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - (vGrid.Left + 300)
vGrid.Height = Me.Height - (vGrid.top + StatusBarX.Height + 450)

imgBanner.Height = Me.Height


End Sub




Private Sub TimerX_Timer()
TimerX.Interval = 0

On Error GoTo vError

strSQL = "select COD_GRUPO as 'IdX', DESCRIPCION AS 'itmX'" _
       & "  From TES_BANCOS_GRUPOS" _
       & " Where ACTIVO = 1 order by DESCRIPCION"

vPaso = True
    Call sbCbo_Llena_New(cboBanco, strSQL, True, True)
vPaso = False

Call sbCuentas_Load
Call sbConceptos_Load

Exit Sub

vError:

End Sub

Private Sub sbCuentas_Load()

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

txtFiltroCta.Text = fxSysCleanTxtInject(txtFiltroCta.Text)

lswBancos.ListItems.Clear

strSQL = "select id_Banco as IdX, rtrim(Descripcion) as ItmX" _
       & " from Tes_Bancos where estado = 'A' and descripcion like '%" & txtFiltroCta.Text & "%'"
       
If cboBanco.Text <> "TODOS" Then
    strSQL = strSQL & " and Cod_Grupo = '" & cboBanco.ItemData(cboBanco.ListIndex) & "'"
End If

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswBancos.ListItems.Add(, , rs!itmX)
     itmX.Tag = rs!IdX
     itmX.Checked = chkBancos.Value
 rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbConceptos_Load()

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

txtFiltroConceptos.Text = fxSysCleanTxtInject(txtFiltroConceptos.Text)

lswConceptos.ListItems.Clear

strSQL = "select COD_CONCEPTO as IdX, rtrim(Descripcion) as ItmX" _
       & " from TES_CONCEPTOS where estado = 'A' and descripcion like '%" & txtFiltroConceptos.Text & "%'"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswConceptos.ListItems.Add(, , rs!itmX)
     itmX.Tag = rs!IdX
     itmX.Checked = chkConceptos.Value
 rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub txtFiltroConceptos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbConceptos_Load
End If
End Sub

Private Sub txtFiltroCta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbCuentas_Load
End If
End Sub

Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim frm As Form

If Row <= 0 Then Exit Sub
If vGrid.MaxRows <= 0 Then Exit Sub

vGrid.Row = Row
vGrid.col = 2

If vGrid.Text = "" Then Exit Sub

 Call sbFormsCall("frmTES_Transacciones")
 For Each frm In Forms
   If UCase(frm.Name) = UCase("frmTES_Transacciones") Then
     Call frm.sbTESDocConsulta(vGrid.Text)
     Exit For
   End If
 Next frm
 
End Sub

Private Sub vGrid_DblClick(ByVal col As Long, ByVal Row As Long)
Dim frm As Form

If Row <= 0 Then Exit Sub
If vGrid.MaxRows <= 0 Then Exit Sub

vGrid.Row = Row
vGrid.col = 2

If vGrid.Text = "" Then Exit Sub

 Call sbFormsCall("frmTES_Transacciones")
 For Each frm In Forms
   If UCase(frm.Name) = UCase("frmTES_Transacciones") Then
     Call frm.sbTESDocConsulta(vGrid.Text)
     Exit For
   End If
 Next frm

End Sub
