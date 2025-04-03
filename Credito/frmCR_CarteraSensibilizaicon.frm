VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCR_CarteraSensibilizacion 
   Caption         =   "Sensibilizacion de Cartera"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10245
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmCR_CarteraSensibilizaicon.frx":0000
   ScaleHeight     =   7425
   ScaleWidth      =   10245
   Begin VB.Frame fraFiltros 
      Caption         =   "Filtros Adicionales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   3840
      TabIndex        =   2
      Top             =   3600
      Width           =   5055
      Begin VB.CheckBox chkPlazos 
         Appearance      =   0  'Flat
         Caption         =   "Plazos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtPlazoInicio 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtPlazoCorte 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         MaxLength       =   4
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox chkTasasAdd 
         Appearance      =   0  'Flat
         Caption         =   "Tasas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtTasasInicio 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtTasasCorte 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   1080
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   4
         Left            =   2760
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   7
         Left            =   1080
         TabIndex        =   10
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   8
         Left            =   2760
         TabIndex        =   9
         Top             =   1560
         Width           =   975
      End
   End
   Begin MSComctlLib.ProgressBar PrgBarX 
      Height          =   135
      Left            =   1080
      TabIndex        =   40
      Top             =   3600
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtPtsAdd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8880
      TabIndex        =   38
      Text            =   "0"
      Top             =   1920
      Width           =   855
   End
   Begin VB.ComboBox cboInstitucion 
      Appearance      =   0  'Flat
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   2475
      Width           =   3975
   End
   Begin VB.TextBox txtCodigo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   21
      Top             =   1395
      Width           =   855
   End
   Begin VB.CheckBox chkLineas 
      Appearance      =   0  'Flat
      Caption         =   "Todas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6240
      TabIndex        =   20
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox cboRecurso 
      Appearance      =   0  'Flat
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2115
      Width           =   3975
   End
   Begin VB.ComboBox cboDestino 
      Appearance      =   0  'Flat
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1755
      Width           =   3975
   End
   Begin VB.CheckBox chkFechas 
      Appearance      =   0  'Flat
      Caption         =   "Todas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   17
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7800
      Picture         =   "frmCR_CarteraSensibilizaicon.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9000
      Picture         =   "frmCR_CarteraSensibilizaicon.frx":D0A4
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2400
      Width           =   855
   End
   Begin VB.CheckBox chkFiltrosAdd 
      Appearance      =   0  'Flat
      Caption         =   ">>> ver más filtros <<<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox txtTasa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      TabIndex        =   1
      Top             =   1920
      Width           =   855
   End
   Begin VB.CheckBox chkTBPPtsAdd 
      Appearance      =   0  'Flat
      Caption         =   "TBP + Pts Adicionales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7800
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   7170
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Casos"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Cuotas Actuales"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Nuevas Cuotas "
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   2160
      TabIndex        =   23
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   111345667
      CurrentDate     =   36278
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   4800
      TabIndex        =   24
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   111345667
      CurrentDate     =   36278
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3135
      Left            =   120
      TabIndex        =   25
      Top             =   3840
      Width           =   9855
      _Version        =   524288
      _ExtentX        =   17383
      _ExtentY        =   5530
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
      MaxCols         =   499
      SpreadDesigner  =   "frmCR_CarteraSensibilizaicon.frx":138F6
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pts. Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8880
      TabIndex        =   39
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Resultados"
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
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   37
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtros"
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
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   36
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Sensibilización de la Cartera de Crédito"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   35
      Top             =   240
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   9960
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   9960
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Institución"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   18
      Left            =   1320
      TabIndex        =   34
      Top             =   2475
      Width           =   855
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   33
      Top             =   1395
      Width           =   3975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Recurso"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   15
      Left            =   1320
      TabIndex        =   32
      Top             =   2115
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Destino"
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   13
      Left            =   1320
      TabIndex        =   31
      Top             =   1755
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Línea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   12
      Left            =   120
      TabIndex        =   30
      Top             =   1395
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   6
      Left            =   3840
      TabIndex        =   29
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   5
      Left            =   1320
      TabIndex        =   28
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Formalizadas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   27
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblTasa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tasa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8040
      TabIndex        =   26
      Top             =   1680
      Width           =   855
   End
End
Attribute VB_Name = "frmCR_CarteraSensibilizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkFechas_Click()

If chkFechas.Value = vbChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub

Private Sub chkFiltrosAdd_Click()
If chkFiltrosAdd.Value = vbChecked Then
   fraFiltros.Visible = True
Else
   fraFiltros.Visible = False
End If
End Sub

Private Sub chkLineas_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If chkLineas.Value = vbChecked Then
  
  txtCodigo.Enabled = False
  
  strSQL = "select cod_grupo + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  catalogo_grupos"
  Call sbLlenaCbo(cboRecurso, strSQL)
  
  strSQL = "select cod_destino + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  catalogo_destinos"
  Call sbLlenaCbo(cboDestino, strSQL)
  
Else
  txtCodigo.Enabled = True

  strSQL = "select (R.cod_grupo) + ' - ' + rtrim(R.descripcion) as ItmX" _
         & " from catalogo_grupos R inner join catalogo_AsignaGrp A on R.cod_grupo = A.cod_grupo" _
         & " where A.codigo = '" & txtCodigo & "'"
  Call sbLlenaCbo(cboRecurso, strSQL)
  
  strSQL = "select (R.cod_destino) + ' - ' + rtrim(R.descripcion) as ItmX" _
         & " from catalogo_destinos R inner join catalogo_destinosAsg A on R.cod_destino = A.cod_destino" _
         & " where A.codigo = '" & txtCodigo & "'"
  Call sbLlenaCbo(cboDestino, strSQL)

End If

End Sub



Private Sub chkPlazos_Click()

If chkPlazos.Value = vbChecked Then
  txtPlazoInicio.Enabled = True
  txtPlazoInicio.BackColor = vbWhite
  txtPlazoInicio.SetFocus
Else
  txtPlazoInicio.Enabled = False
  txtPlazoInicio.BackColor = lblTasa.BackColor
End If

txtPlazoCorte.Enabled = txtPlazoInicio.Enabled
txtPlazoCorte.BackColor = txtPlazoInicio.BackColor

End Sub

Private Sub chkTasasAdd_Click()

If chkTasasAdd.Value = vbChecked Then
   txtTasasInicio.Enabled = True
   txtTasasInicio.BackColor = vbWhite
   txtTasasInicio.SetFocus
Else
   txtTasasInicio.Enabled = False
End If

txtTasasCorte.Enabled = txtTasasInicio.Enabled
txtTasasCorte.BackColor = txtTasasInicio.BackColor

End Sub

Private Sub chkTBPPtsAdd_Click()
Dim strSQL As String, rs As New ADODB.Recordset

vGrid.MaxRows = 0

If chkTBPPtsAdd.Value = vbChecked Then
    txtTasa.Enabled = False
    strSQL = "select cr_tbp from par_ahcr"
    Call OpenRecordSet(rs, strSQL)
      txtTasa.Text = rs!cr_tbp
    rs.Close
Else
    txtTasa.Enabled = True
End If

End Sub

Private Sub CmdAplicar_Click()
Dim strSQL As String, i As Long, vTasa As Currency
Dim vOperacion As Long, vCodigo As String, vCuota As Currency


i = MsgBox("Esta seguro que desea cambiar la Tasa a las operaciones mostradas?", vbYesNo)
If i = vbNo Then
  Exit Sub
End If

Me.MousePointer = vbHourglass

On Error GoTo vError


For i = 1 To vGrid.MaxRows
 vGrid.Row = i
 vGrid.col = 7 'Verifica que la tasa cambie
 If CCur(vGrid.Text) <> CCur(txtTasa) Then
    vGrid.col = 1
    vOperacion = vGrid.Text
    vCodigo = vGrid.CellTag
    
    vGrid.col = 9
    vCuota = CCur(vGrid.Text)
    
    vGrid.col = 12
    vTasa = CCur(vGrid.Text)
    
    strSQL = "update reg_creditos set interesv = " & vTasa _
           & ",cuota = " & vCuota _
           & " where id_solicitud = " & vOperacion
    Call ConectionExecute(strSQL)
    
    vGrid.col = 7
    Call sbBitacoraCredito("02", ("De: " & vGrid.Text & " A: " & vTasa), "C", vOperacion, vCodigo)
    
 End If
Next i

Call Bitacora("Aplica", "Cambio de Tasa a Linea : " & txtCodigo & " a " & txtTasa & "%")

Me.MousePointer = vbDefault

MsgBox "Cambio de Tasas Realizado Satisfactoriamente...", vbInformation

cmdBuscar_Click



Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lngCasos As Long, curCuotas As Currency, curCuotasNew As Currency
Dim i As Byte, vTasa As Currency

'If txtTasa = "" Or txtTasa = "0" Or Not IsNumeric(txtTasa) Then
'   MsgBox "Indique la nueva Tasa ?", vbInformation
'   Exit Sub
'End If
'
'If CCur(txtTasa) < 0 Or CCur(txtTasa) > 100 Then
'   MsgBox "La Tasa Indicada no es válida, verifique...", vbExclamation
'   Exit Sub
'End If


If txtPtsAdd.Text = "" Then
   MsgBox "Indique los puntos de variación en la tasa...", vbExclamation
   Exit Sub
End If

Me.MousePointer = vbHourglass

lngCasos = 0
curCuotas = 0
curCuotasNew = 0

strSQL = "select Top 100 R.id_solicitud,R.cedula,S.nombre,R.montoapr,R.Saldo - isnull(V.amortiza,0) as Saldo" _
       & ",R.cuota,R.plazo,R.interesv,R.prideduc,R.codigo,R.fechaforp,R.int as TasaOriginal, C.Liq_Valor" _
       & ",R.plazo + DATEDIFF(mm,  dbo.MyGetdate(), CONVERT(DATETIME, substring(convert(varchar(6), R.prideduc), 1,4) + '/' + substring(convert(varchar(6), R.prideduc), 5,2) + '/28' )) as PlazoFaltante" _
       & ",isnull(R.liqTasa,0) as LiqTasa, isnull(R.TBP_PuntosAdd,0) as TBPPuntosAdd, isnull(R.Tasa_Piso,0) as Tasa_Piso" _
       & " from socios S inner join reg_creditos R on S.cedula = R.cedula" _
       & " inner join catalogo C on R.codigo = C.codigo" _
       & " left join vista_morosidad V on R.id_solicitud = V.id_solicitud" _
       & " where R.estado = 'A' and R.proceso = 'N' and R.saldo > 0"
       
If chkLineas.Value = vbUnchecked Then
   strSQL = strSQL & " and R.codigo = '" & txtCodigo & "'"
End If
       
If cboDestino.Text <> "TODOS" Then
   strSQL = strSQL & " and R.cod_destino = '" & fxCodigoCbo(cboDestino) & "'"
End If
       
If cboRecurso.Text <> "TODOS" Then
   strSQL = strSQL & " and R.cod_grupo = '" & fxCodigoCbo(cboRecurso) & "'"
End If
       
If cboInstitucion.Text <> "TODOS" Then
   strSQL = strSQL & " and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If
       
If chkPlazos.Value = vbChecked Then
   strSQL = strSQL & " and R.plazo between " & txtPlazoInicio & " and " & txtPlazoCorte
End If
       
If chkTasasAdd.Value = vbChecked Then
   strSQL = strSQL & " and R.interesv between " & txtTasasInicio & " and " & txtTasasCorte
End If
       
If chkFechas.Value = vbUnchecked Then
  strSQL = strSQL & " and R.fechaforp between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00" _
         & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End If

If chkTBPPtsAdd.Value = vbChecked Then
   strSQL = strSQL & " and R.TBP_PuntosAdd is not null"
Else
   strSQL = strSQL & " and R.TBP_PuntosAdd is null"
End If
       
Call OpenRecordSet(rs, strSQL)
       
       
prgBarX.Value = 1
prgBarX.Max = rs.RecordCount + 1
prgBarX.Visible = True

vGrid.Sheet = 1
vGrid.MaxRows = 0
       
Do While Not rs.EOF
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  
  txtTasa.Text = rs!interesv
  
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
       Case 1
          vGrid.Text = CStr(rs!Id_Solicitud)
          vGrid.CellTag = CStr(rs!Codigo)
       Case 2
          vGrid.Text = CStr(rs!Cedula)
       Case 3
          vGrid.Text = CStr(rs!Nombre)
       Case 4
          vGrid.Text = CStr(rs!montoapr)
       Case 5
          vGrid.Text = CStr(rs!Saldo)
       Case 6
          vGrid.Text = CStr(rs!Plazo)
       Case 7
          vGrid.Text = CStr(rs!interesv)
       Case 8
          vGrid.Text = CStr(rs!Cuota)
       Case 9 'Cuota 1
          
        
          vTasa = CCur(txtTasa) + (CCur(txtPtsAdd.Text) * 1)
       
          If vTasa < rs!Tasa_Piso Then vTasa = rs!Tasa_Piso
          vGrid.Text = CStr(fxCalcula_Cuota(rs!Saldo, rs!PlazoFaltante, vTasa))
       
       
       
       Case 10 'Cuota 2
          
'          If rs!LiqTasa = 0 Then
'              vTasa = CCur(txtTasa) + CCur(IIf((txtPtsAdd.Text = ""), 0, txtPtsAdd.Text)) + rs!TBPPuntosAdd
'          Else
'              vTasa = CCur(txtTasa) + CCur(IIf((txtPtsAdd.Text = ""), 0, txtPtsAdd.Text)) + rs!Liq_Valor + rs!TBPPuntosAdd
'          End If
          
          vTasa = CCur(txtTasa) + (CCur(txtPtsAdd.Text) * 2)
                      
          If vTasa < rs!Tasa_Piso Then vTasa = rs!Tasa_Piso
          vGrid.Text = CStr(fxCalcula_Cuota(rs!Saldo, rs!PlazoFaltante, vTasa))
       
       
       Case 11 'Cuota 3
          
'          If rs!LiqTasa = 0 Then
'              vTasa = CCur(txtTasa) + CCur(IIf((txtPtsAdd.Text = ""), 0, txtPtsAdd.Text)) + rs!TBPPuntosAdd
'          Else
'              vTasa = CCur(txtTasa) + CCur(IIf((txtPtsAdd.Text = ""), 0, txtPtsAdd.Text)) + rs!Liq_Valor + rs!TBPPuntosAdd
'          End If
          
          vTasa = CCur(txtTasa) + (CCur(txtPtsAdd.Text) * 3)
                      
          If vTasa < rs!Tasa_Piso Then vTasa = rs!Tasa_Piso
          vGrid.Text = CStr(fxCalcula_Cuota(rs!Saldo, rs!PlazoFaltante, vTasa))
       
       Case 12 'Cuota 4
          
'          If rs!LiqTasa = 0 Then
'              vTasa = CCur(txtTasa) + CCur(IIf((txtPtsAdd.Text = ""), 0, txtPtsAdd.Text)) + rs!TBPPuntosAdd
'          Else
'              vTasa = CCur(txtTasa) + CCur(IIf((txtPtsAdd.Text = ""), 0, txtPtsAdd.Text)) + rs!Liq_Valor + rs!TBPPuntosAdd
'          End If
          
          vTasa = CCur(txtTasa) + (CCur(txtPtsAdd.Text) * 4)
                      
          If vTasa < rs!Tasa_Piso Then vTasa = rs!Tasa_Piso
          vGrid.Text = CStr(fxCalcula_Cuota(rs!Saldo, rs!PlazoFaltante, vTasa))
       
       
       
       Case 13
          vGrid.Text = CStr(rs!FechaForp)
       Case 14
          vGrid.Text = CStr(rs!TasaOriginal)
          vGrid.TextTip = TextTipFixed
          vGrid.TextTipDelay = 1000
          vGrid.CellNote = "Puntos Adicionales a la Tasa Basica Pasiva : " & CStr(rs!TBPPuntosAdd)
          
       Case 15 'Tasa 1
          
'          If rs!LiqTasa = 0 Then
'              vTasa = CCur(txtTasa) + CCur(IIf((txtPtsAdd.Text = ""), 0, txtPtsAdd.Text)) + rs!TBPPuntosAdd
'          Else
'              vTasa = CCur(txtTasa) + CCur(IIf((txtPtsAdd.Text = ""), 0, txtPtsAdd.Text)) + rs!Liq_Valor + rs!TBPPuntosAdd
'          End If
          vTasa = CCur(txtTasa) + (CCur(txtPtsAdd.Text) * 1)
          If vTasa < rs!Tasa_Piso Then vTasa = rs!Tasa_Piso
          vGrid.Text = CStr(vTasa)
          
          
       Case 16 'Tasa 2
          
'          If rs!LiqTasa = 0 Then
'              vTasa = CCur(txtTasa) + CCur(IIf((txtPtsAdd.Text = ""), 0, txtPtsAdd.Text)) + rs!TBPPuntosAdd
'          Else
'              vTasa = CCur(txtTasa) + CCur(IIf((txtPtsAdd.Text = ""), 0, txtPtsAdd.Text)) + rs!Liq_Valor + rs!TBPPuntosAdd
'          End If
                      
          vTasa = CCur(txtTasa) + (CCur(txtPtsAdd.Text) * 2)
                      
          If vTasa < rs!Tasa_Piso Then vTasa = rs!Tasa_Piso
          vGrid.Text = CStr(vTasa)
          
       Case 17 'Tasa 3
          
'          If rs!LiqTasa = 0 Then
'              vTasa = CCur(txtTasa) + CCur(IIf((txtPtsAdd.Text = ""), 0, txtPtsAdd.Text)) + rs!TBPPuntosAdd
'          Else
'              vTasa = CCur(txtTasa) + CCur(IIf((txtPtsAdd.Text = ""), 0, txtPtsAdd.Text)) + rs!Liq_Valor + rs!TBPPuntosAdd
'          End If
                      
          vTasa = CCur(txtTasa) + (CCur(txtPtsAdd.Text) * 3)
                      
          If vTasa < rs!Tasa_Piso Then vTasa = rs!Tasa_Piso
          vGrid.Text = CStr(vTasa)
          
          
       Case 18 'Tasa 4
          
'          If rs!LiqTasa = 0 Then
'              vTasa = CCur(txtTasa) + CCur(IIf((txtPtsAdd.Text = ""), 0, txtPtsAdd.Text)) + rs!TBPPuntosAdd
'          Else
'              vTasa = CCur(txtTasa) + CCur(IIf((txtPtsAdd.Text = ""), 0, txtPtsAdd.Text)) + rs!Liq_Valor + rs!TBPPuntosAdd
'          End If
                      
          vTasa = CCur(txtTasa) + (CCur(txtPtsAdd.Text) * 4)
                      
          If vTasa < rs!Tasa_Piso Then vTasa = rs!Tasa_Piso
          vGrid.Text = CStr(vTasa)
          
          
          
    End Select
  Next i
   
  lngCasos = lngCasos + 1
  curCuotas = curCuotas + rs!Cuota
  
  vGrid.col = 9
  curCuotasNew = curCuotasNew + CCur(vGrid.Text)
  If prgBarX.Max > prgBarX.Value Then
      prgBarX.Value = prgBarX.Value + 1
  End If
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
prgBarX.Visible = False



stBar.Panels(1) = Format(lngCasos, "###,###,###")
stBar.Panels(2) = Format(curCuotas, "Standard")
stBar.Panels(3) = Format(curCuotasNew, "Standard")

End Sub

Private Sub cmdGenerar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Long

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = "delete CRD_SENSIBILIZA_PF"
Call ConectionExecute(strSQL)

strSQL = "delete CRD_SENSIBILIZA_PL"
Call ConectionExecute(strSQL)

prgBarX.Visible = True
prgBarX.Value = 1
prgBarX.Max = vGrid.MaxRows

With vGrid
  .Sheet = 1
  For i = 1 To .MaxRows
    .Row = i
    .col = 1
    
    strSQL = "insert CRD_SENSIBILIZA_PF(id_solicitud,cuota_01,cuota_02,cuota_03,cuota_04,tasa_01,tasa_02,tasa_03,tasa_04) values(" _
           & .Text & ","
    .col = 9
    strSQL = strSQL & CCur(.Text) & ","
    .col = 10
    strSQL = strSQL & CCur(.Text) & ","
    .col = 11
    strSQL = strSQL & CCur(.Text) & ","
    .col = 12
    strSQL = strSQL & CCur(.Text) & ","
     
     
    .col = 15
    strSQL = strSQL & CCur(.Text) & ","
    .col = 16
    strSQL = strSQL & CCur(.Text) & ","
    .col = 17
    strSQL = strSQL & CCur(.Text) & ","
    .col = 18
    strSQL = strSQL & CCur(.Text) & ")"
     
    Call ConectionExecute(strSQL)
    prgBarX.Value = .Row
    
  Next i

  'Creando Resultados
  
  .Sheet = 2
  .MaxRows = 1


'SELECT R.CEDULA,SUM(CUOTA_01) AS CUOTA01,SUM(CUOTA_02) AS CUOTA02
' ,SUM(CUOTA_03) AS CUOTA03,SUM(CUOTA_04) AS CUOTA04
' ,L.DEVENGADO_MES,L.LIQUIDEZ_SIMPLE,L.LIQUIDEZ_CONFIANZA,L.TOTAL_CARGA_CCSS
' ,L.DEDUCCIONES - (L.REFUNDICIONES_CUOTA + L.DESEMBOLSOS_CUOTA + CRD_TRANSITO_CANCELADOS) AS DEDUCCIONES
' ,dbo.fxCRD_Sensilidad(R.cedula,1,'F') as SaldoFijo, dbo.fxCRD_Sensilidad(R.cedula,0,'F') as CuotaFija
' Into tmpResultadoSen
'FROM CRD_SENSIBILIZA_PF S INNER JOIN REG_CREDITOS R ON S.ID_SOLICITUD = R.ID_SOLICITUD
'  INNER JOIN CRD_SENSIBILIZA_LIQ L ON R.CEDULA = L.CEDULA
'GROUP BY R.CEDULA ,L.DEVENGADO_MES,L.LIQUIDEZ_SIMPLE,L.LIQUIDEZ_CONFIANZA,L.TOTAL_CARGA_CCSS
' ,L.DEDUCCIONES,L.REFUNDICIONES_CUOTA,L.DESEMBOLSOS_CUOTA,CRD_TRANSITO_CANCELADOS
' ,R.TBP_PUNTOSADD,R.CUOTA
  
'
'  strSQL = ""
'
'
'select sum(case when TBP_PUNTOSADD is null then Cuota else 0 end) as CuotaR
'      ,sum(case when TBP_PUNTOSADD is null then Saldo else 0 end) as SaldoR
'      ,sum(case when TBP_PUNTOSADD is not null then Cuota else 0 end) as CuotaI
'      ,sum(case when TBP_PUNTOSADD is not null then Saldo else 0 end) as SaldoI
'From reg_creditos
'where cedula = '601020080'
' and estado = 'A'





End With











Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

vGrid.MaxCols = 18
vGrid.MaxRows = 0

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
chkFechas.Value = vbUnchecked
chkLineas.Value = vbUnchecked

cboInstitucion.Clear

strSQL = "select cod_institucion,descripcion from instituciones"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
   cboInstitucion.AddItem Trim(rs!Descripcion)
   cboInstitucion.ItemData(cboInstitucion.NewIndex) = rs!cod_institucion
   rs.MoveNext
Loop
cboInstitucion.AddItem "TODOS"
cboInstitucion.Text = "TODOS"
rs.Close


Call chkFechas_Click
Call chkLineas_Click
Call chkFiltrosAdd_Click

Me.MousePointer = vbDefault

vModulo = 3

Call Formularios(Me)
Call RefrescaTags(Me)

Me.Width = 10275
Me.Height = 7755

End Sub

Private Sub Form_Resize()
On Error Resume Next


Line1.Item(0).X2 = Me.Width - 250
Line1.Item(1).X2 = Me.Width - 250

vGrid.Width = Me.Width - 350
vGrid.Height = Me.Height - (vGrid.Top + 800)

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then cboDestino.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Resultado = ""
  gBusquedas.Consulta = "select codigo,descripcion from catalogo"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Columna = "descripcion"
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  lblDescripcion.Caption = gBusquedas.Resultado2
End If

End Sub


Private Sub txtCodigo_LostFocus()
 If Len(Trim(txtCodigo)) > 0 Then lblDescripcion.Caption = fxDescribeCodigo(Trim(txtCodigo))
 Call chkLineas_Click
End Sub

Private Sub txtTasa_Change()
' vGrid.MaxRows = 0
End Sub


