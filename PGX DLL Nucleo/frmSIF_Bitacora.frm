VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmSIF_Bitacora 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Bitácora "
   ClientHeight    =   8580
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   17010
   HelpContextID   =   9003
   Icon            =   "frmSIF_Bitacora.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   17010
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   312
      Left            =   4080
      TabIndex        =   28
      Top             =   1920
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Todas las Fechas"
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
   Begin XtremeSuiteControls.PushButton btnConsulta 
      Height          =   375
      Left            =   13800
      TabIndex        =   15
      Top             =   2280
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Consulta"
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
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmSIF_Bitacora.frx":000C
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1455
      Left            =   4920
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   7695
      _Version        =   1441793
      _ExtentX        =   13568
      _ExtentY        =   2561
      _StockProps     =   79
      Caption         =   "Cargando.... [Espere]"
      BackColor       =   16777215
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.ProgressBar prgBarLoad 
         Height          =   132
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   7332
         _Version        =   1441793
         _ExtentX        =   12933
         _ExtentY        =   233
         _StockProps     =   93
         BackColor       =   -2147483633
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label lblLoad 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   4332
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5655
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   16695
      _Version        =   524288
      _ExtentX        =   29448
      _ExtentY        =   9975
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
      SpreadDesigner  =   "frmSIF_Bitacora.frx":070C
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboModulo 
      Height          =   312
      Left            =   1320
      TabIndex        =   16
      Top             =   1080
      Width           =   4452
      _Version        =   1441793
      _ExtentX        =   7858
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
   Begin XtremeSuiteControls.ComboBox cboMovimientos 
      Height          =   312
      Left            =   1320
      TabIndex        =   17
      Top             =   1440
      Width           =   4452
      _Version        =   1441793
      _ExtentX        =   7858
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
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   7200
      TabIndex        =   18
      Top             =   1080
      Width           =   6132
      _Version        =   1441793
      _ExtentX        =   10816
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
   Begin XtremeSuiteControls.FlatEdit txtDetalle 
      Height          =   312
      Left            =   7200
      TabIndex        =   19
      Top             =   1440
      Width           =   6132
      _Version        =   1441793
      _ExtentX        =   10816
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAppNombre 
      Height          =   312
      Left            =   7200
      TabIndex        =   20
      Top             =   1920
      Width           =   2292
      _Version        =   1441793
      _ExtentX        =   4043
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtLogIP 
      Height          =   315
      Left            =   13800
      TabIndex        =   21
      Top             =   1920
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4043
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtLogEquipo 
      Height          =   312
      Left            =   11040
      TabIndex        =   22
      Top             =   1920
      Width           =   2292
      _Version        =   1441793
      _ExtentX        =   4043
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAppVersion 
      Height          =   312
      Left            =   7200
      TabIndex        =   23
      Top             =   2280
      Width           =   2292
      _Version        =   1441793
      _ExtentX        =   4043
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1320
      TabIndex        =   24
      Top             =   1920
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
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
      Left            =   1320
      TabIndex        =   25
      Top             =   2280
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicioTime 
      Height          =   312
      Left            =   2640
      TabIndex        =   26
      Top             =   1920
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
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
      Format          =   2
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorteTime 
      Height          =   312
      Left            =   2640
      TabIndex        =   27
      Top             =   2280
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
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
      Format          =   2
   End
   Begin XtremeSuiteControls.CheckBox chkHoras 
      Height          =   312
      Left            =   4080
      TabIndex        =   29
      Top             =   2280
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Todos los Horarios"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.PushButton btnExporta 
      Height          =   375
      Left            =   14880
      TabIndex        =   30
      Top             =   2280
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Exporta"
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
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmSIF_Bitacora.frx":0F62
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.FlatEdit txtMAC 
      Height          =   315
      Left            =   11040
      TabIndex        =   31
      Top             =   2280
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4043
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Equipo MAC"
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
      Height          =   255
      Index           =   10
      Left            =   9840
      TabIndex        =   32
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Log. IP "
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
      Height          =   255
      Index           =   9
      Left            =   13800
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Equipo"
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
      Height          =   255
      Index           =   8
      Left            =   9840
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "App. Versión"
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
      Height          =   255
      Index           =   7
      Left            =   6000
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "App. Nombre"
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
      Height          =   255
      Index           =   6
      Left            =   6000
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle"
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
      Height          =   255
      Index           =   5
      Left            =   6000
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitácora del Sistema"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   1920
      TabIndex        =   5
      Top             =   360
      Width           =   4452
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
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
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Módulo"
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
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Movimiento"
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
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
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
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
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
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   13440
      X2              =   0
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Image imgBanner 
      Height          =   990
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13560
   End
End
Attribute VB_Name = "frmSIF_Bitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vHeaders As vGridHeaders, vTitulo As String, vEmpresa As String

Private Sub btnConsulta_Click()
Call sbConsulta
End Sub

Private Sub btnExporta_Click()

'Variables del Exporte
vHeaders.Columnas = vGrid.MaxCols
vTitulo = "ProGrX_Bitacora_" & vEmpresa
    
    vHeaders.Headers(1) = "Módulo"
    vHeaders.Headers(2) = "Nombre"
    vHeaders.Headers(3) = "Usuario"
    vHeaders.Headers(4) = "Movimiento"
    vHeaders.Headers(5) = "Fecha/Hora"
    vHeaders.Headers(6) = "Detalle"
    vHeaders.Headers(7) = "App Nombre"
    vHeaders.Headers(8) = "App Versión"
    vHeaders.Headers(9) = "Estación"
    vHeaders.Headers(10) = "Estación MAC"
    vHeaders.Headers(11) = "LOG IP"
    
   Call sbSIFGridExportar(vGrid, vHeaders, vTitulo)


End Sub

Private Sub cboModulo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboMovimientos.SetFocus
End Sub

Private Sub cboMovimientos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUsuario.SetFocus
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

Private Sub chkHoras_Click()
If chkHoras.Value = vbChecked Then
  dtpCorteTime.Enabled = False
  dtpInicioTime.Enabled = False
Else
  dtpCorteTime.Enabled = True
  dtpInicioTime.Enabled = True
End If
End Sub

Private Sub dtpCorte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboModulo.SetFocus
End Sub

Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpCorte.SetFocus
End Sub

Private Sub Form_Activate()
vModulo = 10
End Sub


Private Sub Form_Load()

vModulo = 10

vGrid.MaxRows = 0

Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

dtpInicioTime.Value = dtpInicio.Value
dtpCorteTime.Value = dtpInicioTime.Value


strSQL = "select PAG_NOMCORTO from SIF_EMPRESA"
Call OpenRecordSet(rs, strSQL)
    vEmpresa = Trim(rs!pag_nomCorto & "")
rs.Close


With cboModulo
  .Clear
  
  strSQL = "exec spSEG_Modulos_Consulta"
  Call OpenRecordSet(rs, strSQL, 1)
  Do While Not rs.EOF
   .AddItem Trim(rs!Nombre)
   .ItemData(.ListCount - 1) = CStr(rs!modulo)
   rs.MoveNext
  Loop
  rs.Close
  
  .AddItem "[TODOS]"
  .Text = "[TODOS]"
End With

With cboMovimientos
  .Clear
  .AddItem "Registra"
  .AddItem "Modifica"
  .AddItem "Borra"
  .AddItem "Reversa"
  .AddItem "Aplica"
  .AddItem "Genera"
  .AddItem "Carga"
  .AddItem "Anula"
  .AddItem "Imprime"
  .AddItem "Consulta"
  .AddItem "Exporta"
  .AddItem "Emite"
  .AddItem "[TODOS]"
  
  .Text = "[TODOS]"
End With

Call Formularios(Me)
Call RefrescaTags(Me)

Call chkFechas_Click
Call chkHoras_Click

End Sub

Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - 350
vGrid.Height = Me.Height - (vGrid.Top + 800)

imgBanner.Width = Me.Width

End Sub

Private Sub sbConsulta()
Dim vSubTitulo As String

If dtpInicio.Value > dtpCorte.Value Then
   MsgBox "Verifique El Rango De Fechas", vbExclamation
   Exit Sub
End If

If dtpInicioTime.Value > dtpCorteTime.Value Then
   MsgBox "Verifique El Rango De Horas", vbExclamation
   Exit Sub
End If


On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spSEG_Bitacora_Consulta " & gPortal.Empresa_Id & ",'"
If chkFechas.Value = vbUnchecked Then
   If chkHoras.Value = vbUnchecked Then
      strSQL = strSQL & Format(dtpInicio.Value, "yyyy/mm/dd") & " " & Format(dtpInicioTime.Value, "HH:MM:SS") & "','" _
             & Format(dtpCorte.Value, "yyyy/mm/dd") & " " & Format(dtpCorteTime.Value, "HH:MM:SS") & "'"
   Else
      strSQL = strSQL & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00','" _
             & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
   
   End If

Else
      strSQL = strSQL & "1900/01/01 00:00:00','2100/12/30 23:59:59'"
End If

If txtUsuario.Text = "" Then
    strSQL = strSQL & ",Null"
Else
    strSQL = strSQL & ",'" & Trim(txtUsuario.Text) & "'"
End If

If cboModulo.Text = "[TODOS]" Then
    strSQL = strSQL & ",Null"
Else
    strSQL = strSQL & "," & cboModulo.ItemData(cboModulo.ListIndex)
End If

If cboMovimientos.Text = "[TODOS]" Then
    strSQL = strSQL & ",Null"
Else
    strSQL = strSQL & ",'" & Trim(cboMovimientos.Text) & "'"
End If

If txtDetalle.Text = "" Then
    strSQL = strSQL & ",Null"
Else
    strSQL = strSQL & ",'" & Trim(txtDetalle.Text) & "'"
End If

If txtAppNombre.Text = "" Then
    strSQL = strSQL & ",Null"
Else
    strSQL = strSQL & ",'" & Trim(txtAppNombre.Text) & "'"
End If

If txtAppVersion.Text = "" Then
    strSQL = strSQL & ",Null"
Else
    strSQL = strSQL & ",'" & Trim(txtAppVersion.Text) & "'"
End If

If txtLogEquipo.Text = "" Then
    strSQL = strSQL & ",Null"
Else
    strSQL = strSQL & ",'" & Trim(txtLogEquipo.Text) & "'"
End If

If txtLogIP.Text = "" Then
    strSQL = strSQL & ",Null"
Else
    strSQL = strSQL & ",'" & Trim(txtLogIP.Text) & "'"
End If


If txtMAC.Text = "" Then
    strSQL = strSQL & ",Null"
Else
    strSQL = strSQL & ",'" & Trim(txtMAC.Text) & "'"
End If

GroupBox1.Visible = True

Call OpenRecordSet(rs, strSQL, 1)

prgBarLoad.Max = rs.RecordCount + 1

With vGrid
  .MaxRows = 0
  .MaxCols = 11
  
  
  
  Do While Not rs.EOF
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows
    
    lblLoad.Caption = "Cargando Registro [ " & .MaxRows & " / " & prgBarLoad.Max & " ]"
    lblLoad.Refresh
    DoEvents
    
    
    .Col = 1
    .Text = rs!ModuloDesc & ""
    .Col = 2
    .Text = rs!UsuarioNombre & ""

    .Col = 3
    .Text = rs!Usuario & ""

    .Col = 4
    .Text = rs!Movimiento & ""
    .Col = 5
    .Text = rs!Fecha_FORMAT & ""
    .Col = 6
    .Text = rs!Detalle & ""
    .Col = 7
    .Text = rs!App_Nombre & ""
    .Col = 8
    .Text = rs!App_Version & ""
    .Col = 9
    .Text = rs!App_Equipo & ""
    .Col = 10
    .Text = rs!Equipo_MAC & ""
    
    .Col = 11
    .Text = rs!App_IP & ""
  
    
    rs.MoveNext
    prgBarLoad.Value = .MaxRows

  Loop

End With

rs.Close

GroupBox1.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInicio.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = Trim(txtUsuario)
    gBusquedas.Consulta = "Select Nombre,Descripcion From Usuarios"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    frmBusquedas.Show vbModal
    txtUsuario = Trim(gBusquedas.Resultado)
End If

vError:

End Sub


