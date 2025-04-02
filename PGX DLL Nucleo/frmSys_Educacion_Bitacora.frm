VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmSys_Educacion_Bitacora 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Bitácora de Cartas a Universidades"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17040
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   17040
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   375
      Left            =   6600
      TabIndex        =   27
      Top             =   1320
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Todas las Fecha"
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
      Appearance      =   17
      Alignment       =   1
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   14040
      Top             =   240
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   495
      Left            =   5520
      TabIndex        =   25
      Top             =   1800
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmSys_Educacion_Bitacora.frx":0000
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   16815
      _Version        =   524288
      _ExtentX        =   29660
      _ExtentY        =   12938
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
      MaxCols         =   13
      ScrollBars      =   2
      SpreadDesigner  =   "frmSys_Educacion_Bitacora.frx":0700
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   330
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   240
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
         Weight          =   700
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
      Left            =   1800
      TabIndex        =   4
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   600
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
      Left            =   3720
      TabIndex        =   5
      Top             =   240
      Width           =   4335
      _Version        =   1441793
      _ExtentX        =   7646
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtBeneficiario 
      Height          =   330
      Left            =   3720
      TabIndex        =   6
      Top             =   600
      Width           =   4335
      _Version        =   1441793
      _ExtentX        =   7646
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   330
      Left            =   1800
      TabIndex        =   8
      Top             =   960
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboUniversidad 
      Height          =   330
      Left            =   9960
      TabIndex        =   10
      Top             =   240
      Width           =   3975
      _Version        =   1441793
      _ExtentX        =   7011
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
   Begin XtremeSuiteControls.ComboBox cboNivel 
      Height          =   330
      Left            =   9960
      TabIndex        =   12
      Top             =   600
      Width           =   3975
      _Version        =   1441793
      _ExtentX        =   7011
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
   Begin XtremeSuiteControls.ComboBox cboCarrera 
      Height          =   330
      Left            =   9960
      TabIndex        =   14
      Top             =   960
      Width           =   3975
      _Version        =   1441793
      _ExtentX        =   7011
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
   Begin XtremeSuiteControls.ComboBox cboEspecialidad 
      Height          =   330
      Left            =   9960
      TabIndex        =   16
      Top             =   1320
      Width           =   3975
      _Version        =   1441793
      _ExtentX        =   7011
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
   Begin XtremeSuiteControls.ComboBox cboCiclo 
      Height          =   330
      Left            =   1800
      TabIndex        =   18
      Top             =   1440
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
   Begin XtremeSuiteControls.FlatEdit txtAnio_Inicio 
      Height          =   330
      Left            =   1800
      TabIndex        =   20
      Top             =   1800
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
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
   Begin XtremeSuiteControls.FlatEdit txtAnio_Corte 
      Height          =   330
      Left            =   2760
      TabIndex        =   21
      Top             =   1800
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   330
      Left            =   5400
      TabIndex        =   23
      Top             =   960
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   330
      Left            =   6720
      TabIndex        =   24
      Top             =   960
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
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   495
      Left            =   6840
      TabIndex        =   26
      Top             =   1800
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Picture         =   "frmSys_Educacion_Bitacora.frx":0E90
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   9
      Left            =   3840
      TabIndex        =   22
      Top             =   960
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Fecha Registro"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   19
      Top             =   1800
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ciclo Año"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   17
      Top             =   1440
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ciclo"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   6
      Left            =   8640
      TabIndex        =   15
      Top             =   1320
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Especialidad"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   5
      Left            =   8640
      TabIndex        =   13
      Top             =   960
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Carrera"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   4
      Left            =   8640
      TabIndex        =   11
      Top             =   600
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Nivel Acade."
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   8640
      TabIndex        =   9
      Top             =   240
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Universidad"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cliente"
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
End
Attribute VB_Name = "frmSys_Educacion_Bitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean


Private Sub sbInicializa()

On Error GoTo vError

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("m", -3, dtpCorte.Value)

cboCiclo.Clear
cboCiclo.AddItem "TODOS"
cboCiclo.AddItem "I   Quatrimestre"
cboCiclo.AddItem "II  Quatrimestre"
cboCiclo.AddItem "III Quatrimestre"
cboCiclo.AddItem "IV  Quatrimestre"

cboCiclo.AddItem "I   Semestre"
cboCiclo.AddItem "II  Semestre"

cboCiclo.Text = "TODOS"

txtAnio_Corte.Text = Year(dtpCorte.Value)
txtAnio_Inicio.Text = Year(dtpCorte.Value) - 2


vPaso = True
    
strSQL = "exec spSys_Educacion_List 'U', Null"
Call sbCbo_Llena_New(cboUniversidad, strSQL, True, True)

strSQL = "exec spSys_Educacion_List 'N', ''"
Call sbCbo_Llena_New(cboNivel, strSQL, True, True)

strSQL = "exec spSys_Educacion_List 'C', ''"
Call sbCbo_Llena_New(cboCarrera, strSQL, True, True)

vPaso = False

Call cboCarrera_Click


Exit Sub

vError:



End Sub

Private Sub btnBuscar_Click()
Dim pWhere As Boolean

On Error GoTo vError

Me.MousePointer = vbHourglass

txtUsuario.Text = fxSysCleanTxtInject(txtUsuario.Text)
txtCedula.Text = fxSysCleanTxtInject(txtCedula.Text)
txtNombre.Text = fxSysCleanTxtInject(txtNombre.Text)
txtBeneficiarioId.Text = fxSysCleanTxtInject(txtBeneficiarioId.Text)
txtBeneficiario.Text = fxSysCleanTxtInject(txtBeneficiario.Text)

txtAnio_Inicio.Text = fxSysCleanTxtInject(txtAnio_Inicio.Text)
txtAnio_Corte.Text = fxSysCleanTxtInject(txtAnio_Corte.Text)

pWhere = False

strSQL = "select CEDULA, NOMBRE, REGISTRO_FECHA, REGISTRO_USUARIO, UNIVERSIDAD, NIVEL, CARRERA, ESPECIALIDAD " _
       & ", CICLO, CICLO_ANIO, BENEFICIARIO_ID, BENEFICIARIO, PARENTESCO" _
       & " from  vSys_Educacion_Log" _
       & " Where CICLO_ANIO between '" & txtAnio_Inicio.Text & "' and '" & txtAnio_Corte.Text & "'"


If chkFechas.Value = xtpUnchecked Then
    strSQL = strSQL & " and REGISTRO_FECHA between '" & Format(dtpInicio.Value, "yyyy-mm-dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'"
    
End If

If cboCiclo.Text <> "TODOS" Then
    strSQL = strSQL & " and CICLO = '" & cboCiclo.Text & "'"
End If

If Len(txtUsuario.Text) > 0 Then
    strSQL = strSQL & " and REGISTRO_USUARIO like '%" & txtUsuario.Text & "%'"
End If


If Len(txtCedula.Text) > 0 Then
    strSQL = strSQL & " and Cedula like '%" & txtCedula.Text & "%'"
End If

If Len(txtNombre.Text) > 0 Then
    strSQL = strSQL & " and Nombre like '%" & txtNombre.Text & "%'"
End If

If Len(txtBeneficiarioId.Text) > 0 Then
    strSQL = strSQL & " and BENEFICIARIO_ID like '%" & txtBeneficiarioId.Text & "%'"
End If
If Len(txtBeneficiario.Text) > 0 Then
    strSQL = strSQL & " and BENEFICIARIO like '%" & txtBeneficiario.Text & "%'"
End If

If cboUniversidad.Text <> "TODOS" Then
    strSQL = strSQL & " and COD_UNIVERSIDAD = '" & cboUniversidad.ItemData(cboUniversidad.ListIndex) & "'"
End If
If cboNivel.Text <> "TODOS" Then
    strSQL = strSQL & " and COD_NIVEL = '" & cboNivel.ItemData(cboNivel.ListIndex) & "'"
End If
If cboCarrera.Text <> "TODOS" Then
    strSQL = strSQL & " and COD_CARRERA = '" & cboCarrera.ItemData(cboCarrera.ListIndex) & "'"
End If
If cboEspecialidad.Text <> "TODOS" Then
    strSQL = strSQL & " and COD_ESPECIALIDAD = '" & cboEspecialidad.ItemData(cboEspecialidad.ListIndex) & "'"
End If


Call sbCargaGrid(vGrid, 13, strSQL)
vGrid.MaxRows = vGrid.MaxRows - 1

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders

    vHeaders.Columnas = 13
    vHeaders.Headers(1) = "Identificación"
    vHeaders.Headers(2) = "Nombre"
    vHeaders.Headers(3) = "Fecha"
    vHeaders.Headers(4) = "Usuario"
    vHeaders.Headers(5) = "Universidad"
    vHeaders.Headers(6) = "Nivel"
    vHeaders.Headers(7) = "Carrera"
    vHeaders.Headers(8) = "Especialidad"
    vHeaders.Headers(9) = "Ciclo Lectivo"
    vHeaders.Headers(10) = "Año"
    vHeaders.Headers(11) = "Beneficiario Id"
    vHeaders.Headers(12) = "Beneficiario"
    vHeaders.Headers(13) = "Parentesco"

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Convenio_Universidades_Log")

End Sub

Private Sub cboCarrera_Click()
If vPaso Then Exit Sub

vPaso = True

strSQL = "exec spSys_Educacion_List 'E', '" & cboCarrera.ItemData(cboCarrera.ListIndex) & "'"
Call sbCbo_Llena_New(cboEspecialidad, strSQL, True, True)

vPaso = False

End Sub

Private Sub chkFechas_Click()
If chkFechas.Value = xtpChecked Then
    dtpInicio.Enabled = False
    dtpCorte.Enabled = False
Else
    dtpInicio.Enabled = True
    dtpCorte.Enabled = True
End If

End Sub

Private Sub Form_Load()
vModulo = 10


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - 230
vGrid.Height = Me.Height - (vGrid.Top + 450)

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub

Private Sub txtBeneficiarioId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col2Name = "Nombre"
    gBusquedas.Consulta = "Select Identificacion, Nombre from vSys_Padron_Nacional"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    
    frmBusquedas.Show vbModal
    If gBusquedas.Resultado <> "" Then
        txtBeneficiarioId.Text = gBusquedas.Resultado
        txtBeneficiario.Text = gBusquedas.Resultado2
    End If
End If
End Sub

Private Sub txtBeneficiarioId_LostFocus()
txtBeneficiario.Text = fxPadron_Nacional_Nombre(txtBeneficiarioId.Text)
End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col2Name = "Id Alterno"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    If gBusquedas.Resultado <> "" Then
        txtCedula.Text = Trim(gBusquedas.Resultado)
        txtNombre.Text = gBusquedas.Resultado2
    End If
End If

End Sub
