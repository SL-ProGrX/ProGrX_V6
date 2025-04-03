VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmActivos_PolizasReportes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de Pólizas"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.RadioButton rbInforme 
      Height          =   330
      Index           =   0
      Left            =   3480
      TabIndex        =   12
      Top             =   2640
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5318
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Listado de Pólizas"
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
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1215
      Left            =   0
      TabIndex        =   7
      Top             =   4320
      Width           =   9975
      _Version        =   1572864
      _ExtentX        =   17595
      _ExtentY        =   2143
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   735
         Left            =   7200
         TabIndex        =   8
         Top             =   240
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3831
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Reporte"
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
         Picture         =   "frmActivos_PolizasReportes.frx":0000
      End
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   330
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   5535
      _Version        =   1572864
      _ExtentX        =   9763
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
   Begin XtremeSuiteControls.FlatEdit txtPoliza 
      Height          =   330
      Left            =   1680
      TabIndex        =   5
      Top             =   1680
      Width           =   5535
      _Version        =   1572864
      _ExtentX        =   9763
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
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   330
      Left            =   1680
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2990
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
   Begin XtremeSuiteControls.CheckBox chkPolizas 
      Height          =   315
      Left            =   7440
      TabIndex        =   10
      Top             =   1680
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3196
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Todos"
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
      Appearance      =   21
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkTipos 
      Height          =   315
      Left            =   7440
      TabIndex        =   11
      Top             =   1320
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3196
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Todos"
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
      Appearance      =   21
      Value           =   1
   End
   Begin XtremeSuiteControls.RadioButton rbInforme 
      Height          =   330
      Index           =   1
      Left            =   3480
      TabIndex        =   13
      Top             =   3000
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5318
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Listado de Tipos de Pólizas"
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
   End
   Begin XtremeSuiteControls.RadioButton rbInforme 
      Height          =   330
      Index           =   2
      Left            =   3480
      TabIndex        =   14
      Top             =   3360
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5318
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Listado de Activos Protegidos"
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
   End
   Begin XtremeSuiteControls.RadioButton rbInforme 
      Height          =   330
      Index           =   3
      Left            =   3480
      TabIndex        =   4
      Top             =   3720
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5318
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Listado de Activos Desprotegidos"
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
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Pólizas para Activos Fijos"
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
      Height          =   492
      Index           =   3
      Left            =   1800
      TabIndex        =   9
      Top             =   360
      Width           =   7212
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Polizas"
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
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos"
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
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmActivos_PolizasReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset


Private Sub Form_Load()
Dim strSQL As String

vModulo = 36

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

strSQL = "select rtrim(tipo_poliza) as 'Idx', rtrim(descripcion) as 'ItmX'" _
       & " from Activos_polizas_tipos order by tipo_poliza"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

cboEstado.AddItem "Todas"
cboEstado.AddItem "Activas"
cboEstado.AddItem "Vencidas"
cboEstado.Text = "Todas"


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub txtPoliza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_poliza,descripcion from Activos_polizas"
  
  If chkTipos.Value = xtpChecked Then
      gBusquedas.Filtro = ""
  Else
      gBusquedas.Filtro = " AND TIPO_POLIZA = '" & cbo.ItemData(cbo.ListIndex) & "'"
  End If
  frmBusquedas.Show vbModal
  txtPoliza.Tag = gBusquedas.Resultado
  txtPoliza.Text = gBusquedas.Resultado2

End If

End Sub
