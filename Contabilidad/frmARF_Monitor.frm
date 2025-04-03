VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmARF_Monitor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Monitor de Arrendamientos"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15390
   LinkTopic       =   "Form2"
   ScaleHeight     =   8445
   ScaleWidth      =   15390
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   255
      Left            =   10560
      TabIndex        =   9
      Top             =   600
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Fechas"
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
   Begin XtremeSuiteControls.FlatEdit txtUnidad 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   240
      Width           =   6735
      _Version        =   1572864
      _ExtentX        =   11874
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
   Begin XtremeSuiteControls.FlatEdit txtArrendador 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   600
      Width           =   6735
      _Version        =   1572864
      _ExtentX        =   11874
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
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   7215
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   14775
      _Version        =   524288
      _ExtentX        =   26061
      _ExtentY        =   12726
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
      ScrollBarShowMax=   0   'False
      SpreadDesigner  =   "frmARF_Monitor.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      ScrollBarTrack  =   1
      AppearanceStyle =   1
      ScrollBarStyle  =   2
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   330
      Left            =   9000
      TabIndex        =   5
      Top             =   240
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   330
      Left            =   9000
      TabIndex        =   8
      Top             =   600
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
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
   Begin XtremeSuiteControls.ComboBox cboFecha 
      Height          =   330
      Left            =   10560
      TabIndex        =   10
      Top             =   240
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   615
      Left            =   12600
      TabIndex        =   11
      Top             =   240
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   1080
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
      Appearance      =   21
      Picture         =   "frmARF_Monitor.frx":0F3F
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   615
      Left            =   13920
      TabIndex        =   12
      ToolTipText     =   "Exportar a Excel"
      Top             =   240
      Width           =   615
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   1080
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
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Picture         =   "frmARF_Monitor.frx":195D
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   8280
      TabIndex        =   7
      ToolTipText     =   "Oficina / Agencia"
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   8280
      TabIndex        =   6
      ToolTipText     =   "Oficina / Agencia"
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Arrendador"
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
      Index           =   0
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Oficina / Agencia"
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Oficina/ Ud"
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
      Index           =   14
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Oficina / Agencia"
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmARF_Monitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub btnExportar_Click()
        Dim vHeaders As vGridHeaders
            vHeaders.Columnas = 19
            vHeaders.Headers(1) = "Det."
            vHeaders.Headers(2) = "No.Operación"
            vHeaders.Headers(3) = "Arrendador"
            vHeaders.Headers(4) = "Unidad"
            vHeaders.Headers(5) = "Estado"
            vHeaders.Headers(6) = "Monto"
            vHeaders.Headers(7) = "Frecuencia"
            vHeaders.Headers(8) = "Plazo"
            vHeaders.Headers(9) = "Inicia"
            vHeaders.Headers(10) = "Termina"
            vHeaders.Headers(11) = "Prox.Pago"
            vHeaders.Headers(12) = "Corte"
            vHeaders.Headers(13) = "V.Pasivo"
            vHeaders.Headers(14) = "Dep.Acumulada"
            vHeaders.Headers(15) = "V.Libros"
            vHeaders.Headers(16) = "R-Fecha"
            vHeaders.Headers(17) = "R-Usuario"
            vHeaders.Headers(18) = "A-Fecha"
            vHeaders.Headers(19) = "A-Usuario"

        
         Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Arrendamientos")

End Sub

Private Sub chkFechas_Click()

If chkFechas.Value = xtpChecked Then
    dtpInicio.Enabled = False
Else
    dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub

Private Sub Form_Load()
vModulo = 20

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

cboFecha.AddItem "Registro"
cboFecha.AddItem "Activación"
cboFecha.AddItem "Inicio"
cboFecha.AddItem "Finaliza"
cboFecha.Text = "Registro"

vGrid.MaxRows = 1

Call chkFechas_Click

End Sub


Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - 250
vGrid.Height = Me.Height - (vGrid.Top + 650)

End Sub

Private Sub txtArrendador_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "COD_ACREEDOR"
  gBusquedas.Orden = "COD_ACREEDOR"
  gBusquedas.Consulta = "select COD_ACREEDOR, Descripcion from ARF_ACREEDORES"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtArrendador.Tag = gBusquedas.Resultado
  txtArrendador.Text = gBusquedas.Resultado2

End If
End Sub

Private Sub txtUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "COD_LOCAL"
  gBusquedas.Orden = "COD_LOCAL"
  gBusquedas.Consulta = "select COD_LOCAL, Descripcion from ARF_UNIDADES"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtUnidad.Tag = gBusquedas.Resultado
  txtUnidad.Text = gBusquedas.Resultado2
  
End If
End Sub
