VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCntX_AsientosConsultaAdv 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta Avanzada: Movimientos"
   ClientHeight    =   7980
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5412
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   13932
      _Version        =   524288
      _ExtentX        =   24575
      _ExtentY        =   9546
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
      SpreadDesigner  =   "frmCntX_AsientosConsultaAdv.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.GroupBox fraConsultaAvanzada 
      Height          =   1692
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   13932
      _Version        =   1245187
      _ExtentX        =   24574
      _ExtentY        =   2984
      _StockProps     =   79
      BackColor       =   14737632
      Transparent     =   -1  'True
      Appearance      =   12
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkFechasTodas 
         Height          =   252
         Left            =   11760
         TabIndex        =   3
         Top             =   120
         Width           =   1092
         _Version        =   1245187
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   612
         Left            =   11760
         TabIndex        =   4
         Top             =   600
         Width           =   1332
         _Version        =   1245187
         _ExtentX        =   2350
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmCntX_AsientosConsultaAdv.frx":0B88
      End
      Begin XtremeSuiteControls.FlatEdit txtCentroCosto 
         Height          =   312
         Left            =   720
         TabIndex        =   5
         Top             =   960
         Width           =   852
         _Version        =   1245187
         _ExtentX        =   1503
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtCAsiento 
         Height          =   312
         Left            =   720
         TabIndex        =   6
         Top             =   120
         Width           =   852
         _Version        =   1245187
         _ExtentX        =   1503
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtUnidad 
         Height          =   312
         Left            =   720
         TabIndex        =   7
         Top             =   600
         Width           =   852
         _Version        =   1245187
         _ExtentX        =   1503
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDAsiento 
         Height          =   312
         Left            =   1560
         TabIndex        =   8
         Top             =   120
         Width           =   2772
         _Version        =   1245187
         _ExtentX        =   4890
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUnidadDesc 
         Height          =   312
         Left            =   1560
         TabIndex        =   9
         Top             =   600
         Width           =   2772
         _Version        =   1245187
         _ExtentX        =   4890
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtCentroCostoDesc 
         Height          =   312
         Left            =   1560
         TabIndex        =   10
         Top             =   960
         Width           =   2772
         _Version        =   1245187
         _ExtentX        =   4890
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtNAsiento 
         Height          =   312
         Left            =   5640
         TabIndex        =   11
         Top             =   120
         Width           =   2412
         _Version        =   1245187
         _ExtentX        =   4254
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtReferencia 
         Height          =   312
         Left            =   5640
         TabIndex        =   12
         Top             =   1320
         Width           =   6012
         _Version        =   1245187
         _ExtentX        =   10604
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtCuenta 
         Height          =   312
         Left            =   9000
         TabIndex        =   13
         Top             =   960
         Width           =   2652
         _Version        =   1245187
         _ExtentX        =   4678
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   312
         Left            =   9000
         TabIndex        =   14
         Top             =   600
         Width           =   2652
         _Version        =   1245187
         _ExtentX        =   4678
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDivisa 
         Height          =   312
         Left            =   5640
         TabIndex        =   15
         Top             =   960
         Width           =   2412
         _Version        =   1245187
         _ExtentX        =   4254
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   312
         Left            =   5640
         TabIndex        =   16
         Top             =   600
         Width           =   2412
         _Version        =   1245187
         _ExtentX        =   4254
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtLineas 
         Height          =   312
         Left            =   12360
         TabIndex        =   17
         Top             =   1320
         Width           =   732
         _Version        =   1245187
         _ExtentX        =   1291
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "1000"
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaInicio 
         Height          =   312
         Left            =   9000
         TabIndex        =   18
         Top             =   120
         Width           =   1332
         _Version        =   1245187
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
      Begin XtremeSuiteControls.DateTimePicker dtpFechaCorte 
         Height          =   312
         Left            =   10320
         TabIndex        =   19
         Top             =   120
         Width           =   1332
         _Version        =   1245187
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   5
         Left            =   4560
         TabIndex        =   30
         Top             =   1320
         Width           =   1188
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   4
         Left            =   8220
         TabIndex        =   29
         Top             =   960
         Width           =   828
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Divisa"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   3
         Left            =   4536
         TabIndex        =   28
         Top             =   960
         Width           =   1188
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "C.C."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   708
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   708
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   2
         Left            =   8220
         TabIndex        =   25
         Top             =   600
         Width           =   828
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Documento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   1
         Left            =   4536
         TabIndex        =   24
         Top             =   600
         Width           =   1188
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   8220
         TabIndex        =   23
         Top             =   120
         Width           =   588
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   708
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Asiento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   0
         Left            =   4536
         TabIndex        =   21
         Top             =   120
         Width           =   948
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Líneas:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   11760
         TabIndex        =   20
         Top             =   1356
         Width           =   648
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDebito 
      Height          =   312
      Left            =   9960
      TabIndex        =   31
      Top             =   7560
      Width           =   1932
      _Version        =   1245187
      _ExtentX        =   3408
      _ExtentY        =   550
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
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtCredito 
      Height          =   312
      Left            =   11880
      TabIndex        =   32
      Top             =   7560
      Width           =   1932
      _Version        =   1245187
      _ExtentX        =   3408
      _ExtentY        =   550
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
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Totales:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   8520
      TabIndex        =   1
      Top             =   7596
      Width           =   1368
   End
End
Attribute VB_Name = "frmCntX_AsientosConsultaAdv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnBuscar_Click()
Call sbConsulta
End Sub

Private Sub chkFechasTodas_Click()
If chkFechasTodas.Value = vbChecked Then
   dtpFechaInicio.Enabled = False
Else
   dtpFechaInicio.Enabled = True
End If

dtpFechaCorte.Enabled = dtpFechaInicio.Enabled
End Sub

Private Sub Form_Activate()
vModulo = 20
End Sub

Private Sub Form_Load()
vModulo = 20

  vPaso = False
  dtpFechaInicio.Value = fxFechaServidor
  dtpFechaCorte.Value = dtpFechaInicio.Value
  Call chkFechasTodas_Click
  vGrid.MaxRows = 0
End Sub

Private Sub sbBuscar(Optional vBusca As Integer = 1)

Select Case vBusca
  Case 1 'Tipo de ASiento
     gBusquedas.Columna = "Tipo_Asiento"
     gBusquedas.Orden = "Tipo_Asiento"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     gBusquedas.Consulta = "select Tipo_Asiento,descripcion from CntX_Tipos_Asientos"
     frmBusquedas.Show vbModal
     txtCAsiento = gBusquedas.Resultado
  Case 2 'Descripcion del Tipo de Asiento
     gBusquedas.Columna = "Descripcion"
     gBusquedas.Orden = "Descripcion"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     gBusquedas.Consulta = "select Tipo_Asiento,descripcion from CntX_Tipos_Asientos"
     frmBusquedas.Show vbModal
     txtCAsiento = gBusquedas.Resultado
  Case 3 'Numero de Asiento
     gBusquedas.Columna = "Num_Asiento"
     gBusquedas.Orden = "Num_Asiento"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                       & " and tipo_asiento = '" & txtCAsiento & "'"
                       
                       
  Case 4 'Codigo Unidad
     gBusquedas.Columna = "cod_unidad"
     gBusquedas.Orden = "cod_unidad"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     gBusquedas.Consulta = "select cod_unidad as 'Unidad',Descripcion from CntX_Unidades"
     frmBusquedas.Show vbModal
     txtUnidad.Text = gBusquedas.Resultado
     txtUnidadDesc.Text = gBusquedas.Resultado2
  Case 5 'Descripcion de la Unidad
     gBusquedas.Columna = "Descripcion"
     gBusquedas.Orden = "Descripcion"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     gBusquedas.Consulta = "select cod_unidad as 'Unidad',Descripcion from CntX_Unidades"
     frmBusquedas.Show vbModal
     txtUnidad.Text = gBusquedas.Resultado
     txtUnidadDesc.Text = gBusquedas.Resultado2
                       
                       
  Case 6 'Codigo Centro de Costo
     gBusquedas.Columna = "cod_Centro_Costo"
     gBusquedas.Orden = "cod_Centro_Costo"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     gBusquedas.Consulta = "select cod_Centro_Costo as 'Centro',Descripcion from CntX_Centro_Costos"
     frmBusquedas.Show vbModal
     txtCentroCosto.Text = gBusquedas.Resultado
     txtCentroCostoDesc.Text = gBusquedas.Resultado2
  Case 7 'Descripcion de la Centro de Costo
     gBusquedas.Columna = "Descripcion"
     gBusquedas.Orden = "Descripcion"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     gBusquedas.Consulta = "select cod_Centro_Costo as 'Centro',Descripcion from CntX_Centro_Costos"
     frmBusquedas.Show vbModal
     txtCentroCosto.Text = gBusquedas.Resultado
     txtCentroCostoDesc.Text = gBusquedas.Resultado2
                       
  Case 8 'Divisa
     gBusquedas.Columna = "cod_divisa"
     gBusquedas.Orden = "cod_divisa"
     gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
     gBusquedas.Consulta = "select cod_divisa as 'Divisa',Descripcion from CntX_Divisas"
     frmBusquedas.Show vbModal
     txtDivisa.Text = gBusquedas.Resultado
                       
  Case 9 'Cuenta Contable
     frmCntX_ConsultaCuentas.Show vbModal
     txtCuenta.Text = fxCntX_CuentaFormato(True, gCuenta, 0)
                       
End Select
End Sub


Private Sub sbConsulta()
Dim strSQL As String, curDebitos As Currency, curCreditos As Currency, lng As Long

strSQL = "select Top " & txtLineas.Text & " 0,Asi.TIPO_ASIENTO,Asi.NUM_ASIENTO,Asi.FECHA_ASIENTO" _
       & ",Det.COD_CUENTA,Det.COD_UNIDAD, Det.Cod_Centro_Costo" _
       & ",Det.COD_DIVISA,Det.TIPO_CAMBIO,Det.DOCUMENTO,Det.DETALLE" _
       & ",Det.MONTO_CREDITO,Det.MONTO_DEBITO" _
       & " from CNTX_CUENTAS Cta inner join CNTX_ASIENTOS_DETALLE Det" _
       & " on Cta.COD_CONTABILIDAD = Det.COD_CONTABILIDAD and Cta.COD_CUENTA = Det.COD_CUENTA" _
       & " inner join CNTX_ASIENTOS Asi on Det.COD_CONTABILIDAD = Asi.COD_CONTABILIDAD" _
       & " and Det.TIPO_ASIENTO = Asi.TIPO_ASIENTO and Det.NUM_ASIENTO = Asi.NUM_ASIENTO" _
       & " Where Cta.COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta
       
       
If txtReferencia.Text <> "" Then
   strSQL = strSQL & " and Asi.Referencia like '%" & txtReferencia.Text & "%'"
End If
              
              
If txtCAsiento.Text <> "" Then
   strSQL = strSQL & " and Det.Tipo_Asiento = '" & txtCAsiento.Text & "'"
End If
       
If txtNAsiento.Text <> "" Then
   strSQL = strSQL & " and Det.Num_Asiento like '%" & txtNAsiento.Text & "%'"
End If
       
If txtUnidad.Text <> "" Then
   strSQL = strSQL & " and Det.Cod_Unidad = '" & txtUnidad.Text & "'"
End If
       
If txtCentroCosto.Text <> "" Then
   strSQL = strSQL & " and Det.Cod_Centro_Costo = '" & txtCentroCosto.Text & "'"
End If
       
If txtDivisa.Text <> "" Then
   strSQL = strSQL & " and Det.cod_Divisa = '" & txtDivisa.Text & "'"
End If
       
If txtDocumento.Text <> "" Then
   strSQL = strSQL & " and Det.documento like '%" & txtDocumento.Text & "%'"
End If
       
If txtDetalle.Text <> "" Then
   strSQL = strSQL & " and Det.detalle like '%" & txtDetalle.Text & "%'"
End If
       
If txtCuenta.Text <> "" Then
   strSQL = strSQL & " and Det.cod_cuenta = '" & fxCntX_CuentaFormato(False, txtCuenta.Text, 0) & "'"
End If
       
If chkFechasTodas.Value = vbUnchecked Then
   strSQL = strSQL & " and Asi.Fecha_Asiento between  '" & Format(dtpFechaInicio.Value, "yyyy/mm/dd") _
          & "' and '" & Format(dtpFechaCorte.Value, "yyyy/mm/dd") & "'"
End If
       
vPaso = True
 Call sbCargaGrid(vGrid, 13, strSQL)
 vGrid.MaxRows = vGrid.MaxRows - 1
vPaso = False


curDebitos = 0
curCreditos = 0

For lng = 1 To vGrid.MaxRows
  vGrid.Row = lng
  vGrid.col = 12
  curDebitos = curDebitos + CCur(vGrid.Text)
  vGrid.col = 13
  curCreditos = curCreditos + CCur(vGrid.Text)
Next lng

txtDebito.Text = Format(curDebitos, "Standard")
txtCredito.Text = Format(curCreditos, "Standard")

End Sub

Private Sub txtCAsiento_Change()
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select descripcion from CntX_Tipos_Asientos where cod_contabilidad = " _
       & gCntX_Parametros.CodigoConta & " and tipo_asiento = '" _
       & txtCAsiento.Text & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  txtDAsiento = rs!Descripcion
End If
rs.Close
End Sub


Private Sub txtCAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDAsiento.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(1)
End Sub

Private Sub txtCAsiento_LostFocus()
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select descripcion from CntX_Tipos_Asientos where cod_contabilidad = " _
       & gCntX_Parametros.CodigoConta & " and tipo_asiento = '" _
       & txtCAsiento.Text & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  txtDAsiento = rs!Descripcion
End If
rs.Close

End Sub

Private Sub txtCentroCosto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCentroCostoDesc.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(6)
End Sub

Private Sub txtCentroCostoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDivisa.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(7)
End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCAsiento.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(9)
End Sub

Private Sub txtCuenta_LostFocus()
If Len(txtCuenta.Text) > 0 Then
    txtCuenta.Text = fxCntX_CuentaFormato(True, txtCuenta.Text, 0)
End If
End Sub

Private Sub txtDAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNAsiento.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(2)
End Sub

Private Sub txtDivisa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuenta.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(8)
End Sub


Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDetalle.SetFocus
End Sub

Private Sub txtNAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUnidad.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(3)
End Sub

Private Sub txtUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUnidadDesc.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(4)
End Sub

Private Sub txtUnidadDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus
If KeyCode = vbKeyF4 Then Call sbBuscar(5)
End Sub


Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

Dim frmX As Form, vEncuenta As Boolean

Me.Hide
vGrid.Row = Row
vGrid.col = 2
gCntX_Arbol.AsientoTipo = vGrid.Text
vGrid.col = 3
gCntX_Arbol.AsientoNumr = vGrid.Text
gCntX_Arbol.ArbolActivo = True

vEncuenta = False
For Each frmX In Forms
   If Trim(frmX.Name) = "frmCntX_Asientos" Then
        frmX.sbFormReLoad
        vEncuenta = True
   End If
Next

If Not vEncuenta Then
    Call sbFormsCall("frmCntX_Asientos")
End If
gCntX_Arbol.ArbolActivo = False
    
End Sub
