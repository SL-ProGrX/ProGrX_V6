VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmFND_ConciliacionTeledolarSinpe 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Fondos: Consulta Conciliación Teledolar - SINPE"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14700
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   14700
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   960
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   661
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
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   330
      Left            =   5160
      TabIndex        =   2
      Top             =   960
      Width           =   1935
      _Version        =   1572864
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
   Begin XtremeSuiteControls.GroupBox gbAccion 
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   14295
      _Version        =   1572864
      _ExtentX        =   25215
      _ExtentY        =   1296
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   495
         Index           =   0
         Left            =   4920
         TabIndex        =   7
         Top             =   240
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
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
         Picture         =   "frmFND_ConciliacionTeledolarSinpe.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   495
         Index           =   1
         Left            =   6360
         TabIndex        =   8
         Top             =   240
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
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
         Picture         =   "frmFND_ConciliacionTeledolarSinpe.frx":0700
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   495
         Index           =   2
         Left            =   7920
         TabIndex        =   9
         Top             =   240
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Informe"
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
         Picture         =   "frmFND_ConciliacionTeledolarSinpe.frx":086A
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBarX 
         Height          =   135
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   14895
         _Version        =   1572864
         _ExtentX        =   26273
         _ExtentY        =   238
         _StockProps     =   93
         BackColor       =   -2147483633
         Scrolling       =   1
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6015
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   14295
      _Version        =   524288
      _ExtentX        =   25215
      _ExtentY        =   10610
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
      SpreadDesigner  =   "frmFND_ConciliacionTeledolarSinpe.frx":0F71
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta de Conciliación Teledolar vrs SINPE"
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
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   5
      Top             =   240
      Width           =   6255
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   3
      Top             =   960
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   765
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15135
   End
End
Attribute VB_Name = "frmFND_ConciliacionTeledolarSinpe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Private Sub sbBuscar()

On Error GoTo vError

Me.MousePointer = vbHourglass

Dim pTipo As String

If Mid(cboTipo.Text, 1, 1) = "T" Then
  pTipo = "0"
Else
  pTipo = "1"
End If


strSQL = "exec spFndReporteConciliacionTeledolarAseccss '" & Format(dtpInicio.Value, "yyyy-mm-dd") & "', '" & Format(dtpCorte.Value, "yyyy-mm-dd") _
       & " 23:59', " & pTipo
Call sbCargaGrid(vGrid, 11, strSQL)

'Call OpenRecordSet(rs, strSQL)

'With vGrid
'  .MaxRows = 0
'  Do While Not rs.EOF
'     .MaxRows = .MaxRows + 1
'     .Row = .MaxRows
'     .Col = 1
'     .Value = chkTodas.Value
'     .Col = 2
'     .Text = CStr(rs!cod_Renuncia)
'     .Col = 3
'     .Value = chkS06.Value
'     .Col = 4
'     .Text = Trim(rs!Cedula)
'     .Col = 5
'     .Text = Trim(rs!Nombre)
'     .Col = 6
'     .Text = Trim(rs!Tipo_Desc)
'     .Col = 7
'     .Text = Trim(rs!Causa_Desc)
'     .Col = 8
'     .Text = Trim(rs!Estado_Desc)
'     .Col = 9
'     .Text = rs!Resuelto_Fecha_Mask
'     .Col = 10
'     .Text = Trim(rs!Resuelto_User & "")
'     .Col = 11
'     .Text = rs!Registro_Fecha_Mask
'     .Col = 12
'     .Text = Trim(rs!Registro_User & "")
'     .Col = 13
'     .Text = Trim(rs!Promotor_Desc)
'   rs.MoveNext
'  Loop
'  rs.Close
'End With


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub

Private Sub sbExportar()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 11
    vHeaders.Headers(1) = "Fecha"
    vHeaders.Headers(2) = "Cédula"
    vHeaders.Headers(3) = "Nombre"
    vHeaders.Headers(4) = "Producto"
    vHeaders.Headers(5) = "No. Servicio"
    vHeaders.Headers(6) = "No. Contrato"
    vHeaders.Headers(7) = "Monto Teledolar"
    vHeaders.Headers(8) = "Débito en Cuenta"
    vHeaders.Headers(9) = "Diferencia"
    vHeaders.Headers(10) = "Fondo Negativo"
    vHeaders.Headers(11) = "Estado"

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Conciliacion_Teledolar_Sinpe")

End Sub

Private Sub sbReportes()
Dim strSQL As String, pTipo As Integer


Me.MousePointer = vbHourglass

On Error GoTo vError

With frmContenedor.Crt
  .Reset
  .WindowShowGroupTree = True
  .WindowShowPrintSetupBtn = True
  .WindowShowRefreshBtn = True
  .WindowShowSearchBtn = True
  .WindowState = crptMaximized
  .WindowTitle = "Reportes del Módulo de Fondos"

  .Connect = glogon.ConectRPT
    
  If Mid(cboTipo.Text, 1, 1) = "T" Then
    pTipo = 0
  Else
    pTipo = 1
  End If
         
    .ReportFileName = SIFGlobal.fxPathReportes("Fondos_Conciliacion_Teledolar_Sinpe.rpt")

'    .StoredProcParam(0) = Format(dtpInicio.Value, "yyyy-MM-dd")
'    .StoredProcParam(1) = Format(dtpCorte.Value, "yyyy-MM-dd")
    .StoredProcParam(2) = pTipo
    
    .Formulas(0) = "Subtitulo='Incio " & Format(dtpInicio.Value, "dd/mm/yyyy") & " Corte " & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
    .Formulas(1) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "Usuario='" & Trim(glogon.Usuario) & "'"
    .Formulas(3) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
    
   .Action = 1

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnAccion_Click(Index As Integer)

Select Case Index
    Case 0 'Buscar
      Call sbBuscar
    Case 1 'Exportar
      Call sbExportar
    Case 2 'Informe
      Call sbReportes
    Case Else
    
End Select

End Sub

Private Sub cboTipo_Click()
If vPaso Then Exit Sub

vGrid.MaxRows = 0

End Sub


Private Sub Form_Load()
vModulo = 1
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True
    cboTipo.AddItem "Todos los Casos"
    cboTipo.AddItem "Diferencias"
    cboTipo.Text = "Todos los Casos"
vPaso = False

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -15, dtpCorte.Value)

vGrid.MaxRows = 0

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

gbAccion.Width = Me.Width - 550
ProgressBarX.Width = gbAccion.Width

vGrid.Width = gbAccion.Width

vGrid.Height = Me.Height - (vGrid.top + 650)

End Sub

