VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCntX_CuentaHistorico 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historama de Cuentas"
   ClientHeight    =   8805
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   14670
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.RadioButton OptX 
      CausesValidation=   0   'False
      Height          =   252
      Index           =   0
      Left            =   8400
      TabIndex        =   9
      Top             =   1440
      Width           =   1932
      _Version        =   1310723
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Ultimos 24 meses"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
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
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   12480
      Top             =   480
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6375
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   14415
      _Version        =   524288
      _ExtentX        =   25426
      _ExtentY        =   11245
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
      MaxCols         =   493
      ScrollBars      =   2
      SpreadDesigner  =   "frmCntX_CuentaHistorico.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   615
      Left            =   11280
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Buscar"
      BackColor       =   -2147483633
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
      Picture         =   "frmCntX_CuentaHistorico.frx":07FC
   End
   Begin XtremeSuiteControls.ComboBox cboUnidad 
      Height          =   312
      Left            =   2880
      TabIndex        =   5
      Top             =   1440
      Width           =   5292
      _Version        =   1310723
      _ExtentX        =   9340
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
   Begin XtremeSuiteControls.ComboBox cboCentroCosto 
      Height          =   312
      Left            =   2880
      TabIndex        =   6
      Top             =   1800
      Width           =   5292
      _Version        =   1310723
      _ExtentX        =   9340
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
   Begin XtremeSuiteControls.RadioButton OptX 
      CausesValidation=   0   'False
      Height          =   252
      Index           =   1
      Left            =   8400
      TabIndex        =   10
      Top             =   1800
      Width           =   1932
      _Version        =   1310723
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Histórico"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.FlatEdit txtCuenta 
      Height          =   312
      Left            =   2880
      TabIndex        =   0
      Top             =   480
      Width           =   2532
      _Version        =   1310723
      _ExtentX        =   4466
      _ExtentY        =   556
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
      Height          =   312
      Left            =   5400
      TabIndex        =   1
      Top             =   480
      Width           =   6972
      _Version        =   1310723
      _ExtentX        =   12298
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   615
      Left            =   12720
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Exportar"
      BackColor       =   -2147483633
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
      Picture         =   "frmCntX_CuentaHistorico.frx":121A
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad de Negocio"
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
      Index           =   5
      Left            =   960
      TabIndex        =   8
      Top             =   1440
      Width           =   1692
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Centro de Costo"
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
      Index           =   6
      Left            =   960
      TabIndex        =   7
      Top             =   1800
      Width           =   1692
   End
   Begin VB.Label lblBalanceCtaHistorico 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   828
   End
   Begin VB.Image imgBanner 
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   15855
   End
End
Attribute VB_Name = "frmCntX_CuentaHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Public Sub sbConsultaExterna(pCuenta As String)

txtCuenta.Text = pCuenta
txtCuenta_LostFocus

End Sub

Private Sub btnBuscar_Click()
Call sbBuscar
End Sub

Private Sub btnExport_Click()

 Dim vHeaders As vGridHeaders
    vHeaders.Columnas = vGrid.MaxCols
    
    vHeaders.Headers(1) = "Año"
    vHeaders.Headers(2) = "Mes"
    
    vHeaders.Headers(3) = "Cuenta"
    vHeaders.Headers(4) = "Unidad"
    vHeaders.Headers(5) = "Centro"
    vHeaders.Headers(6) = "Saldo Inicial"
    vHeaders.Headers(7) = "Débitos"
    vHeaders.Headers(8) = "Créditos"
    vHeaders.Headers(9) = "Neto del Mes"
    vHeaders.Headers(10) = "Saldo Final"

 Call sbSIFGridExportar(vGrid, vHeaders, "Cta_Historico_" & txtCuenta.Text)
       


End Sub

Private Sub cboUnidad_Click()
Dim strSQL As String

If vPaso Then Exit Sub

If cboUnidad.Text = "[CONSOLIDADO]" Then
    strSQL = "select rtrim(cod_Centro_Costo) as 'IdX', rtrim(descripcion) as 'ItmX'" _
           & " from CntX_Centro_Costos where cod_contabilidad = " & gCntX_Parametros.CodigoConta
Else
    strSQL = "select rtrim(cod_Centro_Costo) as 'IdX',rtrim(descripcion) as 'ItmX'" _
           & " from CntX_Centro_Costos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and cod_centro_costo in(select cod_centro_costo from CntX_Unidades_CC where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and cod_unidad = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "')"
End If


Call sbCbo_Llena_New(cboCentroCosto, strSQL, True, True)
End Sub

Private Sub Form_Load()

vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

vPaso = True
Call sbCntX_CargaCboUnidades(cboUnidad)
vPaso = False

Call cboUnidad_Click

End Sub



Private Sub OptX_Click(Index As Integer)
Call sbBuscar

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

If GLOBALES.gTag <> txtCuenta.Text Then
   txtCuenta.Text = GLOBALES.gTag
   txtCuentaDesc.SetFocus
End If

End Sub



Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pCuenta As String, pFecha As Date

Me.MousePointer = vbHourglass

On Error GoTo vError


pCuenta = fxCntX_CuentaFormato(False, txtCuenta.Text, 0)

strSQL = "select dbo.fxSys_FechaAnioMesToDatetime(" & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes & ") as 'Fecha'"
Call OpenRecordSet(rs, strSQL)
 pFecha = rs!fecha
rs.Close
   
     'Selecciona la fuente de la información dependiento de los filtros
     Select Case True
       Case (cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS")
       
            strSQL = "select M.Anio, M.Mes, C.cod_Cuenta_Mask, '' AS 'cod_unidad', '' AS 'cod_Centro_Costo'" _
                   & ",M.Saldo_Inicial, abs(M.Total_Debitos) as 'Debitos' , abs(M.Total_Creditos) as 'Creditos'" _
                   & ",M.Neto_Mes as 'Mes'" _
                   & ",M.SALDO_Final 'SaldoFinal'" _
                   & " from vCntX_Mov_Cuentas_General M inner join CntX_Cuentas C on M.cod_Contabilidad = C.cod_Contabilidad and M.cod_Cuenta = C.cod_Cuenta" _
                   & " inner join CntX_Periodos P on M.cod_Contabilidad = P.cod_Contabilidad and M.Anio = P.Anio and M.mes = P.Mes" _
                   & " where M.cod_cuenta = '" & pCuenta & "' and M.cod_contabilidad = " & gCntX_Parametros.CodigoConta
              
     
     
     
       Case (cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS")
          
            strSQL = "select M.Anio, M.Mes, C.cod_Cuenta_Mask,  M.cod_unidad, M.cod_Centro_Costo" _
                   & ",M.Saldo_Inicial, abs(M.Total_Debitos) as 'Debitos' , abs(M.Total_Creditos) as 'Creditos'" _
                   & ",M.Neto_Mes as 'Mes'" _
                   & ",M.SALDO_Final 'SaldoFinal'" _
                   & " from vCntX_Mov_Cuentas_CentroCosto M inner join CntX_Cuentas C on M.cod_Contabilidad = C.cod_Contabilidad and M.cod_Cuenta = C.cod_Cuenta" _
                   & " inner join CntX_Periodos P on M.cod_Contabilidad = P.cod_Contabilidad and M.Anio = P.Anio and M.mes = P.Mes" _
                   & " where M.cod_cuenta = '" & pCuenta & "' and M.cod_contabilidad = " & gCntX_Parametros.CodigoConta
            
            
       Case (cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS")
            
            strSQL = "select M.Anio, M.Mes, C.cod_Cuenta_Mask,  M.cod_unidad, '' as cod_Centro_Costo" _
                   & ",M.Saldo_Inicial, abs(M.Total_Debitos) as 'Debitos' , abs(M.Total_Creditos) as 'Creditos'" _
                   & ",M.Neto_Mes as 'Mes'" _
                   & ",M.SALDO_Final 'SaldoFinal'" _
                   & " from vCntX_Mov_Cuentas_Unidad M inner join CntX_Cuentas C on M.cod_Contabilidad = C.cod_Contabilidad and M.cod_Cuenta = C.cod_Cuenta" _
                   & " inner join CntX_Periodos P on M.cod_Contabilidad = P.cod_Contabilidad and M.Anio = P.Anio and M.mes = P.Mes" _
                   & " where M.cod_cuenta = '" & pCuenta & "' and M.cod_contabilidad = " & gCntX_Parametros.CodigoConta

       Case (cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS")
     
            strSQL = "select M.Anio, M.Mes, C.cod_Cuenta_Mask,  M.cod_unidad, M.cod_Centro_Costo" _
                   & ",abs(M.Total_Debitos) as 'Debitos' , abs(M.Total_Creditos) as 'Creditos'" _
                   & ",M.Saldo_Inicial, M.Total_Debitos + M.Total_Creditos as 'Mes'" _
                   & ",M.SALDO_Inicial + M.Total_Debitos + M.Total_Creditos as 'SaldoFinal'" _
                   & " from CntX_Mov_Cuentas_Detallado M inner join CntX_Cuentas C on M.cod_Contabilidad = C.cod_Contabilidad and M.cod_Cuenta = C.cod_Cuenta" _
                   & " inner join CntX_Periodos P on M.cod_Contabilidad = P.cod_Contabilidad and M.Anio = P.Anio and M.mes = P.Mes" _
                   & " where M.cod_cuenta = '" & pCuenta & "' and M.cod_contabilidad = " & gCntX_Parametros.CodigoConta
     
     End Select
   
   
     'Filtros Finales
     Select Case True
       Case (cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS")
     
       Case (cboUnidad.Text = "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS")
          strSQL = strSQL & " and M.cod_centro_costo = '" & cboCentroCosto.ItemData(cboCentroCosto.ListIndex) & "'"
            
       Case (cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text = "TODOS")
          strSQL = strSQL & " and M.cod_unidad = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "'"

       Case (cboUnidad.Text <> "[CONSOLIDADO]" And cboCentroCosto.Text <> "TODOS")
          strSQL = strSQL & " and M.cod_unidad = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "'" _
                 & " and M.cod_centro_costo = '" & cboCentroCosto.ItemData(cboCentroCosto.ListIndex) & "'"
     End Select

    'Filtros Finales
    
    Select Case True
      Case OptX.Item(0).Value 'Ultimos 24 meses
            
            strSQL = strSQL & " and P.PERIODO_CORTE between '" & Format(DateAdd("m", -24, pFecha), "yyyy/mm/dd") _
                   & "' and '" & Format(pFecha, "yyyy/mm/dd") & " 23:59:59'"
      Case OptX.Item(1).Value 'Historico
            strSQL = strSQL & " and M.mes = " & gCntX_Parametros.PeriodoMes
    End Select
    
    strSQL = strSQL & " Order by M.anio DESC,M.mes desc"
    
Call sbCargaGrid(vGrid, 10, strSQL, True)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaDesc.SetFocus

If KeyCode = vbKeyF4 Then
     frmCntX_ConsultaCuentas.Show vbModal
     txtCuenta.Text = gCuenta
End If
End Sub

Private Sub txtCuenta_LostFocus()

txtCuenta.Text = fxCntX_CuentaFormato(False, txtCuenta.Text, 0)

txtCuentaDesc.Text = fxCntX_Cuenta("D", txtCuenta.Text)

txtCuenta.Text = fxCntX_CuentaFormato(True, txtCuenta.Text, 0)

If txtCuentaDesc.Text <> "" Then
  Call sbBuscar
End If

End Sub
