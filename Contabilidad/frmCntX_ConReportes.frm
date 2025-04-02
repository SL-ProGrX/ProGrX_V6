VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmCntX_ConReportes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reportes Consolidados"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6420
   HelpContextID   =   10
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbo 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   120
      Width           =   5055
   End
   Begin VB.OptionButton optTipo 
      Caption         =   "Detalle"
      Height          =   315
      Index           =   1
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   2055
   End
   Begin VB.OptionButton optTipo 
      Caption         =   "Resumen"
      Height          =   315
      Index           =   0
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.TextBox txtPeriodo 
      Height          =   315
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox txtAnio 
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtMes 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin VB.ComboBox cboReporte 
      Height          =   315
      ItemData        =   "frmCntX_ConReportes.frx":0000
      Left            =   1320
      List            =   "frmCntX_ConReportes.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   5055
   End
   Begin XtremeSuiteControls.PushButton cmdReporte 
      Height          =   612
      Left            =   4920
      TabIndex        =   11
      Top             =   1800
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Reporte"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Picture         =   "frmCntX_ConReportes.frx":0035
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   6360
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      X1              =   0
      X2              =   6360
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      Height          =   255
      Index           =   3
      Left            =   1320
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Consolidación"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Periodo"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Reporte"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "frmCntX_ConReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type xUtilidad
  UB_GENERAL_MES As Currency
  UB_GENERAL_ACUMULADO As Currency
  UN_GENERAL_MES As Currency
  UN_GENERAL_ACUMULADO As Currency
End Type

Private Type xPasivosPatrimonio
  TOTAL_PPNETO As Currency
  TOTAL_PPACTUAL As Currency
End Type

Private Utilidad As xUtilidad, Totales As xPasivosPatrimonio
Private gcurUtilidadMes As Currency, gcurUtilidadActual As Currency

Private Function fxVerificaPeriodoCierre(lngAnio As Long, iMes As Integer, lngCodEmpresa As Long) As Boolean
Dim strSQL As String, rsX As New ADODB.Recordset
'Verifica si el periodo es el último del periodo fiscal y si este ya fue cerrado,
'Entonces la cuenta de perdidas y ganancias fue afecta, por lo tanto las variables
'de utilidad no se deben de utilizar

fxVerificaPeriodoCierre = True

strSQL = "select * from periodos where COD_CONTABILIDAD = " & lngCodEmpresa _
       & " and anio = " & lngAnio & " and mes = " & iMes
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
  rsX.Close
  Exit Function
End If

If rsX!Estado = "P" Then
  rsX.Close
  Exit Function
End If
rsX.Close

strSQL = "select * from cierres where COD_CONTABILIDAD = " & lngCodEmpresa _
       & " and corte_anio = " & lngAnio & " and corte_mes = " & iMes
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
  rsX.Close
  Exit Function
End If

If rsX!Estado = "P" Then
  rsX.Close
  Exit Function
Else
  fxVerificaPeriodoCierre = False
End If
rsX.Close

End Function

Private Sub sbTotalesCPA(lngAnio As Long, iMes As Integer, vConsolida As Long, lngCodigoEmpresa As Long)
Dim rs As New ADODB.Recordset, strSQL As String

Totales.TOTAL_PPNETO = 0
Totales.TOTAL_PPACTUAL = 0


strSQL = "select C.Clasificacion,isnull(sum(saldo_inicial),0) as SI" _
       & ", isnull(sum(total_debitos),0) as TD, isnull(sum(total_creditos),0) as TC " _
       & " from Con_Movimientos A inner join cuentas B" _
       & " On A.COD_CONTABILIDAD = B.COD_CONTABILIDAD and A.cod_cuenta = B.cod_cuenta" _
       & " Inner join tipos_cuentas C " _
       & " On A.COD_CONTABILIDAD = C.COD_CONTABILIDAD and B.tipo_cuenta = C.tipo_cuenta" _
       & " where A.COD_CONTABILIDAD = " & lngCodigoEmpresa _
       & " and A.cod_consolida = " & vConsolida _
       & " and B.cuenta_madre = '' and A.anio = " & lngAnio _
       & " and A.mes = " & iMes & " group by C.clasificacion"


rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  Select Case rs!Clasificacion
    Case "C", "P"
       Totales.TOTAL_PPNETO = Totales.TOTAL_PPNETO + rs!TD + rs!TC
       Totales.TOTAL_PPACTUAL = Totales.TOTAL_PPACTUAL + rs!si + rs!TD + rs!TC
  End Select
  rs.MoveNext

Loop
rs.Close


End Sub

Public Sub sbUtilidad(lngAnio As Long, iMes As Integer, vConsolida As Long, lngCodigoEmpresa As Long)
Dim rs As New ADODB.Recordset, strSQL As String
Dim curGastosMes As Currency, curIngresosMes As Currency
Dim curGastosInicial As Currency, curIngresosInicial As Currency

'Utiliza las varibles publicas de Utilidad (gcurUtilidad...)

curGastosMes = 0
curIngresosMes = 0
curGastosInicial = 0
curIngresosInicial = 0


strSQL = "select C.Clasificacion, isnull(sum(saldo_inicial),0) as SI" _
       & ", isnull(sum(total_debitos),0) as TD, isnull(sum(total_creditos),0) as TC " _
       & " from con_Movimientos A inner join cuentas B " _
       & " on A.COD_CONTABILIDAD = B.COD_CONTABILIDAD and A.cod_cuenta = B.cod_cuenta" _
       & " inner join tipos_cuentas C on A.COD_CONTABILIDAD = C.COD_CONTABILIDAD" _
       & " and C.tipo_cuenta = B.tipo_cuenta" _
       & " where A.COD_CONTABILIDAD = " & lngCodigoEmpresa _
       & " and A.cod_consolida = " & vConsolida _
       & " and B.cuenta_madre = '' and A.anio = " & lngAnio _
       & " and A.mes = " & iMes & " group by C.clasificacion"

rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  Select Case rs!Clasificacion
    Case "I"
      curIngresosMes = rs!TD + rs!TC
      curIngresosInicial = rs!si
    Case "G", "V"
      curGastosMes = rs!TD + rs!TC
      curGastosInicial = rs!si
  End Select
  rs.MoveNext
Loop
rs.Close

gcurUtilidadMes = curIngresosMes - curGastosMes
gcurUtilidadActual = (curIngresosInicial + curIngresosMes) - (curGastosInicial + curGastosMes)

End Sub


Private Sub sbUtilidadER(lngAnio As Long, iMes As Integer, vConsolida As Long, iCodigoEmpresa As Long)
Dim rs As New ADODB.Recordset, strSQL As String
Dim curGastosMes As Currency, curIngresosMes As Currency
Dim curGastosInicial As Currency, curIngresosInicial As Currency
Dim curGastosMesX As Currency, curIngresosMesX As Currency
Dim curGastosInicialX As Currency, curIngresosInicialX As Currency

'Utiliza las varibles de arreglo de Utilidad

Utilidad.UB_GENERAL_ACUMULADO = 0
Utilidad.UB_GENERAL_MES = 0
Utilidad.UN_GENERAL_ACUMULADO = 0
Utilidad.UN_GENERAL_MES = 0

curGastosMes = 0
curIngresosMes = 0
curGastosInicial = 0
curIngresosInicial = 0

curGastosMesX = 0
curIngresosMesX = 0
curGastosInicialX = 0
curIngresosInicialX = 0

strSQL = "select C.Clasificacion,C.ER,isnull(sum(saldo_inicial),0) as SI" _
       & ", isnull(sum(total_debitos),0) as TD, isnull(sum(total_creditos),0) as TC " _
       & " from Con_Movimientos A inner join cuentas B" _
       & " On A.COD_CONTABILIDAD = B.COD_CONTABILIDAD and A.cod_cuenta = B.cod_cuenta" _
       & " Inner Join tipos_cuentas C" _
       & " On A.COD_CONTABILIDAD = C.COD_CONTABILIDAD and B.Tipo_cuenta = C.Tipo_cuenta" _
       & " where A.COD_CONTABILIDAD = " & iCodigoEmpresa _
       & " and A.cod_consolida = " & vConsolida _
       & " and B.cuenta_madre = '' and A.anio = " & lngAnio _
       & " and A.mes = " & iMes & " group by C.clasificacion,C.ER"

rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  Select Case rs!Clasificacion
    Case "I"
      If IsNull(rs!eR) Then
        curIngresosMes = rs!TD + rs!TC
        curIngresosInicial = rs!si
      Else
        curIngresosMesX = rs!TD + rs!TC
        curIngresosInicialX = rs!si
      End If
    
    Case "G", "V"
      If IsNull(rs!eR) Then
        curGastosMes = rs!TD + rs!TC
        curGastosInicial = rs!si
      Else
        curGastosMesX = rs!TD + rs!TC
        curGastosInicialX = rs!si
      End If
  
  End Select
  rs.MoveNext

Loop
rs.Close

Utilidad.UB_GENERAL_MES = curIngresosMes - curGastosMes
Utilidad.UB_GENERAL_ACUMULADO = (curIngresosInicial + curIngresosMes) - (curGastosInicial + curGastosMes)

Utilidad.UN_GENERAL_MES = curIngresosMesX - curGastosMesX
Utilidad.UN_GENERAL_ACUMULADO = (curIngresosInicialX + curIngresosMesX) - (curGastosInicialX + curGastosMesX)

End Sub


Private Sub cmdReporte_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim xMascara As String, lngCodEmpresa As Long
Dim iNivel1 As Integer

On Error GoTo vError

Screen.MousePointer = vbHourglass

strSQL = "select E.*" _
       & " from CNTX_CONTABILIDADES E inner join CNTX_CONSOLIDA_DEFINICION C" _
       & " on E.COD_CONTABILIDAD = C.COD_CONTABILIDAD" _
       & " where C.cod_consolida = " & cbo.ItemData(cbo.ListIndex)
rs.Open strSQL, glogon.Conection, adOpenStatic
    lngCodEmpresa = rs!COD_CONTABILIDAD
    xMascara = rs!Nivel1 & rs!Nivel2 & rs!Nivel3 & rs!Nivel4 & rs!Nivel5
    iNivel1 = rs!Nivel1
rs.Close

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "ProGrX: Contabilidad"
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "yyyy/mm/dd") & "'"
 .Formulas(1) = "Usuario='" & glogon.Usuario & "'"
 .Formulas(2) = "SubTitulo='" & txtPeriodo & " [" & UCase(cbo.Text) & "]'"
 .Connect = glogon.ConectRPT

 Select Case Mid(cboReporte, 1, 2)
   Case "01" 'Balance General
     .ReportFileName = App.Path & "\ConBalanceGeneral.rpt"
     .SelectionFormula = "{CON_MOVIMIENTOS.MES} = " & txtMes & " AND {CON_MOVIMIENTOS.ANIO} = " _
                       & txtAnio & " AND {CON_MOVIMIENTOS.COD_CONSOLIDA} = " & cbo.ItemData(cbo.ListIndex)
     If optTipo(0).Value = True Then .SelectionFormula = .SelectionFormula & " AND {CUENTAS.CUENTA_MADRE} = ''"
     
     If fxVerificaPeriodoCierre(txtAnio, txtMes, lngCodEmpresa) Then
        Call sbUtilidad(txtAnio, txtMes, cbo.ItemData(cbo.ListIndex), lngCodEmpresa)
        Call sbTotalesCPA(txtAnio, txtMes, cbo.ItemData(cbo.ListIndex), lngCodEmpresa)
        .Formulas(3) = "Utilidad = " & gcurUtilidadMes
        .Formulas(4) = "Utilidad_Actual = " & gcurUtilidadActual
        .Formulas(5) = "Total_netopp = " & Totales.TOTAL_PPNETO + gcurUtilidadMes
        .Formulas(6) = "Total_ActualPP = " & Totales.TOTAL_PPACTUAL + gcurUtilidadActual
        
     Else
        Call sbTotalesCPA(txtAnio, txtMes, cbo.ItemData(cbo.ListIndex), lngCodEmpresa)
        .Formulas(3) = "Utilidad = 0"
        .Formulas(4) = "Utilidad_Actual = 0"
        .Formulas(5) = "Total_netopp = " & Totales.TOTAL_PPNETO
        .Formulas(6) = "Total_ActualPP = " & Totales.TOTAL_PPACTUAL
     End If
     
   
   Case "02"

     .ReportFileName = App.Path & "\ConEstadoResultados.rpt"

     Call sbUtilidadER(txtAnio, txtMes, cbo.ItemData(cbo.ListIndex), lngCodEmpresa)
     .Formulas(5) = "Utilidad = " & Utilidad.UB_GENERAL_MES + Utilidad.UN_GENERAL_MES
     .Formulas(6) = "Utilidad_Actual = " & Utilidad.UB_GENERAL_ACUMULADO + Utilidad.UN_GENERAL_ACUMULADO

     .Formulas(7) = "UtilidadX = " & Utilidad.UB_GENERAL_MES
     .Formulas(8) = "Utilidad_ActualX = " & Utilidad.UB_GENERAL_ACUMULADO
     .Formulas(9) = "Estado = 'ESTADO DE RESULTADOS CONSOLIDADOS'"
     
     .SelectionFormula = "{CNTX_CONTABILIDADES.COD_CONTABILIDAD} = " & lngCodEmpresa
     
     .SubreportToChange = "INGRESOS"
     .Connect = glogon.ConectRPT
     .SelectionFormula = "{CON_MOVIMIENTOS.COD_CONTABILIDAD} = {?Pm-CNTX_CONTABILIDADES.COD_CONTABILIDAD} AND {CON_MOVIMIENTOS.MES} = " & txtMes & " AND {CON_MOVIMIENTOS.ANIO} = " _
                       & txtAnio & " AND {CON_MOVIMIENTOS.COD_CONSOLIDA} = " & cbo.ItemData(cbo.ListIndex)
     If optTipo(0).Value = True Then .SelectionFormula = .SelectionFormula & " AND {CUENTAS.CUENTA_MADRE} = ''"
    
     .SelectionFormula = .SelectionFormula & " AND {TIPOS_CUENTAS.CLASIFICACION} = 'I' AND ISNULL({TIPOS_CUENTAS.ER}) = TRUE"

     .SubreportToChange = "GASTOS"
     .Connect = glogon.ConectRPT
     .SelectionFormula = "{CON_MOVIMIENTOS.COD_CONTABILIDAD} = {?Pm-CNTX_CONTABILIDADES.COD_CONTABILIDAD} AND {CON_MOVIMIENTOS.MES} = " & txtMes & " AND {CON_MOVIMIENTOS.ANIO} = " _
                       & txtAnio & " AND {CON_MOVIMIENTOS.COD_CONSOLIDA} = " & cbo.ItemData(cbo.ListIndex)
     If optTipo(0).Value = True Then .SelectionFormula = .SelectionFormula & " AND {CUENTAS.CUENTA_MADRE} = ''"
    
     .SelectionFormula = .SelectionFormula & " AND {TIPOS_CUENTAS.CLASIFICACION} = 'G' AND ISNULL({TIPOS_CUENTAS.ER}) = TRUE"

     .SubreportToChange = "OINGRESOS"
     .Connect = glogon.ConectRPT
     .SelectionFormula = "{CON_MOVIMIENTOS.COD_CONTABILIDAD} = {?Pm-CNTX_CONTABILIDADES.COD_CONTABILIDAD} AND {CON_MOVIMIENTOS.MES} = " & txtMes & " AND {CON_MOVIMIENTOS.ANIO} = " _
                       & txtAnio & " AND {CON_MOVIMIENTOS.COD_CONSOLIDA} = " & cbo.ItemData(cbo.ListIndex)
     If optTipo(0).Value = True Then .SelectionFormula = .SelectionFormula & " AND {CUENTAS.CUENTA_MADRE} = ''"
     
     .SelectionFormula = .SelectionFormula & " AND {TIPOS_CUENTAS.CLASIFICACION} = 'I' AND {TIPOS_CUENTAS.ER} = 'OI'"

     .SubreportToChange = "OGASTOS"
     .Connect = glogon.ConectRPT
     .SelectionFormula = "{CON_MOVIMIENTOS.COD_CONTABILIDAD} = {?Pm-CNTX_CONTABILIDADES.COD_CONTABILIDAD} AND {CON_MOVIMIENTOS.MES} = " & txtMes & " AND {CON_MOVIMIENTOS.ANIO} = " _
                       & txtAnio & " AND {CON_MOVIMIENTOS.COD_CONSOLIDA} = " & cbo.ItemData(cbo.ListIndex)
     If optTipo(0).Value = True Then .SelectionFormula = .SelectionFormula & " AND {CUENTAS.CUENTA_MADRE} = ''"
     
     .SelectionFormula = .SelectionFormula & " AND {TIPOS_CUENTAS.CLASIFICACION} = 'G' AND {TIPOS_CUENTAS.ER} = 'OG'"
     
   Case "03"
     .ReportFileName = App.Path & "\ConBalanceComprobacion.rpt"
     .Formulas(3) = "GrupoCuenta = mid({CUENTAS.COD_CUENTA},1," & iNivel1 & ")"
     .Formulas(4) = "Mascara = '" & xMascara & "'"
     .SelectionFormula = "{CON_MOVIMIENTOS.MES} = " & txtMes & " AND {CON_MOVIMIENTOS.ANIO} = " _
                       & txtAnio & " AND {CON_MOVIMIENTOS.COD_CONSOLIDA} = " & cbo.ItemData(cbo.ListIndex)
     If optTipo(0).Value = True Then .SelectionFormula = .SelectionFormula & " AND {CUENTAS.CUENTA_MADRE} = ''"
     
   
 End Select

 .PrintReport

End With

Screen.MousePointer = vbDefault

Exit Sub

vError:
    Screen.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Set Me.MouseIcon = frmContenedor.MouseIcon

Dim strSQL As String, rs As New ADODB.Recordset, vPaso As Boolean

 
 vPaso = False
 
 strSQL = "select * from CNTX_CONSOLIDA_DEFINICION"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 cbo.Clear
 
 Do While Not rs.EOF
   cbo.AddItem Format(rs!COD_CONSOLIDA, "000") & " - " & rs!Descripcion
   cbo.ItemData(cbo.NewIndex) = rs!COD_CONSOLIDA
   vPaso = True
   rs.MoveNext
 Loop
 
 If vPaso Then
   rs.MoveFirst
   cbo.Text = Format(rs!COD_CONSOLIDA, "000") & " - " & rs!Descripcion
 End If
 rs.Close

 txtMes = Month(fxFechaServidor)
 txtAnio = Year(fxFechaServidor)


cboReporte.Clear
cboReporte.AddItem "01 - Balance General"
cboReporte.AddItem "02 - Estado de Resultados"
cboReporte.AddItem "03 - Balance de Comprobación"

cboReporte.Text = "01 - Balance General"

End Sub

Private Sub txtAnio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdReporte.SetFocus
End Sub

Private Sub txtMes_Change()
On Error GoTo vError
txtPeriodo = fxCntX_PeriodoDesc(txtAnio, txtMes)
vError:
End Sub

Private Sub txtMes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAnio.SetFocus
End Sub

Private Sub txtAnio_Change()
On Error GoTo vError
txtPeriodo = fxCntX_PeriodoDesc(txtAnio, txtMes)
vError:
End Sub

