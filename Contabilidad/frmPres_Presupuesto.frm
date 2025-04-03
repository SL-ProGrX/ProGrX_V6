VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmPres_Presupuesto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Presupuesto"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10785
   HelpContextID   =   11
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboTipo 
      Height          =   315
      ItemData        =   "frmPres_Presupuesto.frx":0000
      Left            =   7200
      List            =   "frmPres_Presupuesto.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   120
      Width           =   1830
   End
   Begin VB.Frame fra 
      Height          =   1095
      Left            =   35
      TabIndex        =   4
      Top             =   4320
      Width           =   10695
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
         Height          =   375
         Left            =   9480
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Reporte"
         Height          =   375
         Left            =   9480
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtPorcentaje 
         Height          =   315
         Left            =   5400
         TabIndex        =   12
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cboProyeccion 
         Height          =   315
         ItemData        =   "frmPres_Presupuesto.frx":004E
         Left            =   3840
         List            =   "frmPres_Presupuesto.frx":005B
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtAnioCorte 
         Height          =   315
         Left            =   4440
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtMesCorte 
         Height          =   315
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblPeriodoCorte 
         Caption         =   "..."
         Height          =   315
         Left            =   5520
         TabIndex        =   17
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label lblPorcentaje 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5895
         TabIndex        =   13
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Proyectar y"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Periodo de Corte"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proyectar el Presupuesto de las Cuentas Marcadas a Otros Periodos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.TextBox txtAnio 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtMes 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3615
      Left            =   0
      TabIndex        =   18
      Top             =   600
      Width           =   10575
      _Version        =   524288
      _ExtentX        =   18653
      _ExtentY        =   6376
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
      MaxCols         =   489
      ScrollBars      =   2
      SpreadDesigner  =   "frmPres_Presupuesto.frx":0081
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   0
   End
   Begin VB.Label Label4 
      Caption         =   "Cuentas de Tipo"
      Height          =   255
      Left            =   5760
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblPeriodo 
      Caption         =   "..."
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Periodo"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmPres_Presupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vInicializa As Boolean
'
'Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer)
'Dim strSQL As String, rs As New ADODB.Recordset, i As Integer
'Dim rsTmp As New ADODB.Recordset
'Dim curDebitos As Currency, curCreditos As Currency
'
'Me.MousePointer = vbHourglass
'
'
'On Error GoTo vError
'
'vGrid.MaxCols = vGridMaxCol
'vGrid.MaxRows = 1
'
'vGrid.Row = vGrid.MaxRows
'
'strSQL = "select C.cod_cuenta,C.descripcion" _
'       & " from Cuentas C inner join tipos_cuentas T" _
'       & " on T.COD_CONTABILIDAD = C.COD_CONTABILIDAD and T.tipo_cuenta = C.tipo_cuenta" _
'       & " where C.acepta_movimientos = 'S' and C.COD_CONTABILIDAD = " & vParametros.CodigoEmpresa
'
'Select Case Mid(cboTipo.Text, 1, 2)
'  Case "01" 'Ingresos
'    strSQL = strSQL & " and T.clasificacion in('I','V')"
'  Case "02" 'Gastos
'    strSQL = strSQL & " and T.clasificacion in('G')"
'  Case "03" 'Capital
'    strSQL = strSQL & " and T.clasificacion in('C')"
'  Case "04" 'Activos
'    strSQL = strSQL & " and T.clasificacion in('A')"
'  Case "05" 'Pasivos
'    strSQL = strSQL & " and T.clasificacion in('P')"
'  Case "06" 'Orden
'    strSQL = strSQL & " and T.clasificacion in('O')"
'End Select
'
'strSQL = strSQL & " order by C.cod_cuenta"
'
'rs.CursorLocation = adUseServer
'Call OpenRecordSet(rs, strSQL, 0)
'
'Do While Not rs.EOF
'  strSQL = "select coalesce(Debitos,0) as Debitos,coalesce(creditos,0) as Creditos" _
'         & " from presupuesto where Anio = " & txtAnio & " and Mes = " & txtMes _
'         & " and COD_CONTABILIDAD = " & vParametros.CodigoEmpresa & " and Cod_cuenta = '" _
'         & rs!COD_Cuenta & "'"
'  rsTmp.Open strSQL, glogon.Conection, adOpenStatic
'  curDebitos = 0
'  curCreditos = 0
'  If Not rsTmp.EOF And Not rsTmp.BOF Then
'     curDebitos = rsTmp!debitos
'     curCreditos = rsTmp!creditos
'  End If
'
'  vGrid.Row = vGrid.MaxRows
'  For i = 1 To vGrid.MaxCols - 1
'    vGrid.col = i
'    Select Case i
'     Case 1
'        vGrid.Text = fxFormatoCuenta(True, CStr(rs.Fields(i - 1).Value))
'     Case 3 'Debitos
'        vGrid.Text = CStr(curDebitos)
'     Case 4 'Creditos
'        vGrid.Text = CStr(curCreditos)
'     Case Else
'        vGrid.Text = CStr(rs.Fields(i - 1).Value)
'    End Select
'  Next i
'  rsTmp.Close
'  vGrid.MaxRows = vGrid.MaxRows + 1
'  rs.MoveNext
'Loop
'
'rs.Close
'
'Me.MousePointer = vbDefault
'Exit Sub
'
'vError:
'  Me.MousePointer = vbDefault
'  MsgBox Err.Description, vbCritical
'End Sub
'
'
'Private Sub cboTipo_Click()
'Call sbCargaGridLocal(vGrid, vGrid.MaxCols)
'End Sub
'
'
'Private Function fxPorcentaje(vPorcentaje As Currency, curMonto As Currency, iPeriodos As Long) As Currency
'Dim curMontoResultado As Currency, iTiempos As Integer
'
'curMontoResultado = 0
'
'For iTiempos = 0 To iPeriodos
'   curMontoResultado = (curMonto - curMontoResultado) * vPorcentaje
'Next
'
'fxPorcentaje = curMontoResultado
'
'End Function
'
'Private Sub cmdAplicar_Click()
'Dim strSQL As String, rs As New ADODB.Recordset
'Dim lng As Long, curMovimiento As Currency
'Dim iMes As Integer, lngAnio As Long
'Dim iContador As Long, vPorcentaje As Currency
'Dim curMonto As Currency
'
''Hay que brincar los errores ya que la llave no permite
''Insertar valores a cuentas existentes y para no perder tiempo
''con verificaciones, mejor se saltan.
'On Error Resume Next
'
'If txtPorcentaje = "" Then txtPorcentaje = 0
'
'If CCur(txtPorcentaje) > 100 Or CCur(txtPorcentaje) < 0 Then
'  MsgBox "El porcentaje de proyeccion no es correcto...", vbCritical
'  Exit Sub
'End If
'
'vPorcentaje = CCur(txtPorcentaje) / 100
'
'For lng = 1 To vGrid.MaxRows
' vGrid.Row = lng
' vGrid.col = 1
'
' If vGrid.Text <> "" Then
'' strSQL = "select coalesce(count(*),0) as Existe from presupuesto" _
''        & " where COD_CONTABILIDAD = " & vParametros.CodigoEmpresa _
''        & " and cod_cuenta = '" & fxFormatoCuenta(False, vGrid.Text) & "'"
'' Call OpenRecordSet(rs, strSQL, 0)
'' If rs!existe = 0 Then
'
'    vGrid.col = 3
'    curMovimiento = CCur(vGrid.Text)
'    vGrid.col = 4
'    curMovimiento = curMovimiento + CCur(vGrid.Text)
'
'    If curMovimiento > 0 Then 'Guarda Solo Las cuentas que tienen movimiento
'        vGrid.col = 1
'        strSQL = "insert presupuesto(COD_CONTABILIDAD,cod_cuenta,anio,mes,debitos,creditos) values(" _
'               & vParametros.CodigoEmpresa & ",'" & fxFormatoCuenta(False, vGrid.Text) & "'," _
'               & txtAnio & "," & txtMes & ","
'        vGrid.col = 3
'        strSQL = strSQL & CCur(vGrid.Text) & ","
'        vGrid.col = 4
'        strSQL = strSQL & CCur(vGrid.Text) & ")"
'        Call ConectionExecute(strSQL, 0)
'
'        vGrid.col = 5
'        If CCur(vGrid.Text) = 1 Then
'            'Proyectar
'            Select Case Mid(cboProyeccion.Text, 1, 2)
'              Case "01" 'Mantener
'                iMes = txtMes
'                lngAnio = txtAnio
'
'                Do While Not ((lngAnio = CLng(txtAnioCorte)) And (iMes = CLng(txtMesCorte)))
'                    If iMes = 12 Then
'                      iMes = 1
'                      lngAnio = lngAnio + 1
'                    Else
'                      iMes = iMes + 1
'                    End If
'                    vGrid.col = 1
'                    strSQL = "insert presupuesto(COD_CONTABILIDAD,cod_cuenta,anio,mes,debitos,creditos) values(" _
'                           & vParametros.CodigoEmpresa & ",'" & fxFormatoCuenta(False, vGrid.Text) & "'," _
'                           & lngAnio & "," & iMes & ","
'                    vGrid.col = 3
'                    strSQL = strSQL & CCur(vGrid.Text) & ","
'                    vGrid.col = 4
'                    strSQL = strSQL & CCur(vGrid.Text) & ")"
'                    Call ConectionExecute(strSQL, 0)
'                Loop
'
'              Case "02" 'Incrementar
'
'                iMes = txtMes
'                lngAnio = txtAnio
'                iContador = 0
'                Do While Not ((lngAnio = txtAnioCorte) And (iMes = txtMesCorte))
'                    If iMes = 12 Then
'                      iMes = 1
'                      lngAnio = lngAnio + 1
'                    Else
'                      iMes = iMes + 1
'                    End If
'                    iContador = iContador + 1
'
'                    vGrid.col = 1
'                    strSQL = "insert presupuesto(COD_CONTABILIDAD,cod_cuenta,anio,mes,debitos,creditos) values(" _
'                           & vParametros.CodigoEmpresa & ",'" & fxFormatoCuenta(False, vGrid.Text) & "'," _
'                           & lngAnio & "," & iMes & ","
'                    vGrid.col = 3
'                    curMonto = fxPorcentaje(vPorcentaje, CCur(vGrid.Text), iContador)
'                    strSQL = strSQL & CCur(vGrid.Text) + curMonto & ","
'                    vGrid.col = 4
'                    curMonto = fxPorcentaje(vPorcentaje, CCur(vGrid.Text), iContador)
'                    strSQL = strSQL & CCur(vGrid.Text) + curMonto & ")"
'                    Call ConectionExecute(strSQL, 0)
'                Loop
'
'
'              Case "03" 'Disminuir
'
'                iMes = txtMes
'                lngAnio = txtAnio
'                iContador = 0
'                Do While Not ((lngAnio = txtAnioCorte) And (iMes = txtMesCorte))
'                    If iMes = 12 Then
'                      iMes = 1
'                      lngAnio = lngAnio + 1
'                    Else
'                      iMes = iMes + 1
'                    End If
'                    iContador = iContador + 1
'
'                    vGrid.col = 1
'                    strSQL = "insert presupuesto(COD_CONTABILIDAD,cod_cuenta,anio,mes,debitos,creditos) values(" _
'                           & vParametros.CodigoEmpresa & ",'" & fxFormatoCuenta(False, vGrid.Text) & "'," _
'                           & lngAnio & "," & iMes & ","
'                    vGrid.col = 3
'                    curMonto = fxPorcentaje(vPorcentaje, CCur(vGrid.Text), iContador)
'                    If (CCur(vGrid.Text) - curMonto) < 0 Then
'                      curMonto = 0
'                    Else
'                      curMonto = (CCur(vGrid.Text) - curMonto)
'                    End If
'                    strSQL = strSQL & curMonto & ","
'                    vGrid.col = 4
'                    curMonto = fxPorcentaje(vPorcentaje, CCur(vGrid.Text), iContador)
'                    If (CCur(vGrid.Text) - curMonto) < 0 Then
'                      curMonto = 0
'                    Else
'                      curMonto = (CCur(vGrid.Text) - curMonto)
'                    End If
'                    strSQL = strSQL & curMonto & ")"
'                    Call ConectionExecute(strSQL, 0)
'                Loop
'
'
'            End Select
'        End If 'Proyeccion
'
'
'    End If
'
'' End If 'Existe
'' rs.Close
'
' End If ' <> ""
'Next lng
'
'
'MsgBox "Cuentas guardadas y Proyectas Satisfactoriamente...", vbInformation
'
'
'End Sub
'
'Private Sub cmdReporte_Click()
'Screen.MousePointer = vbHourglass
'
'
'With frmContenedor.Crt
' .Reset
' .WindowShowGroupTree = True
' .WindowShowPrintSetupBtn = True
' .WindowShowRefreshBtn = True
' .WindowShowSearchBtn = True
' .WindowState = crptMaximized
' .WindowTitle = "ContaExpress"
' .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
' .Formulas(1) = "Empresa='" & vParametros.NombreEmpresa & "'"
' .Formulas(2) = "Usuario='" & glogon.Usuario & "'"
' .Formulas(3) = "Mascara='" & vParametros.MascaraCod & "'"
' .Formulas(4) = "SubTitulo='" & lblPeriodo.Caption & "'"
'
' .ReportFileName = App.Path & "\PreDefinicion.rpt"
' .SelectionFormula = "{PRESUPUESTO.ANIO} = " & txtAnio & " AND {PRESUPUESTO.MES} = " _
'              & txtMes & " AND {PRESUPUESTO.COD_CONTABILIDAD} = " & vCodEmpresa & " AND {CUENTAS.ACEPTA_MOVIMIENTOS}='S'"
' .Connect = glogon.ConectRPT
'
' .PrintReport
'
'End With
'
'Screen.MousePointer = vbDefault
'
'End Sub
'
'Private Sub txtMes_Change()
'On Error GoTo vError
'lblPeriodo.Caption = fxPeriodoRes(txtAnio, txtMes)
'If Not vInicializa Then Call sbCargaGridLocal(vGrid, vGrid.MaxCols)
'vError:
'End Sub
'
'Private Sub txtMes_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAnio.SetFocus
'End Sub
'
'Private Sub txtAnio_Change()
'On Error GoTo vError
'lblPeriodo.Caption = fxPeriodoRes(txtAnio, txtMes)
'If Not vInicializa Then Call sbCargaGridLocal(vGrid, vGrid.MaxCols)
'vError:
'End Sub
'
'Private Sub txtMesCorte_Change()
'On Error GoTo vError
'lblPeriodoCorte.Caption = fxPeriodoRes(txtAnioCorte, txtMesCorte)
'vError:
'End Sub
'
'Private Sub txtMesCorte_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAnioCorte.SetFocus
'End Sub
'
'Private Sub txtAnioCorte_Change()
'On Error GoTo vError
'lblPeriodoCorte.Caption = fxPeriodoRes(txtAnioCorte, txtMesCorte)
'vError:
'End Sub
'


Private Sub Form_Load()
Set Me.MouseIcon = frmContenedor.MouseIcon
vGrid.MaxCols = 5
vGrid.MaxRows = 1
'Solo cuentas de Gastos que acepten movimientos
'Ver si es posible a futuro repartir presupuestos por cierres y por porcentajes
'incrementales

vInicializa = True

txtMes = Month(fxFechaServidor)
txtAnio = Year(fxFechaServidor)

txtMesCorte = txtMes
txtAnioCorte = txtAnio

cboTipo.Clear
cboTipo.AddItem "01 - Ingresos"
cboTipo.AddItem "02 - Gastos"
cboTipo.AddItem "03 - Capital"
cboTipo.AddItem "04 - Activos"
cboTipo.AddItem "05 - Pasivos"
cboTipo.AddItem "06 - Orden"

cboTipo.Text = "01 - Ingresos"

txtPorcentaje = 0

cboProyeccion.Clear
cboProyeccion.AddItem "01 - Mantener"
cboProyeccion.AddItem "02 - Incrementar"
cboProyeccion.AddItem "03 - Disminuir"

cboProyeccion.Text = "01 - Mantener"


vInicializa = False


End Sub



