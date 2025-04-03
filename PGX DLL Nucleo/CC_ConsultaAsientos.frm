VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSIF_AsientosConsultaCtaCor 
   Caption         =   "Consulta : Cola de Asientos"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
   Icon            =   "CC_ConsultaAsientos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   9960
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   16
      Top             =   5745
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   7335
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox chkPendiente 
         Alignment       =   1  'Right Justify
         Caption         =   "Pendientes"
         ForeColor       =   &H00800000&
         Height          =   365
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox chkGenerado 
         Alignment       =   1  'Right Justify
         Caption         =   "Generados"
         ForeColor       =   &H00800000&
         Height          =   365
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtCaso 
         Height          =   315
         Left            =   840
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkTodos 
         Alignment       =   1  'Right Justify
         Caption         =   "Todos los Casos"
         ForeColor       =   &H00800000&
         Height          =   365
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   315
         Left            =   3600
         TabIndex        =   9
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   89522179
         CurrentDate     =   36462
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   315
         Left            =   3600
         TabIndex        =   10
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   89522179
         CurrentDate     =   36462
      End
      Begin VB.Label lblTipo 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblDesde 
         Caption         =   "Inicio"
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblHasta 
         Caption         =   "Corte"
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblCaso 
         Caption         =   "Caso"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   735
      End
      Begin VB.Image imgBuscaRapido 
         Height          =   240
         Index           =   0
         Left            =   2640
         Picture         =   "CC_ConsultaAsientos.frx":6852
         Stretch         =   -1  'True
         Top             =   600
         Width           =   225
      End
   End
   Begin MSComctlLib.ListView lswAsientos 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5953
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Usuario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cuenta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descripcion"
         Object.Width           =   5468
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Débitos"
         Object.Width           =   3775
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Créditos"
         Object.Width           =   3775
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Fecha"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Caso"
         Object.Width           =   2893
      EndProperty
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Reporte"
      Height          =   1000
      Left            =   8640
      Picture         =   "CC_ConsultaAsientos.frx":711C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   1000
      Left            =   7560
      Picture         =   "CC_ConsultaAsientos.frx":D96E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   90
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle de Asientos"
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
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   9735
   End
End
Attribute VB_Name = "frmSIF_AsientosConsultaCtaCor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function DescripcionCuenta(strCuenta As String) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "Select * from cuentas where cod_cuenta='" & strCuenta & "'"

rsX.Open strSQL, glogon.Conection, adOpenStatic

If Not rsX.EOF And Not rsX.BOF Then
   DescripcionCuenta = rsX!descripcion
Else
   DescripcionCuenta = ""
End If

rsX.Close

End Function


Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCaso.SetFocus
vError:
End Sub

Private Sub chkTodos_Click()
If chkTodos.Value = vbChecked Then
  txtCaso.Enabled = False
  dtpDesde.Enabled = False
  dtpHasta.Enabled = False
Else
  txtCaso.Enabled = True
  dtpDesde.Enabled = True
  dtpHasta.Enabled = True
End If
End Sub

Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem, curDebe As Currency, curHaber As Currency


Me.MousePointer = vbHourglass

lswAsientos.ListItems.Clear

strSQL = "Select A.*,C.DESCRIPCION from Asientos_TMP A LEFT JOIN CUENTAS C " _
       & " ON A.TMP_CUENTA = C.COD_CUENTA where TMP_TIPO="

Select Case Trim(cboTipo)
  Case "Ingreso"
       strSQL = strSQL & "'ING'"
  Case "Liquidacion"
       strSQL = strSQL & "'LIQ'"
  Case "Procesos"
       strSQL = strSQL & "'PRM'"
  Case "Traspaso"
       strSQL = strSQL & "'TRA'"
  Case "Formalizacion"
       strSQL = strSQL & "'FRM'"
  Case "Forma.Anulada"
       strSQL = strSQL & "'AFR'"
  Case "Vivienda"
       strSQL = strSQL & "'VIV'"
End Select

If chkTodos = 0 Then
  If Trim(txtCaso) <> "" Then
     strSQL = strSQL & " and TMP_CASO like '" & Trim(txtCaso) & "%'"
  Else
     strSQL = strSQL & " and TMP_FECHA between '" & Format(dtpDesde, "yyyy/mm/dd") _
            & " 00:00:00' and '" & Format(dtpHasta, "yyyy/mm/dd") & " 23:59:59'"
  End If
End If


If (chkPendiente = vbChecked And chkGenerado = vbChecked) _
 Or (chkPendiente = vbChecked And chkGenerado = vbChecked) Then
 'Nada, debe de mostrar todo tipo de asiento generados y no a contabilidad
Else
 'preguntar si es pendiente o generado el que se quiere
 If chkPendiente = vbChecked Then strSQL = strSQL & " and tmp_fechatrp is null"
 If chkGenerado = vbChecked Then strSQL = strSQL & " and tmp_fechatrp is not null"
End If

rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1

Do While Not rs.EOF
   Set itmX = lswAsientos.ListItems.Add(, , rs!TMP_USUARIO)
       itmX.Tag = itmX.Index
       itmX.SubItems(1) = Format(rs!tmp_cuenta, GLOBALES.gstrMascara)
       itmX.SubItems(2) = rs!descripcion & ""  'DescripcionCuenta(Trim(rs!TMP_CUENTA))
       
       If rs!tmp_debehaber = "D" Then
          itmX.SubItems(3) = Format(rs!tmp_monto, "Standard")
          itmX.SubItems(4) = "0"
          curDebe = curDebe + rs!tmp_monto
       Else
          itmX.SubItems(4) = Format(rs!tmp_monto, "Standard")
          itmX.SubItems(3) = "0"
          curHaber = curHaber + rs!tmp_monto
       End If
       itmX.SubItems(5) = IIf(IsNull(rs!tmp_fecha), "", Format(rs!tmp_fecha, "dd/mm/yyyy"))
       itmX.SubItems(6) = IIf(IsNull(rs!tmp_caso), "", rs!tmp_caso)
 prgBar.Value = prgBar.Value + 1
 rs.MoveNext
Loop
rs.Close

prgBar.Value = 1
prgBar.Max = 1

Set itmX = lswAsientos.ListItems.Add(lswAsientos.ListItems.Count + 1, , "")
Set itmX = lswAsientos.ListItems.Add(lswAsientos.ListItems.Count + 1, , "Totales:")
    itmX.SubItems(3) = Format(curDebe, "Standard")
    itmX.SubItems(4) = Format(curHaber, "Standard")
    itmX.ForeColor = vbBlue
    itmX.Bold = True

Me.MousePointer = vbDefault

End Sub

Private Sub cmdReporte_Click()
Dim strSQL As String

Me.MousePointer = vbHourglass


With frmContenedor.Crt
 .Reset
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Reportes - Cola de Asientos"

 .Connect = glogon.ConectRPT

 .ReportFileName = SIFGlobal.fxSIFPathReportes("SIFAsientosTMP.rpt")

 strSQL = "{ASIENTOS_TMP.TMP_TIPO}="

 Select Case Trim(cboTipo)
  Case "Ingreso"
       strSQL = strSQL & "'ING'"
  Case "Liquidacion"
       strSQL = strSQL & "'LIQ'"
  Case "Procesos"
       strSQL = strSQL & "'PRM'"
  Case "Traspaso"
       strSQL = strSQL & "'TRA'"
  Case "Formalizacion"
       strSQL = strSQL & "'FRM'"
  Case "Forma.Anulada"
       strSQL = strSQL & "'AFR'"
  Case "Vivienda"
       strSQL = strSQL & "'VIV'"
 End Select

 If chkTodos = vbUnchecked Then
  If Trim(txtCaso) <> "" Then
     strSQL = strSQL & " and {ASIENTOS_TMP.TMP_CASO}='" & Trim(txtCaso) & "'"
  Else
     strSQL = strSQL & " and Cdate({ASIENTOS_TMP.TMP_FECHA}) in date('" & Format(dtpDesde.Value, "yyyy,mm,dd") _
            & "') to date('" & Format(dtpDesde.Value, "yyyy,mm,dd") & "')"
            
            
             
  End If
 End If

If (chkPendiente = vbChecked And chkGenerado = vbChecked) _
 Or (chkPendiente = vbChecked And chkGenerado = vbChecked) Then
 'Nada, debe de mostrar todo tipo de asiento generados y no a contabilidad
Else
 'preguntar si es pendiente o generado el que se quiere
 If chkPendiente = vbChecked Then strSQL = strSQL & " AND isnull({ASIENTOS_TMP.TMP_FECHATRP})"
 If chkGenerado = vbChecked Then strSQL = strSQL & " AND not isnull({ASIENTOS_TMP.TMP_FECHATRP})"
End If

 .Formulas(0) = "fxFecha = '" & Format(fxFechaServidor, "dd/mm/yyyy hh:mm:ss") & "'"
 .Formulas(1) = "fxUsuario = '" & glogon.Usuario & "'"
 .Formulas(2) = "Empresa = '" & GLOBALES.gstrNombreEmpresa & "'"
 .SelectionFormula = strSQL
 .PrintReport
End With

Me.MousePointer = vbDefault

End Sub



Private Sub dtpDesde_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpHasta.SetFocus
vError:
End Sub

Private Sub dtpHasta_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdBuscar.SetFocus
vError:
End Sub


Private Sub Form_Load()

With cboTipo
     .AddItem "Ingreso"           'ING
     .AddItem "Liquidacion"       'LIQ
     .AddItem "Procesos"          'PRM
     .AddItem "Traspaso"          'TRA
     .AddItem "Formalizacion"     'FRM
     .AddItem "Forma.Anulada"     'AFR
     .AddItem "Vivienda"          'VIV
     .Text = "Ingreso"
End With

dtpDesde.Value = fxFechaServidor
dtpHasta.Value = dtpDesde.Value

Me.Height = 6450
Me.Width = 10200

End Sub


Private Sub Form_Resize()
On Error Resume Next

lbl.Width = Me.Width - 420
lswAsientos.Width = lbl.Width

lswAsientos.Height = (Me.Height - lswAsientos.Top - prgBar.Height - 550)

End Sub

Private Sub imgBuscaRapido_Click(Index As Integer)
    
MsgBox "Las Busquedas Toman como primeros parametros las Fechas y el Tipo de Movimiento " _
    & "que tiene actualmente especificados...", vbInformation
    
    gBusquedas.Filtro = " and TMP_FECHA between '" & Format(dtpDesde.Value, "yyyy/mm/dd") & "' and '" _
                & Format(dtpHasta.Value, "yyyy/mm/dd") & "'"
    gBusquedas.Convertir = "N"
                
Select Case Trim(cboTipo)
  Case "Ingreso"
       gBusquedas.Filtro = gBusquedas.Filtro & " and tmp_tipo = 'ING'"
  Case "Liquidacion"
       gBusquedas.Filtro = gBusquedas.Filtro & " and tmp_tipo = 'LIQ'"
  Case "Proceso Mensual"
       gBusquedas.Filtro = gBusquedas.Filtro & " and tmp_tipo = 'PRM'"
  Case "Traspaso"
       gBusquedas.Filtro = gBusquedas.Filtro & " and tmp_tipo = 'TRA'"
  Case "Abonos"
       gBusquedas.Filtro = gBusquedas.Filtro & " and tmp_tipo = 'ABO'"
  Case "Anulacion Abonos"
       gBusquedas.Filtro = gBusquedas.Filtro & " and tmp_tipo = 'ANU'"
  Case "Anulacion Aportes"
       gBusquedas.Filtro = gBusquedas.Filtro & " and tmp_tipo = 'ANH'"
End Select
                
 '     gbusquedas.filtro = gbusquedas.filtro & " group by tmp_caso"
                
    gBusquedas.Resultado = Trim(txtCaso)
    gBusquedas.Consulta = "Select tmp_caso,tmp_tipo,Tmp_monto,Tmp_fecha,TMP_ESTADO_ASIENTO,tmp_usuario From asientos_tmp"
    gBusquedas.Columna = "tmp_caso"
    gBusquedas.Orden = "tmp_caso"
    frmBusquedas.Show vbModal
    txtCaso = gBusquedas.Resultado
End Sub


Private Sub txtCaso_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpDesde.SetFocus
vError:

End Sub
