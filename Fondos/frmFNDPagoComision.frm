VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.ShortcutBar.v19.1.0.ocx"
Begin VB.Form frmFNDPagoComision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago Automatico de Comisiones (Tesoreria)"
   ClientHeight    =   7308
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   11652
   Icon            =   "frmFNDPagoComision.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7308
   ScaleWidth      =   11652
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4212
      Left            =   0
      TabIndex        =   12
      Top             =   1920
      Width           =   11652
      _Version        =   1245185
      _ExtentX        =   20553
      _ExtentY        =   7429
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.GroupBox gbAccion 
      Height          =   852
      Left            =   0
      TabIndex        =   3
      Top             =   6240
      Width           =   11652
      _Version        =   1245185
      _ExtentX        =   20553
      _ExtentY        =   1503
      _StockProps     =   79
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   492
         Left            =   8400
         TabIndex        =   4
         Top             =   240
         Width           =   1572
         _Version        =   1245185
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Informe"
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
         Picture         =   "frmFNDPagoComision.frx":030A
      End
      Begin XtremeSuiteControls.PushButton cmdGenera 
         Height          =   492
         Left            =   9960
         TabIndex        =   5
         Top             =   240
         Width           =   1572
         _Version        =   1245185
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Generar"
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
         Picture         =   "frmFNDPagoComision.frx":0AC6
      End
      Begin XtremeSuiteControls.DateTimePicker dtpReporte 
         Height          =   312
         Left            =   6480
         TabIndex        =   11
         Top             =   240
         Width           =   1332
         _Version        =   1245185
         _ExtentX        =   2350
         _ExtentY        =   550
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Generación:"
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
         Index           =   2
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   1332
      End
   End
   Begin VB.Frame fraComision 
      Caption         =   "Cargando y Calculando Información [Espere...]"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3240
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   5295
      Begin MSComctlLib.ProgressBar prgBar 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   5055
         _ExtentX        =   8911
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   372
      Left            =   3600
      TabIndex        =   2
      Top             =   1440
      Width           =   1212
      _Version        =   1245185
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmFNDPagoComision.frx":129E
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   840
      TabIndex        =   7
      Top             =   1440
      Width           =   1332
      _Version        =   1245185
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
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
      Left            =   2160
      TabIndex        =   8
      Top             =   1440
      Width           =   1332
      _Version        =   1245185
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   312
      Left            =   5040
      TabIndex        =   13
      Top             =   1440
      Width           =   6492
      _Version        =   1245185
      _ExtentX        =   11451
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   576
      Left            =   0
      TabIndex        =   10
      Top             =   1320
      Width           =   11652
      _Version        =   1245185
      _ExtentX        =   20553
      _ExtentY        =   1023
      _StockProps     =   14
      Caption         =   "Fechas: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   3
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pago de Comisiones de Colocación de Planes:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   0
      Left            =   1920
      TabIndex        =   9
      Top             =   360
      Width           =   7332
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   12252
   End
End
Attribute VB_Name = "frmFNDPagoComision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxMaestroTesoreria(vTipoDocumento As String, vBanco As Long, vMonto As Currency, vCodigo As String _
                              , vBeneficiario As String, vOP As Long, vDetalle1 As String, vReferencia As Long _
                              , vDetalle2 As String, vCuenta As String, vFecha As Date) As Long  'Regresa el NSOLICITUD
Dim rsX As New ADODB.Recordset, strSQL As String
Dim lngSol As Long

strSQL = "insert Tes_Transacciones(id_banco,tipo,codigo,beneficiario,monto,fecha_solicitud,estado,estadoi" _
       & ",modulo,submodulo,cta_ahorros,detalle1,detalle2,referencia,op,genera,actualiza) values(" & vBanco _
       & ",'" & vTipoDocumento & "','" & vCodigo & "','" & vBeneficiario & "'," & vMonto _
       & ",'" & Format(vFecha, "mm/dd/yyyy") & "','P','P','TE','N','" & vCuenta _
       & "','" & vDetalle1 & "','" & vDetalle2 & "'," & vReferencia & "," & vOP & ",'S','S')"
Call ConectionExecute(strSQL)

strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones"
Call OpenRecordSet(rsX, strSQL, 0)
 strSQL = "select * from Tes_Transacciones where nsolicitud = " & rsX!solicitud
rsX.Close

lngSol = 0

Call OpenRecordSet(rsX, strSQL, 0)
If Trim(rsX!Codigo) = Trim(vCodigo) Then lngSol = rsX!NSolicitud
rsX.Close

If lngSol = 0 Then
  strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones where codigo ='" & vCodigo _
         & "' and op = " & vOP
  rsX.CursorLocation = adUseServer
  Call OpenRecordSet(rsX, strSQL, 0)
  lngSol = rsX!solicitud
  rsX.Close
End If

fxMaestroTesoreria = lngSol

End Function

Private Sub sbCreaDetalle(vSolicitud As Long, vCtaConta As String, vMonto As Currency, vDH As String, vLinea As Integer)
Dim strSQL As String

strSQL = "insert Tes_Trans_Asiento(nsolicitud,cuenta_contable,monto,debehaber,linea) values(" _
       & vSolicitud & ",'" & Trim(vCtaConta) & "'," & vMonto & ",'" & vDH _
       & "'," & vLinea & ")"
Call ConectionExecute(strSQL)

End Sub

Private Function fxCuentaBanco(vBanco) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select ctaconta from Tes_Bancos where id_banco = " & vBanco
Call OpenRecordSet(rsX, strSQL, 0)
    fxCuentaBanco = Trim(rsX!ctaConta)
rsX.Close

End Function

Private Sub cmdBuscar_Click()
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListViewItem, vPorcComision As Double

Me.MousePointer = vbHourglass

fraComision.Visible = True

lsw.ListItems.Clear

strSQL = "select V.cod_vendedor, V.cedula,V.nombre,V.cuenta_ahorros,V.Tipo_Pago" _
       & ",V.cod_banco,V.minimo,V.porc_comision,sum(C.monto) as Monto, count(*) as Casos " _
       & "from fnd_contratos C inner join fnd_vendedores V on C.cod_vendedor = V.cod_vendedor " _
       & "where C.ind_comision = 0 and C.estado <> 'L' " _
       & "and C.fecha_inicio between '" & Format(dtpInicio, "yyyy/mm/dd") & " 01:00:00' and '" _
       & Format(dtpCorte, "yyyy/mm/dd") & " 23:59:59' and V.aplica_comision = 1 " _
       & "group by V.cod_vendedor, V.cedula,V.nombre,V.cuenta_ahorros,V.Tipo_Pago,V.cod_banco,V.minimo,V.porc_comision"
       
       
'    strSQL = "select F.cod_contrato,F.Cedula,S.Nombre,F.Estado,F.Plazo,F.Monto" _
'           & ",F.Aportes,F.Rendimiento,F.Fecha_Corte,F.Fecha_Inicio" _
'           & ",dbo.fxSys_Cuentas_Bancarias(F.cedula,B.id_Banco,0) as CuentaAhorroX" _
'           & ",B.id_Banco as BancoX,B.descripcion as BancoDesc,Est.Descripcion as 'EstadoDesc'" _
'           & " from Fnd_Contratos F inner join Socios S on F.Cedula = S.Cedula" _
'           & " inner join Fnd_Planes Pln on F.cod_Operadora = Pln.Cod_Operadora and F.cod_Plan = Pln.cod_Plan " _
'           & " inner join AFI_ESTADOS_PERSONA Est on S.estadoActual = Est.cod_Estado" _
'           & " inner join Tes_Bancos B on B.id_Banco = " & cboBanco.ItemData(cboBanco.ListIndex)
'
'    If cboCuentaFiltro.Text <> "TODOS" Then
'       If cboCuentaFiltro.Text = "Cuenta Interna" Then
'            strSQL = strSQL & " inner join vSys_Personas_Cuenta_Bancaria_Local Cta on F.cedula = Cta.Identificacion" _
'                   & " and Cta.cod_Banco = B.cod_Grupo and Cta.cod_Divisa = Pln.Cod_Moneda"
'       Else
'            strSQL = strSQL & " inner join vSys_Personas_Cuenta_Bancaria_Interbancaria Cta on F.cedula = Cta.Identificacion" _
'                   & " and Cta.cod_Banco = B.cod_Grupo and Cta.cod_Divisa = Pln.Cod_Moneda"
'       End If
       
Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1

Do While Not rs.EOF
  If rs!Monto > 0 Then
    Set itmX = lsw.ListItems.Add(, , rs!cod_vendedor)
        itmX.Tag = rs!cod_vendedor
        itmX.SubItems(1) = rs!Cedula
        itmX.SubItems(2) = rs!Nombre
        itmX.SubItems(3) = Format((rs!Monto * (rs!porc_comision / 100)), "Standard")
        itmX.SubItems(4) = "" ' fxgFNDBancos("D", rs!cod_banco)
        itmX.SubItems(5) = rs!Cuenta_Ahorros & ""
        itmX.SubItems(6) = rs!TIPO_PAGO
        itmX.SubItems(7) = rs!cod_banco
        itmX.SubItems(8) = rs!porc_comision
     
  End If
  
  If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
  
  rs.MoveNext
Loop
rs.Close

'Activa la generación y Reportes
If lsw.ListItems.Count > 0 Then
  cmdGenera.Enabled = True
Else
  cmdGenera.Enabled = False
End If

Me.MousePointer = vbDefault

fraComision.Visible = False

End Sub

Private Sub cmdGenera_Click()
Dim lngSolicitud As Long, lng As Long, vFecha As Date
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCuenta As String, vPorcComision As Double


On Error GoTo vError

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

fraComision.Visible = True
prgBar.Max = lsw.ListItems.Count
 
strSQL = "select cta_comisiones from fnd_parametros"
Call OpenRecordSet(rs, strSQL)
  vCuenta = Trim(rs!cta_comisiones)
rs.Close
 
 
''Inicia Transacción
'glogon.Conection.BeginTrans
'
'For lng = 1 To lsw.ListItems.Count
'
'  prgBar.Value = lng
'
'  lsw.SelectedItem = lsw.ListItems(lng)
'
'  If lsw.SelectedItem.Checked And CCur(lsw.SelectedItem.SubItems(2)) > 0 Then
'   lngSolicitud = fxMaestroTesoreria(lsw.SelectedItem.SubItems(5), lsw.SelectedItem.SubItems(6) _
'                , CCur(lsw.SelectedItem.SubItems(2)), lsw.SelectedItem.Text _
'                , lsw.SelectedItem.SubItems(1), 0, "FONDOS EXTRAORD.", 0 _
'                , "PAGO COMISION", lsw.SelectedItem.SubItems(4), vFecha)
'
'   Call sbCreaDetalle(lngSolicitud, fxCuentaBanco(lsw.SelectedItem.SubItems(6)), CCur(lsw.SelectedItem.SubItems(2)), "H", 1)
'   Call sbCreaDetalle(lngSolicitud, vCuenta, CCur(lsw.SelectedItem.SubItems(2)), "D", 2)
'
'
'   'Actualiza Contratos indicando que ya se proceso la comisión para este vendedor
'   strSQL = "update ht_contratos set ind_comision = 1,comision_fecha = '" & Format(vFecha, "yyyy/mm/dd") _
'          & "',comision_Tesoreria = " & lngSolicitud & ",comision_monto = monto * " & (lsw.SelectedItem.SubItems(7) / 100) _
'          & " where cod_vendedor = '" & lsw.SelectedItem.Text _
'          & "' and ind_comision = 0 and fecha_contrato between '" _
'          & Format(dtpInicio, "yyyy/mm/dd") & " 01:00:00' and '" _
'          & Format(dtpCorte, "yyyy/mm/dd") & " 23:59:59' AND ESTADO <> 'L'"
'   Call ConectionExecute(strSQL)
'
'
'  End If
'Next lng
'
'
''Cierra Transacción
'glogon.Conection.CommitTrans
  
Me.MousePointer = vbDefault


'Actualiza Información
Call cmdBuscar_Click

MsgBox "Comisiones Generadas a Tesoreria...", vbInformation

Exit Sub

vError:
'Reversa Transacción
 glogon.Conection.RollbackTrans
 
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cmdReporte_Click()

Me.MousePointer = vbHourglass


With frmContenedor.Crt
  .Reset
  .WindowShowGroupTree = True
  .WindowShowPrintSetupBtn = True
  .WindowShowRefreshBtn = True
  .WindowShowSearchBtn = True
  .WindowState = crptMaximized
  .WindowTitle = "Hotelería"

  .Formulas(0) = "Empresa = '" & GLOBALES.gstrNombreEmpresa & "'"
  .Formulas(1) = "Usuario = '" & glogon.Usuario & "'"
  .Formulas(2) = "Fecha = '" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
  .Formulas(3) = "SubTitulo = 'Generadas a el día " & Format(dtpReporte, "dd/mm/yyyy") & "'"
  
  .ReportFileName = SIFGlobal.fxPathReportes("Comision_Generadas.rpt")
  
  .SelectionFormula = "{FND_CONTRATOS.COMISION_FECHA} = Datetime(" & Year(dtpReporte) & "," _
                    & Month(dtpReporte) & "," & Day(dtpReporte) & ")"
  .PrintReport

End With

Me.MousePointer = vbDefault

End Sub

Private Sub dtpCorte_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdBuscar.SetFocus
End Sub

Private Sub dtpInicio_Change()
lsw.ListItems.Clear
cmdGenera.Enabled = False
End Sub

Private Sub dtpCorte_Change()
lsw.ListItems.Clear
cmdGenera.Enabled = False
End Sub

Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpCorte.SetFocus
End Sub

Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 18 'Fondo de Inversion

strSQL = "select B.id_banco as 'Idx',B.descripcion as 'ItmX'" _
       & " from tes_banco_asg T inner join Tes_Bancos B on T.id_banco = B.id_banco" _
       & " where T.nombre = '" & glogon.Usuario & "' and B.Estado = 'A'"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)


 dtpInicio.Value = fxFechaServidor
 dtpCorte.Value = dtpInicio.Value
 dtpReporte.Value = dtpInicio.Value
 
 Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture
 
 With lsw.ColumnHeaders
    .Clear
    .Add , , "Agente Id", 1440
    .Add , , "Identificación", 1400
    .Add , , "Nombre", 4000
    .Add , , "Comisión", 1200, vbRightJustify
    .Add , , "Banco", 2000
    .Add , , "Cuenta", 2000
    .Add , , "Emitir", 900, vbCenter
    .Add , , "Banco Id", 100
 End With
 
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
End Sub
