VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCR_AbonosComprobante 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Generacion de Comprobantes de Abonos (Recibos)"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumDoc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   21
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmCR_AbonosComprobante.frx":0000
      Left            =   2280
      List            =   "frmCR_AbonosComprobante.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton CmdAbono 
      Caption         =   "&Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      Picture         =   "frmCR_AbonosComprobante.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtProceso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      TabIndex        =   11
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame fraAbono 
      BorderStyle     =   0  'None
      Caption         =   "Tipo de Abono"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   8655
      Begin VB.OptionButton optAbono 
         BackColor       =   &H80000003&
         Caption         =   "Extra Ordinario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optAbono 
         BackColor       =   &H80000003&
         Caption         =   "Cancelación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optAbono 
         BackColor       =   &H80000003&
         Caption         =   "Ordinario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optAbono 
         BackColor       =   &H80000003&
         Caption         =   "Adelantos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo de Abono   >"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtCedula 
      Height          =   315
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   4
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Nombre Completo del Socio (Apellidos y Nombre)"
      Top             =   840
      Width           =   5415
   End
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   2
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtOperacion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   240
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   3828
      Width           =   9144
      _ExtentX        =   16140
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7832
            MinWidth        =   7832
            Object.ToolTipText     =   "Oficina"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Linea"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Recurso"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Num.- Doc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   20
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo - Doc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   19
      Top             =   2520
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   8400
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Identifica."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   360
      TabIndex        =   17
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2760
      TabIndex        =   16
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Línea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   360
      TabIndex        =   15
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1080
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblOpex 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   7560
      TabIndex        =   13
      Top             =   1200
      Width           =   615
   End
End
Attribute VB_Name = "frmCR_AbonosComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vOperacion As Long, vCuotasDeducidas As Integer, vCuotasDirectas As Integer
Dim vInteres As Currency, vPlazo As Integer, vSaldoMes As Currency, vUltimoRecibo As Long
Dim vRetencion As Boolean, vBaseCalculo As String, vPrideduc As Long, vAnticipoPorc As Currency, vAnticipoMeses As Integer
Dim vDiasActivo As Long, vFechaHoy As Date


Private Sub sbAbono()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lngRecibo As Long, vCuenta As String
Dim vTipo As String, vFecha As Date
Dim i As Integer, vExtraOrdinario As Boolean


Me.MousePointer = vbHourglass

On Error GoTo vError


vFecha = fxFechaServidor
vExtraOrdinario = False

vTipo = fxTipoASEDoc(cboTipo.Text)


vCuenta = Trim(fxDocumentoCuenta(vTipo))

lngRecibo = txtNumDoc.Text

vUltimoRecibo = lngRecibo


If vAseDocValido = False Then
  MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
        & " válida para esta operación...", vbCritical
  Exit Sub
End If

'Genera el Comprobante
Select Case True
  Case optAbono(0) 'Abono Ordinario
      If uRecibos Then lngRecibo = fxDocumentoAbono("ABONO ORDINARIO", vTipo, CStr(lngRecibo), "CRD001", vCuenta)
  Case optAbono(1) 'Abono Extraordinario
      If uRecibos Then lngRecibo = fxDocumentoAbono("ABONO EXTRAORDINARIO", vTipo, CStr(lngRecibo), "CRD002", vCuenta)
  Case optAbono(2) 'Abono De Cancelacion
      If uRecibos Then lngRecibo = fxDocumentoAbono("CANCELACION DE DEUDA", vTipo, CStr(lngRecibo), "CRD003", vCuenta)
  Case optAbono(3) 'Adelanto de Cuotas
      If uRecibos Then lngRecibo = fxDocumentoAbono("ADELANTO DE CUOTAS", vTipo, CStr(lngRecibo), "CRD004", vCuenta)
End Select



'IMPRIMIR RECIBO
If lngRecibo > 0 Then Call sbImprimeRecibo(lngRecibo, vTipo)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Function fxVerifica() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

Dim vMensaje As String, i As Integer

vMensaje = ""

strSQL = "select count(*) as existe from sif_transacciones where tipo_documento = '" & fxTipoASEDoc(cboTipo.Text) _
      & "' and  cod_transaccion = '" & txtNumDoc.Text & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then

  vMensaje = "Ya existe el comprobante creado, verifique...!"

End If
rs.Close


If Len(vMensaje) = 0 Then
  fxVerifica = True
Else
  fxVerifica = False
  MsgBox vMensaje, vbExclamation
End If

End Function


Private Sub CmdAbono_Click()
Dim iRespuesta As Integer

If Not fxVerifica Then Exit Sub

 iRespuesta = MsgBox("Esta seguro de confeccionar el documento...", vbYesNo)
 If iRespuesta = vbYes Then
  
  Call sbAbono
  If vAseDocValido Then MsgBox "Comprobante de Abono Realizado ... " & cboTipo.Text & " #" & vUltimoRecibo, vbInformation
  Call sbConsultaOperacion
 
 Else 'Respuesta
  
  MsgBox "Transacción Cancelada...", vbInformation
 
 End If

End Sub

Private Sub Form_Activate()
 vModulo = 3
 Call RefrescaTags(Me)
End Sub

Private Sub Form_Load()
 vModulo = 3
 vOperacion = 0
 vFechaHoy = fxFechaServidor
 
 Call Formularios(Me)
 Call sbDocumentosCombo(cboTipo)
 Call sbLimpiaDatos
End Sub

Private Sub sbConsultaOperacion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Me.MousePointer = vbHourglass

Call sbLimpiaDatos
 
strSQL = "select R.id_solicitud,R.saldo, R.saldo - isnull(V.amortiza,0) As Saldo_mes,R.proceso" _
       & ",R.interesv,R.int,R.plazo,R.interesc,R.amortiza,R.fecult,R.Prideduc" _
       & ",R.opex,R.cuota,R.codigo,R.cedula,R.cuotas_planilla,R.cuotas_directas, datediff(m,R.fechaforp,dbo.MyGetdate()) as 'Meses'" _
       & ",S.nombre,C.descripcion,C.retencion,C.poliza,R.fechaforp,C.PORC_CARGO_CANCELACION,C.ANTICIPO_MESES,R.Base_Calculo" _
       & ",dbo.fxCrdPlanPagosDiasActivo(" & vOperacion & ") as 'DiasActivo', dbo.fxCrdOperacionTagReg(R.id_solicitud,'S15') as 'AutPagoAnt'" _
       & ",C.descripcion as 'LineaDesc',Ofi.descripcion as 'OficinaDesc',Pre.Descripcion as 'RecursoDesc',dbo.MyGetdate() as 'FechaServer'" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo " _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " left join Sif_Oficinas Ofi on R.cod_Oficina_R = Ofi.cod_Oficina" _
       & " left join CATALOGO_GRUPOS Pre on R.cod_grupo = Pre.cod_grupo" _
       & " left join vista_morosidad V on R.id_solicitud = V.id_solicitud" _
       & " where R.ID_SOLICITUD = " & vOperacion
       
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
  vBaseCalculo = Trim(rs!Base_Calculo)
  vPrideduc = rs!PriDeduc
  vOperacion = rs!Id_Solicitud
  vPlazo = rs!Plazo
  vDiasActivo = rs!DiasActivo
  
  'Indica si Aplica Cargo por Cancelacion Anticipada y no se encuentra autorizado debe de cobrarse
  If rs!Meses <= rs!ANTICIPO_MESES And rs!AutPagoAnt = 0 Then
     vAnticipoPorc = rs!PORC_CARGO_CANCELACION / 100
  Else
     vAnticipoPorc = 0
  End If
  
  vInteres = IIf(IsNull(rs!interesv), rs!Int, rs!interesv)
  If IsNull(rs!saldo_mes) Then
    vSaldoMes = rs!Saldo
    strSQL = "update reg_creditos set saldo_mes = saldo where id_solicitud = " & rs!Id_Solicitud
    Call ConectionExecute(strSQL)
  Else
    If rs!saldo_mes = 0 Then
        vSaldoMes = rs!Saldo
        strSQL = "update reg_creditos set saldo_mes = saldo where id_solicitud = " & rs!Id_Solicitud
        Call ConectionExecute(strSQL)
    Else
       vSaldoMes = rs!saldo_mes
    End If
  
  End If
  
     lblOpex.Caption = IIf((rs!opex = 1), "OPEX", "")
    
     txtOperacion = rs!Id_Solicitud
    txtCedula = rs!Cedula
    txtNombre = rs!Nombre
    txtCodigo = rs!Codigo
    
    
    lblDescripcion.Caption = rs!Descripcion
    If rs!retencion = "S" Or rs!Poliza = "S" Then
      vRetencion = True
    Else
      vRetencion = False
    End If
        
    'Barra de Estado
    
    StatusBarX.Panels.Item(1).Text = rs!OficinaDesc & ""
    StatusBarX.Panels.Item(2).Text = rs!LineaDesc & ""
    StatusBarX.Panels.Item(3).Text = rs!RecursoDesc & ""
        
        
       

Else
 
 vOperacion = 0
 vPlazo = 0
 vInteres = 0
 vSaldoMes = 0
 MsgBox "No se Encontró operación para abonos,puede que se encuentre cancelada ", vbInformation

End If
rs.Close

Me.MousePointer = vbDefault

End Sub

Private Sub sbLimpiaDatos()
 
 lblDescripcion.Caption = ""
 lblOpex.Caption = ""
 txtCedula = ""
 txtCodigo = ""
 txtNombre = ""
 txtOperacion = ""
End Sub

Private Sub sbBusqueda()

On Error GoTo vError

gBusquedas.Convertir = "N"
gBusquedas.Consulta = "Select R.id_solicitud as Operacion,R.Codigo,S.Cedula,S.Nombre,C.Descripcion" _
          & " from REG_CREDITOS R inner join SOCIOS S on R.cedula = S.cedula" _
          & " inner join Catalogo C on R.codigo = C.codigo"
gBusquedas.Columna = "R.CEDULA"
gBusquedas.Orden = "R.CEDULA"
gBusquedas.Filtro = " AND R.ESTADO = 'A'"

frmBusquedas.Show vbModal

txtOperacion = Trim(gBusquedas.Resultado)
vOperacion = txtOperacion

gBusquedas.Consulta = ""
gBusquedas.Columna = ""
gBusquedas.Orden = ""
gBusquedas.Resultado = ""
gBusquedas.Filtro = ""

Call sbConsultaOperacion

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub

Private Sub txtOperacion_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
 vOperacion = txtOperacion
 Call sbConsultaOperacion
End If
End Sub


Private Function fxDocumentoAbono(pTipoAbono As String, pTipoDoc As String, pComprobante As String _
                                , pConcepto As String, pCuenta As String) As Long
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(11) As String
Dim lngRecibo As Long, strCliente As String, vCuenta As String
Dim rsTmp As New ADODB.Recordset, vCuentaPoliza As String
Dim curIntC As Currency, curIntM As Currency, curCargo As Currency, curAmortiza As Currency, curPoliza As Currency

Dim curSaldoAnterior As Currency, curSaldoActual As Currency, vFecha As Date, vUsuario As String, vOficina As String

vCuenta = pCuenta

lngRecibo = CLng(pComprobante)

fxDocumentoAbono = lngRecibo


'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)


strSQL = "exec spCrdDocumentoAfectacion '" & fxTipoASENumero(pTipoDoc) & "','" & pComprobante & "','R'"
Call OpenRecordSet(rsTmp, strSQL, 0)
If rsTmp.EOF And rsTmp.BOF Then
  curIntC = 0
  curIntM = 0
  curAmortiza = 0
  curCargo = 0
  curPoliza = 0
Else
  curIntC = rsTmp!IntCor
  curIntM = rsTmp!IntMor
  curAmortiza = rsTmp!Principal
  curCargo = rsTmp!Cargos
  curPoliza = rsTmp!Polizas
End If
rsTmp.Close

strSQL = "select SALDO_ANTERIOR,SALDO_ACTUAL , COD_CONCEPTO,MOV_USUARIO,MOV_FECHA " _
       & " From CRD_OPERACION_TRANSAC Where ID_SOLICITUD = " & txtOperacion.Text _
       & " and TIPO_DOCUMENTO = '" & fxTipoASENumero(pTipoDoc) & "' and NUM_COMPROBANTE  = '" & pComprobante & "'" _
       & " order by ID_SEQ"
Call OpenRecordSet(rsTmp, strSQL, 0)

If rsTmp.EOF Or rsTmp.BOF Then
   rsTmp.Close
   Me.MousePointer = vbDefault
   MsgBox "No se localizan movimientos registrados con este comprabante!", vbExclamation
   Exit Function
End If

curSaldoAnterior = -1
curSaldoActual = 0

Do While Not rsTmp.EOF
  If curSaldoAnterior = -1 Then
     curSaldoAnterior = rsTmp!Saldo_anterior
     vFecha = rsTmp!Mov_fecha
     vUsuario = Trim(rsTmp!Mov_usuario)
     pConcepto = Trim(rsTmp!cod_Concepto)
  End If
  curSaldoActual = rsTmp!saldo_actual
  rsTmp.MoveNext
Loop
rsTmp.Close

'Cargar Oficinas
strSQL = "exec sbSIFOficinasUsuario '" & vUsuario & "'"
Call OpenRecordSet(rsTmp, strSQL, 0)
  vOficina = rsTmp!Titular
rsTmp.Close



'Lineas de Comprobante
strLinea(1) = "Saldo Anterior    " & Format(curSaldoAnterior, "Standard")
strLinea(2) = "Interes Corriente " & Format(curIntC, "Standard")
strLinea(3) = "Interes Atrasado  " & Format(curIntM, "Standard")
strLinea(4) = "Amortización      " & Format(curAmortiza, "Standard")
strLinea(5) = "Cargos            " & Format(curCargo, "Standard")
strLinea(6) = "Saldo Actual      " & Format(curSaldoActual, "Standard")
strLinea(7) = "Operacion/Línea   " & "Op.:" & txtOperacion.Text & " Lí.:" & txtCodigo.Text & " Ret.:" & IIf(vRetencion, "SI", "NO")
strLinea(8) = Trim(lblDescripcion.Caption)

strLinea(11) = "Póliza            " & Format(curPoliza, "Standard")

strSQL = "exec spCrdOperacionFechaProxPago " & txtOperacion.Text
Call OpenRecordSet(rsTmp, strSQL, 0)
  If Not IsNull(rsTmp!Fecha_Pago) Then
       strLinea(9) = "Prox.Pago..:" & Format(rsTmp!Fecha_Pago, "dd/mm/yyyy") & " Cta.(" & rsTmp!num_cuota & ") " & Format(rsTmp!Cuota, "Standard")
  Else
       strLinea(9) = "Prox.Pago..: >> <<"
  End If
  strLinea(10) = "Notas: " & rsTmp!notas & ""
 
rsTmp.Close
      

'Registro del Comprobante
strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
         & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
         & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,linea11,detalle,documento)" _
         & " values('" & lngRecibo & "','" & pTipoDoc & "','" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "','" & vUsuario & "','" & Trim(txtCedula.Text) _
         & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & curIntC + curIntM + curAmortiza + curCargo + curPoliza & ",'P','" & txtOperacion.Text _
         & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & vOficina & "','" & strLinea(1) & "','" _
         & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
         & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
         & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
         & strLinea(11) & "','" & vAseDocDetalle & "','" & vAseDocDeposito & "')"
 Call ConectionExecute(strSQL)
 
 'ASIENTO
 If curIntC > 0 Then
   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curIntC & ",'C','" & rs!cod_Divisa _
          & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintc _
          & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
   Call ConectionExecute(strSQL)
 End If
 
 If curIntM > 0 Then
   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curIntM & ",'C','" & rs!cod_Divisa _
          & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintm _
          & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
   Call ConectionExecute(strSQL)
 End If
 
 If curCargo > 0 Then
 'Detallar Cargos
   strSQL = "exec spCrdDocumentoAfectacionCargos '" & pTipoDoc & "','" & lngRecibo & "'"
   Call OpenRecordSet(rsTmp, strSQL, 0)
   Do While Not rsTmp.EOF
         strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & IIf(IsNull(rsTmp!Mov_Monto), curCargo, rsTmp!Mov_Monto) & ",'C','" & rs!cod_Divisa _
                & "',1," & GLOBALES.gEnlace & ",'" & rsTmp!cod_unidad & "','" & rsTmp!cod_centro_costo & "','" & rsTmp!cod_cuenta _
                & "','" & rsTmp!Id_Solicitud & "','" & rsTmp!Codigo & "','" & vAseDocDeposito & "'"
         Call ConectionExecute(strSQL)
         rsTmp.MoveNext
   Loop
   rsTmp.Close
 End If
 
 If curPoliza > 0 Then
   strSQL = "select dbo.fxCrdOperacionCtaContaPolizas(" & rs!Id_Solicitud & ") as 'Cuenta'"
   Call OpenRecordSet(rsTmp, strSQL, 0)
     vCuentaPoliza = Trim(rsTmp!Cuenta)
   rsTmp.Close
   
   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curPoliza & ",'C','" & rs!cod_Divisa _
          & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & vCuentaPoliza _
          & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
   Call ConectionExecute(strSQL)
 End If
 
 If curAmortiza > 0 Then
   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curAmortiza & ",'C','" & rs!cod_Divisa _
          & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaamortiza _
          & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
   Call ConectionExecute(strSQL)
 End If
 
 If curIntC + curIntM + curAmortiza + curCargo + curPoliza > 0 Then
   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curIntC + curIntM + curCargo + curAmortiza + curPoliza & ",'D','" & rs!cod_Divisa _
          & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & vCuenta _
          & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
   Call ConectionExecute(strSQL)
 End If
  

rs.Close


End Function



