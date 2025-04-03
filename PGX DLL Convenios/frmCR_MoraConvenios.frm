VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmCR_MoraConvenios 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cancelación de Morosidad de los Convenios"
   ClientHeight    =   7620
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   12984
   Icon            =   "frmCR_MoraConvenios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   12984
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCasos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox txtCargos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox txtAmortiza 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox txtIntMor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox txtIntCor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   6840
      Width           =   1335
   End
   Begin VB.ComboBox cboDestino 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   480
      Width           =   6375
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11160
      Picture         =   "frmCR_MoraConvenios.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1095
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   12735
      _ExtentX        =   22458
      _ExtentY        =   9758
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "# Operación"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Cargos"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Int.Cor"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Int.Mor"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amortización"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Mora Ctas"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Mora Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Cédula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Nombre"
         Object.Width           =   7832
      EndProperty
   End
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Presione F4"
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblLoad 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   21
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Casos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   8760
      TabIndex        =   20
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   6960
      TabIndex        =   18
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cargos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   16
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   3
      Left            =   360
      TabIndex        =   14
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amortización"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   13
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Int.Mor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   12
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Int.Cor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   11
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Línea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   13
      Left            =   1680
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblEstado 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   7320
      Width           =   10215
   End
End
Attribute VB_Name = "frmCR_MoraConvenios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMensaje As String, vPaso As Boolean

Private Function fxVerifica() As Boolean
Dim strSQL As String, rsX As New ADODB.Recordset

vMensaje = ""


strSQL = "select * from catalogo where codigo = '" & txtCodigo & "'"
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
  vMensaje = vMensaje & "No existe el código especificado" & vbCrLf
Else
  If rsX!retencion = "S" Or rsX!Poliza = "S" Then vMensaje = vMensaje & "El código está como Retención o PSD " & vbCrLf
End If
rsX.Close

If lsw.ListItems.Count = 0 Then vMensaje = vMensaje & "No existen operaciones atrasadas en esta línea" & vbCrLf

If Not IsNumeric(txtTotal.Text) Then
    vMensaje = vMensaje & "El monto para aplicación no es válido" & vbCrLf
Else
    If CCur(txtTotal.Text) = 0 Then
        vMensaje = vMensaje & "No existen casos para aplicación " & vbCrLf
    End If
End If

If Len(Trim(vMensaje)) > 0 Then
  fxVerifica = False
Else
  fxVerifica = True
End If

End Function


Private Sub cboDestino_Click()

If Not vPaso Then Exit Sub

lblLoad.Caption = "Cargando Casos Morosos!"
lblLoad.Refresh

Call sbCargaLsw

lblLoad.Caption = ""

End Sub

Private Sub CmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
Dim lngRecibo As Long, vCuenta As String, vFecha As Date, strLinea(10) As String
Dim curAplicado As Currency, vTipoDoc As String, vConcepto As String

'Verificar si hay casos y si el codigo es de convenio
On Error GoTo vError

If Not fxVerifica Then
   MsgBox vMensaje, vbCritical
   Exit Sub
End If

If Not vPaso Then Exit Sub

vFecha = fxFechaServidor
vConcepto = "CRD001"
vCuenta = Trim(fxDocumentoCuenta("NC"))



vTipoDoc = "NC"


If vAseDocValido = False Then
    Me.MousePointer = vbDefault
    MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
          & " válida para esta operación...", vbCritical
    Exit Sub
End If

Me.MousePointer = vbHourglass

lngRecibo = fxDocumentoConsecutivo("NC")




'Detalle del Documento
strLinea(1) = "Abono General a Mora "
strLinea(2) = ""
strLinea(3) = "Casos             " & txtCasos.Text
strLinea(4) = "Amortizacion      " & txtAmortiza.Text
strLinea(5) = "Interes Corriente " & txtIntCor.Text
strLinea(6) = "Interes Moratorio " & txtIntMor.Text
strLinea(7) = "Cargos [General]  " & txtCargos.Text

strLinea(8) = ""
strLinea(9) = "Descripción : " & txtDescripcion.Text
strLinea(10) = "Destino     : " & cboDestino.Text


lblEstado.Caption = "Cancelando Morosidad.... [Espere]"
lblEstado.Refresh

'Cancela la Mora
strSQL = "update morosidad set estado = 'C', abintc = intc,abintm = intm,abamortiza = amortiza" _
       & ",AbCargo = Cargo,tcon = '" & vTipoDoc & "',ncon = '" & lngRecibo & "', cod_Concepto = '" & vConcepto & "'" _
       & ",fecult = dbo.MyGetdate(), cod_Caja = '', usuario = '" & glogon.Usuario & "'" _
       & " where Estado = 'A' and id_solicitud in(select id_solicitud from reg_creditos" _
       & " where codigo = '" & Trim(txtCodigo) & "' and Estado = 'A' and Proceso <> 'J' and cod_destino = '" _
       & SIFGlobal.fxCodText(cboDestino.Text) & "')"
Call ConectionExecute(strSQL)

'Actualiza Saldos con Aplicación
strSQL = "select R.id_solicitud,R.saldo,isnull(sum(M.abintc),0) as Intc,isnull(sum(M.abintm),0) as IntM" _
       & ",isnull(sum(M.abamortiza),0) as Amortiza from reg_creditos R inner join morosidad M" _
       & " on R.id_solicitud = M.id_solicitud" _
       & " where M.tcon = '" & vTipoDoc & "' and M.ncon = '" & lngRecibo & "' group by R.id_solicitud,R.saldo"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  strSQL = "update reg_creditos set saldo = saldo - " & rs!Amortiza _
         & ",interesc = interesc + " & (rs!intc + rs!intm) _
         & ",amortiza = amortiza + " & rs!Amortiza
         
  If rs!Amortiza >= rs!SALDO Then strSQL = strSQL & ",estado = 'C'"
  
  strSQL = strSQL & " where id_solicitud = " & rs!id_Solicitud
  
  Call ConectionExecute(strSQL)
 
  rs.MoveNext
Loop
rs.Close


lblEstado.Caption = "Actualizando Fecha de Ultimo Pago.... [Espere]"
lblEstado.Refresh

curAplicado = 0

strSQL = "select M.id_solicitud,R.cedula,max(FechaP) as 'UltMov'" _
       & " from morosidad M inner join Reg_creditos R on M.id_solicitud = R.id_solicitud" _
       & " where M.estado = 'C' and M.Tcon = '" & vTipoDoc & "' and M.ncon = '" & lngRecibo _
       & "' group by M.id_solicitud,R.cedula"
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 strSQL = "update reg_creditos set fecUlt = " & rs!UltMov _
        & " where id_solicitud = " & rs!id_Solicitud & " and Fecult < " & rs!UltMov
 Call ConectionExecute(strSQL)
 
 rs.MoveNext
Loop
rs.Close


'CREAR ASIENTO (NOTA CREDITO) - RESUMEN
lblEstado.Caption = "Creando Nota de Credito.... [Espere]"
lblEstado.Refresh

'Encabezado del Documento (Maestro)
strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
        & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
        & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento)" _
        & " values('" & lngRecibo & "','" & vTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','-General-" _
        & "','CANCELA MORA - CONVENIOS','" & vConcepto & "'," & CCur(txtTotal.Text) & ",'P','" & txtCodigo.Text _
        & "','" & SIFGlobal.fxCodText(cboDestino.Text) & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
        & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
        & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
        & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
        & vAseDocDetalle & "','" & vAseDocDeposito & "')"
Call ConectionExecute(strSQL)

'Asiento General de la Aplicación (DUAL v1-2 CtrlDoc)
strSQL = "exec spCrdMovAsientoCredito '" & vTipoDoc & "','" & lngRecibo & "','" & vCuenta _
        & "','" & txtCodigo.Text & "','" & SIFGlobal.fxCodText(cboDestino.Text) & "',''"
Call ConectionExecute(strSQL)


Me.MousePointer = vbDefault

Call Bitacora("Aplica", "NC-Cancelando Mora Convenio:" & txtCodigo)

lblEstado.Caption = ""

'Imprime nota
Call sbImprimeRecibo(lngRecibo, "NC")

MsgBox "Se aplicó Nota de Crédito # " & lngRecibo, vbInformation

Call cboDestino_Click

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBusqueda(Index As Integer)

'Set GLOBALES.gfrmFormulario = Me
gBusquedas.Convertir = "N"
gBusquedas.Resultado = ""
gBusquedas.Filtro = " and poliza = 'N' and Retencion = 'N' and codigo in(select A.codigo" _
          & " from catalogo_destinos D inner join catalogo_destinosASG A " _
          & " on D.cod_destino = A.cod_destino and D.envio_tesoreria = 0 group by A.codigo)"
gBusquedas.Consulta = "select codigo,descripcion from catalogo"

Select Case Index
  Case 0 'txtCodigo
        gBusquedas.Orden = "codigo"
        gBusquedas.Columna = "codigo"
  Case 1 'txtCodigo
        gBusquedas.Orden = "descripcion"
        gBusquedas.Columna = "descripcion"
End Select

frmBusquedas.Show vbModal
txtCodigo = gBusquedas.Resultado

If Len(Trim(txtCodigo)) > 0 Then
  txtDescripcion = fxDescribeCodigo(Trim(txtCodigo))
  Call sbCargaLsw
Else
  txtDescripcion = ""
  lsw.ListItems.Clear
End If

End Sub

Private Sub Form_Activate()
vModulo = 16
Call RefrescaTags(Me)
End Sub

Private Sub Form_Load()

vModulo = 16

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String

If KeyCode = vbKeyF4 Then Call sbBusqueda(0)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDestino.SetFocus

End Sub

Private Sub txtCodigo_LostFocus()
Dim strSQL As String

If txtCodigo.Text = "" Then Exit Sub

lblLoad.Caption = "Cargando destinos en Mora...ESPERE!"
lblLoad.Refresh

vPaso = False
  strSQL = "select rtrim(R.cod_destino) + ' - ' + rtrim(R.descripcion) as ItmX" _
         & " from catalogo_destinos R inner join catalogo_destinosAsg A on R.cod_destino = A.cod_destino" _
         & " where A.codigo = '" & txtCodigo & "' and R.envio_tesoreria = 0 and R.cod_Destino in(select Reg.cod_destino" _
         & " from reg_creditos Reg inner join Vista_Morosidad Vm on Reg.id_Solicitud = Vm.id_Solicitud where Reg.codigo = '" _
         & txtCodigo.Text & "' group by Reg.cod_destino)"
  Call sbLlenaCbo(cboDestino, strSQL, False)
vPaso = True

lblLoad.Caption = ""

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(1)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDestino.SetFocus
End Sub

Private Sub sbCargaLsw()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curIntCor As Currency, curIntMor As Currency, curAmortiza As Currency, curCargos As Currency
Dim itmX As ListItem

Me.MousePointer = vbHourglass

On Error GoTo vError

curIntCor = 0
curIntMor = 0
curAmortiza = 0
curCargos = 0


strSQL = "select R.id_Solicitud,R.codigo,S.cedula,S.Nombre,isnull(sum(M.intc),0) as IntCor,isnull(sum(M.intm),0) as IntMor" _
       & ",isnull(sum(M.amortiza),0) as Amortiza,isnull(sum(M.Cargo),0) as 'Cargos', count(*) as Cuotas" _
       & " from reg_creditos R inner join Morosidad M on R.id_Solicitud = M.id_Solicitud" _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " Where R.Estado = 'A' and R.Proceso <> 'J' and R.codigo = '" & Trim(txtCodigo.Text) & "' and cod_destino = '" _
       & SIFGlobal.fxCodText(cboDestino.Text) & "' and M.Estado = 'A'" _
       & " group by R.id_Solicitud,R.codigo,S.cedula,S.Nombre"
       
Call OpenRecordSet(rs, strSQL)
lsw.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!id_Solicitud)
      itmX.SubItems(1) = Format(rs!Cargos, "Standard")
      itmX.SubItems(2) = Format(rs!IntCor, "Standard")
      itmX.SubItems(3) = Format(rs!IntMor, "Standard")
      itmX.SubItems(4) = Format(rs!Amortiza, "Standard")
      
      
      itmX.SubItems(5) = Format(rs!Amortiza + rs!Cargos + rs!IntCor + rs!IntMor, "Standard")
      itmX.SubItems(6) = rs!Cuotas
      itmX.SubItems(7) = rs!Cedula & ""
      itmX.SubItems(8) = rs!Nombre & ""
      
    curIntCor = curIntCor + rs!IntCor
    curIntMor = curIntMor + rs!IntMor
    curAmortiza = curAmortiza + rs!Amortiza
    curCargos = curCargos + rs!Cargos
      
  rs.MoveNext
Loop
rs.Close

txtIntCor.Text = Format(curIntCor, "Standard")
txtIntMor.Text = Format(curIntMor, "Standard")
txtAmortiza.Text = Format(curAmortiza, "Standard")
txtCargos.Text = Format(curCargos, "Standard")

txtTotal.Text = Format(curIntCor + curIntMor + curAmortiza + curCargos, "Standard")
txtCasos.Text = Format(lsw.ListItems.Count, "###,###")


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtDescripcion_LostFocus()
Dim strSQL As String

vPaso = False
  
  strSQL = "select rtrim(R.cod_destino) + ' - ' + rtrim(R.descripcion) as ItmX" _
         & " from catalogo_destinos R inner join catalogo_destinosAsg A on R.cod_destino = A.cod_destino" _
         & " where A.codigo = '" & txtCodigo & "' and R.envio_tesoreria = 0"
  Call sbLlenaCbo(cboDestino, strSQL, True)

vPaso = True

End Sub
