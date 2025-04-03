VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPosFacturaAnulacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulación de Factura"
   ClientHeight    =   5064
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   6468
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5064
   ScaleWidth      =   6468
   Begin VB.TextBox txtPago 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1680
      Width           =   5415
   End
   Begin VB.CommandButton cmdAnular 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Anular"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   21
      Top             =   4560
      Width           =   975
   End
   Begin VB.ComboBox cboDev 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   3960
      Width           =   5415
   End
   Begin VB.TextBox txtDetalle 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   960
      TabIndex        =   18
      Top             =   2880
      Width           =   5415
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   5040
      TabIndex        =   16
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   550
      _Version        =   393216
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
      Format          =   198180867
      CurrentDate     =   37830
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   2520
      Width           =   1935
   End
   Begin VB.ComboBox cboCaja 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2160
      Width           =   5415
   End
   Begin VB.TextBox txtCedula 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox txtCaja 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   5415
   End
   Begin VB.TextBox txtCodigo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdComprobante 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Comprobante Anulación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   24
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox txtEstado 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtFecha 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1320
      Width           =   3855
   End
   Begin VB.ComboBox cboTipo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      ItemData        =   "frmPosFacturaAnulacion.frx":0000
      Left            =   2520
      List            =   "frmPosFacturaAnulacion.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Pago"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   120
      TabIndex        =   22
      Top             =   1680
      Width           =   852
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Devolución"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   19
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Anulación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   3480
      TabIndex        =   15
      Top             =   2520
      Width           =   1572
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Monto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   852
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   852
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   6480
      X2              =   0
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   6480
      X2              =   0
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   12
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   852
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   852
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   852
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "# Factura"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   852
   End
End
Attribute VB_Name = "frmPosFacturaAnulacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub sbLimpiaPantalla(Optional vInicializa As Boolean = False)
Dim strSQL As String, rs As New ADODB.Recordset

txtCodigo.Tag = 0

If vInicializa Then
  Call sbPosCombosCarga("cajasFull", cboCaja)
  cboTipo.Clear
  cboTipo.AddItem "M ¦ Facturas Manuales"
  cboTipo.AddItem "A ¦ Facturas Automaticas"
  cboTipo.Text = "A ¦ Facturas Automaticas"
  
  Call sbPosCombosCarga("FormaPago", cboDev, " where clasificacion <> '03'")
End If


txtMonto = ""
txtEstado = ""
txtFecha = ""
txtCedula = ""
txtNombre = ""
txtCaja = ""
txtFecha = ""
txtPago = ""

dtpFecha.Value = fxFechaServidor
txtDetalle = ""

End Sub


Private Sub sbConsulta(vFactura As String, vTipo As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select F.cod_factura,F.tipo,F.total,F.fecha,F.estado,F.cedula,C.nombre" _
       & ",(rtrim(X.cod_caja) + ' - ' + X.usuario + ' ¦ ' + X.nombre) as Caja" _
       & ",F.cod_forma_pago,P.descripcion as FormaPagoDesc,P.clasificacion" _
       & " from pv_facturacion F inner join pv_clientes C on F.cedula = C.cedula" _
       & " inner join pv_cajas X on F.cod_caja = X.cod_caja and F.usuario = X.usuario" _
       & " inner join pv_formas_pago P on F.cod_forma_pago = P.cod_forma_pago" _
       & " where F.tipo = '" & vTipo & "' and F.cod_factura = '" & vFactura & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  MsgBox "No se encontró numero de factura digitada...", vbExclamation
  txtCodigo.Tag = 0
Else
  
  txtCodigo.Tag = rs!cod_Factura
  
  txtCaja = rs!Caja
  txtFecha = rs!fecha
  txtCedula = rs!Cedula
  txtNombre = rs!Nombre
  txtPago = rs!FormaPagoDesc
  txtPago.Tag = rs!Cod_Forma_Pago
  
  txtMonto = Format(rs!Total, "Standard")
  
  Select Case rs!Estado
    Case "A"
      txtEstado = "Anulada"
    Case "I"
      txtEstado = "Impresa"
    Case "P"
      txtEstado = "Procesada"
    Case "S"
      txtEstado = "Pend.Imprimir"
  End Select
  txtEstado.Tag = rs!Estado
End If

rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboTipo_Click()
Call sbLimpiaPantalla
End Sub

Private Sub sbAnulaFactura(vFactura As String, vTipo As String)
Dim strSQL As String, rs As New ADODB.Recordset, vFecha As Date
Dim vCaja As String, vUsuario As String, vApertura As Integer
Dim i As Integer, x As Integer

'Actualiza el Estado de la Factura
'Registra Movimiento en Cajas
'Reversa Salida de Inventarios, Ingresa a Inventarios segun al fecha de la Anulacion
'Vuelve a Cargar la Factura con el Estado Actual


On Error GoTo vError

Me.MousePointer = vbHourglass

'Verificar Periodo ***********************************************

vFecha = dtpFecha.Value

x = 1
For i = 1 To Len(cboCaja.Text)
 If x = 1 Then
   If Mid(cboCaja.Text, i, 1) = "-" Then
     vCaja = Trim(Mid(cboCaja.Text, 1, i - 1))
     x = i + 1
   End If
 Else
   If Mid(cboCaja.Text, i, 1) = "¦" Then
     vUsuario = Trim(Mid(cboCaja.Text, x, i - (x + 1)))
   End If
 End If
Next i

'Selecciona la Ultima Apertura, de no encontrarse terminar proceso aqui
strSQL = "select cod_ac from pv_cajas_ac where cod_caja = '" & vCaja _
       & "' and usuario = '" & vUsuario & "' and estado = 'A'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   vApertura = rs!cod_ac
Else
   rs.Close
   Me.MousePointer = vbDefault
   MsgBox "No se encontró ninguna Apertura de Caja, para la caja seleccionada...", vbExclamation
   Exit Sub
End If
rs.Close


'Actualiza Estado de la Factura
strSQL = "update pv_facturacion set estado = 'A',Anu_Detalle = '" & txtDetalle _
       & "',Anu_Fecha = '" & Format(vFecha, "yyyy/mm/dd") & "',Anu_CajaUser = '" _
       & vUsuario & "',Anu_CajaCod = '" & vCaja & "', Anu_CajaAP = " & vApertura _
       & ",Anu_Forma_Pago = " & fxCodigoCbo(cboDev) _
       & " where cod_factura = '" & vFactura & "' and tipo = '" & vTipo & "'"
Call ConectionExecute(strSQL)


'Registra Movimiento en Cajas
Call sbPosCajaMovRegistra(IIf((vTipo = "A"), "AA", "AM"), vCaja, vUsuario, vApertura _
         , CCur(txtMonto), fxCodigoCbo(cboDev), CStr(vFactura), CStr(vFactura) & " (" & txtCedula & " - " & txtNombre & ")")

'Guardar Detalle de la Factura y Registra Inventario
strSQL = "select cod_producto,cod_bodega,cantidad,cod_factura,precio,imp_ventas" _
       & ",isnull(imp_consumo,0) as imp_consumo" _
       & " from pv_factura_detalle where tipo = '" & vTipo _
       & "' and cod_factura = '" & vFactura & "' order by linea"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
    'Actualizar Aqui el Inventario
    Call sbInvInventario(rs!cod_producto, rs!cantidad, rs!cod_bodega, rs!cod_Factura, IIf((vTipo = "M"), "Anula.Fact.Man", "Anula.Fact.Auto") _
          , Format(vFecha, "yyyy/mm/dd hh:mm:ss"), rs!Precio, rs!imp_consumo, rs!imp_ventas, "E")
 rs.MoveNext
Loop
rs.Close

Call Bitacora("Anula", "Factura (" & vTipo & ") : " & vFactura)

Call sbConsulta(vFactura, vTipo)

'Imprimir Comprobante de Anulacion
Call cmdComprobante_Click

Me.MousePointer = vbDefault
MsgBox "Anulación de Factura Realizada Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdAnular_Click()
Dim i As Integer

If txtCodigo.Tag <> "" And txtEstado.Tag <> "A" Then
 If Not fxInvPeriodos(dtpFecha.Value) Then
   MsgBox " - El periodo de Anulacion para la factura ya fue cerrado o no es válido ..."
 Else
   Call sbAnulaFactura(txtCodigo.Tag, Mid(cboTipo.Text, 1, 1))
 End If
Else
  MsgBox "No se ha consultado la factura, o esta ya se encuentra Anulada...", vbInformation
End If
End Sub

Private Sub cmdComprobante_Click()
Dim strSQL As String
Me.MousePointer = vbHourglass
strSQL = "{PV_FACTURACION.COD_FACTURA} = '" & txtCodigo.Tag _
       & "' AND {PV_FACTURACION.TIPO} = '" & Mid(cboTipo.Text, 1, 1) & "'"
Call sbPosReportes("Fact.Anu", "COMPROBANTE DE ANULACION", "FACTURA #" & txtCodigo.Tag, strSQL)

Me.MousePointer = vbDefault
End Sub


Private Sub Form_Activate()
vModulo = 33
End Sub

Private Sub Form_Load()
vModulo = 33
Call Formularios(Me)
Call RefrescaTags(Me)

Call sbLimpiaPantalla(True)
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  Call sbConsulta(txtCodigo, Mid(cboTipo.Text, 1, 1))
  cboCaja.SetFocus
End If
vError:
End Sub
