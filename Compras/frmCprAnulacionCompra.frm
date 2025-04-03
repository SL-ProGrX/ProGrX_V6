VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCprAnulacionCompra 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulación de Compras"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4380
   ScaleWidth      =   9750
   Begin XtremeSuiteControls.PushButton cmdComprobante 
      Height          =   495
      Left            =   6120
      TabIndex        =   21
      Top             =   3600
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "&Comprobante"
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
      Picture         =   "frmCprAnulacionCompra.frx":0000
   End
   Begin VB.ComboBox cbo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   2640
      Width           =   3735
   End
   Begin VB.TextBox txtProvDesc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      MaxLength       =   38
      TabIndex        =   6
      ToolTipText     =   "Nombre del Proveedor"
      Top             =   600
      Width           =   6135
   End
   Begin VB.TextBox txtProvCod 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      ToolTipText     =   "Codigo Proveedor"
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2925
      Width           =   1815
   End
   Begin VB.TextBox txtImpuestos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2565
      Width           =   1815
   End
   Begin VB.TextBox txtDescuento 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2235
      Width           =   1815
   End
   Begin VB.TextBox txtSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1875
      Width           =   1815
   End
   Begin VB.TextBox txtCompra 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "Código de Entrada"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtFactura 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtFecha 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1320
      Width           =   1695
   End
   Begin XtremeSuiteControls.PushButton cmdAnular 
      Height          =   495
      Left            =   8040
      TabIndex        =   22
      Top             =   3600
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "&Anular"
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
      Picture         =   "frmCprAnulacionCompra.frx":07B9
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   315
      Left            =   1560
      TabIndex        =   23
      Top             =   2040
      Width           =   2652
      _Version        =   1441793
      _ExtentX        =   4678
      _ExtentY        =   556
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
      CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
      Format          =   3
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   675
      Left            =   3360
      TabIndex        =   24
      Top             =   960
      Width           =   6135
      _Version        =   1441793
      _ExtentX        =   10821
      _ExtentY        =   1191
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
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cargo Periodico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   5280
      X2              =   240
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha Anulación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   17
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Factura"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   6000
      TabIndex        =   12
      Top             =   2925
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(+) Impuestos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   6000
      TabIndex        =   11
      Top             =   2565
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(-) Descuento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   6000
      TabIndex        =   10
      Top             =   2235
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   6000
      TabIndex        =   9
      Top             =   1875
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Compra"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   5295
   End
End
Attribute VB_Name = "frmCprAnulacionCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCompra  As String, vMascara As String

Private Sub cmdAnular_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curMonto As Currency, rsTmp As New ADODB.Recordset
Dim i As Integer


On Error GoTo vError

'Preguntar si ya fue programada
'0. Verificar fecha de la Anulacion
'1. Si No se ha programado entonces simplemente anular
'2. Si ya se Programo, preguntar si se ha realizado algun pago.
'2.1 Si no se ha efectuado ningun pago, simplemente remueve programacion y anula
'2.2 Si se a efectuado pago, remover los pendientes y generar cargo periodico (Monto) x lo el bruto a pagar.
'    sin deducciones de cargos directos ni periodicos anteriomente aplicados.
'3. Reversar Inventario

If Not fxInvPeriodos(dtpFecha.Value) Then
   MsgBox "El periodo de Afectacion de Inventarios no es válido..., verifique...", vbExclamation
   Exit Sub
End If

strSQL = "select * from cpr_compras where cod_compra = '" & Format(txtCompra, vMascara) & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
   rs.Close
   MsgBox "Compra No Existe, verifique...", vbExclamation
   Exit Sub
Else
  If rs!Estado <> "P" Then
     rs.Close
     MsgBox "Compra ya Se encuentra Anulada o se han realizado devoluciones de mercaderia, verifique...", vbExclamation
     Exit Sub
  End If
End If

'Inicia Proceso de Anulacion
Me.MousePointer = vbHourglass

If rs!forma_pago = "CR" Then
   strSQL = "update cxp_proveedores set saldo = isnull(saldo,0) - " & CCur(txtTotal) _
          & " where cod_proveedor = " & txtProvCod
   Call ConectionExecute(strSQL)
End If

strSQL = "update cpr_compras set estado = 'A'" _
       & ",anula_fecha = dbo.MyGetdate(), anula_fec_afecta = '" & Format(dtpFecha.Value, "yyyy/mm/dd") _
       & "',anula_user = '" & glogon.Usuario & "'" _
       & " where cod_compra = '" & Format(txtCompra, vMascara) & "'"
Call ConectionExecute(strSQL)

curMonto = 0

If rs!CxP_Estado = "G" Then
  'Calcula Monto Programado Bruto, pagado
  strSQL = "select isnull(sum(monto),0) as MontoX from cxp_pagoProv" _
         & " where cod_proveedor = " & rs!cod_Proveedor & " and cod_factura = '" & rs!cod_Factura _
         & "' and tesoreria is not null"
  Call OpenRecordSet(rsTmp, strSQL, 0)
    curMonto = rsTmp!montox
  rsTmp.Close

  If rs!forma_pago = "CR" Then
    'Genera Cargo Periodico x el monto Programado (Bruto), pagado
     strSQL = "select isnull(max(ID),0) as ultimo from cxp_cargosper where cod_proveedor = " & rs!cod_Proveedor
     Call OpenRecordSet(rsTmp, strSQL, 0)
       i = rsTmp!ultimo + 1
     rsTmp.Close
     
     If curMonto > 0 Then
        strSQL = "insert cxp_cargosper(id,cod_proveedor,cod_cargo,tipo,valor,vence,saldo,concepto,detalle,recaudado)" _
               & " values(" & i & "," & rs!cod_Proveedor & ",'" & fxCodigoCbo(cbo) & "','M'," & curMonto _
               & ",'" & Format(fxFechaServidor, "yyyy/mm/dd") & "'," & curMonto & ",'ANULACION DE FACTURA DE COMPRA','" _
               & "FACTURA : " & rs!cod_Factura & vbCrLf & "USUARIO :" & glogon.Usuario & "',0)"
        Call ConectionExecute(strSQL)
         
        Call Bitacora("Registra", "Cargo Adicional a Prov:" & rs!cod_Proveedor & " Sec: " & i)
     End If
  End If
  
  'Elimina Programacion Pendiente de Pago de la Factura
  strSQL = "delete cxp_pagoProv  where cod_proveedor = " & rs!cod_Proveedor _
         & " and cod_factura = '" & rs!cod_Factura & "' and tesoreria is null"
  Call ConectionExecute(strSQL)
  
End If

'Reversa Inventario
strSQL = "select * from cpr_compras_detalle where cod_factura = '" & rs!cod_Factura & "' and cod_proveedor = " & rs!cod_Proveedor
rs.Close
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Call sbInvInventario(rs!Cod_Producto, rs!Cantidad, rs!cod_bodega, Format(txtCompra, vMascara), "Compra.Anu", dtpFecha.Value _
             , rs!Precio, rs!imp_consumo, rs!imp_ventas, "S")
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
MsgBox "Anulación realizada Satisfactoriamente...", vbInformation
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdComprobante_Click()
Dim strSQL As String
Me.MousePointer = vbHourglass
strSQL = "{CPR_COMPRAS.COD_COMPRA} = '" & txtCompra & "'"
Call sbInvReportes("Compra.Anu", "COMPROBANTE DE ANULACION", "FACTURA #" & txtFactura, strSQL)

Me.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()
 vModulo = 35
End Sub

Private Sub Form_Load()
 vModulo = 35
 vCompra = ""
 vMascara = "0000000000"
 
 Call sbCprCboCargosPer(cbo)
 
 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub

Private Sub txtCompra_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCompra <> "" Then Call sbConsulta(txtCompra)
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_compra"
  gBusquedas.Orden = "cod_compra"
  gBusquedas.Consulta = "select E.cod_compra,E.cod_orden,E.cod_factura,P.descripcion as Proveedor" _
          & " from cpr_compras E inner join cxp_proveedores P on E.cod_proveedor = P.cod_proveedor"
  gBusquedas.Filtro = ""
  gBusquedas.Mascara = vMascara
  frmBusquedas.Show vbModal
  txtCompra = gBusquedas.Resultado
  If txtCompra <> "" Then Call sbConsulta(gBusquedas.Resultado, gBusquedas.Resultado2)
End If

End Sub


Private Sub sbConsulta(xCodigo As String, Optional xOrden As String = "")
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select E.*,(rtrim(C.Tipo_Orden) + ' - ' + C.descripcion) as Causa" _
       & ",P.descripcion as Proveedor,O.nota" _
       & " from cpr_ordenes O inner join cpr_Tipo_Orden C on O.Tipo_Orden = C.Tipo_Orden" _
       & " inner join cpr_compras E on O.cod_orden = E.cod_orden" _
       & " inner join cxp_proveedores P on E.cod_proveedor = P.cod_proveedor" _
       & " where E.cod_compra = '" & Format(xCodigo, vMascara) & "'"
       
If xOrden <> "" Then
  strSQL = strSQL & " and E.cod_orden = '" & Format(xOrden, vMascara) & "'"
Else
    If txtProvCod <> "" Then
       strSQL = strSQL & " and E.cod_proveedor = " & txtProvCod
    End If
End If

Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  
  txtCompra = Format(xCodigo, vMascara)
  vCompra = Format(xCodigo, vMascara)
  
  txtFactura = rs!cod_Factura
  
  txtProvCod = rs!cod_Proveedor
  txtProvDesc = rs!Proveedor
  
  txtFecha = Format(rs!fecha, "yyyy/mm/dd hh:mm:ss")
  dtpFecha.Value = fxFechaServidor
  
  txtNotas = rs!nota & ""
  
  txtImpuestos.Text = Format(rs!imp_ventas, "Standard")
  txtDescuento.Text = Format(rs!descuento, "Standard")
  txtSubTotal.Text = Format(rs!sub_Total, "Standard")
  txtTotal.Text = Format(rs!Total, "Standard")
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Call RefrescaTags(Me)
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtFactura_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   strSQL = "select cod_compra from cpr_compras where cod_proveedor = " & txtProvCod _
          & " and cod_factura = '" & txtFactura & "'"
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
     txtCompra = Trim(rs!cod_compra)
     Call txtCompra_KeyDown(vbKeyReturn, 0)
   Else
     MsgBox "No se encontro, la factura digitada...", vbExclamation
   End If
   rs.Close
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub txtProvCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProvDesc.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = " and estado = 'A'"
  frmBusquedas.Show vbModal
  txtProvCod = gBusquedas.Resultado
  txtProvDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtProvCod_LostFocus()
txtProvDesc = fxSIFCCodigos("D", txtProvCod, "proveedores")
End Sub

Private Sub txtProvDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = " and estado = 'A'"
  frmBusquedas.Show vbModal
  txtProvCod = gBusquedas.Resultado
  txtProvDesc = gBusquedas.Resultado2
End If

End Sub

