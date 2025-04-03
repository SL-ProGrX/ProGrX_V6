VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmPosDevoluciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devoluciones de Mercadería"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   10170
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   360
      Top             =   4680
   End
   Begin VB.TextBox txtCaja 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   480
      Width           =   3735
   End
   Begin VB.TextBox txtNotas 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   840
      TabIndex        =   19
      Top             =   1320
      Width           =   9255
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      TabIndex        =   17
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtDevolucion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   7
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   480
      Width           =   1575
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
      Left            =   8160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   4635
      Width           =   1692
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
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   8160
      TabIndex        =   5
      Top             =   4995
      Width           =   1692
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
      Left            =   8160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   5325
      Width           =   1692
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
      Left            =   8160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   5685
      Width           =   1692
   End
   Begin VB.ComboBox cboTipo 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      ItemData        =   "frmPosDevoluciones.frx":0000
      Left            =   840
      List            =   "frmPosDevoluciones.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   3735
   End
   Begin VB.TextBox txtDocumento 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7920
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   1005
      ButtonWidth     =   487
      ButtonHeight    =   466
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repBoleta"
                  Text            =   "Boleta "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repListadoGeneral"
                  Text            =   "Listado General"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repSep1"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repAntiguedadSaldos"
                  Text            =   "Antiguedad Saldos"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repPagosPendientes"
                  Text            =   "Pagos Pendientes"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgAux01 
      Left            =   5400
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosDevoluciones.frx":0004
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosDevoluciones.frx":08E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosDevoluciones.frx":0BFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosDevoluciones.frx":0F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosDevoluciones.frx":123C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2292
      Left            =   0
      TabIndex        =   23
      Top             =   2160
      Width           =   10092
      _Version        =   524288
      _ExtentX        =   17801
      _ExtentY        =   4043
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
      MaxCols         =   488
      ScrollBars      =   2
      SpreadDesigner  =   "frmPosDevoluciones.frx":1558
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   22
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   20
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "# Factura"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   18
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7320
      TabIndex        =   16
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "# Dev."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
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
      Height          =   252
      Index           =   6
      Left            =   6600
      TabIndex        =   14
      Top             =   4632
      Width           =   1332
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H000000FF&
      Height          =   252
      Index           =   7
      Left            =   6600
      TabIndex        =   13
      Top             =   4992
      Width           =   1332
   End
   Begin VB.Label Label1 
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
      Height          =   252
      Index           =   8
      Left            =   6600
      TabIndex        =   12
      Top             =   5328
      Width           =   1212
   End
   Begin VB.Label Label1 
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
      Height          =   252
      Index           =   9
      Left            =   6600
      TabIndex        =   11
      Top             =   5688
      Width           =   1212
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   10080
      X2              =   0
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "# Doc."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   7320
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "frmPosDevoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vDevolucion As Long

Private Sub Form_Activate()
vModulo = 33
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vModulo = 33
vGrid.AppearanceStyle = fxGridStyle

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpiaPantalla()
Dim i As Integer

vCodigo = ""
txtCodigo = ""
vDevolucion = 0
txtDevolucion = ""

txtCaja = gCajas.Caja & " ¦ " & gCajas.Nombre & " ¦ " & gCajas.Usuario

txtFecha = Format(fxFechaServidor, "yyyy/mm/dd hh:mm:ss")
txtNotas = ""
txtDocumento = ""

vGrid.MaxRows = 1
vGrid.MaxCols = 7
For i = 1 To vGrid.MaxCols
  vGrid.col = i
  vGrid.Text = ""
Next

txtSubTotal = 0
txtDescuento = 0
txtImpuestos = 0
txtTotal = 0

txtCodigo.Enabled = True
txtDevolucion.Enabled = True

cboTipo.Clear
cboTipo.AddItem "01 - Facturas Automáticas"
cboTipo.AddItem "02 - Facturas Manuales"
cboTipo.Text = "01 - Facturas Automáticas"


End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim strSQL As String
'Bloquear Caja al Entrar / Desbloquear al Salir
strSQL = "update pv_cajas set bloqueo = 0" _
       & " where cod_caja = '" & gCajas.Caja & "' and usuario = '" & gCajas.Usuario & "'"
Call ConectionExecute(strSQL)
End Sub

Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

TimerX.Interval = 0

If gCajas.Apertura > 0 Then

    'No puede darse EOF porque ya verificado
    strSQL = "select nombre,def_cliente,def_bodega,def_precio,def_agente" _
           & ",modifica_precio,modifica_fechas from pv_cajas" _
           & " where cod_caja = '" & gCajas.Caja & "' and usuario = '" & gCajas.Usuario & "'"
    Call OpenRecordSet(rs, strSQL)
        gCajas.Agente = rs!def_agente
        gCajas.Bodega = rs!def_bodega
        gCajas.Cliente = rs!def_cliente
        gCajas.Precio = rs!def_precio
        gCajas.Nombre = rs!Nombre
        gCajas.BodegaDesc = fxSIFCCodigos("D", rs!def_bodega, "bodegas")
        gCajas.ModFechas = IIf((rs!modifica_fechas = 1), True, False)
        gCajas.ModPrecios = IIf((rs!modifica_precio = 1), True, False)
    rs.Close
    
    'Bloquear Caja al Entrar / Desbloquear al Salir
    strSQL = "update pv_cajas set bloqueo = 1" _
           & " where cod_caja = '" & gCajas.Caja & "' and usuario = '" & gCajas.Usuario & "'"
    Call ConectionExecute(strSQL)

    Me.Caption = "POS: Devoluaciones" & Space(10) & "Caja: " & Trim(gCajas.Caja) & Space(5) & "Usuario: " _
               & gCajas.Usuario & Space(5) & " Nombre: " & gCajas.Nombre

Else
  Me.MousePointer = vbDefault
  MsgBox "Debes Ingresar a una Caja con una apertura de Cajas Válida!", vbExclamation
  Unload Me

End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      
      txtCodigo.Enabled = True
      txtCodigo.SetFocus
      txtDevolucion.Enabled = False
      
      vGrid.Enabled = True
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      vGrid.Enabled = False
      vGrid.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsultaFac(vCodigo)
      End If

    Case "CONSULTAR"
'       gBusquedas.Columna = "descripcion"
'       gBusquedas.Orden = "descripcion"
'       gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
'       frmBusquedas.Show vbModal
'       txtCodigo.SetFocus
'       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
'       txtNombre.SetFocus

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Function sbDesActProdCancelados()
Dim lng As Long

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 3
 If vGrid.Text <> "" Then
   vGrid.CellTag = CCur(vGrid.Text)
   If CCur(vGrid.Text) <= 0 Then
      vGrid.col = 1
      vGrid.Text = ""
   End If
 End If
Next lng

End Function
Private Sub sbConsultaFac(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from pv_facturacion where cod_factura = '" & xCodigo _
       & "' and tipo = '" & IIf((Mid(cboTipo.Text, 1, 2) = "01"), "A", "M") & "'"
       
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
'  Call sbToolBar(tlb, "edicion")
  vEdita = False 'False
  Call sbLimpiaPantalla
  
  vCodigo = rs!cod_Factura
  txtCodigo = rs!cod_Factura
  
  If rs!Tipo = "A" Then
    cboTipo.Text = "01 - Facturas Automáticas"
  Else
    cboTipo.Text = "02 - Facturas Manuales"
  End If
  
  
 'Tengo que conservar visibles aquellos que ya fueron despachados para conservar el consecutivo
 'de la linea del detalle, indica la bodega por defecto
  strSQL = "select D.cod_producto,P.descripcion,(D.cantidad - isnull(D.cantidad_devuelta,0)) as Cantidad" _
         & ",'" & gCajas.Bodega & "',D.precio,D.imp_ventas,(((D.cantidad - isnull(D.cantidad_devuelta,0)) * D.precio) + " _
         & "(((D.cantidad - isnull(D.cantidad_devuelta,0)) * D.precio) * (D.imp_ventas / 100))) as Total" _
         & " from pv_factura_detalle D inner join pv_productos P on D.cod_producto = P.cod_producto" _
         & " where D.cod_factura = '" & rs!cod_Factura & "' and D.tipo = '" & rs!Tipo _
         & "' order by D.Linea"
  
  Call sbCargaGrid(vGrid, 7, strSQL)
  Call sbDesActProdCancelados
  
  
  Call sbCalculaTotales
  
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


Private Sub sbConsulta(lngCodigo As Long)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select *  from pv_devoluciones where cod_devolucion = " & lngCodigo
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = rs!cod_Factura
  txtCodigo = rs!cod_Factura
  
  txtDevolucion = lngCodigo
  vDevolucion = lngCodigo
  
  
  Select Case UCase(rs!Tipo)
    Case "A"
      cboTipo.Text = "01 - Facturas Automáticas"
    Case "M"
      cboTipo.Text = "02 - Facturas Manuales"
  End Select
  
  txtFecha = Format(rs!fecha, "yyyy/mm/dd hh:mm:ss")
  txtNotas = rs!nota & ""
  txtDocumento = rs!Documento & ""
  
  txtImpuestos = Format(rs!imp_ventas, "Standard")

  strSQL = "select D.cod_producto,P.descripcion,D.cantidad,D.cod_bodega,D.precio,D.imp_ventas," _
         & "(D.cantidad * D.precio) + (D.cantidad * D.precio * (D.imp_ventas / 100)) as Total" _
         & " from pv_devolucion_detalle D inner join pv_productos P on D.cod_producto = P.cod_producto" _
         & " where D.cod_devolucion = " & rs!COD_DEVOLUCION _
         & " order by D.Linea"
  Call sbCargaGrid(vGrid, 7, strSQL)
  
  Call sbCalculaTotales
  
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


Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer
Dim vMensaje As String


On Error GoTo vError

vMensaje = ""
fxValida = True

'Verifica Bodegas y Articulos
vMensaje = fxInvVerificaLineaDetalle(vGrid, 3, "E", 1, 4)

'Verifica Periodo
If Not fxInvPeriodos(fxFechaServidor) Then vMensaje = vMensaje & vbCrLf & " - El Periodo del Movimiento no es válido ..."

'If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Proveedor no es válido ..."
'Verifiqua que exista la factura y que no se encuentre anulada

''If IsNumeric(txtProvCod) Then
''    strSQL = "select estado from cpr_compras where cod_factura = '" & txtCodigo _
''           & "' and cod_proveedor = " & txtProvCod & " and estado in('P','D')"
''    Call OpenRecordSet(rs, strSQL)
''    If rs.EOF And rs.BOF Then
''       vMensaje = vMensaje & vbCrLf & " - No se encontró registro de la factura, o se encuentra Anulada, verifique..."
''    End If
''    rs.Close
''Else
''   vMensaje = vMensaje & vbCrLf & " - El codigo del Proveedor no es válido, verifique..."
''End If

'Verifica que las cantidades de las devoluciones no sean mayores al original pendiente
For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.col = 1
  If vGrid.Text <> "" Then
    vGrid.col = 3
    If CCur(vGrid.Text) > CCur(vGrid.CellTag) Then
     vMensaje = vMensaje & vbCrLf & " - Las Cantidad devoluciones en la Linea " & i & ", es mayor al remanente..."
    End If
  End If
Next i

If CCur(txtTotal) = 0 Then vMensaje = vMensaje & vbCrLf & " - El total de la mercadería es cero, verifique..."


vError:

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If


End Function

Private Function fxConsecDev() As Long
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select (isnull(sum(cod_devolucion),0) + 1) as Ultimo from pv_devoluciones"
Call OpenRecordSet(rs, strSQL)
  fxConsecDev = rs!ultimo
rs.Close
End Function


Private Sub sbGuardar()
Dim strSQL As String, i As Integer, vFecha As Date
Dim curCantidad As Currency, vCodPro As String, vCodBodega As String
Dim curPrecio As Currency, curImpVentas As Currency, curImpConsumo As Currency

On Error GoTo vError

If vEdita Then
   MsgBox "No se puede editar una Devolución Guardada...", vbInformation
   Exit Sub
Else
   
   vFecha = fxFechaServidor
   vDevolucion = fxConsecDev
   txtDevolucion = vDevolucion
   
   txtCodigo = vCodigo
   vCodigo = txtCodigo
   
   strSQL = "insert pv_devoluciones(cod_caja,usuario,cod_devolucion,cod_factura,tipo,fecha,sub_total,descuento,imp_ventas" _
          & ",imp_consumo,total,documento,nota,asiento_estado) values('" & gCajas.Caja & "','" & gCajas.Usuario & "'," & vDevolucion _
          & ",'" & vCodigo & "','" & IIf((Mid(cboTipo.Text, 1, 2) = "01"), "A", "M") & "',dbo.MyGetdate()," _
          & CCur(txtSubTotal) & "," & CCur(txtDescuento) & "," & CCur(txtImpuestos) & ",0," _
          & CCur(txtTotal) & ",'" & txtDocumento & "','" & txtNotas & "','P')"
   Call ConectionExecute(strSQL)


  'Registra Movimiento en Cajas, en Efectivo 99999 indefinido
  Call sbPosCajaMovRegistra("DV", gCajas.Caja, gCajas.Usuario, CInt(gCajas.Apertura) _
            , CCur(txtTotal), 999999, CStr(vDevolucion), txtDocumento & " # Factura " _
            & IIf((Mid(cboTipo.Text, 1, 2) = "01"), "AUTO :", "MANUAL :") & vCodigo)


  Call Bitacora("Registra", "Devolucion Fact.: " & vCodigo & " Dev: " & vDevolucion)

  txtDevolucion.Enabled = True

End If

'Guardar Detalle de la Orden
strSQL = "delete pv_devolucion_detalle" _
         & " where cod_devolucion = " & vDevolucion
Call ConectionExecute(strSQL)

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  
  vGrid.col = 3
  curCantidad = CCur(IIf((vGrid.Text = ""), 0, vGrid.Text))
  
  vGrid.col = 1
  
  If vGrid.Text <> "" And curCantidad > 0 Then
    
    vGrid.col = 1
    vCodPro = Trim(vGrid.Text)
    strSQL = "insert pv_devolucion_detalle(linea,cod_devolucion,cod_producto,cantidad,cod_bodega" _
           & ",precio,imp_ventas,imp_consumo) values(" & i & "," & vDevolucion & ",'" _
           & vGrid.Text & "'," & curCantidad & ",'"
    vGrid.col = 4
    vCodBodega = Trim(vGrid.Text)
    strSQL = strSQL & vGrid.Text & "',"
    vGrid.col = 5
    curPrecio = CCur(vGrid.Text)
    strSQL = strSQL & CCur(vGrid.Text) & ","
    vGrid.col = 6
    curImpVentas = CCur(vGrid.Text)
    curImpConsumo = 0
    strSQL = strSQL & CCur(vGrid.Text) & ",0)"
    Call ConectionExecute(strSQL)
  
    'Actualizar Aqui el Inventario y la Factura
    vGrid.col = 3
    strSQL = "update pv_factura_detalle set cantidad_devuelta = isnull(cantidad_devuelta,0) + " _
           & CCur(vGrid.Text) & " where linea = " & i & " and cod_factura = '" & vCodigo _
           & "' and tipo = '" & IIf((Mid(cboTipo.Text, 1, 2) = "01"), "A", "M") & "'"
    Call ConectionExecute(strSQL)
    
    Call sbInvInventario(vCodPro, curCantidad, vCodBodega, CStr(vDevolucion), "Dev.Fact.", vFecha _
            , curPrecio, curImpConsumo, curImpVentas, "E")
    
  End If
Next i

Call sbToolBar(tlb, "activo")
Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
   'no se pueden Ejecutar Borrados en Ordenes
'  strSQL = "delete cxp_proveedores where cod_proveedor = " & vCodigo
'  Call ConectionExecute(strSQL)

'  Call Bitacora("Elimina", "ER ESPECIAL : " & vCodigo & " EMP: " & vParametros.CodigoEmpresa)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_factura"
  gBusquedas.Orden = "cod_factura"
  gBusquedas.Consulta = "select cod_factura,tipo,total,cedula from pv_facturacion"
  gBusquedas.Filtro = " and tipo = '" & IIf((Mid(cboTipo.Text, 1, 2) = "01"), "A", "M") & "'"
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsultaFac(gBusquedas.Resultado)
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" Then Call sbConsultaFac(txtCodigo)
End Sub

Private Sub txtDescuento_GotFocus()
On Error GoTo vError
txtDescuento = CCur(txtDescuento)
Exit Sub
vError:
  MsgBox "Información del Descuento no es válida...", vbCritical
End Sub

Private Sub txtDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTotal.SetFocus
End Sub

Private Sub txtDescuento_LostFocus()
On Error GoTo vError

txtDescuento = Format(CCur(txtDescuento), "Standard")
txtTotal = Format(CCur(txtSubTotal) + CCur(txtImpuestos) - CCur(txtDescuento), "Standard")

Exit Sub
vError:
  MsgBox "Información del Descuento no es válida...", vbCritical
End Sub

Private Sub txtDevolucion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_devolucion"
  gBusquedas.Orden = "cod_devolucion"
  gBusquedas.Consulta = "select Cod_devolucion,cod_factura,tipo,nota,fecha" _
          & " from pv_devoluciones"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtDevolucion = gBusquedas.Resultado
  If txtDevolucion <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub


Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
vError:
End Sub


Private Sub sbCalculaTotales()
Dim curSubTotal As Currency, curIV As Currency
Dim curTmpPrecio As Currency, curTmpIV As Currency, curTmpCant As Currency
Dim i As Integer, lng As Long

'**********************************************    OJO
'Revisar esta formula por la situacion del descuento, si es antes o despues del
'impuesto de ventas, por ahora está despues del impuesto
curSubTotal = 0
curIV = 0

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 3
 If vGrid.Text <> "" Then
   curTmpCant = CCur(vGrid.Text)
   If curTmpCant > 0 Then
    vGrid.col = 5
    curTmpPrecio = CCur(vGrid.Text)
    vGrid.col = 6
    curTmpIV = CCur(vGrid.Text)

    curSubTotal = curSubTotal + (curTmpCant * curTmpPrecio)
    curIV = curIV + ((curTmpCant * curTmpPrecio) * (curTmpIV / 100))
   End If
 End If
Next lng

txtSubTotal = Format(curSubTotal, "Standard")
txtImpuestos = Format(curIV, "Standard")
txtTotal = Format(curSubTotal + curIV - CCur(txtDescuento), "Standard")

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(7) As Variant, x As Integer

'Abrir Nueva Linea
If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    Call sbCalculaTotales
  End If
End If

'Consular Articulo
'If vGrid.ActiveCol = 1 And KeyCode = vbKeyF4 Then
'   frmBusquedaArticulos.Show vbModal
'   vGrid.Row = vGrid.ActiveRow
'   vGrid.Col = 1
'   vGrid.Text = gBusquedas.Resultado
'End If

'Consular Bodegas
If vGrid.ActiveCol = 4 And KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_bodega"
   gBusquedas.Orden = "cod_bodega"
   gBusquedas.Consulta = "select cod_bodega,descripcion from pv_bodegas"
   gBusquedas.Filtro = " and permite_entradas = 1"
   frmBusquedas.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 4
   vGrid.Text = gBusquedas.Resultado
End If


''Borrar una linea
'If KeyCode = vbKeyDelete Then
'  vGrid.Row = vGrid.ActiveRow
'  vGrid.Col = 7
'  For lng = vGrid.ActiveRow To vGrid.MaxRows
'     vGrid.Row = lng + 1
'     For x = 1 To 7
'        vGrid.Col = x
'        vTemp(x) = vGrid.Text
'     Next x
'
'     vGrid.Row = lng
'     For x = 1 To 7
'       vGrid.Col = x
'       vGrid.Text = vTemp(x)
'     Next x
'  Next lng
'  vGrid.MaxRows = vGrid.MaxRows - 1
'  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
'  Call sbCalculaTotales
'End If


End Sub



Private Sub vGrid_KeyPress(KeyAscii As Integer)
Dim curCantidad As Currency, curPrecio As Currency, curIV As Currency

On Error GoTo vError
'Calcula Total
Select Case vGrid.ActiveCol
  Case 3, 5, 6
    vGrid.Row = vGrid.ActiveRow
    vGrid.col = 3
    curCantidad = CCur(vGrid.Text)
    vGrid.col = 5
    curPrecio = CCur(vGrid.Text)
    vGrid.col = 6
    curIV = CCur(vGrid.Text)
    vGrid.col = 7
    vGrid.Text = (curPrecio * curCantidad) + ((curPrecio * curCantidad) * (curIV / 100))
   Call sbCalculaTotales
  Case Else 'No Aplica
End Select
vError:
End Sub






