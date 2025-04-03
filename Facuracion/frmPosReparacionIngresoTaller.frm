VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmPosReparacionIngresoTaller 
   Caption         =   "Ingresos de Taller"
   ClientHeight    =   5136
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   11328
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5136
   ScaleWidth      =   11328
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo 
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
      ItemData        =   "frmPosReparacionIngresoTaller.frx":0000
      Left            =   1800
      List            =   "frmPosReparacionIngresoTaller.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Estado Actual del proveedor"
      Top             =   360
      Width           =   6255
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   708
      Left            =   9120
      TabIndex        =   1
      Top             =   120
      Width           =   1932
      _ExtentX        =   3408
      _ExtentY        =   1249
      ButtonWidth     =   1376
      ButtonHeight    =   1249
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Recibir"
            Key             =   "recibir"
            Object.ToolTipText     =   "Recibir articulos del Taller"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Boleta"
            Key             =   "boleta"
            Object.ToolTipText     =   "Boleta de Envio"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   120
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosReparacionIngresoTaller.frx":0022
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosReparacionIngresoTaller.frx":6884
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3852
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   11052
      _Version        =   524288
      _ExtentX        =   19495
      _ExtentY        =   6795
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
      MaxCols         =   10
      SpreadDesigner  =   "frmPosReparacionIngresoTaller.frx":D0E6
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   15240
      X2              =   0
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmPosReparacionIngresoTaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, strUltimaSeleccion As String


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim xSQL As String, rs As New ADODB.Recordset
Dim i As Integer, strResultado As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

xSQL = "select Tipo_Servicio,descripcion from Pos_reparacion_tipos_servicios" _
       & " where activo = 1 order by tipo_servicio"
rs.Open xSQL, glogon.Conection, adOpenStatic

If Not rs.EOF And strUltimaSeleccion = "" Then strUltimaSeleccion = rs!Descripcion

strResultado = ""

Do While Not rs.EOF
  If Len(strResultado) = 0 Then
    strResultado = Chr$(9) & rs!Tipo_Servicio & "-" & rs!Descripcion
  Else
    strResultado = strResultado & Chr$(9) & rs!Tipo_Servicio & "-" & rs!Descripcion
  End If
  rs.MoveNext
Loop
rs.Close

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
     Case 3 'Articulo
        
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = rs!ProductoDesc
        vGrid.TextTip = TextTipFixed
        vGrid.TextTipDelay = 1000
        vGrid.Text = CStr(rs.Fields(i - 1).Value)
     
     Case 9  'Tipo de Servicio
        vGrid.CellType = CellTypeComboBox
        vGrid.TypeComboBoxList = strResultado
        vGrid.TypeComboBoxEditable = False
        vGrid.Text = strUltimaSeleccion
     
     Case Else
        vGrid.Text = CStr(IIf(IsNull(rs.Fields(i - 1).Value), "", rs.Fields(i - 1).Value))
    End Select
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

Me.MousePointer = vbDefault

End Sub


Private Sub cbo_Click()
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "select D.cod_orden,D.linea,D.cod_producto,D.detalle,D.nserie,D.cod_factura,D.boleta_Sr" _
       & ",D.detalle_taller,'',0,P.descripcion as ProductoDesc" _
       & " from pos_reparacion_detalle D inner join pv_productos P on D.cod_producto = P.cod_producto" _
       & " where D.cod_proveedor = " & cbo.ItemData(cbo.ListIndex) & " and D.Fecha_Recibo is null and D.estado = 'T'"
Call sbCargaGridLocal(vGrid, 10, strSQL)
vGrid.MaxRows = vGrid.MaxRows - 1

Exit Sub

vError:
 vGrid.MaxRows = 0


End Sub


Private Sub Form_Activate()
vModulo = 33
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vPaso = True

vModulo = 33
vGrid.AppearanceStyle = fxGridStyle

strSQL = "select cod_proveedor,descripcion from cxp_proveedores" _
       & " where cod_proveedor in(select cod_proveedor from pos_reparacion_detalle" _
       & " where boleta_ingreso is null and estado = 'T')"
Call OpenRecordSet(rs, strSQL)
cbo.Clear
Do While Not rs.EOF
 cbo.AddItem rs!Descripcion
 cbo.ItemData(cbo.NewIndex) = rs!cod_proveedor
 
 If rs.AbsolutePosition = 1 Then cbo.Text = rs!Descripcion

 rs.MoveNext
Loop
rs.Close

vPaso = False
Call cbo_Click

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Function fxBoletaInv() As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCodCausa As String, vCodBodega As String, xTipo As String
Dim vCodigo As String, vMascara As String, vFecha As Date, i As Integer

vMascara = "0000000000"
vFecha = fxFechaServidor

xTipo = "E"
vCodBodega = fxPosSRParametro("03")
vCodCausa = fxPosSRParametro("04")


'Inserta Boleta
strSQL = "select isnull(max(Boleta),0)+1 as Ultimo from pv_InvTranSac where Tipo = '" & xTipo & "'"
Call OpenRecordSet(rs, strSQL)
  vCodigo = Format(rs!ultimo, vMascara)
rs.Close

strSQL = "insert pv_InvTranSac(Boleta,Tipo,cod_entsal,genera_fecha,documento,notas" _
       & ",genera_user,estado, plantilla,fecha,fecha_sistema,Autoriza_user,autoriza_fecha)" _
       & " values('" & vCodigo & "','" & xTipo & "','" & vCodCausa _
       & "',dbo.MyGetdate(),'SR." & vCodigo & "','Generado Automaticamente x POS : Servicio de Reparacion" _
       & "','" & glogon.Usuario & "','A',0,'" _
       & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "',dbo.MyGetdate(),'" & glogon.Usuario & "',dbo.MyGetdate())"
Call ConectionExecute(strSQL)

vModulo = 32
Call Bitacora("Registra", "TranSac Inv.Tipo Cod.: " & vCodigo)

fxBoletaInv = vCodigo

vModulo = 34

End Function


Private Sub sbRecibir()
Dim vExiste As Boolean, i As Integer, strSQL As String
Dim vBoletaSR As String, vLinea As Integer, vBoletaInv As String
Dim vMoverInventarios As Boolean

'Cuenta Cuantos Registros Pendientes existen
'de lo contrario no procesa nada

On Error GoTo vError

Me.MousePointer = vbHourglass

vExiste = False

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.col = vGrid.MaxCols
  
  If vGrid.Value = 1 Then
    vExiste = True
    Exit For
  End If
Next i

If fxPosSRParametro("01") <> "S" Then
   vMoverInventarios = False
Else
   vMoverInventarios = True
End If


If vExiste Then
    'Procesa Todas las Lineas Pendientes del proveedor actual para todas
    'las boletas marcadas
    
    If vMoverInventarios Then
        'Crea Boleta Maestra de Afectacion de Inventarios
        vBoletaInv = fxBoletaInv
    Else
        'Consecutivo de Movimientos de Servicio
        vBoletaInv = fxPosSRConsecInterno
    End If
    
    For i = 1 To vGrid.MaxRows
      vGrid.Row = i
      vGrid.col = vGrid.MaxCols
      
      If vGrid.Value = 1 Then
         vGrid.col = 1
         vBoletaSR = vGrid.Text
         vGrid.col = 2
         vLinea = vGrid.Text
         
         Call sbPosSRAfectaInv(vBoletaSR, "R", vLinea, cbo.ItemData(cbo.ListIndex), vBoletaInv, False)
      
      End If
    Next i
        
    If vMoverInventarios Then
        'Afectar Inventarios Aqui
        strSQL = "exec spINVTranProcesa 'E','" & vBoletaInv & "','" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)
    End If

    Me.MousePointer = vbDefault
    MsgBox "Registros Recibidos del Taller Satisfactoriamente...", vbInformation
    Call cbo_Click
    Call sbBoleta
 
Else
    Me.MousePointer = vbDefault
    MsgBox "No hay Registros Pendientes para Recibir...", vbInformation
End If


Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbBoleta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vSQL As String, vOrden As String

vSQL = ""
vOrden = ""

strSQL = "select max(BOLETA_INGRESO) as Boleta" _
       & " From pos_reparacion_detalle" _
       & " where BOLETA_INGRESO is not null" _
       & " and cod_proveedor = " & cbo.ItemData(cbo.ListIndex)
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  
  vSQL = "{POS_REPARACION_DETALLE.BOLETA_INGRESO} = '" & rs!Boleta & "'"
  Call sbPosReportesSR("BoletaRecibo", "SERVICIO DE REPARACION", "RECIBO DE TALLER", vSQL, vOrden)

End If
rs.Close
End Sub


Private Sub Form_Resize()

On Error Resume Next


vGrid.Width = Me.Width - 320
vGrid.Height = Me.Height - 1800


End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "recibir"
    Call sbRecibir
  Case "boleta"
    Call sbBoleta
End Select

End Sub



Private Sub vGrid_LeaveCell(ByVal col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim strSQL As String, vTmpBoleta As String, vTmpDetalle As String
Dim Orden As String, Linea As Integer

On Error GoTo vError

If vGrid.MaxRows = 0 Then Exit Sub

Me.MousePointer = vbHourglass


If col = 7 Or col = 8 Then
 vGrid.col = col
 vGrid.Row = Row
 If vGrid.Text <> "" Then
    vGrid.col = 1
    Orden = vGrid.Text
    vGrid.col = 2
    Linea = vGrid.Text
    
    vGrid.col = 7
    vTmpBoleta = vGrid.Text
    vGrid.col = 8
    vTmpDetalle = vGrid.Text
    
    strSQL = "update pos_reparacion_detalle set Boleta_sr = '" & vTmpBoleta & "',detalle_taller = '" & vTmpDetalle _
           & "' where cod_orden = '" & Orden & "' and linea = " & Linea
    Call ConectionExecute(strSQL)
 
 End If
End If

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
