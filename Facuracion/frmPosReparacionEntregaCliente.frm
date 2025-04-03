VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmPosReparacionEntregaCliente 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Entrega a Cliente"
   ClientHeight    =   5868
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   9612
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPosReparacionEntregaCliente.frx":0000
   ScaleHeight     =   5868
   ScaleWidth      =   9612
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDetalleT 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5400
      Width           =   7455
   End
   Begin VB.TextBox txtDetalleI 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5040
      Width           =   7455
   End
   Begin VB.TextBox txtArticulo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4680
      Width           =   7455
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   708
      Left            =   7320
      TabIndex        =   0
      Top             =   120
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   1249
      ButtonWidth     =   1588
      ButtonHeight    =   1249
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Entregar"
            Key             =   "entregar"
            Object.ToolTipText     =   "Entregar Articulos Reparados al Cliente"
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
      Left            =   360
      Top             =   -120
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
            Picture         =   "frmPosReparacionEntregaCliente.frx":169B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosReparacionEntregaCliente.frx":1D214
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3372
      Left            =   0
      TabIndex        =   8
      Top             =   1200
      Width           =   9612
      _Version        =   524288
      _ExtentX        =   16955
      _ExtentY        =   5948
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
      MaxCols         =   484
      ScrollBars      =   2
      SpreadDesigner  =   "frmPosReparacionEntregaCliente.frx":23A76
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Articulo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Nota del Taller"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Detalle Interno"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label lbl 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   6495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   15240
      X2              =   0
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "frmPosReparacionEntregaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim i As Integer, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
     Case 1 'Boleta
        vGrid.CellTag = rs!Detalle & ""
        vGrid.Text = CStr(IIf(IsNull(rs.Fields(i - 1).Value), "", rs.Fields(i - 1).Value))
     
     Case 3 'Articulo

        vGrid.CellTag = rs!ProductoDesc & ""
        vGrid.Text = CStr(IIf(IsNull(rs.Fields(i - 1).Value), "", rs.Fields(i - 1).Value))
     
     Case 6 'Boleta de Servicio
        vGrid.CellTag = rs!detalle_taller & ""
        vGrid.Text = CStr(IIf(IsNull(rs.Fields(i - 1).Value), "", rs.Fields(i - 1).Value))

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


Private Sub sbCarga()
Dim strSQL As String

On Error GoTo vError

strSQL = "select D.cod_orden,D.linea,D.cod_producto,D.nserie,D.cod_factura,D.boleta_Sr" _
       & ",isnull(S.descripcion,''), 0 ,D.detalle_taller,P.descripcion as ProductoDesc,D.detalle" _
       & " from pos_reparacion_detalle D inner join pv_productos P on D.cod_producto = P.cod_producto" _
       & " left join pos_reparacion_Tipos_Servicios S on D.tipo_servicio = S.tipo_servicio" _
       & " where D.cod_orden = '" & GLOBALES.gTag & "' and D.Fecha_Recibo is not null and D.estado = 'R'"
Call sbCargaGridLocal(vGrid, 8, strSQL)
vGrid.MaxRows = vGrid.MaxRows - 1

txtArticulo.Text = ""
txtDetalleI.Text = ""
txtDetalleT.Text = ""


Exit Sub

vError:
 vGrid.MaxRows = 0


End Sub


Private Sub Form_Activate()
vModulo = 33
End Sub

Private Sub Form_Load()
vModulo = 33

lbl.Caption = "Entrega Articulos Reparados [Boleta : " & GLOBALES.gTag & "]"
vGrid.AppearanceStyle = fxGridStyle

Call sbCarga

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Function fxBoletaInv() As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCodCausa As String, vCodBodega As String, xTipo As String
Dim vCodigo As String, vMascara As String, vFecha As Date, i As Integer

vMascara = "0000000000"
vFecha = fxFechaServidor

xTipo = "S"
vCodBodega = fxPosSRParametro("03")
vCodCausa = fxPosSRParametro("05")


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


Private Sub sbEntregar()
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
         
         Call sbPosSRAfectaInv(vBoletaSR, "E", vLinea, 0, vBoletaInv, False)
      
      End If
    Next i
        
    If vMoverInventarios Then
        'Afectar Inventarios Aqui
        strSQL = "exec spINVTranProcesa 'S','" & vBoletaInv & "','" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)
    End If


    Me.MousePointer = vbDefault
    MsgBox "Articulos Entregados Satisfactoriamente...", vbInformation
    Call sbCarga
    Call sbBoleta
 
Else
    Me.MousePointer = vbDefault
    MsgBox "No hay Registros Pendientes para Entrega...", vbInformation
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

strSQL = "select max(Boleta_entrega) as Boleta" _
       & " From pos_reparacion_detalle" _
       & " where cod_orden = '" & GLOBALES.gTag & "' and estado = 'E'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then

 vSQL = "{POS_REPARACION_DETALLE.BOLETA_ENTREGA} = '" & rs!Boleta & "'"
 Call sbPosReportesSR("BoletaEntrega", "SERVICIO DE REPARACION", "ENTREGA A CLIENTE", vSQL, vOrden)
 
End If
rs.Close
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "entregar"
    Call sbEntregar
  Case "boleta"
    Call sbBoleta
End Select

End Sub



Private Sub vGrid_LeaveCell(ByVal col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

vGrid.Row = Row

vGrid.col = 1
txtDetalleI.Text = vGrid.CellTag

vGrid.col = 3
txtArticulo.Text = vGrid.CellTag

vGrid.col = 6
txtDetalleT.Text = vGrid.CellTag

End Sub


