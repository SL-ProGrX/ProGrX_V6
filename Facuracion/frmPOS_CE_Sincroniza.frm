VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmPOS_CE_Sincroniza 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Sincroniza Comprobantes Electrónicos"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   14235
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5532
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   11892
      _Version        =   1310723
      _ExtentX        =   20976
      _ExtentY        =   9758
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin VB.Timer Timer_Sinc 
      Left            =   360
      Top             =   0
   End
   Begin XtremeSuiteControls.PushButton btnDatos 
      Height          =   372
      Index           =   0
      Left            =   4080
      TabIndex        =   3
      Top             =   1140
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Clientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Checked         =   -1  'True
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.PushButton btnDatos 
      Height          =   372
      Index           =   1
      Left            =   5400
      TabIndex        =   4
      Top             =   1140
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Productos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnDatos 
      Height          =   372
      Index           =   2
      Left            =   6720
      TabIndex        =   5
      Top             =   1140
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Facturas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnSincroniza 
      Height          =   372
      Index           =   0
      Left            =   8400
      TabIndex        =   6
      Top             =   1140
      Width           =   1692
      _Version        =   1310723
      _ExtentX        =   2984
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Sinc. Manual"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmPOS_CE_Sincroniza.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnSincroniza 
      Height          =   372
      Index           =   1
      Left            =   10080
      TabIndex        =   7
      Top             =   1140
      Width           =   1692
      _Version        =   1310723
      _ExtentX        =   2984
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Sinc. Auto"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmPOS_CE_Sincroniza.frx":061C
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   375
      Left            =   12000
      TabIndex        =   9
      Top             =   1140
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Exportar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmPOS_CE_Sincroniza.frx":0D35
   End
   Begin XtremeShortcutBar.ShortcutCaption scEstado 
      Height          =   492
      Left            =   120
      TabIndex        =   8
      Top             =   7200
      Width           =   11892
      _Version        =   1310723
      _ExtentX        =   20976
      _ExtentY        =   868
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   12495
      _Version        =   1310723
      _ExtentX        =   22040
      _ExtentY        =   873
      _StockProps     =   14
      Caption         =   "Consulta Información Pendiente de Sincronizar: "
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
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "POS: Sincronización de Comprobantes Electrónicos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   6612
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmPOS_CE_Sincroniza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Dim mClienteIdSys As Long, mClienteId As String, mClienteTipoId As String
Dim mClienteEmail As String, mCabys As String

Dim mConServer As String, mConDb As String, mConUser As String, mConKey As String, db As New ADODB.Connection



Private Sub btnDatos_Click(Index As Integer)
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

btnDatos.Item(Index).Checked = True

lsw.ListItems.Clear

With lsw.ColumnHeaders
    .Clear

    Select Case Index
        Case 0 'Clientes
            .Add , , "Identificación", 1800
            .Add , , "Tipo Id", 1200, vbCenter
            .Add , , "Nombre", 4000
            .Add , , "Email", 3000
            .Add , , "Móvil", 1500, vbCenter
                
            btnDatos.Item(1).Checked = False
            btnDatos.Item(2).Checked = False
            
            strSQL = "exec spPOS_FE_Clientes_List '" & glogon.Usuario & "'"
            Call OpenRecordSet(rs, strSQL)
            Do While Not rs.EOF
              Set itmX = lsw.ListItems.Add(, , rs!Cedula)
                  itmX.SubItems(1) = rs!Tipo_ID_FE
                  itmX.SubItems(2) = rs!Nombre
                  itmX.SubItems(3) = rs!Email
                  itmX.SubItems(4) = rs!Celular
                  
              rs.MoveNext
            Loop
            rs.Close
        
        Case 1 'Productos
            .Add , , "Código", 1800
            .Add , , "Cabys", 1800, vbCenter
            .Add , , "Nombre", 4000
            .Add , , "Unidad", 1500, vbCenter
            .Add , , "Precio", 1400, vbRightJustify
            .Add , , "Impuesto Id", 1400, vbCenter
        
            btnDatos.Item(0).Checked = False
            btnDatos.Item(2).Checked = False
        
        
            strSQL = "exec spPOS_FE_Productos_List '" & glogon.Usuario & "'"
            Call OpenRecordSet(rs, strSQL)
            Do While Not rs.EOF
              Set itmX = lsw.ListItems.Add(, , rs!Cod_Producto)
                  itmX.SubItems(1) = rs!Cabys
                  itmX.SubItems(2) = rs!Descripcion
                  itmX.SubItems(3) = rs!Cod_Unidad
                  itmX.SubItems(4) = Format(rs!Precio_Col, "Standard")
                  itmX.SubItems(5) = rs!Impuesto_Cod
                  
              rs.MoveNext
            Loop
            rs.Close
        
        
        Case 2 'Facturas
            .Add , , "No.Factura", 1800
            .Add , , "C.E. Tipo", 1200, vbCenter
            .Add , , "C.E.No", 2200, vbCenter
            .Add , , "C.E.Llave", 4800, vbCenter
            .Add , , "Nombre", 4000
            .Add , , "Total", 1600, vbRightJustify
            .Add , , "Fecha", 1600, vbCenter
            .Add , , "Sub Total", 1600, vbRightJustify
            .Add , , "Descuentos", 1600, vbRightJustify
            .Add , , "IVA", 1600, vbRightJustify
        
            btnDatos.Item(0).Checked = False
            btnDatos.Item(1).Checked = False
            
            strSQL = "exec spPOS_FE_Facturas_List '" & glogon.Usuario & "'"
            Call OpenRecordSet(rs, strSQL)
            Do While Not rs.EOF
              Set itmX = lsw.ListItems.Add(, , rs!cod_Factura)
                  itmX.SubItems(1) = rs!FE_TIPO
                  itmX.SubItems(2) = rs!FE_Numero
                  itmX.SubItems(3) = rs!FE_CLAVE
                  itmX.SubItems(4) = rs!Nombre
                  itmX.SubItems(5) = Format(rs!Total_Comprobante, "Standard")
                  itmX.SubItems(6) = Format(rs!fecha, "yyyy-mm-dd")
                  itmX.SubItems(7) = Format(rs!TOTAL_VENTA_NETA, "Standard")
                  itmX.SubItems(8) = Format(rs!Total_Descuentos, "Standard")
                  itmX.SubItems(9) = Format(rs!TOTAL_IMPUESTO, "Standard")
                  
              rs.MoveNext
            Loop
            rs.Close
            
    End Select


End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass


Call Excel_Exportar_Lsw(lsw)


Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnSincroniza_Click(Index As Integer)

On Error GoTo vError


Select Case Index
    Case 0 ' Sincronizar
    
        Select Case True
            Case btnDatos.Item(0).Checked 'Clientes
                Call sbSinc_Clientes
            
            Case btnDatos.Item(1).Checked 'Productos
                Call sbSinc_Productos
                
            Case btnDatos.Item(2).Checked 'Facturas (+Clientes +Productos)
                Call sbSinc_Clientes
                Call sbSinc_Productos
                Call sbSinc_Facturas
        End Select
            
        MsgBox "Sincronización Finalizada!", vbInformation
    
    Case 1 'Activar Temporizador
         If btnSincroniza.Item(1).Checked Then
             btnSincroniza.Item(1).Checked = False
             Timer_Sinc.Interval = 0
             Timer_Sinc.Enabled = False
         Else
             btnSincroniza.Item(1).Checked = True
             Timer_Sinc.Interval = 6000
             Timer_Sinc.Enabled = True
             
             Call Timer_Sinc_Timer
         End If
         
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation


End Sub

Private Sub Form_Load()

 vModulo = 33
 
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

scTitulo.Width = Me.Width - (scTitulo.Left + 200)

scEstado.Width = scTitulo.Width

lsw.Width = scTitulo.Width

lsw.Height = Me.Height - (lsw.top + 850 + scEstado.Height)

scEstado.top = lsw.top + lsw.Height + 150

End Sub

Private Sub Timer_Sinc_Timer()

On Error GoTo vError

Timer_Sinc.Interval = 0

Me.MousePointer = vbHourglass

Call sbSinc_Clientes
Call sbSinc_Productos
Call sbSinc_Facturas
                
Me.MousePointer = vbDefault

Timer_Sinc.Interval = 6000

Exit Sub
    
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call btnDatos_Click(0)

Call sbInicializa_Portal

End Sub



Private Sub sbInicializa_Portal()

On Error GoTo vError

strSQL = "select P.*, isnull(C.descripcion,'') as 'Cabys_Desc' " _
       & " from sys_FE_Parametros P left join vINV_Cabys C on P.Cabys = C.COD_BYS"
Call OpenRecordSet(rs, strSQL)

mClienteId = rs!Cod_Cliente
mClienteTipoId = rs!Tipo_id
mClienteId = rs!Cedula

mClienteEmail = rs!NOTIFICA_EMAIL

mConServer = rs!ACC_SERVER
mConDb = rs!ACC_DB
mConUser = rs!ACC_USR
mConKey = rs!ACC_KEY


mCabys = rs!Cabys & ""

rs.Close

'Estableciendo Conexion
Call FE_Portal_Access


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub




Public Sub FE_Portal_Access()
Dim strSQL As String, vPaso As Boolean

On Error GoTo pConPortalError

Me.MousePointer = vbHourglass


strSQL = "PROVIDER=MSDASQL;Driver={SQL Server};Server=" & RTrim(mConServer) _
       & ";Database=" & RTrim(mConDb) & ";APP=PGX_Facturacion;tcp:" & RTrim(mConServer) _
       & "," & SIFGlobal.PuertosDisponibles & ";"


db.Close

vPaso = False
 
  
Conexion_Portal_Inicial:
  
With db
  
  vPaso = True
  .CommandTimeout = 15
  .Mode = adModeReadWrite
  .CursorLocation = adUseClient
  
  .Open strSQL, RTrim(mConUser), RTrim(mConKey)
  .CommandTimeout = 360
End With

Me.MousePointer = vbDefault

Exit Sub

pConPortalError:
  If Not vPaso Then GoTo Conexion_Portal_Inicial
  
  Screen.MousePointer = vbDefault
  MsgBox "No se tiene Conexión con el Servidor de Facturación!", vbCritical, "Contacte a su Administrador"

End Sub


Private Sub sbSinc_Clientes()
Dim rsRes As New ADODB.Recordset
Dim sResult As String, i As Long, iTotal As Long

Dim pActividad As String, pMoneda As String

On Error GoTo vError

Me.MousePointer = vbHourglass


scEstado.Caption = "Cargando parámetros..."

strSQL = "select CABYS, ACTIVIDAD_ECONOMICA, MONEDA from SYS_FE_PARAMETROS"
Call OpenRecordSet(rs, strSQL)
    pActividad = Trim(rs!Actividad_Economica)
    pMoneda = Trim(rs!Moneda)
rs.Close


scEstado.Caption = "Sincronizando Clientes"

sResult = ""

strSQL = "exec spPOS_FE_Clientes_List '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
    i = 0
    iTotal = rs.RecordCount
End If

Do While Not rs.EOF
  
 strSQL = "exec sp_IW_CLIENTEInsert_ProGrX '" & Trim(rs!Cod_Cliente) & "', '" & rs!Tipo_ID_FE & "', '" & rs!CLIENTE_ID & "','" & Trim(rs!Cedula) _
        & "','" & rs!Nombre & "', '" & rs!Email & "', '', '" & Trim(rs!Celular) & "', ''" _
        & ",'" & rs!Direccion & "',1,1,1,1, '" & Format(rs!fecha, "yyyy/mm/dd") & "', 30, 1"
 rsRes.Open strSQL, db, adOpenStatic
    sResult = sResult & Space(10) & "exec spPOS_FE_Clientes_Result '" & rs!Cod_Cliente & "','" & rs!Cedula _
            & "', '" & glogon.Usuario & "' ," & rsRes!CODIGO_INTERNO
 rsRes.Close
  i = i + 1
  scEstado.Caption = "Registrando Clientes Nuevos [" & i & ", " & iTotal & "]"
  DoEvents
 If Len(sResult) > 20000 Then
   Call ConectionExecute(sResult)
   sResult = ""
 End If
  
 rs.MoveNext
Loop
rs.Close

'Lote Final
If Len(sResult) > 0 Then
  Call ConectionExecute(sResult)
  sResult = ""
End If


scEstado.Caption = ""

Me.MousePointer = vbDefault


Select Case True
    Case btnDatos.Item(0).Checked
        Call btnDatos_Click(0)
    Case btnDatos.Item(1).Checked
        Call btnDatos_Click(1)
        
    Case btnDatos.Item(2).Checked
        Call btnDatos_Click(2)
End Select

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbSinc_Productos()
Dim rsRes As New ADODB.Recordset
Dim sResult As String, i As Long, iTotal As Long

On Error GoTo vError

Me.MousePointer = vbHourglass


scEstado.Caption = "Sincronizando Productos"

sResult = ""

strSQL = "exec spPOS_FE_Productos_List '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
    i = 0
    iTotal = rs.RecordCount
End If

'sp_IW_PRODUCTOInsert](
'    @ID_CLIENTE_ORIGEN INT,
'    @CODIGO varchar(15),
'    @NOMBRE varchar(80),
'    @NOMBRE_EXTENDIDO varchar(250),
'    @UNIDAD_MEDIDA varchar(10),
'    @PRECIO_COL decimal(10, 2),
'    @PRECIO_DOL decimal(10, 2),
'    @ESTADO smallint,
'    @COD_IMPUESTO VARCHAR(2),
'    @COD_IMPUESTO_BASE VARCHAR(2),
'    @CABYS varchar(13)

Do While Not rs.EOF
  
 strSQL = "exec sp_IW_PRODUCTOInsert '" & Trim(rs!Cod_Cliente) & "', '" & Trim(rs!Cod_Producto) _
        & "', '" & Trim(rs!Descripcion) & "','" & Trim(rs!Descripcion) _
        & "','" & Trim(rs!Cod_Unidad) & "', " & rs!Precio_Col & "," & rs!Precio_Dol _
        & ", 1, '" & rs!Impuesto_Cod & "','" & rs!Impuesto_Cod & "', '" & Trim(rs!Cabys) & "'"
 
 db.Execute strSQL

 sResult = sResult & Space(10) & "exec spPOS_FE_Productos_Result '" & rs!Cod_Cliente & "','" & rs!Cod_Producto _
            & "', '" & glogon.Usuario & "', ''"
  
  i = i + 1
  scEstado.Caption = "Registrando Productos Nuevos [" & i & ", " & iTotal & "]"
  DoEvents
 If Len(sResult) > 20000 Then
   Call ConectionExecute(sResult)
   sResult = ""
 End If
  
 rs.MoveNext
Loop
rs.Close

'Lote Final
If Len(sResult) > 0 Then
  Call ConectionExecute(sResult)
  sResult = ""
End If


scEstado.Caption = ""

Me.MousePointer = vbDefault


Select Case True
    Case btnDatos.Item(0).Checked
        Call btnDatos_Click(0)
        
    Case btnDatos.Item(1).Checked
        Call btnDatos_Click(1)
        
    Case btnDatos.Item(2).Checked
        Call btnDatos_Click(2)
End Select

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub






Private Sub sbSinc_Facturas()
Dim rsRes As New ADODB.Recordset, rsDet As New ADODB.Recordset
Dim sResult As String, i As Long, iTotal As Long

Dim pFactura As String, pTipoFact As String

Dim pActividad As String, pMoneda As String

On Error GoTo vError

Me.MousePointer = vbHourglass


scEstado.Caption = "Cargando parámetros..."
DoEvents

strSQL = "select INCLUYE_POLIZAS, INCLUYE_PRINCIPAL, CABYS, ACTIVIDAD_ECONOMICA, MONEDA from SYS_FE_PARAMETROS"
Call OpenRecordSet(rs, strSQL)
    pActividad = Trim(rs!Actividad_Economica)
    pMoneda = Trim(rs!Moneda)
rs.Close



scEstado.Caption = "Procesando Facturas"
DoEvents

sResult = ""

strSQL = "exec spPOS_FE_Facturas_List '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
    i = 0
    iTotal = rs.RecordCount
End If


Dim CodPais As String, FechaTransac As Date, idEmpresa As String, pCedula As String _
    , codSucursal As String, TerminalPOS As String, ComprobanteInterno As String _
    , SituacionComprobante As String, TipoComprobante As String
Dim pClave50 As String, pClave20 As String, pFecha As Date, pLinea As Integer

Dim pTotalGravado As Currency, pTotalExento As Currency, pImpuesto As Currency, pDescuento As Currency

'Default
pFecha = fxFechaServidor

CodPais = "CRC"
codSucursal = "2"
TerminalPOS = "00001"
SituacionComprobante = "1"

TipoComprobante = "01" 'Factura Electronica

Do While Not rs.EOF
    
    idEmpresa = rs!Cod_Cliente
    
    CodPais = Trim(rs!Pais)
    codSucursal = Trim(rs!Sucursal)
    TerminalPOS = Trim(rs!Terminal)
    
    ComprobanteInterno = rs!cod_Factura
    pTipoFact = rs!Tipo
    
    If Abs(DateDiff("d", rs!fecha, pFecha)) > 1 Then
        FechaTransac = pFecha ' rs!fecha
    Else
        FechaTransac = rs!fecha
    End If
    
    pClave50 = Trim(rs!FE_CLAVE)
    pClave20 = Trim(rs!FE_Numero)

    TipoComprobante = Trim(rs!FE_TIPO)
    TipoComprobante = Format(TipoComprobante, "00")


 strSQL = "select dbo.fxProGrX_Facturas_Existe_Id(" & idEmpresa & ",'" & TipoComprobante & "','" & pClave50 & "') as 'FacturaId'"
 rsRes.Open strSQL, db, adOpenStatic
   pFactura = rsRes!FacturaId
 rsRes.Close
 
 If pFactura = 0 Then
         'Moneda= 'CRC' , Condicion de Venta = '02', Metodo de Pago1 = '05', Situacion = 1
         strSQL = "exec sp_IW_ENC_FACTURAInsert " & rs!Cod_Cliente & ",'" & pClave50 & "','" & pClave20 & "','" & rs!CLIENTE_ID _
                 & "','" & CodPais & "','" & codSucursal & "','02','" & ComprobanteInterno & "',30,'" _
                 & rs!Email & "',1,'" & Format(FechaTransac, "yyyy/mm/dd hh:mm:ss") & "',0,'05','','',''" _
                & ",'', Null, '', '', ''" _
                & ",'1','" & TerminalPOS & "'," & rs!Tipo_Cambio & ",'" & TipoComprobante & "','" & Format(rs!Tipo_ID_FE, "00") _
                & "'," & rs!Servicios_Gravados & ", " & rs!Servicios_Exentos & ", " & rs!Servicios_Exonerados _
                & ", " & rs!Mercaderia_Gravada & ", " & rs!MERCADERIA_EXENTA & ", " & rs!Mercaderia_Exonerada _
                & ", " & rs!Total_Gravado & ", " & rs!Total_Exento & ", " & rs!Total_Exonerado _
                & ", " & rs!Total_Venta & ", " & rs!Total_Descuentos _
                & ", " & rs!TOTAL_VENTA_NETA & "," & rs!TOTAL_IMPUESTO & ", " & rs!TOTAL_IVA_DEVUELTO & ", " & rs!TOTAL_OTROS_CARGOS _
                & ", " & rs!Total_Comprobante _
                & ", 1,'" & glogon.Usuario & "','" & Format(FechaTransac, "yyyy/mm/dd hh:mm:ss") _
                & "','','','Fact. ProGrX','" & pActividad & "'"
        
         rsRes.Open strSQL, db, adOpenStatic
           pFactura = rsRes!id_Factura
         rsRes.Close
        
          '--Factura (Detalle)
         strSQL = "exec spPOS_FE_Facturas_Detalle_List '" & ComprobanteInterno & "', '" & pTipoFact & "'"
         Call OpenRecordSet(rsDet, strSQL)
         Do While Not rsDet.EOF
            strSQL = "exec sp_IW_DET_FACTURAInsert " & rsDet!Cod_Cliente & "," & pFactura & "," & rsDet!Linea _
                   & ",'" & rsDet!Cod_Producto & "'," & rsDet!Cantidad & ", '" & rsDet!Cod_Unidad & "','" & rsDet!Producto_Desc _
                   & "'," & rsDet!Precio_Unitario & "," & rsDet!Total _
                   & "," & rsDet!Descuento_Monto & ", 'DESCUENTO CLIENTES'," & rsDet!SubTotal _
                   & ", '" & rsDet!Impuesto_Codigo & "', " & rsDet!IVA_PORC & ", " & rsDet!Impuesto_Monto _
                   & ", Null, Null, Null, Null, Null, Null, Null" _
                   & ",'" & rsDet!IMPUESTO_CODIGO_BASE & "','" & pClave50 & "'"
            db.Execute strSQL
            
            rsDet.MoveNext
         Loop
         rsDet.Close
  
  End If 'If pFactura = 0 Then


  '--Factura Procesada: ProGrX
  sResult = sResult & Space(10) & "exec spPOS_FE_Factura_Result '" & idEmpresa _
         & "','" & ComprobanteInterno & "','" & pTipoFact & "','" & pFactura & "','" & glogon.Usuario & "'"
  i = i + 1
  scEstado.Caption = " Registrando Facturas [" & i & ", " & iTotal & "]"
  DoEvents

 If Len(sResult) > 20000 Then
   Call ConectionExecute(sResult)
   sResult = ""
 End If

 rs.MoveNext
Loop
rs.Close

'Lote Final
If Len(sResult) > 0 Then
  Call ConectionExecute(sResult)
  sResult = ""
End If

scEstado.Caption = ""

Me.MousePointer = vbDefault


Select Case True
    Case btnDatos.Item(0).Checked
        Call btnDatos_Click(0)
        
    Case btnDatos.Item(1).Checked
        Call btnDatos_Click(1)
        
    Case btnDatos.Item(2).Checked
        Call btnDatos_Click(2)
End Select

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
    'Lote Final
    If Len(sResult) > 0 Then
      Call ConectionExecute(sResult)
      sResult = ""
    End If

End Sub



