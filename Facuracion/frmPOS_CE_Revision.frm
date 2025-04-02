VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmPOS_CE_Revision 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Revisión de Comprobantes Electrónicos"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   12570
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5532
      Left            =   120
      TabIndex        =   0
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   1140
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Buscar"
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   315
      Left            =   2640
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.PushButton btnReenvio 
      Height          =   372
      Left            =   8400
      TabIndex        =   7
      Top             =   1140
      Width           =   3612
      _Version        =   1310723
      _ExtentX        =   6371
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Renumerar y Volver a Enviar"
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "POS: Revisión de Comprobantes Electrónicos"
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
      TabIndex        =   4
      Top             =   360
      Width           =   6612
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   492
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   11892
      _Version        =   1310723
      _ExtentX        =   20976
      _ExtentY        =   868
      _StockProps     =   14
      Caption         =   "Consulta: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.95
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scEstado 
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   7200
      Width           =   11892
      _Version        =   1310723
      _ExtentX        =   20976
      _ExtentY        =   868
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.95
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmPOS_CE_Revision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset


Dim mClienteIdSys As Long, mClienteId As String, mClienteTipoId As String
Dim mClienteEmail As String, mCabys As String

Dim mConServer As String, mConDb As String, mConUser As String, mConKey As String, db As New ADODB.Connection



Private Function fxEstadoHacienda(pClave50 As String) As String
Dim pResultado As String

pResultado = ""

On Error GoTo vError

With glogon.Recordset
    strSQL = "select XML_RESPUESTA From AVS_INTEGRAFAST_01" _
           & " Where ID_CLIENTE_ORIGEN = " & mClienteIdSys _
           & " and CLAVE50 = '" & Trim(pClave50) & "'"
    .Open strSQL, db, adOpenStatic
    
    If Not .BOF And Not .EOF Then
        pResultado = Trim(!XML_RESPUESTA)
    End If
    .Close

End With


fxEstadoHacienda = pResultado
Exit Function

vError:

End Function


Private Sub btnBuscar_Click()
Dim itmX As ListViewItem, pEstado As String

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

With lsw.ColumnHeaders
    .Clear
    .Add , , "No.Factura", 1800
    .Add , , "C.E. Tipo", 1200, vbCenter
    .Add , , "C.E.No", 2200, vbCenter
    .Add , , "C.E.Llave", 3200, vbCenter
    .Add , , "Identificación", 2000
    .Add , , "Nombre", 4000
    .Add , , "Fecha", 1600, vbCenter
    .Add , , "Sub Total", 1600, vbRightJustify
    .Add , , "Descuentos", 1600, vbRightJustify
    .Add , , "IVA", 1600, vbRightJustify
    .Add , , "Total", 1600, vbRightJustify
End With
    
Call FE_Portal_Access
    
strSQL = "select F.*, C.nombre" _
       & " from vpos_factura F left join pv_clientes C on F.cedula = C.cedula" _
       & " Where Estado = 'P' and FE_Tipo = '01' and FE_Estado = 'E'" _
       & " and FECHA between '" & Format(dtpInicio.Value, "yyyy-mm-dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  
  scEstado.Caption = "Revisando Factura.: " & rs!cod_Factura & ", " & rs!Nombre
  DoEvents
  
  pEstado = Mid(fxEstadoHacienda(rs!FE_CLAVE), 1, 1)
  If pEstado = "R" Or pEstado = "E" Then
  
    Set itmX = lsw.ListItems.Add(, , rs!cod_Factura)
        itmX.SubItems(1) = rs!FE_TIPO
        itmX.SubItems(2) = rs!FE_Numero
        itmX.SubItems(3) = rs!FE_CLAVE
        itmX.SubItems(4) = rs!Cedula
        itmX.SubItems(5) = rs!Nombre
        itmX.SubItems(6) = Format(rs!fecha, "yyyy-mm-dd")
        itmX.SubItems(7) = Format(rs!sub_Total, "Standard")
        itmX.SubItems(8) = Format(rs!descuento, "Standard")
        itmX.SubItems(9) = Format(rs!imp_ventas, "Standard")
        itmX.SubItems(10) = Format(rs!Total, "Standard")
        itmX.Tag = rs!Tipo
  End If
  rs.MoveNext
Loop
rs.Close
    
'Cierra Conexion al Portal de FE
db.Close

scEstado.Caption = ""
 
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnReenvio_Click()
Dim i As Long, pTipo As String, pFactura As String

Dim CodPais As String, FechaTransac As Date, pCedula As String _
    , codSucursal As String, TerminalPOS As String, ComprobanteInterno As String _
    , SituacionComprobante As String, TipoComprobante As String
Dim pClave50 As String, pClave20 As String, vFecha As Date


On Error GoTo vError

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

With lsw.ListItems
 For i = 1 To .Count
    If .Item(i).Checked Then
            
            
        pTipo = .Item(i).Tag
        pFactura = .Item(i).Text
            

        strSQL = "exec spPOS_FE_SECUENCIA '01', 'FE'"
        Call OpenRecordSet(rs, strSQL)
        
        CodPais = Trim(rs!Pais)
        codSucursal = Trim(rs!Sucursal)
        TerminalPOS = Trim(rs!Terminal)
        
        SituacionComprobante = "1"
        
        TipoComprobante = Trim(rs!Tipo_Comprobante)
     
        ComprobanteInterno = rs!Consecutivo
        FechaTransac = vFecha
        
        pCedula = Trim(rs!CEDULA_ID)
        pClave50 = fxHacienda_Clave50("506", FechaTransac, pCedula, codSucursal, TerminalPOS, ComprobanteInterno, SituacionComprobante, TipoComprobante)
        pClave20 = fxHacienda_Clave20(codSucursal, TerminalPOS, ComprobanteInterno, TipoComprobante)
          
        rs.Close
    
        'Actualiza Factura
        strSQL = "update pv_facturacion set FE_NUMERO = '" & pClave20 & "', FE_CLAVE = '" & pClave50 _
               & "', FE_ESTADO = 'P', FE_TIPO = '" & TipoComprobante & "'" _
               & " Where COD_FACTURA = '" & pFactura & "' AND TIPO = '" & pTipo & "'"
        Call ConectionExecute(strSQL)
            
    End If
 Next i

End With


Me.MousePointer = vbDefault

MsgBox "Comprobantes Activados!", vbInformation

Call btnBuscar_Click


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Load()

 vModulo = 33
 
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -3, dtpCorte.Value)


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


Private Sub sbPortal_Config_Load()

On Error GoTo vError

strSQL = "select P.*, isnull(C.descripcion,'') as 'Cabys_Desc' " _
       & " from sys_FE_Parametros P left join vINV_Cabys C on P.Cabys = C.COD_BYS"
Call OpenRecordSet(rs, strSQL)

mClienteIdSys = rs!cod_cliente
mClienteTipoId = rs!Tipo_id
mClienteId = rs!Cedula

mClienteEmail = rs!NOTIFICA_EMAIL

mConServer = rs!ACC_SERVER
mConDb = rs!ACC_DB
mConUser = rs!ACC_USR
mConKey = rs!ACC_KEY


mCabys = rs!Cabys & ""

rs.Close

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



Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False


Call sbPortal_Config_Load

End Sub
