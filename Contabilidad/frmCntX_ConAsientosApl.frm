VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmCntX_ConAsientosApl 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aplicación de Asientos Consolidados"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   492
      Left            =   5280
      TabIndex        =   7
      Top             =   5520
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Picture         =   "frmCntX_ConAsientosApl.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   492
      Left            =   6720
      TabIndex        =   8
      Top             =   5520
      Width           =   1572
      _Version        =   1245185
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Aplicar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Picture         =   "frmCntX_ConAsientosApl.frx":0A1E
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   4572
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   8292
      _ExtentX        =   14631
      _ExtentY        =   8070
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod.Asiento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descripción"
         Object.Width           =   7832
      EndProperty
   End
   Begin VB.TextBox txtMes 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   4
      ToolTipText     =   "(F4) Mes del Periodo"
      Top             =   480
      Width           =   315
   End
   Begin VB.TextBox txtAnio 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2496
      MaxLength       =   4
      TabIndex        =   3
      ToolTipText     =   "(F4) Año del Periodo"
      Top             =   480
      Width           =   528
   End
   Begin VB.TextBox txtPeriodo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3084
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "(F4) Descripción del Periodo"
      Top             =   480
      Width           =   4185
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   5172
   End
   Begin VB.Label Label1 
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   480
      Width           =   1428
   End
   Begin VB.Label Label1 
      Caption         =   "Consolidación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   1452
   End
End
Attribute VB_Name = "frmCntX_ConAsientosApl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMensaje As String

Private Function fxVerificaPeriodo()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select cod_consolida from con_periodos where cod_consolida = " & cbo.ItemData(cbo.ListIndex) _
       & " and Estado = 'P' and Anio = " & txtAnio & " and mes = " & txtMes
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs.BOF And rs.EOF Then
  vMensaje = " - El periodo no existe, o ya fue cerrado en la consolidación..."
Else
  vMensaje = ""
End If
rs.Close


If Len(vMensaje) = 0 Then
    fxVerificaPeriodo = True
Else
    fxVerificaPeriodo = False
End If
End Function


Private Sub cbo_Click()
lsw.ListItems.Clear
End Sub


Private Function fxExisteMovimiento(lngAnio As Long, iMes As Integer, strCuenta As String, vConsolida As Long) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String



strSQL = "select isnull(count(*),0) as existe from con_movimientos where COD_CONTABILIDAD = " _
       & gCntX_Parametros.CodigoConta & " and anio = " & lngAnio & " and mes = " & iMes _
       & " and cod_cuenta = '" & strCuenta & "' and cod_consolida = " & vConsolida

rsX.Open strSQL, glogon.Conection, adOpenStatic
fxExisteMovimiento = IIf((rsX!Existe = 0), False, True)
rsX.Close

End Function



Private Sub sbGuardaMovimiento(lngAnio As Long, iMes As Integer, vCuentaActual As String _
                                , curDebe As Currency, curHaber As Currency, vConsolida As Long, iCodEmpresa As Long)
Dim strSQL As String

If fxExisteMovimiento(lngAnio, iMes, vCuentaActual, vConsolida) Then
 
 strSQL = "update con_movimientos set total_debitos = total_debitos + " & curDebe _
        & ", total_creditos = total_creditos + " & curHaber _
        & " where COD_CONTABILIDAD = " & iCodEmpresa & " and anio = " _
        & lngAnio & " and mes = " & iMes & " and cod_cuenta = '" _
        & vCuentaActual & "' and cod_consolida = " & vConsolida
 glogon.Conection.Execute strSQL
 
Else
  strSQL = "insert into con_movimientos(anio,mes,COD_CONTABILIDAD,cod_cuenta,saldo_inicial" _
         & ",total_debitos,total_creditos,cod_consolida) values(" & lngAnio & "," & iMes & "," _
         & iCodEmpresa & ",'" & vCuentaActual & "',0," & curDebe _
         & "," & curHaber & "," & vConsolida & ")"
 glogon.Conection.Execute strSQL

End If
End Sub

Private Sub sbMayorizar(vNumeroAsiento As String, vFecha As Date, vConsolida As Long)
Dim rs As New ADODB.Recordset, strSQL As String, curDebe As Currency, curHaber As Currency
Dim lngAnio As Long, iMes As Integer, vCuentaMadre As String, rsTemp As New ADODB.Recordset
Dim iCodEmpresa As Long


Screen.MousePointer = vbHourglass

iMes = Month(vFecha)
lngAnio = Year(vFecha)

strSQL = "select COD_CONTABILIDAD from CNTX_CONSOLIDA_DEFINICION where cod_consolida = " & vConsolida
rs.Open strSQL, glogon.Conection, adOpenStatic
iCodEmpresa = rs!COD_CONTABILIDAD
rs.Close

strSQL = "select A.*,B.Tipo_Cuenta,B.Cuenta_Madre,T.clasificacion" _
       & " from con_asientos_detalle A inner join cuentas B" _
       & " On A.cod_cuenta = B.cod_cuenta inner join tipos_cuentas T on B.tipo_cuenta = T.tipo_cuenta" _
       & " where B.COD_CONTABILIDAD = " & iCodEmpresa _
       & " and A.cod_asiento = '" & vNumeroAsiento & "' and A.cod_consolida = " & vConsolida
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  curDebe = 0
  curHaber = 0
  Select Case rs!Clasificacion
     Case "G", "A", "O", "V"
        curDebe = rs!Debitos
        curHaber = rs!Creditos * -1
     Case "I", "C", "P"
        curDebe = rs!Debitos * -1
        curHaber = rs!Creditos
  End Select
   
  vCuentaMadre = rs!cuenta_madre
  
  'Registra movimientos a la Cuenta y afecta en cascada a las cuentas madres
  Call sbGuardaMovimiento(lngAnio, iMes, rs!cod_cuenta, curDebe, curHaber, vConsolida, iCodEmpresa)
  
  Do While vCuentaMadre <> ""
   Call sbGuardaMovimiento(lngAnio, iMes, vCuentaMadre, curDebe, curHaber, vConsolida, iCodEmpresa)
   strSQL = "select cuenta_madre from cuentas where COD_CONTABILIDAD = " _
          & iCodEmpresa & " and cod_cuenta = '" & vCuentaMadre & "'"
   rsTemp.CursorLocation = adUseServer
   rsTemp.Open strSQL, glogon.Conection, adOpenStatic
   If rsTemp.EOF And rsTemp.BOF Then
     vCuentaMadre = ""
   Else
     vCuentaMadre = rsTemp!cuenta_madre
   End If
   rsTemp.Close
  Loop
  
  rs.MoveNext
Loop


'Actualiza el Estado del Asiento
strSQL = "update con_asientos set Aplicado = 'S'" _
       & " where cod_consolida = " & vConsolida _
       & " and cod_asiento = '" & vNumeroAsiento & "'"
glogon.Conection.Execute strSQL

rs.Close

Screen.MousePointer = vbDefault

End Sub


Private Sub cmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

If Not fxVerificaPeriodo Then
   MsgBox vMensaje, vbCritical
   Exit Sub
End If

strSQL = "select * from con_asientos where aplicado = 'N' and cod_consolida = " _
       & cbo.ItemData(cbo.ListIndex) & " and datepart(mm,fecha) = " & txtMes _
       & " and datepart(yyyy,fecha) = " & txtAnio
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  Call sbMayorizar(rs!cod_asiento, Format(rs!fecha, "yyyy/mm/dd"), cbo.ItemData(cbo.ListIndex))
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

MsgBox "Asientos de Consolidados Mayorizados Satisfactoriamente...", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
'  Resume
End Sub

Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

lsw.ListItems.Clear

strSQL = "select * from con_asientos where aplicado = 'N' and cod_consolida = " _
       & cbo.ItemData(cbo.ListIndex) & " and datepart(mm,fecha) = " & txtMes _
       & " and datepart(yyyy,fecha) = " & txtAnio
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!cod_asiento)
      itmX.SubItems(1) = rs!fecha
      itmX.SubItems(2) = rs!Descripcion & ""
      itmX.Checked = True
  rs.MoveNext
Loop
rs.Close
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
       
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset, vPaso As Boolean

 
 vPaso = False
 
 strSQL = "select * from CNTX_CONSOLIDA_DEFINICION"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 cbo.Clear
 
 Do While Not rs.EOF
   cbo.AddItem Format(rs!COD_CONSOLIDA, "000") & " - " & rs!Descripcion
   cbo.ItemData(cbo.NewIndex) = rs!COD_CONSOLIDA
   vPaso = True
   rs.MoveNext
 Loop
 
 If vPaso Then
   rs.MoveFirst
   cbo.Text = Format(rs!COD_CONSOLIDA, "000") & " - " & rs!Descripcion
 End If
 rs.Close

 txtMes = Month(fxFechaServidor)
 txtAnio = Year(fxFechaServidor)


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub txtAnio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdBuscar.SetFocus
End Sub

Private Sub txtMes_Change()
On Error GoTo vError
txtPeriodo = fxCntX_PeriodoDesc(txtAnio, txtMes)
vError:
End Sub

Private Sub txtMes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAnio.SetFocus
End Sub

Private Sub txtAnio_Change()
On Error GoTo vError
txtPeriodo = fxCntX_PeriodoDesc(txtAnio, txtMes)
vError:
End Sub


Private Sub txtPeriodo_Change()
lsw.ListItems.Clear
End Sub
