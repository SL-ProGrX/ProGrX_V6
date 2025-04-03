VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmCC_DocCKCajas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Documentos : Registro de Cheques de Cajas Chicas"
   ClientHeight    =   4044
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7740
   Icon            =   "frmCC_DocCKCajas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4044
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTab 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7575
      _ExtentX        =   13356
      _ExtentY        =   6795
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Registro de Cheque"
      TabPicture(0)   =   "frmCC_DocCKCajas.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboBanco"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtDocumento"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtMonto"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtDetalle"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ImageList1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "tlbDocumento"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtBeneficiario"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cbo"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Parámetros"
      TabPicture(1)   =   "frmCC_DocCKCajas.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Line3"
      Tab(1).Control(2)=   "Line4"
      Tab(1).Control(3)=   "Label3(0)"
      Tab(1).Control(4)=   "Label3(1)"
      Tab(1).Control(5)=   "ImageList2"
      Tab(1).Control(6)=   "tlbParametros"
      Tab(1).Control(7)=   "txtCuentaCod"
      Tab(1).Control(8)=   "txtCuentaDesc"
      Tab(1).Control(9)=   "lsw"
      Tab(1).ControlCount=   10
      Begin VB.ComboBox cbo 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   840
         Width           =   5295
      End
      Begin VB.TextBox txtBeneficiario 
         Height          =   315
         Left            =   1440
         TabIndex        =   17
         Top             =   1320
         Width           =   5295
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   16
         Top             =   1440
         Width           =   7335
         _ExtentX        =   12933
         _ExtentY        =   4043
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Banco"
            Object.Width           =   9596
         EndProperty
      End
      Begin VB.TextBox txtCuentaDesc 
         Height          =   315
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox txtCuentaCod 
         Height          =   315
         Left            =   -74160
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   720
         Width           =   1455
      End
      Begin MSComctlLib.Toolbar tlbDocumento 
         Height          =   708
         Left            =   6120
         TabIndex        =   9
         Top             =   3000
         Width           =   732
         _ExtentX        =   1291
         _ExtentY        =   1249
         ButtonWidth     =   910
         ButtonHeight    =   1249
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Envio"
               Key             =   "envio"
               Object.ToolTipText     =   "Traslado de Documento a Tesoreria"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6840
         Top             =   1560
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
               Picture         =   "frmCC_DocCKCajas.frx":0342
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCC_DocCKCajas.frx":065C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtDetalle 
         Height          =   795
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2040
         Width           =   5295
      End
      Begin VB.TextBox txtMonto 
         Height          =   315
         Left            =   4560
         TabIndex        =   7
         ToolTipText     =   "# del CK"
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txtDocumento 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         ToolTipText     =   "# del CK"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.ComboBox cboBanco 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   5295
      End
      Begin MSComctlLib.Toolbar tlbParametros 
         Height          =   570
         Left            =   -68400
         TabIndex        =   13
         Top             =   600
         Width           =   855
         _ExtentX        =   1503
         _ExtentY        =   995
         ButtonWidth     =   1561
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Guardar"
               Key             =   "guardar"
               Object.ToolTipText     =   "Guarda Parametros"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   -69240
         Top             =   3120
         _ExtentX        =   995
         _ExtentY        =   995
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCC_DocCKCajas.frx":0976
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Beneficiario"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Indique los Bancos Autorizados a Realizar este tipo de Movimiento"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   15
         Top             =   1200
         Width           =   7335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Indique la cuenta puente utilizada para cerrar el Asiento del Documento en Tesoreria"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   14
         Top             =   360
         Width           =   7335
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74880
         X2              =   -67560
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   -67560
         X2              =   -74880
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta"
         Height          =   375
         Left            =   -74880
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   6840
         X2              =   240
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label1 
         Caption         =   "Detalle"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Monto"
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   4
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Banco"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCC_DocCKCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub sbLimpiaDatos()
Dim strSQL As String, rs As New ADODB.Recordset


On Error GoTo vError

ssTab.Tab = 0

vPaso = True
cboBanco.Clear

strSQL = "select B.id_banco,B.descripcion" _
       & " from tes_banco_asg T inner join Tes_Bancos B on T.id_banco = B.id_banco" _
       & " where T.nombre = '" & glogon.Usuario & "' and B.ckCaja = 1"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  MsgBox "No existen Bancos [Creados/Asignados], verifique en Tesoreria...", vbCritical

Else
 Do While Not rs.EOF
   cboBanco.AddItem IIf(IsNull(rs!Descripcion), "SIN DESCRIPCION", rs!Descripcion)
   cboBanco.ItemData(cboBanco.NewIndex) = rs!id_banco
   rs.MoveNext
 Loop
 rs.MoveFirst
 cboBanco.Text = IIf(IsNull(rs!Descripcion), "SIN DESCRIPCION", rs!Descripcion)
End If
rs.Close

strSQL = "select descripcion from usuarios where nombre = '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  txtBeneficiario = ""
Else
  txtBeneficiario = UCase(Trim(rs!Descripcion & ""))
End If
rs.Close

txtDocumento = ""
txtDetalle = ""
txtMonto = 0

vPaso = False

Call cboBanco_Click

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Public Sub sbgCargaCboConceptos(cbo As ComboBox, vUsuario As String, vBanco As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

cbo.Clear
strSQL = "select rtrim(C.cod_concepto) + ' - ' + C.descripcion as ItmX" _
       & " from tes_conceptos_ASG A inner join Tes_Conceptos C on A.cod_concepto = C.cod_concepto" _
       & " Where A.id_Banco = " & vBanco & " and A.nombre = '" & vUsuario & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cbo.AddItem rs!itmX
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cbo.Text = rs!itmX
End If
rs.Close
End Sub



Private Sub cboBanco_Click()
If vPaso Then Exit Sub
Call sbgCargaCboConceptos(cbo, glogon.Usuario, cboBanco.ItemData(cboBanco.ListIndex))
End Sub

Private Sub Form_Load()
vModulo = 10 'Cuentas Corrientes
Call sbLimpiaDatos
Call Formularios(Me)
Call RefrescaTags(Me)
End Sub

Private Sub sbCargaParametros()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

strSQL = "select * from ase_ck_parametros"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  txtCuentaCod = ""
  txtCuentaDesc = ""
Else
  txtCuentaCod = rs!cod_cuenta
  txtCuentaDesc = fxgCntCuentaDesc(rs!cod_cuenta)
End If
rs.Close

lsw.ListItems.Clear
strSQL = "select id_banco,descripcion,isnull(CKCAJA,0) as Estado from Tes_Bancos order by ckCaja desc"
rs.Open strSQL, glogon.Conection, adOpenForwardOnly
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!id_banco)
     itmX.SubItems(1) = rs!Descripcion
     itmX.Checked = IIf((rs!Estado = 0), vbUnchecked, vbChecked)
 If itmX.Checked Then itmX.ForeColor = vbBlue
 rs.MoveNext
Loop
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lsw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

If Item.Checked Then
  strSQL = "update tes_bancos set ckCaja = 1 where id_banco = " & Item.Text
Else
  strSQL = "update tes_bancos set ckCaja = 0 where id_banco = " & Item.Text
End If
Call ConectionExecute(strSQL)

End Sub

Private Sub ssTab_Click(PreviousTab As Integer)

Select Case ssTab.Tab
 Case 0 'Documentos
 Case 1 'Parametros
   Call sbCargaParametros
End Select
End Sub


Private Function fxVerificaDocumento() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

fxVerificaDocumento = True
vMensaje = ""

'If Len(txtDocumento) = 0 Then vMensaje = vMensaje & vbCrLf & " - No se especifico el # de Documento..."
If Not IsNumeric(txtMonto) Then vMensaje = vMensaje & vbCrLf & " - Monto no es valido..."
If cboBanco.ListCount <= 0 Then vMensaje = vMensaje & vbCrLf & " - No existe Banco Definido..."
If cbo.ListCount <= 0 Then vMensaje = vMensaje & vbCrLf & " - No existen Conceptos para este Banco..."

strSQL = "select isnull(count(*),0) as Existe from cheques where Tipo = 'CK' and Ndocumento = '" _
       & Trim(txtDocumento) & "' and id_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
  vMensaje = vMensaje & vbCrLf & " - Ya Existe un Documento con el Mismo Numero en Tesoreria, verifique..."
End If
rs.Close

If Len(vMensaje) > 0 Then
 fxVerificaDocumento = False
 MsgBox vMensaje, vbExclamation
End If

End Function


Private Sub tlbDocumento_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset, vFecha As Date, lngSol As Long
Dim vDetalle1 As String, vDetalle2 As String, vDetalle3 As String
Dim vDetalle4 As String, vDetalle5 As String, i As Integer
Dim vCtaBanco As String, vCtaPuente As String, vConcepto As String

On Error GoTo vError

If Not fxVerificaDocumento Then Exit Sub

Me.MousePointer = vbHourglass

txtDetalle = UCase(txtDetalle)

vDetalle1 = Mid(txtDetalle, 1, 26)
vDetalle2 = Mid(txtDetalle, 27, 53)
vDetalle3 = Mid(txtDetalle, 54, 80)
vDetalle4 = Mid(txtDetalle, 81, 107)
vDetalle5 = Mid(txtDetalle, 108, 134)

vConcepto = SIFGlobal.fxCodText(cbo.Text)

vFecha = Format(fxFechaServidor, "yyyy/mm/dd")

strSQL = "select ctaConta from Bancos where id_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vCtaBanco = Trim(rs!ctaConta)
rs.Close

''strSQL = "select cod_cuenta from ase_ck_parametros"
''Call OpenRecordSet(rs, strSQL)
''  vCtaPuente = Trim(rs!cod_cuenta)
''rs.Close

strSQL = "select cod_cuenta from tes_conceptos where cod_concepto = '" & vConcepto & "'"
Call OpenRecordSet(rs, strSQL)
  vCtaPuente = Trim(rs!cod_cuenta)
rs.Close


strSQL = "insert tes_transacciones(id_banco,tipo,codigo,beneficiario,monto,fecha_solicitud,estado,estadoi" _
       & ",modulo,submodulo,cta_ahorros,detalle1,detalle2,detalle3,detalle4,detalle5,referencia,op,genera,actualiza" _
       & ",cod_unidad,cod_concepto,user_solicita,user_autoriza,fecha_autorizacion,autoriza)" _
       & " values(" & cboBanco.ItemData(cboBanco.ListIndex) & ",'CK','00001','" _
       & UCase(Trim(txtBeneficiario)) & "'," & CCur(txtMonto) & ",'" & Format(vFecha, "yyyy/mm/dd") _
       & "','P','P','CC','C','','" & vDetalle1 & "','" & vDetalle2 & "','" & vDetalle3 & "','" _
       & vDetalle4 & "','" & vDetalle5 & "',0,0,'S','S'" _
       & ",'OC','" & vConcepto & "','" & glogon.Usuario & "','" & glogon.Usuario _
       & "',dbo.MyGetdate(),'S')"
Call ConectionExecute(strSQL)

'Recupera Consecutivo Tesoreria
strSQL = "select max(nsolicitud) as Solicitud from cheques" _
       & " where tipo = 'CK' and codigo ='00001' and id_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
Call OpenRecordSet(rs, strSQL)
  lngSol = rs!solicitud
rs.Close


'Bancos
strSQL = "insert TES_TRANS_ASIENTO(nsolicitud,cuenta_contable,monto,debehaber,linea,cod_unidad) values(" _
       & lngSol & ",'" & Trim(vCtaBanco) & "'," & CCur(txtMonto) & ",'H',1,'OC')"
Call ConectionExecute(strSQL)

'puente
strSQL = "insert TES_TRANS_ASIENTO (nsolicitud,cuenta_contable,monto,debehaber,linea,cod_unidad) values(" _
       & lngSol & ",'" & Trim(vCtaPuente) & "'," & CCur(txtMonto) & ",'D',2,'OC')"
Call ConectionExecute(strSQL)


'insertar en Documentos

strSQL = "insert ASE_CK_CAJA(id_banco,documento,beneficiario,monto,fecha,usuario,nsolicitud,detalle) values(" _
       & cboBanco.ItemData(cboBanco.ListIndex) & ",'" & Trim(txtDocumento) & "','" & txtBeneficiario _
       & "'," & CCur(txtMonto) & ",dbo.MyGetdate(),'" & glogon.Usuario & "'," & lngSol & ",'" & txtDetalle & "')"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

Call sbLimpiaDatos

MsgBox "Documento Enviado a Tesoreria Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub tlbParametros_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset


On Error GoTo vError

strSQL = "select isnull(count(*),0) as Existe from ase_ck_parametros"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  strSQL = "insert ase_ck_parametros(cod_cuenta) values('" & txtCuentaCod & "')"
Else
  strSQL = "update ase_ck_parametros set cod_cuenta = '" & Trim(txtCuentaCod) & "'"
End If
rs.Close

Call ConectionExecute(strSQL)

MsgBox "Parametros Actualizados...", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub txtCuentaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cod_cuenta"
  gBusquedas.Orden = "cod_cuenta"
  gBusquedas.Consulta = "select cod_cuenta,descripcion from CntX_cuentas"
  gBusquedas.Filtro = " and acepta_movimientos = 'S'"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtCuentaCod = gBusquedas.Resultado
  txtCuentaDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtCuentaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_cuenta,descripcion from CntX_cuentas"
  gBusquedas.Filtro = " and acepta_movimientos = 'S'"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtCuentaCod = gBusquedas.Resultado
  txtCuentaDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus
End Sub

Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub

Private Sub txtMonto_GotFocus()
On Error GoTo vError
txtMonto = CCur(txtMonto)
vError:
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDetalle.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError
txtMonto = Format(CCur(txtMonto), "Standard")
vError:
End Sub
