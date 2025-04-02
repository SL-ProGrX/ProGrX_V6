VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCxC_Conceptos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conceptos de Cuentas por Cobrar"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   15315
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5535
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   15255
      _Version        =   1572864
      _ExtentX        =   26908
      _ExtentY        =   9763
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Conceptos"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Contratos y Estados de Facturas Asociadas"
      Item(1).ControlCount=   11
      Item(1).Control(0)=   "Label2(1)"
      Item(1).Control(1)=   "Label2(0)"
      Item(1).Control(2)=   "lswContratos"
      Item(1).Control(3)=   "lswEstados"
      Item(1).Control(4)=   "Label2(2)"
      Item(1).Control(5)=   "chkAdelantoInfo"
      Item(1).Control(6)=   "cboConcepto"
      Item(1).Control(7)=   "Label2(3)"
      Item(1).Control(8)=   "txtPagadorId"
      Item(1).Control(9)=   "txtPagadorDesc"
      Item(1).Control(10)=   "btnPagador"
      Begin XtremeSuiteControls.ListView lswEstados 
         Height          =   3492
         Left            =   -63040
         TabIndex        =   6
         Top             =   1920
         Visible         =   0   'False
         Width           =   6732
         _Version        =   1572864
         _ExtentX        =   11874
         _ExtentY        =   6159
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswContratos 
         Height          =   3492
         Left            =   -69880
         TabIndex        =   5
         Top             =   1920
         Visible         =   0   'False
         Width           =   6732
         _Version        =   1572864
         _ExtentX        =   11874
         _ExtentY        =   6159
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnPagador 
         Height          =   315
         Left            =   -63040
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Actualizar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtPagadorId 
         Height          =   312
         Left            =   -68680
         TabIndex        =   11
         ToolTipText     =   "Presiones F4 para Consultar"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2984
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkAdelantoInfo 
         Height          =   492
         Left            =   -63040
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   3372
         _Version        =   1572864
         _ExtentX        =   5948
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Utiliza Adelantos como Informativos?"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4935
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   15015
         _Version        =   524288
         _ExtentX        =   26485
         _ExtentY        =   8705
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
         MaxCols         =   499
         ScrollBars      =   2
         SpreadDesigner  =   "frmCxC_Conceptos.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboConcepto 
         Height          =   312
         Left            =   -68680
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   5532
         _Version        =   1572864
         _ExtentX        =   9763
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtPagadorDesc 
         Height          =   312
         Left            =   -67000
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   3852
         _Version        =   1572864
         _ExtentX        =   6794
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pagador (Omisión)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   3
         Left            =   -68680
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Estados de Facturas Admitidos en el registro inicial:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   2
         Left            =   -60400
         TabIndex        =   8
         Top             =   1680
         Visible         =   0   'False
         Width           =   3972
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Contratos Asociados:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   -65680
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   2532
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   -69760
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   1092
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Conceptos/Tipos de Cuentas por Cobrar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   1884
      TabIndex        =   0
      Top             =   360
      Width           =   8892
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   15975
   End
End
Attribute VB_Name = "frmCxC_Conceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub btnPagador_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "Update CxC_Conceptos set Pagador_Default = '" & txtPagadorId.Text _
       & "' where cod_concepto = '" & cboConcepto.ItemData(cboConcepto.ListIndex) & "'"
Call ConectionExecute(strSQL)
If Not glogon.error Then
  Call Bitacora("Registra", "Pagador Defaul: " & txtPagadorId.Text & " al Concepto:" & cboConcepto.ItemData(cboConcepto.ListIndex))
End If

MsgBox "Pagador Default: Registrado Satifactoriamente!", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub cboConcepto_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


If vPaso Or cboConcepto.ListCount <= 0 Then
   Exit Sub
End If


vPaso = True

lswContratos.ListItems.Clear
lswEstados.ListItems.Clear


strSQL = "select C.ADELANTO_INFORMATIVO" _
       & ", isnull(C.PAGADOR_DEFAULT,'') as 'PagadorId', isnull(P.Nombre,'') as 'PagadorDesc'" _
       & " from CxC_Conceptos C left join CxC_Personas P on C.PAGADOR_DEFAULT = P.cedula" _
       & " where C.cod_Concepto = '" & cboConcepto.ItemData(cboConcepto.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
  chkAdelantoInfo.Value = rs!ADELANTO_INFORMATIVO
  txtPagadorId.Text = rs!PagadorId
  txtPagadorDesc.Text = rs!PagadorDesc
rs.Close

strSQL = "select Cnt.Cod_Contrato,Cnt.Descripcion,Asg.registro_Fecha,ASg.Registro_Usuario" _
       & " from CxC_Contratos Cnt left join CXC_CONCEPTOS_CONTRATOS Asg on Cnt.cod_contrato = Asg.cod_contrato" _
       & " and Asg.cod_Concepto = '" & cboConcepto.ItemData(cboConcepto.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswContratos.ListItems.Add(, , rs!COD_CONTRATO)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = rs!REGISTRO_USUARIO & ""
     itmX.SubItems(3) = rs!REGISTRO_FECHA & ""
     
     If Not IsNull(rs!REGISTRO_FECHA) Then itmX.Checked = True
 rs.MoveNext
Loop
rs.Close


strSQL = "select Cnt.FACTURA_ESTADO,Cnt.Descripcion,Asg.registro_Fecha,ASg.Registro_Usuario" _
       & " from CXC_FACTURAS_ESTADOS Cnt left join CXC_CONCEPTOS_FACTURA_ESTADO Asg on Cnt.FACTURA_ESTADO = Asg.FACTURA_ESTADO" _
       & " and Asg.cod_Concepto = '" & cboConcepto.ItemData(cboConcepto.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswEstados.ListItems.Add(, , rs!FACTURA_ESTADO)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = rs!REGISTRO_USUARIO & ""
     itmX.SubItems(3) = rs!REGISTRO_FECHA & ""
     
     If Not IsNull(rs!REGISTRO_FECHA) Then itmX.Checked = True
 rs.MoveNext
Loop
rs.Close


vPaso = False


End Sub


Private Sub chkAdelantoInfo_Click()
Dim strSQL As String, vCodigo As String

If vPaso Then Exit Sub

On Error GoTo vError

vCodigo = cboConcepto.ItemData(cboConcepto.ListIndex)

strSQL = "update cxc_Conceptos set ADELANTO_INFORMATIVO = " & chkAdelantoInfo.Value _
       & " where cod_concepto = '" & vCodigo & "'"
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
vModulo = 31
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 31
vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


With lswContratos.ColumnHeaders
  .Add , , "Código", 1200
  .Add , , "Descripción", 2000
  .Add , , "Usuario", 1500
  .Add , , "Fecha", 1900
End With

With lswEstados.ColumnHeaders
  .Add , , "Estado", 1200
  .Add , , "Descripción", 2000
  .Add , , "Usuario", 1500
  .Add , , "Fecha", 1900
End With

Call Formularios(Me)
Call RefrescaTags(Me)

chkAdelantoInfo.Enabled = vGrid.Enabled
lswContratos.Checkboxes = vGrid.Enabled
lswEstados.Checkboxes = vGrid.Enabled
btnPagador.Enabled = vGrid.Enabled

tcMain.Item(0).Selected = True

strSQL = "select * from CxC_Conceptos" _
      & " order by cod_concepto"
Call sbCargaGridLocal(vGrid, 10, strSQL)



End Sub

Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
     Case 1 'Codigo
        vGrid.Text = CStr(rs!cod_Concepto)
     Case 2 'descripcion
        vGrid.Text = CStr(rs!Descripcion)
     Case 3 'Cuenta
        vGrid.Text = fxgCntCuentaFormato(True, rs!cod_cuenta & "")
     Case 4 'Cuenta Salida
        vGrid.Text = fxgCntCuentaFormato(True, rs!cod_Cuenta_Salida & "")
     Case 5 'Requiere Contrato
        vGrid.Value = rs!Requiere_Contrato
     Case 6 'Requiere Documento
        vGrid.Value = rs!Requiere_Documento
     Case 7 'Genera Desembolso
        vGrid.Value = rs!Genera_Desembolso
     Case 8 'Proceso de Descuento
        vGrid.Value = rs!Proceso_Descuento
     Case 9 'Disponible por Persona
        vGrid.Text = Format(rs!MONTO_MAX & "Standard")
     Case 10 'Concepto Activo
        vGrid.Value = rs!activo
    End Select
  
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

Me.MousePointer = vbDefault

End Sub



Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCuenta As String, vCuentaSalida As String

On Error GoTo vError

vGrid.Col = 1
fxGuardar = 0
If Trim(vGrid.Text) = "" Then Exit Function

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 3
vCuenta = fxgCntCuentaFormato(False, vGrid.Text)
vGrid.Col = 4
vCuentaSalida = fxgCntCuentaFormato(False, vGrid.Text)


vGrid.Col = 1
strSQL = "select isnull(count(*),0) as Existe from CxC_Conceptos " _
       & " where cod_concepto = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar

  strSQL = "insert into CxC_Conceptos(cod_concepto,descripcion,cod_cuenta,cod_cuenta_Salida,requiere_contrato,requiere_documento" _
         & ",genera_desembolso,proceso_descuento, MONTO_MAX, Activo, ADELANTO_INFORMATIVO, REGISTRO_FECHA, REGISTRO_USUARIO) values('" & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "', '" & vCuenta & "', '" & vCuentaSalida & "',"
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 6
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 7
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 8
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 9
  strSQL = strSQL & CCur(vGrid.Text) & ","
  
  vGrid.Col = 10
  strSQL = strSQL & vGrid.Value & ", 0, dbo.MyGetdate(), '" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Concepto de Cuenta x Cobrar: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update CxC_Conceptos set descripcion = '" & vGrid.Text & "', cod_Cuenta = '" & vCuenta _
        & "', cod_cuenta_salida = '" & vCuentaSalida & "', requiere_contrato ="
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Value & ", requiere_documento = "
 vGrid.Col = 6
 strSQL = strSQL & vGrid.Value & ", genera_desembolso = "
 vGrid.Col = 7
 strSQL = strSQL & vGrid.Value & ", proceso_descuento = "
 vGrid.Col = 8
 strSQL = strSQL & vGrid.Value & ", MONTO_MAX = "
 vGrid.Col = 9
 strSQL = strSQL & CCur(vGrid.Text) & ", Activo = "
 vGrid.Col = 10
 strSQL = strSQL & vGrid.Value & " where cod_concepto = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Concepto de Cuenta x Cobrar: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function




Private Sub lswContratos_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, vCodigo As String

If vPaso Then Exit Sub

On Error GoTo vError

vCodigo = cboConcepto.ItemData(cboConcepto.ListIndex)

If Item.Checked Then
   strSQL = "insert CxC_Conceptos_Contratos(cod_contrato,cod_concepto,registro_usuario,registro_fecha) values('" _
          & Item.Text & "','" & vCodigo & "','" & glogon.Usuario & "',dbo.MyGetdate())"
   Item.SubItems(2) = glogon.Usuario
   Item.SubItems(3) = Date

Else
   strSQL = "delete CxC_Conceptos_Contratos where cod_contrato  = '" & Item.Text & "' and cod_concepto = '" & vCodigo & "'"
   
   Item.SubItems(2) = ""
   Item.SubItems(3) = ""
   
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswEstados_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, vCodigo As String

If vPaso Then Exit Sub

On Error GoTo vError

vCodigo = cboConcepto.ItemData(cboConcepto.ListIndex)

If Item.Checked Then
   strSQL = "insert CXC_CONCEPTOS_FACTURA_ESTADO(factura_estado,cod_concepto,registro_usuario,registro_fecha) values('" _
          & Item.Text & "','" & vCodigo & "','" & glogon.Usuario & "',dbo.MyGetdate())"
   Item.SubItems(2) = glogon.Usuario
   Item.SubItems(3) = Date

Else
   strSQL = "delete CXC_CONCEPTOS_FACTURA_ESTADO where factura_estado  = '" & Item.Text & "' and cod_concepto = '" & vCodigo & "'"
   
   Item.SubItems(2) = ""
   Item.SubItems(3) = ""
   
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String

If Item.Index = 0 Then Exit Sub

vPaso = True
strSQL = "select rtrim(cod_concepto) as 'IdX',  rTrim(descripcion) as 'ItmX' from CxC_Conceptos order by cod_concepto"
    Call sbCbo_Llena_New(cboConcepto, strSQL, False, True)
vPaso = False

Call cboConcepto_Click

End Sub

Private Sub txtPagadorId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from CxC_Personas"
  gBusquedas.Filtro = " and Rol_Pagador = 1"
  frmBusquedas.Show vbModal
  txtPagadorId.Text = gBusquedas.Resultado
  txtPagadorDesc.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Cuenta Contable (Formato)
If (vGrid.ActiveCol = 3 Or vGrid.ActiveCol = 4) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = fxgCntCuentaFormato(True, vGrid.Text)
End If

'Cuenta Contable (Busqueda)
If (vGrid.ActiveCol = 3 Or vGrid.ActiveCol = 4) And KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gCuenta
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete CxC_Conceptos where cod_concepto = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Concepto de Cuenta x Cobrar: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



