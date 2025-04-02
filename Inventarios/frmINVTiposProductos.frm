VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmInvTiposProductos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Productos"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   10995
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5892
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   10812
      _Version        =   1441792
      _ExtentX        =   19071
      _ExtentY        =   10393
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
      Item(0).Caption =   "General"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Sub Categorías"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "vGridSub"
      Item(1).Control(1)=   "Label1"
      Item(1).Control(2)=   "cboCategoria"
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5172
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   10572
         _Version        =   524288
         _ExtentX        =   18648
         _ExtentY        =   9123
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
         MaxCols         =   6
         ScrollBars      =   2
         SpreadDesigner  =   "frmINVTiposProductos.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridSub 
         Height          =   4572
         Left            =   -69760
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   10212
         _Version        =   524288
         _ExtentX        =   18013
         _ExtentY        =   8064
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
         MaxCols         =   4
         ScrollBars      =   2
         SpreadDesigner  =   "frmINVTiposProductos.frx":06B6
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboCategoria 
         Height          =   312
         Left            =   -66880
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   7332
         _Version        =   1441792
         _ExtentX        =   12938
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label1 
         Caption         =   "Categoria:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   -68080
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   972
      End
   End
End
Attribute VB_Name = "frmInvTiposProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub cboCategoria_Click()

If vPaso Or cboCategoria.ListCount = 0 Then Exit Sub

vGridSub.MaxRows = 0
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select COD_LINEA_SUB,  DESCRIPCION, ACTIVO, CABYS" _
    & " From PV_PROD_CLASIFICA_SUB where COD_PRODCLAS = " & cboCategoria.ItemData(cboCategoria.ListIndex)
Call sbCargaGrid(vGridSub, 4, strSQL, True)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical



End Sub

Private Sub Form_Activate()
vModulo = 32
End Sub

Private Sub Form_Load()
Dim strSQL As String

Call sbToolBarIconos(tlb)

vModulo = 32
vGrid.AppearanceStyle = fxGridStyle

Call Formularios(Me)
Call RefrescaTags(Me)

If tlb.Buttons(1).Enabled = False Then
    vGrid.Enabled = False
    vGridSub.Enabled = False
End If

tcMain.Item(0).Selected = True

strSQL = "select T.cod_prodclas,T.descripcion,T.costeo,T.valuacion,T.cod_cuenta,T.cod_Alter,C.descripcion as CtaDesc" _
       & " from pv_prod_clasifica T left join CntX_cuentas C on T.cod_cuenta = C.cod_cuenta and C.cod_contabilidad = " & GLOBALES.gEnlace _
       & " order by T.cod_prodclas"
Call sbCargaGridLocal(vGrid, 6, strSQL)

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
     Case 1
        vGrid.Text = CStr(rs!cod_prodclas)
     Case 2
        vGrid.Text = CStr(rs!Descripcion)
     Case 3
        vGrid.Text = CStr(rs!Costeo)
     Case 4
        vGrid.Text = CStr(rs!valuacion)
     Case 5
        vGrid.Text = fxgCntCuentaFormato(True, rs!cod_cuenta)
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = rs!CtaDesc & ""
        vGrid.TextTip = TextTipFixed
     Case 6
        vGrid.Text = CStr(rs!Cod_Alter & "")
      
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
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1
If vGrid.Text = "" Then 'Insertar
  vGrid.col = 2
  If Trim(vGrid.Text) = "" Then Exit Function
  strSQL = "insert into pv_prod_clasifica(descripcion,costeo,valuacion,cod_cuenta,cod_alter) values('"
  strSQL = strSQL & UCase(vGrid.Text) & "','"
  vGrid.col = 3
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.col = 4
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.col = 5
  strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "','"
  vGrid.col = 6
  strSQL = strSQL & vGrid.Text & "')"
  

  Call ConectionExecute(strSQL)

  vGrid.col = 2
  strSQL = "select max(cod_prodclas) as Ultimo from pv_prod_clasifica " _
         & "where descripcion = '" & vGrid.Text & "'"
  Call OpenRecordSet(rs, strSQL)
  fxGuardar = IIf(IsNull(rs!ultimo), 0, rs!ultimo)

  vGrid.col = 1
  Call Bitacora("Registra", "Tipo Producto Cod: " & vGrid.Text)

  rs.Close

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update pv_prod_clasifica set descripcion = '" & vGrid.Text & "',costeo = '"
 vGrid.col = 3
 strSQL = strSQL & vGrid.Text & "',valuacion = '"
 vGrid.col = 4
 strSQL = strSQL & vGrid.Text & "',cod_cuenta = '"
 vGrid.col = 5
 strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "', Cod_Alter = '"
 vGrid.col = 6
 strSQL = strSQL & vGrid.Text & "'  where cod_prodclas = "
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 fxGuardar = vGrid.Text
 
 Call Bitacora("Modifica", "Tipo Producto Cod: " & vGrid.Text)

End If

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Function fxGuardarSub() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardarSub = 0

With vGridSub

.Row = .ActiveRow
.col = 1
  strSQL = "select count(*) as Existe from pv_prod_clasifica_Sub " _
         & "where COD_PRODCLAS = " & cboCategoria.ItemData(cboCategoria.ListIndex) & " and COD_LINEA_SUB = '" & .Text & "'"
  Call OpenRecordSet(rs, strSQL)


If rs!Existe = 0 Then 'Insertar
  If Trim(.Text) = "" Then Exit Function
  strSQL = "insert into pv_prod_clasifica_Sub(COD_PRODCLAS,COD_LINEA_SUB, DESCRIPCION, Activo, CABYS, REGISTRO_FECHA, REGISTRO_USUARIO)" _
         & " values(" & cboCategoria.ItemData(cboCategoria.ListIndex) & ",'" & .Text & "','"
  .col = 2
  strSQL = strSQL & .Text & "',"
  .col = 3
  strSQL = strSQL & .Value & ",'"
  .col = 4
  strSQL = strSQL & .Text & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  .col = 1
  fxGuardarSub = 1

  .col = 1
  Call Bitacora("Registra", "Tipo Producto, Cat:" & cboCategoria.ItemData(cboCategoria.ListIndex) & " - Sub: " & .Text)


Else 'Actualizar

 .col = 2
 strSQL = "update pv_prod_clasifica_Sub set descripcion = '" & .Text & "',Activo = "
 .col = 3
 strSQL = strSQL & .Value & ", CABYS = '"
 .col = 4
 strSQL = strSQL & .Text & "' Where cod_prodclas = " & cboCategoria.ItemData(cboCategoria.ListIndex)
 .col = 1
 strSQL = strSQL & " and COD_LINEA_SUB = '" & .Text & "'"
 Call ConectionExecute(strSQL)

 fxGuardarSub = 1
 
  Call Bitacora("Modifica", "Tipo Producto, Cat:" & cboCategoria.ItemData(cboCategoria.ListIndex) & " - Sub: " & .Text)

End If
  rs.Close

End With

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

If Item.Index = 1 Then
    strSQL = "select COD_PRODCLAS as 'IdX', DESCRIPCION AS 'itmX' " _
           & " From PV_PROD_CLASIFICA"
    vPaso = True
        Call sbCbo_Llena_New(cboCategoria, strSQL, False, True)
    vPaso = False
    
    Call cboCategoria_Click
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer, strSQL As String

On Error Resume Next

Select Case UCase(Button.Key)
  Case "NUEVO"
    vGrid.MaxRows = vGrid.MaxRows + 1

  Case "BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = 6 Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        strSQL = "delete pv_prod_clasifica where cod_prodclas = " & vGrid.Text
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.col = 2
        Call Bitacora("Elimina", "Tipo Producto : " & strSQL & " - " & vGrid.Text)
        vGrid.col = 1
        strSQL = "select T.cod_prodclas,T.descripcion,T.costeo,T.valuacion,T.cod_cuenta, T.Cod_Alter,C.descripcion as CtaDesc" _
               & " from pv_prod_clasifica T left join CntX_cuentas C on T.cod_cuenta = C.cod_cuenta and C.cod_contabilidad = " & GLOBALES.gEnlace _
               & " order by T.cod_prodclas"
        
        Call sbCargaGridLocal(vGrid, 6, strSQL)
     End If
  Case "REPORTES"
     Call sbInvReportes("TiposProductos", "Tipos de Productos", "Listado", "")

  Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 1
  vGrid.Text = i
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Formato de Cuenta Contable
If vGrid.ActiveCol = 5 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = fxgCntCuentaFormato(True, vGrid.Text)
End If

'Consulta Cuentas Contables
If vGrid.ActiveCol = 5 And KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gCuenta
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


End Sub

Private Sub vGridSub_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

With vGridSub

    If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
      i = fxGuardarSub
      If i = 0 Then Exit Sub
      .Row = .ActiveRow
      .col = 1
'      .Text = i
      If .MaxRows <= .ActiveRow Then
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
      End If
    End If
    
    If KeyCode = vbKeyF4 Then
            .Row = .ActiveRow
            .col = 4
        
            gBusquedas.Convertir = "N"
            gBusquedas.Columna = "Cod_ByS"
            gBusquedas.Orden = "Cod_ByS"
            gBusquedas.Consulta = "select Cod_ByS,Descripcion from vINV_Cabys"
            gBusquedas.Filtro = ""
            frmBusquedas.Show vbModal
            .Text = gBusquedas.Resultado
    End If
    
    'Inserta Linea
    If KeyCode = vbKeyInsert Then
        .MaxRows = .MaxRows + 1
        .InsertRows .ActiveRow, 1
        .Row = .ActiveRow
    End If

End With

End Sub

