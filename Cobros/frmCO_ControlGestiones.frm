VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmCO_ControlGestiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Gestiones"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   14925
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5892
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   14652
      _Version        =   1310723
      _ExtentX        =   25844
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
      Item(0).Caption =   "Gestiones"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Seguridad"
      Item(1).ControlCount=   5
      Item(1).Control(0)=   "lswGestiones"
      Item(1).Control(1)=   "lswUsuarios"
      Item(1).Control(2)=   "scSeguridad(0)"
      Item(1).Control(3)=   "scSeguridad(1)"
      Item(1).Control(4)=   "Label2"
      Begin XtremeSuiteControls.ListView lswUsuarios 
         Height          =   4812
         Left            =   -62680
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   7092
         _Version        =   1310723
         _ExtentX        =   12509
         _ExtentY        =   8488
         _StockProps     =   77
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.ListView lswGestiones 
         Height          =   4812
         Left            =   -69880
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   7092
         _Version        =   1310723
         _ExtentX        =   12509
         _ExtentY        =   8488
         _StockProps     =   77
         BackColor       =   -2147483643
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
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5412
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   14532
         _Version        =   524288
         _ExtentX        =   25633
         _ExtentY        =   9546
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
         MaxCols         =   503
         ScrollBars      =   2
         SpreadDesigner  =   "frmCO_ControlGestiones.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   12
         Left            =   -69640
         TabIndex        =   9
         Top             =   1920
         Visible         =   0   'False
         Width           =   132
      End
      Begin XtremeShortcutBar.ShortcutCaption scSeguridad 
         Height          =   312
         Index           =   1
         Left            =   -62680
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   7092
         _Version        =   1310723
         _ExtentX        =   12509
         _ExtentY        =   550
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption scSeguridad 
         Height          =   312
         Index           =   0
         Left            =   -69880
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   7212
         _Version        =   1310723
         _ExtentX        =   12721
         _ExtentY        =   550
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   264
      Left            =   12120
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   2436
      _ExtentX        =   4286
      _ExtentY        =   476
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   300
         Left            =   1200
         TabIndex        =   3
         Top             =   1560
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "frmCO_ControlGestiones.frx":097F
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Tab 1"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Tab 2"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
      End
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   492
      Left            =   2400
      TabIndex        =   6
      Top             =   240
      Width           =   6372
      _Version        =   1310723
      _ExtentX        =   11239
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Tipos de Gestiones de Cobros"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15132
   End
End
Attribute VB_Name = "frmCO_ControlGestiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub Form_Activate()
vModulo = 4
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 4


tcMain.Item(0).Selected = True

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lswGestiones.ColumnHeaders
    .Add , , "Código", 2100
    .Add , , "Descripción", 4950
End With

With lswUsuarios.ColumnHeaders
    .Add , , "Usuario", 2100
    .Add , , "Nombre", 4950
End With

strSQL = "select cod_gestion,descripcion,codigo_referencia,monto,MODIFICA_USUARIO,MODIFICA_DESVIACION,isnull(COD_CUENTA,'') as COD_CUENTA" _
      & ", case when NIVEL_GESTION = 'U' then 'Usuario' else 'Sistema' end as NIVEL_GESTION,ACCESO_RESTRINGIDO" _
      & ",MRECUPERACION,IVA_PORCENTAJE, ESTADO" _
      & " from cbr_gestiones" _
      & " order by cod_gestion"
Call sbCargaGridLocal(12, strSQL)


Call Formularios(Me)
Call RefrescaTags(Me)

If tlb.Buttons(1).Enabled = False Then vGrid.Enabled = False
If tlb.Buttons(1).Enabled = False Then lswUsuarios.Enabled = False


End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from cbr_gestiones " _
       & " where cod_gestion = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

vGrid.col = 7

If Not fxgCntCuentaValida(fxgCntCuentaFormato(False, vGrid.Text)) Then
  MsgBox "Cuenta Contable no es válida...", vbCritical
  vGrid.Text = Empty
  Exit Function
End If

vGrid.col = 1

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into cbr_gestiones(cod_gestion,descripcion,codigo_referencia,monto,MODIFICA_USUARIO,MODIFICA_DESVIACION" _
         & ",COD_CUENTA,NIVEL_GESTION,ACCESO_RESTRINGIDO,MRECUPERACION,IVA_PORCENTAJE, ESTADO) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.col = 2 'Descipcion
  strSQL = strSQL & UCase(vGrid.Text) & "','"
 
  vGrid.col = 3 ' Cod Referencia
  strSQL = strSQL & UCase(vGrid.Text) & "',"
  
  vGrid.col = 4 ' Monto
  strSQL = strSQL & CCur(vGrid.Text) & ","
  
  vGrid.col = 5 ' Modifica Usuario
  strSQL = strSQL & vGrid.Value & ","
  
  vGrid.col = 6 ' Modifica Desviacion
  strSQL = strSQL & CCur(vGrid.Text) & ",'"
  
  vGrid.col = 7 ' Código de Cuenta Contable
  strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "','"
  
  vGrid.col = 8 ' NIVEL_GESTION
  strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "',"
  
  vGrid.col = 9 ' ACCESO_RESTRINGIDO
  strSQL = strSQL & vGrid.Value & ","

  vGrid.col = 10 'Monitorea Recuperacion
  strSQL = strSQL & vGrid.Value & ","

  vGrid.col = 11 ' IVA
  strSQL = strSQL & CCur(vGrid.Text) & ","


  vGrid.col = 12 ' Estado
  strSQL = strSQL & vGrid.Value & ")"

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Tipo Gestión Cobros Id:" & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update cbr_gestiones set descripcion = '" & vGrid.Text & "',codigo_referencia = '"
 vGrid.col = 3
 strSQL = strSQL & UCase(vGrid.Text) & "', monto = "
 vGrid.col = 4
 strSQL = strSQL & CCur(vGrid.Text) & ", MODIFICA_USUARIO = "
 vGrid.col = 5
 strSQL = strSQL & vGrid.Value & ", MODIFICA_DESVIACION = "
 vGrid.col = 6
 strSQL = strSQL & CCur(vGrid.Text) & ", COD_CUENTA = '"
 vGrid.col = 7
 strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "', NIVEL_GESTION = '"
 vGrid.col = 8
 strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "', ACCESO_RESTRINGIDO = "
  vGrid.col = 9
 strSQL = strSQL & vGrid.Value & ", MRECUPERACION = "
 vGrid.col = 10
 strSQL = strSQL & vGrid.Value & ", IVA_PORCENTAJE = "
 vGrid.col = 11
 strSQL = strSQL & CCur(vGrid.Text) & ", ESTADO = "
 
 vGrid.col = 12
 strSQL = strSQL & vGrid.Value & " where cod_gestion = '"
 
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

  Call Bitacora("Modifica", "Tipo Gestión Cobros Id:" & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub lswGestiones_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

scSeguridad.Item(1).Caption = "Gestion: " & Trim(Item.SubItems(1))
scSeguridad.Item(1).Tag = Trim(Item.Text)

With lswUsuarios
 .ListItems.Clear
  
 strSQL = "select U.USUARIO,U.NOMBRE,GU.USUARIO as asignado " _
        & " from CBR_USUARIOS U left join CBR_GESTIONES_USUARIOS GU on GU.USUARIO = U.USUARIO " _
        & " and GU.COD_GESTION = '" & scSeguridad.Item(1).Tag _
        & "' where U.ESTADO = 1 " _
        & " order by U.NOMBRE"
 Call OpenRecordSet(rs, strSQL, 0)
   
  vPaso = True
 
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!Usuario)
      itmX.SubItems(1) = rs!Nombre
      If Not IsNull(rs!Asignado) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
 Loop
 rs.Close
 
 vPaso = False
 
End With

End Sub


Private Sub lswUsuarios_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vPaso Or scSeguridad.Item(1).Tag = "" Then
    Exit Sub
End If


If Item.Checked Then
  strSQL = "insert CBR_GESTIONES_USUARIOS(COD_GESTION,USUARIO) values('" & scSeguridad.Item(1).Tag _
         & "','" & Item.Text & "')"
Else
  strSQL = "delete CBR_GESTIONES_USUARIOS where COD_GESTION = '" & scSeguridad.Item(1).Tag _
         & "' and USUARIO = '" & Item.Text & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub
    
vError:
      MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case Item.Index
    Case 0
    Case 1
        Call sbGestiones_Load
    End Select

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer, strSQL As String

On Error GoTo vError

Select Case UCase(Button.Key)
  Case "NUEVO"
    vGrid.MaxRows = vGrid.MaxRows + 1
'    Call sbCargaCboTipos(7, vGrid.MaxRows, vGrid)

  Case "BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = 6 Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        strSQL = "delete cbr_gestiones where cod_gestion = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.col = 2
        Call Bitacora("Elimina", "Tipo Gestión Cobros Id:" & strSQL & " - " & vGrid.Text)
        
        vGrid.col = 1
        strSQL = "select cod_gestion,descripcion,codigo_referencia,monto,MODIFICA_USUARIO,MODIFICA_DESVIACION,isnull(COD_CUENTA,'') as COD_CUENTA" _
              & ", case when NIVEL_GESTION = 'U' then 'Usuario' else 'Sistema' end as NIVEL_GESTION,ACCESO_RESTRINGIDO" _
              & ",MRECUPERACION,IVA_PORCENTAJE, ESTADO" _
              & " from cbr_gestiones" _
              & " order by cod_gestion"
        Call sbCargaGridLocal(12, strSQL)

     End If
  
  Case "REPORTES"
     'Call sbReportesInv("TiposPrecios", "Tipos de Precios", "Listado", "")

  Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Formato de Cuenta Contable
If vGrid.ActiveCol = 7 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = fxgCntCuentaFormato(True, vGrid.Text)
  vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
  vGrid.CellNote = fxgCntCuentaDesc(fxgCntCuentaFormato(False, vGrid.Text))
  vGrid.TextTip = TextTipFixed
End If


'Activa Busquedas
If KeyCode = vbKeyF4 Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = vGrid.ActiveCol
  If vGrid.col = 7 Then
        Call sbgCntCuentaConsulta
        vGrid.Text = gBusquedas.Resultado
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
'    Call sbCargaCboTipos(7, vGrid.MaxRows, vGrid)
End If




End Sub


Private Sub sbGestiones_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

scSeguridad.Item(0).Caption = "Seleccione una Gestión"

scSeguridad.Item(1).Caption = "Gestion:"
scSeguridad.Item(1).Tag = ""

With lswGestiones
    .ListItems.Clear
     
    strSQL = "select COD_GESTION,DESCRIPCION" _
           & " from CBR_GESTIONES where ACCESO_RESTRINGIDO = 1 and NIVEL_GESTION = 'U'" _
           
    vPaso = True
    
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
     Set itmX = .ListItems.Add(, , rs!COD_GESTION)
         itmX.SubItems(1) = rs!Descripcion
     rs.MoveNext
    Loop
    rs.Close
    
    vPaso = False

End With
    
End Sub

Private Sub sbCargaGridLocal(vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
     Case 1
        vGrid.Text = CStr(rs!COD_GESTION)
     Case 2
        vGrid.Text = CStr(rs!Descripcion)
     Case 3
        vGrid.Text = CStr(rs!codigo_referencia)
     Case 4
        vGrid.Text = CStr(rs!Monto)
     Case 5
        vGrid.Text = rs!MODIFICA_USUARIO
     Case 6
        vGrid.Text = CStr(rs!MODIFICA_DESVIACION)
     Case 7
        If rs!cod_cuenta <> Empty Then
            vGrid.Text = CStr(fxgCntCuentaFormato(True, rs!cod_cuenta))
        End If
     Case 8
        vGrid.Text = rs!NIVEL_GESTION
     Case 9
        vGrid.Value = rs!ACCESO_RESTRINGIDO
     
     Case 10
        vGrid.Value = rs!MRECUPERACION
     
     Case 11
        vGrid.Text = CStr(rs!IVA_PORCENTAJE)
     
     Case 12
        vGrid.Value = rs!Estado

    End Select
  
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext
Loop

rs.Close

Me.MousePointer = vbDefault

End Sub
