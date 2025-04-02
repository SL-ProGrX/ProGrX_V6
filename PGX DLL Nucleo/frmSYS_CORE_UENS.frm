VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmSYS_CORE_UENS 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Definición de Estructura Organizacional"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   11055
      _Version        =   1441793
      _ExtentX        =   19500
      _ExtentY        =   11245
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
      ItemCount       =   3
      SelectedItem    =   2
      Item(0).Caption =   "UENs"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Miembros"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "lswMiembros"
      Item(1).Control(1)=   "txtFiltro"
      Item(2).Caption =   "Roles"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "vgRoles"
      Item(2).Control(1)=   "txtRolesFiltro"
      Begin XtremeSuiteControls.ListView lswMiembros 
         Height          =   5655
         Left            =   -70000
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   11055
         _Version        =   1441793
         _ExtentX        =   19500
         _ExtentY        =   9975
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5895
         Left            =   -69880
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   10815
         _Version        =   524288
         _ExtentX        =   19076
         _ExtentY        =   10398
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
         SpreadDesigner  =   "frmSYS_CORE_UENS.frx":0000
         VScrollSpecialType=   2
         Appearance      =   1
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   330
         Left            =   -70000
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   11055
         _Version        =   1441793
         _ExtentX        =   19500
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin FPSpreadADO.fpSpread vgRoles 
         Height          =   5535
         Left            =   0
         TabIndex        =   6
         Top             =   720
         Width           =   10935
         _Version        =   524288
         _ExtentX        =   19288
         _ExtentY        =   9763
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
         SpreadDesigner  =   "frmSYS_CORE_UENS.frx":0680
         VScrollSpecialType=   2
         Appearance      =   1
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtRolesFiltro 
         Height          =   330
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   11055
         _Version        =   1441793
         _ExtentX        =   19500
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   11055
      _Version        =   1441793
      _ExtentX        =   19500
      _ExtentY        =   661
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CORE: Unidades Estratégicas de Negocios"
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
      Height          =   492
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "frmSYS_CORE_UENS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub sbMiembros_List()

If vPaso Then Exit Sub
If scTitulo.Tag = "" Then Exit Sub

Me.MousePointer = vbHourglass

vPaso = True

txtFiltro.Text = fxSysCleanTxtInject(txtFiltro.Text)

With lswMiembros
 .ListItems.Clear
  
strSQL = "exec spSys_UENS_Miembros_Consultas '" & scTitulo.Tag & "', '" & txtFiltro.Text & "'"
 Call OpenRecordSet(rs, strSQL)
 
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!CORE_USUARIO)
      itmX.SubItems(1) = rs!Nombre
      itmX.SubItems(2) = rs!Usuario_Ref & ""
      If rs!ASIGNADO = 1 Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
 Loop
 rs.Close

End With


vPaso = False


Me.MousePointer = vbDefault

End Sub


Private Sub sbRoles_List()

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

txtRolesFiltro.Text = fxSysCleanTxtInject(txtRolesFiltro.Text)
 
strSQL = "exec spSys_UENS_Roles_Consultas '" & scTitulo.Tag & "', '" & txtRolesFiltro.Text & "'"
Call OpenRecordSet(rs, strSQL)

With vgRoles
    .MaxRows = 0
    
    Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     
     .Col = 1
     .Text = rs!CORE_USUARIO
     .Col = 2
     .Text = rs!Nombre
     .Col = 3
     .Value = rs!ROL_SOLICITA
     .Col = 4
     .Value = rs!ROL_CONSULTA
     .Col = 5
     .Value = rs!ROL_AUTORIZA
     .Col = 6
     .Value = rs!ROL_ENCARGADO
     
     rs.MoveNext
    Loop
    rs.Close

End With

Me.MousePointer = vbDefault

vPaso = False


Exit Sub

vError:
    Me.MousePointer = vbDefault
    vPaso = False
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Activate()
vModulo = 10
End Sub

Private Sub Form_Load()

vModulo = 10

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vGrid.AppearanceStyle = fxGridStyle

With lswMiembros.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2100
    .Add , , "Nombre", 3200
    .Add , , "Usuario Ref", 2800, vbCenter
End With

tcMain.Item(0).Selected = True
Call sbUENS_Load

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbUENS_Load()

strSQL = "select COD_UNIDAD, descripcion,  CntX_Unidad, CntX_Centro_Costo, Activa, 0 as 'btn' from CORE_UENS" _
      & " order by COD_UNIDAD"
Call sbCargaGrid_Local(vGrid, 6, strSQL)

End Sub

Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1
If Trim(vGrid.Text) = "" Then
  MsgBox "Indique un Código Válido!", vbExclamation
  Exit Function
End If

strSQL = "select isnull(count(*),0) as Existe from CORE_UENS " _
       & " where COD_UNIDAD = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  
  strSQL = "insert into CORE_UENS(COD_UNIDAD, descripcion, CntX_Unidad, CntX_Centro_Costo, Activa, Registro_Fecha, Registro_Usuario) values('" _
         & vGrid.Text & "', '"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "', '"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "', '"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Text & "', "
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ", dbo.MyGetdate(), '" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Unidad Estratégica de Negocios Id: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update CORE_UENS set descripcion = '" & vGrid.Text & "', CntX_Unidad = '"
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Text & "', CntX_Centro_Costo = '"
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Text & "', Activa = "
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Value & ", Modifica_Fecha = dbo.MyGetdate(), Modifica_Usuario = '" _
        & glogon.Usuario & "' where COD_UNIDAD = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Unidad Estratégica de Negocios Id: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub lswMiembros_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spSys_UENS_Miembros_Registro '" & scTitulo.Tag & "', '" & Item.Text _
       & "', '" & glogon.Usuario & "', '" & IIf((Item.Checked), "A", "E") & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Pass = 1 Then
Else
    MsgBox rs!Mensaje, vbExclamation
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Index > 0 And scTitulo.Tag = "" Then
    MsgBox "Consulte una UENs primero!", vbInformation
    tcMain(0).Selected = True
    Exit Sub
End If

Select Case Item.Index
    Case 1
        Call sbMiembros_List
    Case 2
        Call sbRoles_List
End Select
End Sub

Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbMiembros_List
End If
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

vGrid.Row = Row
If Col = 6 Then
    vGrid.Col = 1
    scTitulo.Tag = vGrid.Text
    vGrid.Col = 2
    scTitulo.Caption = vGrid.Text
    tcMain.Item(1).Selected = True
End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols - 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

If KeyCode = vbKeyF4 And vGrid.ActiveCol = 3 Then
   gBusquedas.Columna = "cod_unidad"
   gBusquedas.Consulta = "select cod_unidad as unidad, descripcion from CntX_Unidades"
   gBusquedas.Orden = "cod_unidad"
   gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = vGrid.ActiveCol
   vGrid.Text = gBusquedas.Resultado
End If

If KeyCode = vbKeyF4 And vGrid.ActiveCol = 4 Then
   gBusquedas.Columna = "cod_centro_costo"
   gBusquedas.Consulta = "select cod_centro_costo as CentroCosto, descripcion from CNTX_CENTRO_COSTOS"
   gBusquedas.Orden = "cod_centro_costo"
   gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = vGrid.ActiveCol
   vGrid.Text = gBusquedas.Resultado
End If


'Elimina
If KeyCode = vbKeyDelete Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
     i = MsgBox("Está Seguro que desea borrar este registro?", vbYesNo)
     If i = vbYes Then
        strSQL = "delete CORE_UENS where COD_UNIDAD = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Unidad Estratégica de Negocios Id: " & vGrid.Text)
        
        Call sbUENS_Load
     End If

End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


End Sub



Private Sub sbCargaGrid_Local(vGrid As Object, vGridMaxCol As Integer, strSQL As String, Optional vBorra As Boolean = True)
Dim rs As New ADODB.Recordset, i As Integer

On Error GoTo vErrorLoad

If vBorra Then
    vGrid.MaxCols = vGridMaxCol
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
    For i = 1 To vGrid.MaxCols
     vGrid.Col = i
     vGrid.Text = ""
    Next i
End If

Call OpenRecordSet(rs, strSQL, 0)
  
vGrid.MaxRows = 1
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
  
    vGrid.Col = i
    If i < 6 Then
        Select Case vGrid.CellType
            Case CellTypeDate
                vGrid.Text = Format(IIf(IsNull(rs.Fields(i - 1).Value), "", rs.Fields(i - 1)), "dd/mm/yyyy")
            Case Else
                vGrid.Text = CStr(IIf(IsNull(rs.Fields(i - 1).Value), "", rs.Fields(i - 1)))
        End Select
    End If

  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

Exit Sub

vErrorLoad:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub


Private Sub vgRoles_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If vPaso Then Exit Sub

On Error GoTo vError

Dim pUser As String
Dim pRSolicita As Integer, pRConsulta As Integer, pRAutoriza As Integer, pREncargado As Integer

With vgRoles
    .Row = Row
    .Col = 1
    pUser = .Text
    .Col = 3
    pRSolicita = .Value
    .Col = 4
    pRConsulta = .Value
    .Col = 5
    pRAutoriza = .Value
    .Col = 6
    pREncargado = .Value
    
    strSQL = "exec spSys_UENS_Roles_Registro '" & scTitulo.Tag & "', '" & pUser & "', " & pRSolicita _
           & ", " & pRConsulta & ", " & pRAutoriza & ", " & pREncargado & ", '" & glogon.Usuario & "'"
    
    Call ConectionExecute(strSQL)
    
End With

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
