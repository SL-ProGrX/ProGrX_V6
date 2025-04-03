VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmSIF_TagsGrupos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Grupos de Trabajo yAsignación en Tags"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   9735
      _Version        =   1572864
      _ExtentX        =   17171
      _ExtentY        =   10398
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
      Item(0).Caption =   "Grupos"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Miembros"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "lswMiembros"
      Item(1).Control(1)=   "ShortcutCaption1(0)"
      Item(1).Control(2)=   "cboMiembros"
      Item(1).Control(3)=   "txtFiltro"
      Item(2).Caption =   "Tag's"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "lswGruposTags"
      Item(2).Control(1)=   "ShortcutCaption1(1)"
      Item(2).Control(2)=   "cboTags"
      Begin XtremeSuiteControls.ListView lswMiembros 
         Height          =   4695
         Left            =   -68080
         TabIndex        =   2
         Top             =   1140
         Visible         =   0   'False
         Width           =   7095
         _Version        =   1572864
         _ExtentX        =   12515
         _ExtentY        =   8281
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
      End
      Begin XtremeSuiteControls.ListView lswGruposTags 
         Height          =   4935
         Left            =   -68080
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   7100
         _Version        =   1572864
         _ExtentX        =   12524
         _ExtentY        =   8705
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
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5295
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   6735
         _Version        =   524288
         _ExtentX        =   11880
         _ExtentY        =   9340
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
         MaxCols         =   496
         ScrollBars      =   2
         SpreadDesigner  =   "frmSIF_TagsGrupos.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboMiembros 
         Height          =   330
         Left            =   -68080
         TabIndex        =   4
         Top             =   375
         Visible         =   0   'False
         Width           =   7095
         _Version        =   1572864
         _ExtentX        =   12515
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
      Begin XtremeSuiteControls.ComboBox cboTags 
         Height          =   330
         Left            =   -68080
         TabIndex        =   7
         Top             =   375
         Visible         =   0   'False
         Width           =   7095
         _Version        =   1572864
         _ExtentX        =   12515
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   330
         Left            =   -68080
         TabIndex        =   9
         Top             =   760
         Visible         =   0   'False
         Width           =   7095
         _Version        =   1572864
         _ExtentX        =   12515
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   1
         Left            =   -70000
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1572864
         _ExtentX        =   17171
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Grupo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   0
         Left            =   -70000
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1572864
         _ExtentX        =   17171
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Grupo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Grupos de Trabajo y Asignación de etiquetas"
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
      Height          =   480
      Index           =   0
      Left            =   1560
      TabIndex        =   8
      Top             =   360
      Width           =   7335
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmSIF_TagsGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub sbInicializa()
    Me.MousePointer = vbHourglass
    
    tcMain.Item(0).Selected = True
    
    strSQL = "select cod_grupo,descripcion from sif_grupos order by cod_grupo"
    Call sbCargaGrid(vGrid, 2, strSQL)
    
    
    Me.MousePointer = vbDefault

End Sub

Private Function fxIndiceCodigo(xkey As String) As String
    xkey = Mid(xkey, 4, Len(xkey))
    xkey = Mid(xkey, 1, Len(xkey) - 1)
    fxIndiceCodigo = xkey
End Function

Private Sub cboMiembros_Click()

If vPaso Then Exit Sub
If cboMiembros.ListCount <= 0 Then Exit Sub

On Error GoTo vError

txtFiltro.Text = fxSysCleanTxtInject(txtFiltro.Text)

With lswMiembros
 .ListItems.Clear
 
 strSQL = "select U.nombre,U.descripcion,A.usuario" _
        & " from Usuarios U left join sif_grpusers A on U.nombre = A.usuario" _
        & " and U.estado = 'A'  and A.cod_grupo = '" & cboMiembros.ItemData(cboMiembros.ListIndex) _
        & "' and ( U.Nombre like '%" & txtFiltro.Text & "%' or U.Descripcion like '%" & txtFiltro.Text & "%' )" _
        & " order by A.usuario desc,U.nombre asc"
 Call OpenRecordSet(rs, strSQL)
 
 vPaso = True
 
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!Nombre)
      itmX.SubItems(1) = rs!Descripcion
      If Not IsNull(rs!Usuario) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
 Loop
 rs.Close
 
 vPaso = False
 
End With

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub



Private Sub cboTags_Click()

If vPaso Then Exit Sub
If cboTags.ListCount <= 0 Then Exit Sub

On Error GoTo vError

With lswGruposTags
 .ListItems.Clear
  
 strSQL = "select T.TAG_CODIGO,T.DESCRIPCION,TG.TAG_CODIGO as asignado" _
        & " from sif_TAGS T left join sif_TAGS_GRUPOS TG on TG.TAG_CODIGO = T.TAG_CODIGO" _
        & " and TG.COD_GRUPO = '" & cboTags.ItemData(cboTags.ListIndex) & "'" _
        & " order  by TG.TAG_CODIGO desc,T.DESCRIPCION asc "
        
 Call OpenRecordSet(rs, strSQL)
 
 vPaso = True
 
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!TAG_CODIGO)
      itmX.SubItems(1) = rs!Descripcion
      If Not IsNull(rs!Asignado) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
 Loop
 rs.Close

 vPaso = False

End With

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 8
End Sub

Private Sub Form_Load()
vModulo = 8

tcMain.Item(0).Selected = True

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lswMiembros.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2500
    .Add , , "Descripción", 4500
End With

With lswGruposTags.ColumnHeaders
    .Clear
    .Add , , "Tag", 2500
    .Add , , "Descripción", 4500
End With


Call sbInicializa

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub lswGruposTagsx_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

    If Item.Checked Then
      strSQL = "insert sif_tags_grupos(tag_codigo,cod_grupo) values('" & Item.Text _
             & "','" & cboTags.ItemData(cboTags.ListIndex) & "')"
    Else
      strSQL = "delete sif_tags_grupos where tag_codigo = '" & Item.Text _
             & "' and cod_grupo = '" & cboTags.ItemData(cboTags.ListIndex) & "'"
    End If
    Call ConectionExecute(strSQL)
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswMiembros_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswMiembros.SortKey = ColumnHeader.Index - 1
  If lswMiembros.SortOrder = 0 Then lswMiembros.SortOrder = 1 Else lswMiembros.SortOrder = 0
  lswMiembros.Sorted = True
End Sub

Private Sub lswGruposTags_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswGruposTags.SortKey = ColumnHeader.Index - 1
  If lswGruposTags.SortOrder = 0 Then lswGruposTags.SortOrder = 1 Else lswGruposTags.SortOrder = 0
  lswGruposTags.Sorted = True
End Sub

Private Sub lswMiembros_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError


If Item.Checked Then
  'Preguntar si ya Existe el Usuario en Otro Grupo. / de ser asi no continuar
  strSQL = "select isnull(count(*),0) as Existe from sif_grpUsers where cod_grupo <> '" _
         & cboMiembros.ItemData(cboMiembros.ListIndex) & "' and usuario = '" & Item.Text & "'"
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe > 0 Then
     vPaso = True
         Item.Checked = False
     vPaso = False
     MsgBox "El Usuario ya ha sido asignado a otro grupo, proceda a excluirlo primero del otro grupo antes de agregarlo a este", vbExclamation
     Exit Sub
  End If
End If


If Item.Checked Then
  strSQL = "insert SIf_grpusers(cod_grupo,usuario) values('" & cboMiembros.ItemData(cboMiembros.ListIndex) _
         & "','" & Item.Text & "')"
Else
  strSQL = "delete sif_grpusers where cod_grupo = '" & cboMiembros.ItemData(cboMiembros.ListIndex) _
         & "' and usuario = '" & Item.Text & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub




Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from SIF_Grupos" _
       & " where cod_grupo = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into SIF_Grupos(cod_grupo,descripcion) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Grupo de Usuarios: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update SIF_Grupos set descripcion = '" & vGrid.Text & "'"
 strSQL = strSQL & " where cod_grupo = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 Call Bitacora("Modifica", "Grupo de Usuarios : " & vGrid.Text)


End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
  Case 0 'Grupos

  Case 1 'Miembros
    strSQL = "select cod_grupo as 'IdX', rtrim(descripcion) as 'ItmX'" _
         & " from  SIF_grupos order by descripcion"
    vPaso = True
        Call sbCbo_Llena_New(cboMiembros, strSQL, False, True)
    vPaso = False
    
    Call cboMiembros_Click
    
  Case 2 'Tags
  
    strSQL = "select cod_grupo as 'IdX' , rtrim(descripcion) as 'ItmX'" _
         & " from  SIF_grupos order by descripcion"
    vPaso = True
      Call sbCbo_Llena_New(cboTags, strSQL, False, True)
    vPaso = False
    
    Call cboTags_Click
    
End Select
End Sub

Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    Call cboMiembros_Click
End If
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

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

End Sub

