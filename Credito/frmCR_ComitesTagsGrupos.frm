VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCR_GruposTrabajo 
   Caption         =   "Definición de Grupos de Trabajo & Asignación de Tags/Comités"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9465
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab ssTab 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Grupos"
      TabPicture(0)   =   "frmCR_ComitesTagsGrupos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "vGrid"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Miembros"
      TabPicture(1)   =   "frmCR_ComitesTagsGrupos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lswMiembros"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cboMiembros"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Tags"
      TabPicture(2)   =   "frmCR_ComitesTagsGrupos.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2(4)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lswGrupos"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cboTags"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Comites"
      TabPicture(3)   =   "frmCR_ComitesTagsGrupos.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label2(5)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label2(6)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lswComites"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cboComites"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin VB.ComboBox cboComites 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72480
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   480
         Width           =   6375
      End
      Begin VB.ComboBox cboTags 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72480
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   420
         Width           =   6375
      End
      Begin VB.ComboBox cboMiembros 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   6375
      End
      Begin MSComctlLib.ListView lswMiembros 
         Height          =   4095
         Left            =   -72240
         TabIndex        =   1
         Top             =   780
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   7223
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
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Usuario"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6068
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4335
         Left            =   2520
         TabIndex        =   3
         Top             =   540
         Width           =   6615
         _Version        =   524288
         _ExtentX        =   11668
         _ExtentY        =   7646
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   496
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_ComitesTagsGrupos.frx":0070
         VScrollSpecialType=   2
         AppearanceStyle =   0
      End
      Begin MSComctlLib.ListView lswGrupos 
         Height          =   3975
         Left            =   -72480
         TabIndex        =   8
         Top             =   900
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   7011
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
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Usuario"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6068
         EndProperty
      End
      Begin MSComctlLib.ListView lswComites 
         Height          =   3975
         Left            =   -72480
         TabIndex        =   13
         Top             =   960
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   7011
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
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Usuario"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grupos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   6
         Left            =   -73680
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   5
         Left            =   -73680
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tag"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   4
         Left            =   -73680
         TabIndex        =   11
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grupos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   -73680
         TabIndex        =   10
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grupo de Usuarios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Index           =   1
         Left            =   1440
         TabIndex        =   6
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grupo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   -73440
         TabIndex        =   5
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Miembros"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   -73440
         TabIndex        =   4
         Top             =   780
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   0
      Left            =   120
      Picture         =   "frmCR_ComitesTagsGrupos.frx":0535
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label4 
      Caption         =   "Grupos de Trabajo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "frmCR_GruposTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mModoSif As Boolean

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

    Me.MousePointer = vbHourglass
    
    SSTab.Tab = 0
    
    strSQL = "select cod_grupo,descripcion from crd_grupos order by cod_grupo"
    Call sbCargaGrid(vGrid, 2, strSQL)
    
    
    Me.MousePointer = vbDefault

End Sub


Private Function fxIndiceCodigo(xkey As String) As String
xkey = Mid(xkey, 4, Len(xkey))
xkey = Mid(xkey, 1, Len(xkey) - 1)
fxIndiceCodigo = xkey
End Function



Private Sub cboComites_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If vPaso Then Exit Sub
If cboComites.ListCount <= 0 Then Exit Sub

With lswComites
 .ListItems.Clear
  
 strSQL = "select G.cod_grupo,G.descripcion,T.cod_grupo as asignado" _
        & " from CRD_GRUPOS G left join CRD_COMITES_GRUPOS T on G.cod_grupo = T.cod_grupo" _
        & " and T.id_comite = '" & fxCodigoCbo(cboComites) & "'" _
        & " order by G.descripcion"
        
 rs.Open strSQL, glogon.Conection, adOpenForwardOnly
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!cod_grupo)
      itmX.SubItems(1) = rs!Descripcion
      If Not IsNull(rs!asignado) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
 Loop
 rs.Close
End With
End Sub

Private Sub cboMiembros_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If vPaso Then Exit Sub
If cboMiembros.ListCount <= 0 Then Exit Sub

With lswMiembros
 .ListItems.Clear
  
 strSQL = "select U.nombre,U.descripcion,A.usuario" _
        & " from Usuarios U left join crd_grpusers A on U.nombre = A.usuario" _
        & " and U.estado = 'A'  and A.cod_grupo = '" & fxCodigoCbo(cboMiembros) & "'" _
        & " order by A.usuario desc,U.nombre asc"
 rs.Open strSQL, glogon.Conection, adOpenForwardOnly
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
End With

End Sub



Private Sub cboTags_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If vPaso Then Exit Sub
If cboTags.ListCount <= 0 Then Exit Sub

With lswGrupos
 .ListItems.Clear
  
 strSQL = "select G.cod_grupo,G.descripcion,T.cod_grupo as asignado" _
        & " from CRD_GRUPOS G left join CRD_TAGS_GRUPOS T on G.cod_grupo = T.cod_grupo" _
        & " and T.tag_codigo = '" & fxCodigoCbo(cboTags) & "'" _
        & " order by G.descripcion"
        
 rs.Open strSQL, glogon.Conection, adOpenForwardOnly
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!cod_grupo)
      itmX.SubItems(1) = rs!Descripcion
      If Not IsNull(rs!asignado) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
 Loop
 rs.Close
End With
End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()


vGrid.AppearanceStyle = vGrid.AppearanceStyle

Call sbInicializa

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub





Private Sub lswComites_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Item.Checked Then
  strSQL = "insert crd_comites_grupos(id_comite,cod_grupo) values('" & fxCodigoCbo(cboComites) _
         & "','" & Item.Text & "')"
Else
  strSQL = "delete crd_comites_grupos where id_comite = '" & fxCodigoCbo(cboComites) _
         & "' and cod_grupo = '" & Item.Text & "'"
End If
glogon.Conection.Execute strSQL

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub lswGrupos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


'If Item.Checked Then
'  'Preguntar si ya Existe el Usuario en Otro Grupo. / de ser asi no continuar
'  strSQL = "select coalesce(count(*),0) as Existe from crd_tags_grupos where tag_codigo <> '" _
'         & fxCodigoCbo(cboTags) & "' and cod_grupo = '" & Item.Text & "'"
'  rs.Open strSQL, glogon.Conection, adOpenStatic
'  If rs!existe > 0 Then
'     rs.Close
'     Item.Checked = False
'     MsgBox "El Grupo ya ha sido asignado a un tag, proceda a excluirlo primero del otro grupo antes de agregarlo a este", vbExclamation
'     Exit Sub
'  End If
'  rs.Close
'End If


If Item.Checked Then
  strSQL = "insert crd_tags_grupos(tag_codigo,cod_grupo) values('" & fxCodigoCbo(cboTags) _
         & "','" & Item.Text & "')"
Else
  strSQL = "delete crd_tags_grupos where tag_codigo = '" & fxCodigoCbo(cboTags) _
         & "' and cod_grupo = '" & Item.Text & "'"
End If
glogon.Conection.Execute strSQL

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub lswMiembros_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If Item.Checked Then
  'Preguntar si ya Existe el Usuario en Otro Grupo. / de ser asi no continuar
  strSQL = "select coalesce(count(*),0) as Existe from crd_grpUsers where cod_grupo <> '" _
         & fxCodigoCbo(cboMiembros) & "' and usuario = '" & Item.Text & "'"
  rs.Open strSQL, glogon.Conection, adOpenStatic
  If rs!existe > 0 Then
     rs.Close
     Item.Checked = False
     MsgBox "El Usuario ya ha sido asignado a otro grupo, proceda a excluirlo primero del otro grupo antes de agregarlo a este", vbExclamation
     Exit Sub
  End If
  rs.Close
End If


If Item.Checked Then
  strSQL = "insert crd_grpusers(cod_grupo,usuario) values('" & fxCodigoCbo(cboMiembros) _
         & "','" & Item.Text & "')"
Else
  strSQL = "delete crd_grpusers where cod_grupo = '" & fxCodigoCbo(cboMiembros) _
         & "' and usuario = '" & Item.Text & "'"
End If
glogon.Conection.Execute strSQL

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
  
End Sub




Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select coalesce(count(*),0) as Existe from crd_Grupos" _
       & " where cod_grupo = '" & vGrid.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

If rs!existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into crd_Grupos(cod_grupo,descripcion) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & UCase(vGrid.Text) & "')"

  glogon.Conection.Execute strSQL

  vGrid.Col = 1
  Call Bitacora("Registra", "Grupo de Usuarios: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update crd_Grupos set descripcion = '" & vGrid.Text & "'"
 strSQL = strSQL & " where cod_grupo = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 glogon.Conection.Execute strSQL

 Call Bitacora("Modifica", "Grupo de Usuarios : " & vGrid.Text)


End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical

End Function

Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String

Select Case SSTab.Tab
  Case 0 'Grupos

  Case 1 'Miembros
    vPaso = True
    strSQL = "select cod_grupo + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  crd_grupos"
    Call sbLlenaCbo(cboMiembros, strSQL, False)
    vPaso = False
    
    Call cboMiembros_Click
    
  Case 2 'Tags
  
      vPaso = True
    strSQL = "select TAG_CODIGO + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  crd_tags"
    Call sbLlenaCbo(cboTags, strSQL, False)
    vPaso = False
    
    Call cboTags_Click
    
  Case 3 'Comites
  
      vPaso = True
    strSQL = "select cast(ID_COMITE as varchar) + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  comites"
    Call sbLlenaCbo(cboComites, strSQL, False)
    vPaso = False
    
    Call cboComites_Click
    
End Select
End Sub



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
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

Private Function fxGuardarGrpAccss() As Long
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

fxGuardarGrpAccss = 0
vGridGrpAccss.Row = vGridGrpAccss.ActiveRow
vGridGrpAccss.Col = 1

If vGridGrpAccss.Text = "" Then 'Insertar
  vGridGrpAccss.Col = 2
  strSQL = "insert into crd_reportes_grp(descripcion,activo) values('" _
         & UCase(vGridGrpAccss.Text) & "',"
  vGridGrpAccss.Col = 3
  strSQL = strSQL & vGridGrpAccss.Value & ")"
  
  glogon.Conection.Execute strSQL

  vGridGrpAccss.Col = 1
  
  strSQL = "select coalesce(max(cod_grupo),0) as Ultimo from crd_reportes_grp"
  rs.Open strSQL, glogon.Conection, adOpenStatic
   vGridGrpAccss.Text = CStr(rs!Ultimo)
  rs.Close
  
  Call Bitacora("Registra", "Reportes > Grupo de Acceso: " & vGridGrpAccss.Text)

Else 'Actualizar

 vGridGrpAccss.Col = 2
 strSQL = "update crd_reportes_grp set descripcion = '" & UCase(vGridGrpAccss.Text) & "',activo = "
 vGridGrpAccss.Col = 3
 strSQL = strSQL & vGridGrpAccss.Value & " where cod_grupo = "
 vGridGrpAccss.Col = 1
 strSQL = strSQL & vGridGrpAccss.Text
 
 glogon.Conection.Execute strSQL

 Call Bitacora("Modifica", "Reportes > Grupo de Acceso: " & vGridGrpAccss.Text)


End If

fxGuardarGrpAccss = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical
 

End Function

