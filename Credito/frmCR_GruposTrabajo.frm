VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCR_GruposTrabajo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Definición de Grupos Trabajos & Asignación Funcional"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10785
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6012
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   10572
      _Version        =   1572864
      _ExtentX        =   18648
      _ExtentY        =   10604
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
      ItemCount       =   4
      Item(0).Caption =   "Grupos"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "Label1(0)"
      Item(1).Caption =   "Miembros"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "Label1(1)"
      Item(1).Control(1)=   "Label1(2)"
      Item(1).Control(2)=   "cboMiembros"
      Item(1).Control(3)=   "lswMiembros"
      Item(2).Caption =   "Etiquetas"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "Label1(3)"
      Item(2).Control(1)=   "Label1(4)"
      Item(2).Control(2)=   "cboTags"
      Item(2).Control(3)=   "lswTags"
      Item(3).Caption =   "Comités"
      Item(3).ControlCount=   4
      Item(3).Control(0)=   "Label1(5)"
      Item(3).Control(1)=   "Label1(6)"
      Item(3).Control(2)=   "cboComites"
      Item(3).Control(3)=   "lswComites"
      Begin XtremeSuiteControls.ListView lswMiembros 
         Height          =   4692
         Left            =   -67600
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1572864
         _ExtentX        =   11663
         _ExtentY        =   8276
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
      Begin XtremeSuiteControls.ListView lswTags 
         Height          =   4692
         Left            =   -67600
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1572864
         _ExtentX        =   11663
         _ExtentY        =   8276
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
      Begin XtremeSuiteControls.ListView lswComites 
         Height          =   4692
         Left            =   -67600
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1572864
         _ExtentX        =   11663
         _ExtentY        =   8276
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5292
         Left            =   2400
         TabIndex        =   2
         Top             =   600
         Width           =   6612
         _Version        =   524288
         _ExtentX        =   11663
         _ExtentY        =   9334
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
         SpreadDesigner  =   "frmCR_GruposTrabajo.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboMiembros 
         Height          =   312
         Left            =   -67600
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1572864
         _ExtentX        =   11668
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboTags 
         Height          =   312
         Left            =   -67600
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1572864
         _ExtentX        =   11668
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboComites 
         Height          =   312
         Left            =   -67600
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1572864
         _ExtentX        =   11668
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   6
         Left            =   -69400
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Grupo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   732
         Index           =   5
         Left            =   -69400
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Comités de Resolución Vinculados"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   4
         Left            =   -69400
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Grupo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   732
         Index           =   3
         Left            =   -69400
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Etiquetas (Tag's) Vinculados"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   2
         Left            =   -69400
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Grupo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   1
         Left            =   -69400
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Miembros"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   492
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Grupos o Roles de Trabajo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Grupos de Trabajo"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   6612
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10812
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
Dim strSQL As String

Me.MousePointer = vbHourglass

tcMain.Item(0).Selected = True

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
Dim itmX As ListViewItem

If vPaso Then Exit Sub
If cboComites.ListCount <= 0 Then Exit Sub

vPaso = True

With lswComites
 .ListItems.Clear
  
 strSQL = "select C.ID_COMITE,C.DESCRIPCION,CG.ID_COMITE as asignado" _
        & " from COMITES C left join CRD_COMITES_GRUPOS CG on CG.ID_COMITE = C.ID_COMITE" _
        & " and CG.cod_grupo = '" & cboComites.ItemData(cboComites.ListIndex) & "'" _
        & " order by CG.ID_COMITE desc, C.descripcion"
        
 Call OpenRecordSet(rs, strSQL, 0)
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!id_Comite)
      itmX.SubItems(1) = rs!Descripcion
      If Not IsNull(rs!Asignado) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
 Loop
 rs.Close
End With

vPaso = False

End Sub

Private Sub cboMiembros_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If vPaso Then Exit Sub
If cboMiembros.ListCount <= 0 Then Exit Sub

vPaso = True

With lswMiembros
 .ListItems.Clear
  
 strSQL = "select U.nombre,U.descripcion,A.usuario" _
        & " from Usuarios U left join crd_grpusers A on U.nombre = A.usuario" _
        & " and A.cod_grupo = '" & cboMiembros.ItemData(cboMiembros.ListIndex) & "'" _
        & " Where U.estado = 'A'" _
        & " order by A.usuario desc,U.nombre asc"
 Call OpenRecordSet(rs, strSQL, 0)
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

vPaso = False

End Sub



Private Sub cboTags_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If vPaso Then Exit Sub
If cboTags.ListCount <= 0 Then Exit Sub

vPaso = True

With lswTags
 .ListItems.Clear
  
 strSQL = "select T.TAG_CODIGO,T.DESCRIPCION,TG.TAG_CODIGO as asignado" _
        & " from CRD_TAGS T left join CRD_TAGS_GRUPOS TG on TG.TAG_CODIGO = T.TAG_CODIGO" _
        & " and TG.COD_GRUPO = '" & cboTags.ItemData(cboTags.ListIndex) & "'" _
        & " order by TG.TAG_CODIGO desc, T.descripcion"
        
 Call OpenRecordSet(rs, strSQL, 0)
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
End With

vPaso = False

End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()
vModulo = 3

vGrid.AppearanceStyle = vGrid.AppearanceStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lswMiembros.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2000
    .Add , , "Descripción", lswMiembros.Width - (2000 + 250)
End With

With lswTags.ColumnHeaders
    .Clear
    .Add , , "Etiqueta", 2000
    .Add , , "Descripción", lswTags.Width - (2000 + 250)
End With

With lswComites.ColumnHeaders
    .Clear
    .Add , , "Comité Id", 2000
    .Add , , "Descripción", lswComites.Width - (2000 + 250)
End With

Call sbInicializa

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub lswComites_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

Dim strSQL As String

On Error GoTo vError
    
    If Item.Checked Then
      strSQL = "insert crd_comites_grupos(id_comite,cod_grupo) values(" & Item.Text _
             & ",'" & cboComites.ItemData(cboComites.ListIndex) & "')"
    Else
      strSQL = "delete crd_comites_grupos where id_comite = " & Item.Text _
             & " and cod_grupo = '" & cboComites.ItemData(cboComites.ListIndex) & "'"
    End If
    Call ConectionExecute(strSQL)
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub lswTags_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

Dim strSQL As String

On Error GoTo vError

    If Item.Checked Then
      strSQL = "insert crd_tags_grupos(tag_codigo,cod_grupo) values('" & Item.Text _
             & "','" & cboTags.ItemData(cboTags.ListIndex) & "')"
    Else
      strSQL = "delete crd_tags_grupos where tag_codigo = '" & Item.Text _
             & "' and cod_grupo = '" & cboTags.ItemData(cboTags.ListIndex) & "'"
    End If
    Call ConectionExecute(strSQL)
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswMiembros_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset


On Error GoTo vError


If Item.Checked Then
  'Preguntar si ya Existe el Usuario en Otro Grupo. / de ser asi no continuar
  strSQL = "select isnull(count(*),0) as Existe from crd_grpUsers where cod_grupo <> '" _
         & cboMiembros.ItemData(cboMiembros.ListIndex) & "' and usuario = '" & Item.Text & "'"
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe > 0 Then
     rs.Close
     Item.Checked = False
     MsgBox "El Usuario ya ha sido asignado a otro grupo, proceda a excluirlo primero del otro grupo antes de agregarlo a este", vbExclamation
     Exit Sub
  End If
  rs.Close
End If


If Item.Checked Then
  strSQL = "insert crd_grpusers(cod_grupo,usuario) values('" & cboMiembros.ItemData(cboMiembros.ListIndex) _
         & "','" & Item.Text & "')"
Else
  strSQL = "delete crd_grpusers where cod_grupo = '" & cboMiembros.ItemData(cboMiembros.ListIndex) _
         & "' and usuario = '" & Item.Text & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from crd_Grupos" _
       & " where cod_grupo = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into crd_Grupos(cod_grupo,descripcion) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Grupo de Trabajo: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update crd_Grupos set descripcion = '" & vGrid.Text & "'"
 strSQL = strSQL & " where cod_grupo = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 Call Bitacora("Modifica", "Grupo de Trabajo : " & vGrid.Text)


End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String

Select Case Item.Index
  Case 0 'Grupos

  Case 1 'Miembros
    vPaso = True
    strSQL = "select cod_grupo as 'IdX', rtrim(descripcion) as 'ItmX'" _
         & " from  crd_grupos"
    Call sbCbo_Llena_New(cboMiembros, strSQL, False, True)
    vPaso = False
    
    Call cboMiembros_Click
    
  Case 2 'Tags
  
      vPaso = True
    strSQL = "select cod_grupo as 'IdX', rtrim(descripcion) as 'ItmX'" _
         & " from  crd_grupos"
    Call sbCbo_Llena_New(cboTags, strSQL, False, True)
    vPaso = False
    
    Call cboTags_Click
    
  Case 3 'Comites
  
    vPaso = True
    strSQL = "select cod_grupo as 'IdX', rtrim(descripcion) as 'ItmX'" _
         & " from  crd_grupos"
    Call sbCbo_Llena_New(cboComites, strSQL, False, True)
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


