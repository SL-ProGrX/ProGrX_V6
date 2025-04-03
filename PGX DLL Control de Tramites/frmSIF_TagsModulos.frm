VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmSIF_TagsModulos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Procesos y Etiquetas (Control de Tramites)"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7695
      _Version        =   1441793
      _ExtentX        =   13573
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
      ItemCount       =   2
      Item(0).Caption =   "Procesos"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Etiquetas por Proceso"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "lsw"
      Item(1).Control(1)=   "ShortcutCaption1"
      Item(1).Control(2)=   "cboModulos"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4935
         Left            =   -68080
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   5460
         _Version        =   1441793
         _ExtentX        =   9631
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
         Height          =   5175
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   7215
         _Version        =   524288
         _ExtentX        =   12726
         _ExtentY        =   9128
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
         MaxCols         =   2
         ScrollBars      =   2
         SpreadDesigner  =   "frmSIF_TagsModulos.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboModulos 
         Height          =   330
         Left            =   -68080
         TabIndex        =   4
         Top             =   375
         Visible         =   0   'False
         Width           =   5535
         _Version        =   1441793
         _ExtentX        =   9763
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   -70000
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   7695
         _Version        =   1441793
         _ExtentX        =   13573
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Proceso"
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
      Caption         =   "Procesos y Asignación de Etiquetas"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   360
      Width           =   6255
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmSIF_TagsModulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean


Private Sub cboModulos_Click()
If vPaso Then Exit Sub
If cboModulos.ListCount = 0 Then Exit Sub
Call sbModulos_Load
End Sub

Private Sub Form_Activate()
vModulo = 8
End Sub

Private Sub Form_Load()

vModulo = 8

tcMain.Item(0).Selected = True

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Usuario", 1500
    .Add , , "Descripción", 3500
End With

strSQL = "select cod_modulo,descripcion from SIF_MODULOS_TAGS"
Call sbCargaGrid(vGrid, 2, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)
     
End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert SIF_TAGS_MODULOS(tag_codigo,cod_modulo) values('" & Item.Text _
            & "', '" & cboModulos.ItemData(cboModulos.ListIndex) & "')"
Else
   strSQL = "Delete SIF_TAGS_MODULOS where tag_codigo ='" & Item.Text _
          & "' and cod_modulo ='" & cboModulos.ItemData(cboModulos.ListIndex) & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbModulos_Load()

On Error GoTo vError

lsw.ListItems.Clear

strSQL = "select S.TAG_CODIGO,S.descripcion,M.tag_codigo as codigo" _
       & " from SIF_TAGS S left join SIF_TAGS_MODULOS M on S.TAG_CODIGO = M.TAG_CODIGO" _
       & " and M.cod_Modulo = '" & cboModulos.ItemData(cboModulos.ListIndex) & "'" _
       & " where S.ACTIVO = '1'  order by M.tag_codigo desc"
Call OpenRecordSet(rs, strSQL)

vPaso = True
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!TAG_CODIGO)
     itmX.SubItems(1) = rs!Descripcion
 
 If Not IsNull(rs!Codigo) Then
     itmX.Checked = True
     itmX.ForeColor = vbBlue
 End If
 
 rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub
  
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Index = 1 Then
    
    strSQL = "select cod_modulo as 'IdX', rtrim(Descripcion) as 'ItmX' " _
           & " from sif_modulos_tags order by Descripcion"
    vPaso = True
    Call sbCbo_Llena_New(cboModulos, strSQL, False, True)
    vPaso = False
    Call cboModulos_Click
    
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

Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from sif_modulos_tags" _
       & " where cod_MODULO = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into sif_modulos_tags(cod_modulo,descripcion) values('" _
         & vGrid.Text & "', '"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Procesos y Tag's: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update sif_modulos_tags set descripcion = '" & vGrid.Text & "'"
 strSQL = strSQL & " where cod_grupo = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 Call Bitacora("Modifica", "Procesos y Tag's: " & vGrid.Text)


End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function


