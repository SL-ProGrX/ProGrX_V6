VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmAF_BeneficiosGrupos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupos de Beneficios y Ayudas"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   10785
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   10575
      _Version        =   1572864
      _ExtentX        =   18653
      _ExtentY        =   11668
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
      Item(0).Caption =   "Categorías"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Asignación"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "cboGrupos"
      Item(1).Control(1)=   "Label1(1)"
      Item(1).Control(2)=   "lsw"
      Item(1).Control(3)=   "btnAsigna(0)"
      Item(1).Control(4)=   "btnAsigna(1)"
      Item(1).Control(5)=   "btnAsigna(2)"
      Item(1).Control(6)=   "btnAsigna(3)"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5415
         Left            =   -68200
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   8655
         _Version        =   1572864
         _ExtentX        =   15266
         _ExtentY        =   9551
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
      Begin XtremeSuiteControls.PushButton btnAsigna 
         Height          =   375
         Index           =   0
         Left            =   -69880
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Estados"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Checked         =   -1  'True
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6015
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   10215
         _Version        =   524288
         _ExtentX        =   18018
         _ExtentY        =   10610
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
         MaxCols         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "frmAF_BeneficiosGrupos.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboGrupos 
         Height          =   315
         Left            =   -68200
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   5175
         _Version        =   1572864
         _ExtentX        =   9128
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnAsigna 
         Height          =   375
         Index           =   1
         Left            =   -69880
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Requisitos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnAsigna 
         Height          =   375
         Index           =   2
         Left            =   -69880
         TabIndex        =   8
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Motivos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnAsigna 
         Height          =   375
         Index           =   3
         Left            =   -69880
         TabIndex        =   9
         Top             =   2400
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Accesos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Categorías"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   -70120
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Grupos de Beneficios y Ayudas Sociales"
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
      Height          =   492
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   -120
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmAF_BeneficiosGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim mTipos As String, mUltimoSel As String
Dim vPaso As Boolean


Private Sub btnAsigna_Click(Index As Integer)

On Error GoTo vError

Dim i As Integer

For i = 0 To btnAsigna.Count - 1
    btnAsigna(i).Checked = False
Next i
btnAsigna(Index).Checked = True

Select Case Index
    Case 0 'Estados
        strSQL = "exec spAFI_Bene_Grupos_Estados_List  " & cboGrupos.ItemData(cboGrupos.ListIndex)
    
    Case 1 'Requisitos
        strSQL = "exec spAFI_Bene_Grupos_Requisitos_List  " & cboGrupos.ItemData(cboGrupos.ListIndex)
    
    Case 2 'Motivos
        strSQL = "exec spAFI_Bene_Grupos_Motivos_List  " & cboGrupos.ItemData(cboGrupos.ListIndex)
    
    Case 3 'Accesos
        strSQL = "exec spAFI_Bene_Grupos_Accesos_List  " & cboGrupos.ItemData(cboGrupos.ListIndex)
End Select

Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear

vPaso = True

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Codigo)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = rs!Registro_Fecha & ""
     itmX.SubItems(3) = rs!Registro_Usuario & ""
     itmX.Checked = IIf((rs!Asigna = 1), True, False)
 rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub


vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboGrupos_Click()

If vPaso Then Exit Sub

Call btnAsigna_Click(0)
'Call sbCargaBeneficios

End Sub

Private Sub Form_Activate()
vModulo = 7
End Sub

Private Sub Form_Load()

vModulo = 7

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Código", 2100, vbCenter
    .Add , , "Descripción", 3500
    .Add , , "R. Fecha", 2500
    .Add , , "R. Usuario", 2500, vbCenter
End With

tcMain.Item(0).Selected = True


mTipos = ""
mUltimoSel = ""

strSQL = "select rtrim(COD_CATEGORIA) as 'Tipo' From AFI_BENE_CATEGORIAS" _
       & " Where Activo = 1 ORDER BY COD_CATEGORIA"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
   If mTipos = "" Then
    mUltimoSel = rs!Tipo
   End If
   
   mTipos = mTipos & Chr$(9) & rs!Tipo

   rs.MoveNext
Loop
rs.Close


strSQL = "select cod_grupo, Descripcion, Cod_Categoria, Monto, Estado, User_Registra, Fecha from AFI_BENE_GRUPOS"
Call sbCargaGridLocal(vGrid, 5, strSQL)
    

Call Formularios(Me)
Call RefrescaTags(Me)
End Sub


Private Sub sbCargaGridLocal(ByRef vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

Call OpenRecordSet(rs, strSQL)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  vGrid.Col = 3
  vGrid.CellType = CellTypeComboBox
  
  vGrid.TypeComboBoxList = mTipos
  vGrid.TypeComboBoxEditable = False
  
  vGrid.Text = mUltimoSel
  
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
     Case 1
        vGrid.Text = rs!Cod_Grupo
     Case 2
        vGrid.Text = rs!Descripcion
     Case 3
        vGrid.Text = rs!Cod_Categoria
     Case 4
        vGrid.Text = Format(rs!MONTO, "Standard")
     Case 5
        vGrid.Value = rs!Estado
    End Select
  Next i
  
    vGrid.Col = 5
    vGrid.TextTip = TextTipFixed
    vGrid.TextTipDelay = 1000
    vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
    vGrid.CellNote = "Usuario " & IIf(IsNull(rs!User_Registra), "...!", rs!User_Registra) _
                     & vbCrLf & " Fecha " & IIf(IsNull(rs!fecha), "...!", rs!fecha)
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

  vGrid.Row = vGrid.MaxRows
  vGrid.Col = 3
  vGrid.CellType = CellTypeComboBox
  
  vGrid.TypeComboBoxList = mTipos
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = mUltimoSel

Me.MousePointer = vbDefault

End Sub

Private Sub sbCboCategorias(vCol As Integer, vRow As Long, vGrid As Object)

vGrid.Col = vCol
vGrid.Row = vRow
vGrid.CellType = CellTypeComboBox

vGrid.TypeComboBoxList = mTipos
vGrid.TypeComboBoxEditable = False
vGrid.Text = mUltimoSel

End Sub



Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

Select Case True
    Case btnAsigna(0).Checked 'Estados
        strSQL = "exec spAFI_Bene_Grupos_Estados_Add " & cboGrupos.ItemData(cboGrupos.ListIndex)
    
    Case btnAsigna(1).Checked 'Requisitos
        strSQL = "exec spAFI_Bene_Grupos_Requisitos_Add " & cboGrupos.ItemData(cboGrupos.ListIndex)
        
    Case btnAsigna(2).Checked 'Motivos
        strSQL = "exec spAFI_Bene_Grupos_Motivos_Add " & cboGrupos.ItemData(cboGrupos.ListIndex)
    
    Case btnAsigna(3).Checked 'Accesos
        strSQL = "exec spAFI_Bene_Grupos_Accesos_Add " & cboGrupos.ItemData(cboGrupos.ListIndex)

End Select

strSQL = strSQL & ",'" & Item.Text & "', '" & glogon.Usuario & "', '" & IIf((Item.Checked), "A", "E") & "'"
Call ConectionExecute(strSQL)


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String

If Item.Index = 1 Then

   vPaso = True
   strSQL = "select cod_grupo as 'IdX',rtrim(descripcion) as 'ItmX'" _
          & " from afi_bene_grupos where estado = 1 "
   
   Call sbCbo_Llena_New(cboGrupos, strSQL, False, True)
   
   vPaso = False
   
   Call cboGrupos_Click

End If

End Sub

Private Function fxGuardar() As Long
Dim pCodigo As Long
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1


If vGrid.Text = "" Then
    vGrid.Col = 1
      strSQL = "select isnull(max(cod_grupo),0) + 1 as consec" _
             & " from afi_bene_grupos"
      Call OpenRecordSet(rs, strSQL)
        pCodigo = rs!consec
      rs.Close
      vGrid.Text = CStr(pCodigo)
      
      vGrid.Col = 2
      strSQL = "insert afi_bene_grupos(cod_grupo, descripcion, cod_categoria, monto, estado, fecha, user_registra)" _
          & " values(" & pCodigo & ",'" & vGrid.Text & "', '"
      vGrid.Col = 3
      strSQL = strSQL & vGrid.Text & "', "
      vGrid.Col = 4
      strSQL = strSQL & CCur(vGrid.Text) & ","
      vGrid.Col = 5
      strSQL = strSQL & vGrid.Value & ",dbo.MyGetdate(),'" & glogon.Usuario & "')"
      
      Call ConectionExecute(strSQL)
    
      vGrid.Col = 1
      Call Bitacora("Registra", "Grupo de Beneficios Id: " & vGrid.Text)
   
        vGrid.Col = 5
        vGrid.TextTip = TextTipFixed
        vGrid.TextTipDelay = 1000
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = "Usuario " & glogon.Usuario _
                         & vbCrLf & " Fecha " & Format(Date, "dd/mm/yyyy")
   
   Else 'Actualizar

      vGrid.Col = 1
      pCodigo = vGrid.Text
      
      vGrid.Col = 2
      strSQL = "update afi_bene_grupos set descripcion = '" & vGrid.Text & "', cod_Categoria = '"
      vGrid.Col = 3
      strSQL = strSQL & vGrid.Text & "', Monto = "
      vGrid.Col = 4
      strSQL = strSQL & CCur(vGrid.Text) & ", Estado = "
      vGrid.Col = 5
      strSQL = strSQL & vGrid.Value & " where cod_grupo = " & pCodigo
      
      Call ConectionExecute(strSQL)
      
      
      vGrid.Col = 1
      Call Bitacora("Modifica", "Grupo de Beneficios Id: " & vGrid.Text)
    
End If

   vGrid.Col = 1
   fxGuardar = vGrid.Text
   
Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Function


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, i As Integer

On Error GoTo vError


On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i > 0 Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        If vGrid.MaxRows <= vGrid.ActiveRow Then
          vGrid.MaxRows = vGrid.MaxRows + 1
          vGrid.Row = vGrid.MaxRows
          Call sbCboCategorias(3, vGrid.MaxRows, vGrid)
        End If
  End If 'Actualiza o Inserta
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
    Call sbCboCategorias(3, vGrid.ActiveRow, vGrid)
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then

     vGrid.Row = vGrid.ActiveRow
     vGrid.Col = 1

     If vGrid.Text = "" Then Exit Sub

     i = MsgBox("Está Seguro que desea borrar este registro?", vbYesNo)
     If i = vbYes Then
       
       strSQL = "delete afi_bene_grupos where cod_Grupo = " & vGrid.Text
       Call ConectionExecute(strSQL)
        
        vGrid.Col = 1
        Call Bitacora("Elimina", "Grupo de Beneficios Id: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
        
     End If

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub



Private Sub sbCargaBeneficios()
Dim strSQL As String, rs As New ADODB.Recordset

Dim itmX As ListViewItem, i As Integer

If vPaso Then Exit Sub
If cboGrupos.ListCount <= 0 Then Exit Sub


lsw.ListItems.Clear

strSQL = "select B.cod_beneficio,B.descripcion,isnull(G.Cod_grupo,-1) as 'Marca'" _
       & " from  afi_beneficios B left join afi_Grupo_Beneficio G" _
       & " on B.Cod_beneficio = G.cod_beneficio and G.cod_grupo = " & cboGrupos.ItemData(cboGrupos.ListIndex) & "" _
       & " order by G.cod_beneficio desc,B.descripcion"
          
Call OpenRecordSet(rs, strSQL)

vPaso = True
Do While Not rs.EOF

    Set itmX = lsw.ListItems.Add(, , Trim(rs!Cod_Beneficio))
        itmX.SubItems(1) = rs!Descripcion
        
    If rs!Marca > 0 Then
        itmX.Checked = True
        itmX.ForeColor = vbBlue
    End If

 rs.MoveNext
Loop
rs.Close

vPaso = False

End Sub

