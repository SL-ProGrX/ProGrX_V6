VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_BeneficioRoles 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Roles de Usuarios para Beneficios y Ayudas Sociales"
   ClientHeight    =   6240
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4812
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   9492
      _Version        =   1441793
      _ExtentX        =   16743
      _ExtentY        =   8488
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
      Color           =   128
      ItemCount       =   2
      Item(0).Caption =   "Categorías"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Asignación"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "cboGrupos"
      Item(1).Control(1)=   "Label1(1)"
      Item(1).Control(2)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3972
         Left            =   -69880
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   9372
         _Version        =   1441793
         _ExtentX        =   16531
         _ExtentY        =   7006
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
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4092
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   7452
         _Version        =   524288
         _ExtentX        =   13144
         _ExtentY        =   7218
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
         SpreadDesigner  =   "frmAF_BeneficioRoles.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboGrupos 
         Height          =   312
         Left            =   -67600
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   5172
         _Version        =   1441793
         _ExtentX        =   9128
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupos"
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
         Height          =   372
         Index           =   1
         Left            =   -69520
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1572
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Roles de Usuarios para Beneficios y Ayudas Sociales"
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
      Left            =   2280
      TabIndex        =   5
      Top             =   360
      Width           =   7572
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmAF_BeneficioRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vPaso As Boolean

Private Sub sbCarga_Usuarios()
Dim itmX As ListViewItem, i As Integer
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub
If cboGrupos.ListCount <= 0 Then Exit Sub

lsw.ListItems.Clear

strSQL = "select U.nombre,U.descripcion,A.usuario" _
        & " from Usuarios U left join AFI_BENE_USERG A on U.nombre = A.usuario" _
        & "         and A.cod_grupo = '" & cboGrupos.ItemData(cboGrupos.ListIndex) & "' " _
        & " Where U.estado = 'A'" _
        & " order by A.usuario desc,U.nombre asc"
Call OpenRecordSet(rs, strSQL)

vPaso = True

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Nombre)
     itmX.SubItems(1) = rs!Descripcion
 If Not IsNull(rs!Usuario) Then
     itmX.Checked = True
     itmX.ForeColor = vbBlue
 End If
 
 rs.MoveNext
Loop
rs.Close

vPaso = False

End Sub



Private Sub cboGrupos_Click()

If vPaso Then Exit Sub

Call sbCarga_Usuarios

End Sub

Private Sub Form_Activate()
vModulo = 7
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 7

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2300, vbCenter
    .Add , , "Descripción", 4500
End With

tcMain.Item(0).Selected = True

strSQL = "select cod_grupo,descripcion from AFI_BENEFICIO_GRUPOS"
Call sbCargaGrid(vGrid, 2, strSQL)
    

Call Formularios(Me)
Call RefrescaTags(Me)
End Sub


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert AFI_BENE_USERG(usuario,cod_grupo) values('" & Item.Text _
            & "','" & cboGrupos.ItemData(cboGrupos.ListIndex) & "')"
Else
   strSQL = "Delete AFI_BENE_USERG where usuario ='" & Item.Text _
          & "' and cod_grupo ='" & cboGrupos.ItemData(cboGrupos.ListIndex) & "'"
End If
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
          & " from AFI_BENEFICIO_GRUPOS"
   
   Call sbCbo_Llena_New(cboGrupos, strSQL, False, True)
   
   vPaso = False
   
   Call sbCarga_Usuarios

End If

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

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from AFI_BENEFICIO_GRUPOS" _
       & " where cod_grupo = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into AFI_BENEFICIO_GRUPOS(cod_grupo,descripcion) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & UCase(vGrid.Text) & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Grupo de Usuarios: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update AFI_BENEFICIO_GRUPOS set descripcion = '" & vGrid.Text & "'"
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



