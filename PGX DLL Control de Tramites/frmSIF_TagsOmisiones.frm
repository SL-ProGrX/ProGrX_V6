VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmSIF_TagsOmisiones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Definición de Omisiones"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   13695
      _Version        =   1441793
      _ExtentX        =   24156
      _ExtentY        =   10821
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
      Item(0).Caption =   "Lista de Errores y Omisiones"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Asignación de Etiquetas"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "lsw"
      Item(1).Control(1)=   "ShortcutCaption1"
      Item(1).Control(2)=   "cboModulos"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5175
         Left            =   -66160
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   5700
         _Version        =   1441793
         _ExtentX        =   10054
         _ExtentY        =   9128
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
         Height          =   5415
         Left            =   0
         TabIndex        =   1
         Top             =   480
         Width           =   13335
         _Version        =   524288
         _ExtentX        =   23521
         _ExtentY        =   9551
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
         MaxCols         =   485
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "frmSIF_TagsOmisiones.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboModulos 
         Height          =   330
         Left            =   -66160
         TabIndex        =   4
         Top             =   375
         Visible         =   0   'False
         Width           =   5775
         _Version        =   1441793
         _ExtentX        =   10186
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
         Width           =   13695
         _Version        =   1441793
         _ExtentX        =   24156
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
      Caption         =   "Errores y Omsiones + Asignación de Etiquetas"
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
      Width           =   11055
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   14055
   End
End
Attribute VB_Name = "frmSIF_TagsOmisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean
Dim itmX As ListViewItem

Private Sub cboModulos_Click()
If vPaso Then Exit Sub
If cboModulos.ListCount = 0 Then Exit Sub

Call sbModulos_Load

End Sub

Private Sub Form_Activate()
vModulo = 8
End Sub

Private Sub Form_Load()

tcMain.Item(0).Selected = True

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Usuario", 1500
    .Add , , "Descripción", 3500
End With

strSQL = "SELECT ID_ERROR,DESCRIPCION,MENSAJE,ACTIVO FROM  SIF_OMISIONES" _
       & " order by ID_ERROR"
Call sbCargaGrid(vGrid, 4, strSQL)
 
Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from SIF_OMISIONES " _
       & " where ID_ERROR = " & vGrid.Text
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into SIF_OMISIONES(ID_ERROR,DESCRIPCION,MENSAJE,ACTIVO) values(" _
         & vGrid.Text & ",'"
  vGrid.Col = 2
  strSQL = strSQL & Trim(vGrid.Text) & "','"
  vGrid.Col = 3
  strSQL = strSQL & Trim(vGrid.Text) & "','"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Control de Errores: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update SIF_OMISIONES set DESCRIPCION = '" & Trim(UCase(vGrid.Text)) & "', MENSAJE = '"
 vGrid.Col = 3
 strSQL = strSQL & Trim(vGrid.Text) & "', ACTIVO = '"
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Value & "' where ID_ERROR = "

 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Control de Errores: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert SIF_OMISIONES_MODULOS(ID_ERROR,cod_modulo) values('" & Item.Text _
            & "','" & cboModulos.ItemData(cboModulos.ListIndex) & "')"
Else
   strSQL = "Delete SIF_OMISIONES_MODULOS where ID_ERROR ='" & Item.Text _
          & "' and cod_modulo = '" & cboModulos.ItemData(cboModulos.ListIndex) & "'"
End If
Call ConectionExecute(strSQL)

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
Dim i As Integer, strSQL As String

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

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este error", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete SIF_OMISIONES where ID_ERROR = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Elimina", "Control de Errores: " & vGrid.Text)
        
        vGrid.Col = 1
        strSQL = "SELECT ID_ERROR,DESCRIPCION,MENSAJE,ACTIVO FROM  SIF_OMISIONES" _
               & " order by ID_ERROR"
        Call sbCargaGrid(vGrid, 4, strSQL)
     End If
End If


End Sub



Private Sub sbModulos_Load()

On Error GoTo vError

lsw.ListItems.Clear

strSQL = "select O.ID_ERROR,O.descripcion,M.ID_ERROR as codigo" _
       & " from SIF_OMISIONES O left join SIF_OMISIONES_MODULOS M on O.ID_ERROR = M.ID_ERROR" _
       & " and M.cod_Modulo = '" & cboModulos.ItemData(cboModulos.ListIndex) & "'" _
       & " where O.ACTIVO = '1'  order by M.ID_ERROR desc"
          
Call OpenRecordSet(rs, strSQL)

vPaso = True

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!ID_ERROR)
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

