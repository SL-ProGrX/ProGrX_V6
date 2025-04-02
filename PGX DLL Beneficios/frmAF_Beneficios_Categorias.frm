VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_Beneficios_Categorias 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Categorías para Configuración de Beneficios"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   14880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   14655
      _Version        =   524288
      _ExtentX        =   25850
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
      MaxCols         =   9
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_Beneficios_Categorias.frx":0000
      VScrollSpecialType=   2
      Appearance      =   1
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Categorías para Beneficios"
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
      TabIndex        =   1
      Top             =   360
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   15015
   End
End
Attribute VB_Name = "frmAF_Beneficios_Categorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub Form_Activate()
vModulo = 7
End Sub

Private Sub Form_Load()

vModulo = 7

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vGrid.AppearanceStyle = fxGridStyle


Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "select COD_CATEGORIA, descripcion, I_Apremiante, I_Desastres, I_Sepelio, I_Fena, I_Reconocimientos, I_Crece" _
      & ", Activo from AFI_BENE_CATEGORIAS order by COD_CATEGORIA"
Call sbCargaGrid(vGrid, 9, strSQL)

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

strSQL = "select isnull(count(*),0) as Existe from AFI_BENE_CATEGORIAS " _
       & " where COD_CATEGORIA = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  
  strSQL = "insert into AFI_BENE_CATEGORIAS(COD_CATEGORIA, descripcion, I_Apremiante, I_Desastres, I_Sepelio, I_Fena, I_Reconocimientos, I_Crece, Activo" _
         & ", Registro_Fecha, Registro_Usuario) values('" _
         & vGrid.Text & "', '"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "', "
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ", "
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ", "
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ", "
  vGrid.Col = 6
  strSQL = strSQL & vGrid.Value & ", "
  vGrid.Col = 7
  strSQL = strSQL & vGrid.Value & ", "
  vGrid.Col = 8
  strSQL = strSQL & vGrid.Value & ", "
  vGrid.Col = 9
  strSQL = strSQL & vGrid.Value & ", dbo.MyGetdate(), '" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Categorías para Beneficios Id: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update AFI_BENE_CATEGORIAS set descripcion = '" & vGrid.Text & "', I_Apremiante = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & ", I_Desastres = "
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Value & ", I_Sepelio = "
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Value & ", I_Fena = "
 vGrid.Col = 6
 strSQL = strSQL & vGrid.Value & ", I_Reconocimientos = "
 vGrid.Col = 7
 strSQL = strSQL & vGrid.Value & ", I_Crece = "
 vGrid.Col = 8
 strSQL = strSQL & vGrid.Value & ", Activo = "
 vGrid.Col = 9
 strSQL = strSQL & vGrid.Value & ", Modifica_Fecha = dbo.MyGetdate(), Modifica_Usuario = '" _
        & glogon.Usuario & "' where COD_CATEGORIA = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Categorías para Beneficios Id: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


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

'Elimina
If KeyCode = vbKeyDelete Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
     i = MsgBox("Está Seguro que desea borrar este registro?", vbYesNo)
     If i = vbYes Then
        strSQL = "delete AFI_BENE_CATEGORIAS where COD_CATEGORIA = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Categorías para Beneficios Id: " & vGrid.Text)
        
        vGrid.Col = 1
        strSQL = "select COD_CATEGORIA, descripcion, I_Apremiante, I_Desastres, I_Sepelio, I_Fena, I_Reconocimientos, I_Crece" _
              & ", Activo from AFI_BENE_CATEGORIAS order by COD_CATEGORIA"
        Call sbCargaGrid(vGrid, 9, strSQL)
    End If

End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


End Sub


