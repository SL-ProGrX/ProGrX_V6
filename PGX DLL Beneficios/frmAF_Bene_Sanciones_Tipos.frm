VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_Bene_Sanciones_Tipos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Beneficios: Tipos de Sanciones"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   11655
      _Version        =   524288
      _ExtentX        =   20558
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
      MaxCols         =   5
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_Bene_Sanciones_Tipos.frx":0000
      VScrollSpecialType=   2
      Appearance      =   1
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Sanciones para Beneficios"
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
      Width           =   12615
   End
End
Attribute VB_Name = "frmAF_Bene_Sanciones_Tipos"
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

strSQL = "select TIPO_SANCION, descripcion, CODIGO_COBRO, PLAZO_MAXIMO, Activo from AFI_BENE_SANCIONES_TIPOS" _
      & " order by TIPO_SANCION"
Call sbCargaGrid(vGrid, 5, strSQL)

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

strSQL = "select isnull(count(*),0) as Existe from AFI_BENE_SANCIONES_TIPOS " _
       & " where TIPO_SANCION = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  
  strSQL = "insert into AFI_BENE_SANCIONES_TIPOS(TIPO_SANCION, descripcion, CODIGO_COBRO, PLAZO_MAXIMO, Activo, Registro_Fecha, Registro_Usuario) values('" _
         & UCase(vGrid.Text) & "', '"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "', '"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "', "
  vGrid.Col = 4
  strSQL = strSQL & CLng(vGrid.Text) & ", "
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ", dbo.MyGetdate(), '" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Tipos de Sanciones para Beneficios Id: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update AFI_BENE_SANCIONES_TIPOS set descripcion = '" & vGrid.Text & "', CODIGO_COBRO = '"
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Text & "', PLAZO_MAXIMO ="
 vGrid.Col = 4
 strSQL = strSQL & CLng(vGrid.Text) & ", Activo ="
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Value & ", Modifica_Fecha = dbo.MyGetdate(), Modifica_Usuario = '" _
        & glogon.Usuario & "' where TIPO_SANCION = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Tipos de Sanciones para Beneficios Id: " & vGrid.Text)

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

If vGrid.ActiveCol = 3 And KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "Codigo"
    gBusquedas.Orden = "Codigo"
    gBusquedas.Consulta = "select Codigo, Descripcion From Catalogo"
    gBusquedas.Filtro = " and Retencion = 'S' and Poliza = 'N' and Activo = 1"
    frmBusquedas.Show vbModal
    If gBusquedas.Resultado <> "" Then
       vGrid.Row = vGrid.ActiveRow
       vGrid.Col = 3
       vGrid.Text = gBusquedas.Resultado
    End If
End If

'Elimina
If KeyCode = vbKeyDelete Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
     i = MsgBox("Está Seguro que desea borrar este registro?", vbYesNo)
     If i = vbYes Then
        strSQL = "delete AFI_BENE_SANCIONES_TIPOS where TIPO_SANCION = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Tipos de Sanciones para Beneficios Id: " & vGrid.Text)
        
        vGrid.Col = 1
        strSQL = "select TIPO_SANCION,descripcion,Activo from AFI_BENE_SANCIONES_TIPOS" _
              & " order by TIPO_SANCION"
        Call sbCargaGrid(vGrid, 3, strSQL)
     End If

End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


End Sub


