VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmFNDUsuariosAutorizados 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios Autorizados"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   11685
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   7095
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   11295
      _Version        =   524288
      _ExtentX        =   19923
      _ExtentY        =   12515
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
      SpreadDesigner  =   "frmFNDUsuariosAutorizados.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios Autorizados para administrar cuentas de colaboradores"
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
      Height          =   720
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   7452
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   12252
   End
End
Attribute VB_Name = "frmFNDUsuariosAutorizados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Activate()
vModulo = 18
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 18
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "select USUARIO,activo,registro_fecha,registro_usuario,actualiza_fecha,actualiza_usuario" _
      & "  from FND_COLABORADORES_CC order by USUARIO"
Call sbCargaGrid(vGrid, 6, strSQL)


End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from FND_COLABORADORES_CC " _
       & " where USUARIO = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function

  strSQL = "insert into FND_COLABORADORES_CC(USUARIO,activo,registro_fecha,registro_usuario) values('" _
         & UCase(vGrid.Text) & "',"
  vGrid.col = 2
  strSQL = strSQL & vGrid.Value & ",dbo.MyGetdate(),'" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Usuario Aut. Colaboradores: " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update FND_COLABORADORES_CC set activo = "
 strSQL = strSQL & vGrid.Value & ",Actualiza_Fecha = dbo.MyGetdate(), Actualiza_Usuario = '" & glogon.Usuario & "'" _
        & " where USUARIO = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Usuario Aut. Colaboradores: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function




Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

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
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        strSQL = "delete FND_COLABORADORES_CC where USUARIO = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)

        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Usuario Aut. Colaboradores: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


