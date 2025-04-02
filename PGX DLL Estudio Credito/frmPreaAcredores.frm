VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmPreaAcredores 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de Acredores Autorizados"
   ClientHeight    =   6588
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   11892
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6588
   ScaleWidth      =   11892
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5292
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   11772
      _Version        =   524288
      _ExtentX        =   20765
      _ExtentY        =   9335
      _StockProps     =   64
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
      ScrollBars      =   2
      SpreadDesigner  =   "frmPreaAcredores.frx":0000
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Acredores Autorizados para Desembolsos Ordinarios"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   10095
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmPreaAcredores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 3 'Modulo de Credito
End Sub

Private Sub Form_Load()

vModulo = 3 'Modulo de Credito

Call Formularios(Me)
Call RefrescaTags(Me)

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call sbCargar

End Sub


Private Sub sbCargar()
Dim strSQL As String

On Error GoTo vError

strSQL = "select cod_acredor,nombre,nombre_giro,ISNULL(MODIFICA_NOMBRE_GIRO,0),activo from Crd_Prea_Acredores" _
      & " order by cod_acredor"
Call sbCargaGrid(vGrid, 5, strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from Crd_Prea_Acredores " _
       & " where cod_acredor = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into Crd_Prea_Acredores(cod_acredor,nombre,nombre_giro,MODIFICA_NOMBRE_GIRO,activo) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ")"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "PreAnalisis / Acredor : " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update Crd_Prea_Acredores set nombre = '" & vGrid.Text & "',Nombre_Giro = '"
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Text & "',MODIFICA_NOMBRE_GIRO = "
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Value & ",Activo = "
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Value & " where cod_acredor = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "PreAnalisis / Acredor : " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
   strSQL = "delete Crd_Prea_Acredores where cod_acredor = '" & vGrid.Text & "'"
   Call ConectionExecute(strSQL)
   strSQL = vGrid.Text
   vGrid.Col = 1
   Call Bitacora("Elimina", "PreAnalisis / Acredor : " & vGrid.Text)
   
   vGrid.DeleteRows vGrid.ActiveRow, 1
   vGrid.MaxRows = vGrid.MaxRows - 1
   If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

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

'Borrar una linea
If KeyCode = vbKeyDelete Then
  Call sbBorrar
End If

End Sub

Private Sub vGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim x As String

If Col = 2 And NewCol = 3 And (Row = NewRow) Then
   vGrid.Row = Row
   vGrid.Col = Col
   x = vGrid.Text
   vGrid.Col = NewCol
   If vGrid.Text = "" Then
     vGrid.Text = x
   End If

End If

End Sub
