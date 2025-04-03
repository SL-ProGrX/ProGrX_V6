VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_TiposIds 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipos de Identificaciones"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12870
   Icon            =   "frmAF_TiposIds.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   8040
      Top             =   720
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   12615
      _Version        =   524288
      _ExtentX        =   22251
      _ExtentY        =   8705
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
      SpreadDesigner  =   "frmAF_TiposIds.frx":6852
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Identificaciones"
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
      Left            =   1880
      TabIndex        =   0
      Top             =   360
      Width           =   6855
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmAF_TiposIds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub Form_Load()

vModulo = 1

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture



End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from afi_Tipos_IDs " _
       & " where Tipo_ID = " & vGrid.Text
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into afi_Tipos_IDs(Tipo_ID, descripcion, Tipo_Personeria, Largo_Minimo, Mascara" _
         & ", CODIGO_SUGEF, CODIGO_PIN, CODIGO_HACIENDA, CODIGO_SINPE, Usuario, Fecha) values(" _
         & vGrid.Text & ",'"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "',"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Text & ",'"
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Text & "', "
  vGrid.Col = 6 'SUGEF
  strSQL = strSQL & vGrid.Text & ", "
  vGrid.Col = 7 'PIN
  strSQL = strSQL & vGrid.Text & ", "
  vGrid.Col = 8 'HACIENDA
  strSQL = strSQL & vGrid.Text & ", "
  vGrid.Col = 9 'SINPE
  strSQL = strSQL & vGrid.Text & ", '" & glogon.Usuario & "',dbo.MyGetdate())"
  
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Tipo de Idenficiación : " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update afi_Tipos_IDs set descripcion = '" & vGrid.Text & "', Tipo_Personeria = '"
 vGrid.Col = 3
 strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "', Largo_Minimo = "
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Text & ", Mascara = '"
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Text & "', CODIGO_SUGEF = "
 vGrid.Col = 6
 strSQL = strSQL & vGrid.Text & ", CODIGO_PIN = "
 vGrid.Col = 7
 strSQL = strSQL & vGrid.Text & ", CODIGO_HACIENDA = "
 vGrid.Col = 8
 strSQL = strSQL & vGrid.Text & ", CODIGO_SINPE = "
 vGrid.Col = 9
 
 vGrid.Col = 9
 strSQL = strSQL & vGrid.Text & ", MODIFICA_USUARIO = '" & glogon.Usuario & "', MODIFICA_FECHA = dbo.MyGetdate() where Tipo_ID = "
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Tipo de Idenficiación : " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False


Dim strSQL As String

strSQL = "select Tipo_ID,descripcion, Tipo_Personeria_Desc" _
      & ", Largo_Minimo, Mascara, CODIGO_SUGEF, CODIGO_PIN, CODIGO_HACIENDA, CODIGO_SINPE" _
      & " from vSys_Tipos_Ids" _
      & " order by Tipo_ID"
Call sbCargaGrid(vGrid, 9, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

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
        vGrid.Col = 1
        strSQL = "delete afi_Tipos_IDs where Tipo_ID = " & vGrid.Text
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Tipo de Idenficiación : " & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

