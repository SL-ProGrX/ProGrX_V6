VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_Bienes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipos de Bienes (Activos Relevantes) de la Persona"
   ClientHeight    =   6684
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9528
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6684
   ScaleWidth      =   9528
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5172
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   9252
      _Version        =   524288
      _ExtentX        =   16320
      _ExtentY        =   9123
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
      MaxCols         =   490
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_Bienes.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Bienes de las Personas"
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
      Height          =   480
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   7812
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10092
   End
End
Attribute VB_Name = "frmAF_Bienes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub sbInicial()
Dim strSQL As String
      
strSQL = "select BIEN_TIPO,descripcion,ACTIVO,Registro_Fecha,Registro_Usuario" _
       & " from AFI_BIENES_TIPOS" _
       & " order by BIEN_TIPO"
vPaso = True
    Call sbCargaGrid(vGrid, 5, strSQL)
vPaso = False
End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 1
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from AFI_BIENES_TIPOS " _
       & " where BIEN_TIPO = '" & vGrid.Text & "'"

Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function

  strSQL = "insert into AFI_BIENES_TIPOS(BIEN_TIPO,descripcion,ACTIVO,registro_fecha,registro_usuario) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",Getdate(),'" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 4
  vGrid.Text = fxFechaServidor
  vGrid.Col = 5
  vGrid.Text = glogon.Usuario
  
  Call Bitacora("Registra", "Tipo de Bien: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update AFI_BIENES_TIPOS set Descripcion = '" & vGrid.Text & "',ACTIVO= "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & " where BIEN_TIPO = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Tipo de Bien: " & vGrid.Text)

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
Call sbInicial
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If (vGrid.ActiveCol = vGrid.MaxCols Or vGrid.ActiveCol = 3) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
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
        strSQL = "delete AFI_BIENES_TIPOS where BIEN_TIPO = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Tipo de Bien: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub






