VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmPreaTiposSalarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipos de Salarios"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4092
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   11532
      _Version        =   524288
      _ExtentX        =   20341
      _ExtentY        =   7218
      _StockProps     =   64
      BorderStyle     =   0
      DisplayRowHeaders=   0   'False
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
      SpreadDesigner  =   "frmPreaTiposSalarios.frx":0000
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Salarios"
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
      Height          =   612
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmPreaTiposSalarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 3 'Modulo de Credito
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 3 'Modulo de Credito

Call Formularios(Me)
Call RefrescaTags(Me)

vGrid.AppearanceStyle = AppearanceStyleVisualStyles

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

strSQL = "select tipo_salario,descripcion,prioridad,meses,modifica_devengado" _
       & ",modifica_rebajo_extras,modifica_extras_fijas,operacion, activo" _
       & " from Crd_Prea_Tipo_Salario order by tipo_salario"
       
Call sbCargaGridLocal(vGrid, 9, strSQL)

End Sub
Public Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
For i = 1 To vGrid.MaxCols
 vGrid.Col = i
 vGrid.Text = ""
Next i

Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    If (vGrid.Col = 5) Or (vGrid.Col = 6) Or (vGrid.Col = 7) Or (vGrid.Col = 9) Then
        vGrid.Value = CInt(rs.Fields(i - 1).Value)
    Else
        vGrid.Text = CStr(rs.Fields(i - 1).Value)
    End If
    
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext
Loop

rs.Close
End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow


vGrid.Col = 2 'Descripcion
If Len(vGrid.Text) = 0 Then
    MsgBox "No se ha indicado una descripción válida!", vbExclamation
    Exit Function
End If


vGrid.Col = 1




strSQL = "select isnull(count(*),0) as Existe from Crd_Prea_Tipo_Salario " _
       & " where tipo_salario = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into Crd_Prea_Tipo_Salario(tipo_salario,descripcion,prioridad" _
         & ",meses,modifica_devengado,modifica_rebajo_extras,modifica_extras_fijas,operacion,activo) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2 'Descripcion
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3 'Prioridad
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 4 'Meses
  strSQL = strSQL & vGrid.Text & ","
  vGrid.Col = 5 'Modifica Devengado
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 6 'Mod. Rebajo Extras
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 7 'Mod. Extras Fijas
  strSQL = strSQL & vGrid.Value & ",'"
  vGrid.Col = 8 'Operacion
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 9 'Activo
  strSQL = strSQL & vGrid.Value & ")"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "PreAnalisis Tipo de Salario Cod: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update Crd_Prea_Tipo_Salario set descripcion = '" & vGrid.Text & "',prioridad = '"
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Text & "',Meses = "
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Text & ",Modifica_devengado = "
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Value & ",Modifica_rebajo_extras = "
 vGrid.Col = 6
 strSQL = strSQL & vGrid.Value & ",Modifica_extras_fijas = "
 vGrid.Col = 7
 strSQL = strSQL & vGrid.Value & ",Operacion = '"
 vGrid.Col = 8
 strSQL = strSQL & vGrid.Text & "',activo = "
 vGrid.Col = 9
 strSQL = strSQL & vGrid.Value & " where tipo_salario = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 Call Bitacora("Modifica", "PreAnalisis Tipo de Salario Cod : " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


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







