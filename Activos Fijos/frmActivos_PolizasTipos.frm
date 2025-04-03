VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmActivos_PolizasTipos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipos de Pólizas"
   ClientHeight    =   7080
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7404
   Icon            =   "frmActivos_PolizasTipos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   7404
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5412
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   6852
      _Version        =   524288
      _ExtentX        =   12086
      _ExtentY        =   9546
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
      MaxCols         =   3
      ScrollBars      =   2
      SpreadDesigner  =   "frmActivos_PolizasTipos.frx":000C
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Pólizas"
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
      Height          =   372
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   5052
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmActivos_PolizasTipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 36

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 36


Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

strSQL = "select * from Activos_polizas_tipos" _
      & " order by Tipo_Poliza"
Call sbCargaGridLocal(vGrid, strSQL)

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

strSQL = "select coalesce(count(*),0) as Existe from Activos_polizas_tipos " _
       & " where Tipo_Poliza = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL, 0)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into Activos_polizas_tipos(Tipo_Poliza,descripcion,activo,registro_usuario,registro_fecha) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',getdate())"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Tipo de Póliza : " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update Activos_polizas_tipos set descripcion = '" & vGrid.Text & "',Activo = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & ",modifica_usuario = '" & glogon.Usuario & "',modifica_fecha = getdate()" _
        & " where Tipo_Poliza = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Modifica", "Tipo de Póliza : " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub sbCargaGridLocal(ByRef pGrid As Object, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String
Dim strUltimaSeleccion As String



Me.MousePointer = vbHourglass

On Error GoTo vError

pGrid.MaxRows = 0
pGrid.MaxRows = 1
pGrid.Row = pGrid.MaxRows

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)

With pGrid
Do While Not rs.EOF
  .Row = pGrid.MaxRows
  .Col = 1
  
    For i = 1 To 3
      .Col = i
      Select Case i
       Case 1 'Codigo
          .Text = rs!Tipo_Poliza
          .TextTip = TextTipFixed
          .TextTipDelay = 1000
          .CellNote = "Registrado: " & rs!registro_usuario & vbCrLf & "Fecha: " & rs!registro_fecha & vbCrLf & vbCrLf _
                    & "Modificado: " & rs!Modifica_Usuario & vbCrLf & "Fecha: " & rs!Modifica_Fecha
       Case 2 'Descripcion
          .Text = rs!Descripcion
       Case 3 'Activo
          .Value = rs!activo
      End Select
    Next i
  
  pGrid.MaxRows = pGrid.MaxRows + 1
  
  rs.MoveNext

Loop

End With

rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, i As Integer

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
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

'Borrar Línea
If KeyCode = vbKeyDelete Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  strSQL = "delete Activos_polizas_tipos where Tipo_Poliza = '" & vGrid.Text & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Tipo de Póliza : " & vGrid.Text)
    
  vGrid.DeleteRows vGrid.ActiveRow, 1
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub




