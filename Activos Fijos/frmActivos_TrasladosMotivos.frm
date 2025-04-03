VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmActivos_TrasladosMotivos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Motivos de Traslasdos de Responsabilidad de Activos"
   ClientHeight    =   7392
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   7428
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7392
   ScaleWidth      =   7428
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5892
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   6972
      _Version        =   524288
      _ExtentX        =   12298
      _ExtentY        =   10393
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
      SpreadDesigner  =   "frmActivos_TrasladosMotivos.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Motivos para cambio de Responsable"
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
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   5892
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmActivos_TrasladosMotivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_activote()
vModulo = 8
End Sub

Private Sub Form_Activate()
vModulo = 36

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 36


Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

 
strSQL = "select cod_motivo,descripcion,activo from ACTIVOS_TRASLADOS_MOTIVOS"
Call sbCargaGrid(vGrid, 3, strSQL)
   
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


If KeyCode = vbKeyDelete Then
   'Aqui codigo de Borrado
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
   If Trim(vGrid.Text) <> "" Then
    strSQL = "Delete ACTIVOS_TRASLADOS_MOTIVOS where cod_motivo =  '" & UCase(vGrid.Text) & "'"
    Call ConectionExecute(strSQL)
   End If
   vGrid.DeleteRows vGrid.ActiveRow, 1
   vGrid.MaxRows = vGrid.MaxRows - 1
End If


If KeyCode = vbKeyInsert Then
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.InsertRows vGrid.ActiveRow, 1
  vGrid.Row = vGrid.ActiveRow
End If


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
If vGrid.Text = "" Then vGrid.Text = 0
strSQL = "select coalesce(count(*),0) as Existe from ACTIVOS_TRASLADOS_MOTIVOS" _
       & " where cod_motivo ='" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL, 0)

If rs!Existe = 0 Then 'Insertar
    If Trim(vGrid.Text) = "" Then Exit Function
    strSQL = "insert into ACTIVOS_TRASLADOS_MOTIVOS(cod_motivo,descripcion,activo,registro_usuario,registro_fecha)" _
           & " values('" & vGrid.Text & "',"
    vGrid.Col = 2
    strSQL = strSQL & "'" & vGrid.Text & "',"
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',getdate())"
    
    Call ConectionExecute(strSQL)
    
    vGrid.Col = 1
    Call Bitacora("Registra", "Motivo de Traslado: " & vGrid.Text)

Else 'Actualizar
    
    vGrid.Col = 2
    strSQL = "update ACTIVOS_TRASLADOS_MOTIVOS set descripcion= '" & vGrid.Text & "',activo = "
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Text & " where cod_motivo =  '"
    vGrid.Col = 1
    strSQL = strSQL & UCase(vGrid.Text) & "'"
    
     
    
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Motivo de Traslado: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

