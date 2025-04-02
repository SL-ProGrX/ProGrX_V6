VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCO_CJ_Proceso 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cobros: Proceso (Etapas)"
   ClientHeight    =   6612
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   11172
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6612
   ScaleWidth      =   11172
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4932
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   10812
      _Version        =   524288
      _ExtentX        =   19071
      _ExtentY        =   8700
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
      MaxCols         =   497
      ScrollBars      =   2
      SpreadDesigner  =   "frmCO_CJ_Proceso.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Etapas del Proceso de Cobro Judicial"
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
      Height          =   612
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   7332
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   11292
   End
End
Attribute VB_Name = "frmCO_CJ_Proceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Activate()
vModulo = 6
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 6
vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

strSQL = "select cod_proceso,descripcion,honorarios_aplica,honorarios_Monto,orden,Activo from Cbr_Cj_Proceso" _
      & " order by cod_proceso"
Call sbCargaGrid(vGrid, 6, strSQL)

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
vGrid.Col = 4
If vGrid.Text = "" Then
   vGrid.Text = 0
End If


vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from Cbr_Cj_Proceso " _
       & " where cod_proceso = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert Cbr_Cj_Proceso(cod_proceso,descripcion,honorarios_aplica,honorarios_monto,orden,Activo,Registro_Usuario,Registro_Fecha) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 4
  strSQL = strSQL & IIf((vGrid.Text = ""), 0, CCur(vGrid.Text)) & ",'"
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 6
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
  
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Proceso de Cbr.Jud.: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update Cbr_Cj_Proceso set descripcion = '" & vGrid.Text & "',Honorarios_Aplica = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & ",Honorarios_Monto = "
 vGrid.Col = 4
 strSQL = strSQL & IIf((vGrid.Text = ""), 0, CCur(vGrid.Text)) & ",Orden = '"
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Text & "',Activo = "
 vGrid.Col = 6
 strSQL = strSQL & vGrid.Value & " where cod_proceso = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Proceso de Cbr.Jud.: " & vGrid.Text)

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
        vGrid.Col = 1
        strSQL = "delete Cbr_Cj_Proceso where cod_proceso = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Proceso de Cbr.Jud.: " & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
