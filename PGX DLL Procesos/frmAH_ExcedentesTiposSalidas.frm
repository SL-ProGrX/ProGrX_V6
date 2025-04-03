VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAH_ExcedentesTiposSalidas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Excedentes: Tipos de Salidas"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   15375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   240
      Top             =   840
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   7455
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   15135
      _Version        =   524288
      _ExtentX        =   26696
      _ExtentY        =   13150
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
      MaxCols         =   483
      ScrollBars      =   2
      SpreadDesigner  =   "frmAH_ExcedentesTiposSalidas.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Salidas"
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
      Height          =   492
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   5172
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   17295
   End
End
Attribute VB_Name = "frmAH_ExcedentesTiposSalidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub Form_Activate()
vModulo = 2
End Sub


Private Sub sbConsulta()

On Error GoTo vError
                       
                      
strSQL = "select cod_salida,descripcion, activa, opcion_sistema, destino_operadora, Destino_Plan , Destino_Banco " _
       & ", TIPO_APLICACION, PERMITE_RECLASIFICAR, REQUIERE_PORCENTAJE, TIPO_APLICACION_DESC, PLAN_DESC, BANCO_DESC" _
       & " from vExc_Salidas_Tipos order by  Activa desc, cod_salida"
Call OpenRecordSet(rs, strSQL)

With vGrid
    .MaxRows = 0
    .MaxCols = 10
    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      .Col = 1
      .Text = rs!Cod_Salida
      .Col = 2
      .Text = rs!Descripcion
      .Col = 3
      .Value = rs!Activa
      .Col = 4
      .Value = rs!Opcion_Sistema
      .Col = 5
      .Text = rs!Tipo_Aplicacion_Desc
      .Col = 6
      .Text = CStr(rs!Destino_Operadora)
      .Col = 7
      .Text = rs!Destino_Plan
      If rs!Plan_Desc <> "" Then
        .TextTip = TextTipFixed
        .TextTipDelay = 1000
        .CellNoteIndicatorShape = CellNoteIndicatorShapeSquare
        .CellNoteIndicatorColor = vbRed
        .CellNote = rs!Plan_Desc
      End If
      .Col = 8
      .Text = CStr(rs!Destino_Banco)
      If rs!BANCO_DESC <> "" Then
        .TextTip = TextTipFixed
        .TextTipDelay = 1000
        .CellNoteIndicatorShape = CellNoteIndicatorShapeSquare
        .CellNoteIndicatorColor = vbRed
        .CellNote = rs!BANCO_DESC
      End If
      
      .Col = 9
      .Value = rs!REQUIERE_PORCENTAJE
      .Col = 10
      .Value = rs!PERMITE_RECLASIFICAR
      
    
      rs.MoveNext
    Loop
    rs.Close
    
    'Abre Linea Nueva
    .MaxRows = .MaxRows + 1
End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

vModulo = 2

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vGrid.AppearanceStyle = fxGridStyle

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub



Private Function fxGuardar() As Long
Dim vCuenta As String, vTipo As String

On Error GoTo vError

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

fxGuardar = 0


If Trim(vGrid.Text) = "" Then
    MsgBox "No se especificó el código de salida (verifique!)", vbExclamation
    Exit Function
End If

strSQL = "select isnull(count(*),0) as Existe from EXC_TIPOS_SALIDAS" _
       & " where cod_salida = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  
  strSQL = "insert into EXC_TIPOS_SALIDAS(cod_salida,descripcion, activa, opcion_sistema, TIPO_APLICACION" _
         & ", destino_operadora, Destino_Plan , Destino_Banco, REQUIERE_PORCENTAJE, PERMITE_RECLASIFICAR" _
         & ", Registro_fecha, Registro_usuario) values('" & vGrid.Text
  vGrid.Col = 2
  strSQL = strSQL & "', '" & vGrid.Text
  vGrid.Col = 3
  strSQL = strSQL & "', " & vGrid.Value
  vGrid.Col = 4
  strSQL = strSQL & ", " & vGrid.Value
  vGrid.Col = 5
  strSQL = strSQL & ", '" & Mid(vGrid.Text, 1, 1)
  vGrid.Col = 6
  strSQL = strSQL & "', " & vGrid.Text
  vGrid.Col = 7
  strSQL = strSQL & ", '" & vGrid.Text
  vGrid.Col = 8
  strSQL = strSQL & "', " & vGrid.Text
  vGrid.Col = 9
  strSQL = strSQL & ", " & vGrid.Value
  vGrid.Col = 10
  strSQL = strSQL & ", " & vGrid.Value & ", dbo.MyGetdate(), '" & glogon.Usuario & "')"
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  
  Call Bitacora("Registra", "Excedentes: Tipo de Salida: " & vGrid.Text)
  
Else 'Actualizar

    vGrid.Col = 2
    strSQL = "update EXC_TIPOS_SALIDAS set descripcion = '" & vGrid.Text & "', Activa = "
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Value & ", Opcion_Sistema = "
    vGrid.Col = 4
    strSQL = strSQL & vGrid.Value & ", Tipo_Aplicacion = '"
    vGrid.Col = 5
    strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "', Destino_Operadora = "
    vGrid.Col = 6
    strSQL = strSQL & vGrid.Text & ", Destino_Plan = '"
    vGrid.Col = 7
    strSQL = strSQL & vGrid.Text & "', Destino_Banco = "
    vGrid.Col = 8
    strSQL = strSQL & vGrid.Text & ", Requiere_Porcentaje = "
    vGrid.Col = 9
    strSQL = strSQL & vGrid.Value & ", Permite_Reclasificar = "
    vGrid.Col = 10
    strSQL = strSQL & vGrid.Value & ", Modifica_Fecha = dbo.MyGetdate(), Modifica_Usuario = '" & glogon.Usuario & "' "
    
    vGrid.Col = 1
    strSQL = strSQL & " where cod_salida = '" & vGrid.Text & "'"
    Call ConectionExecute(strSQL)
 
    vGrid.Col = 1
   
    Call Bitacora("Modifica", "Excedentes: Tipo de Salida: " & vGrid.Text)
 
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
Call sbConsulta
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, i As Long


On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Consulta Plan
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 7 Then
    vGrid.Row = vGrid.ActiveRow
    
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_plan, descripcion, cod_Operadora from fnd_Planes"
    gBusquedas.Filtro = " AND Estado = 'A' and TIPO_CDP = 0 AND PATRIMONIO_ENLACE = 0 AND TIPO_DEDUC = 'M'"
    frmBusquedas.Show vbModal
    If gBusquedas.Resultado <> "" Then
        vGrid.Col = 6
        vGrid.Text = gBusquedas.Resultado3
        vGrid.Col = 7
        vGrid.Text = gBusquedas.Resultado
        vGrid.TextTip = TextTipFixed
        vGrid.TextTipDelay = 1000
        vGrid.CellNoteIndicatorShape = CellNoteIndicatorShapeSquare
        vGrid.CellNoteIndicatorColor = vbRed
        vGrid.CellNote = gBusquedas.Resultado2
    End If
End If

'Consulta Bancos
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 8 Then
    vGrid.Row = vGrid.ActiveRow
    
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select Id_Banco, Descripcion, Cta from Tes_Bancos"
    gBusquedas.Filtro = " AND Estado = 'A'"
    frmBusquedas.Show vbModal
    If gBusquedas.Resultado <> "" Then
        vGrid.Col = 8
        vGrid.Text = gBusquedas.Resultado
        vGrid.TextTip = TextTipFixed
        vGrid.TextTipDelay = 1000
        vGrid.CellNoteIndicatorShape = CellNoteIndicatorShapeSquare
        vGrid.CellNoteIndicatorColor = vbRed
        vGrid.CellNote = gBusquedas.Resultado2
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
        
        strSQL = "delete EXC_TIPOS_SALIDAS where cod_salida = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Excedentes: Tipo de Salida: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


