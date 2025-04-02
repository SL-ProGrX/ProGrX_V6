VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmActivos_Departamentos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Departamentos y Secciones"
   ClientHeight    =   7236
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   8544
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7236
   ScaleWidth      =   8544
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTab 
      Height          =   5772
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   8292
      _ExtentX        =   14626
      _ExtentY        =   10181
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Departamentos"
      TabPicture(0)   =   "frmActivos_Departamentos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vGrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Secciones"
      TabPicture(1)   =   "frmActivos_Departamentos.frx":6862
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cbo"
      Tab(1).Control(1)=   "vGridSec"
      Tab(1).Control(2)=   "Label2(0)"
      Tab(1).ControlCount=   3
      Begin VB.ComboBox cbo 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   -73560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   6135
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5172
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   6732
         _Version        =   524288
         _ExtentX        =   11874
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
         MaxCols         =   3
         ScrollBars      =   2
         SpreadDesigner  =   "frmActivos_Departamentos.frx":D0C4
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridSec 
         Height          =   4452
         Left            =   -74880
         TabIndex        =   4
         Top             =   1080
         Width           =   7932
         _Version        =   524288
         _ExtentX        =   13991
         _ExtentY        =   7853
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
         SpreadDesigner  =   "frmActivos_Departamentos.frx":D68B
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Departamento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   0
         Left            =   -74760
         TabIndex        =   5
         Top             =   600
         Width           =   1212
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Departamentos y Secciones"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   5052
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmActivos_Departamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vDepartamento As String
Dim vPaso As Boolean


Private Sub cbo_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If Not vPaso Then Exit Sub

strSQL = "select * from Activos_secciones" _
      & " where cod_departamento = '" & SIFGlobal.fxCodText(cbo.Text) _
      & "' order by cod_seccion"
Call sbCargaGridLocal(vGridSec, strSQL, "S")

End Sub

Private Sub Form_Activate()
vModulo = 36
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 36


Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

ssTab.Tab = 0

strSQL = "select * from Activos_departamentos" _
      & " order by cod_departamento"
Call sbCargaGridLocal(vGrid, strSQL, "D")

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub sbCargaGridLocal(ByRef pGrid As Object, strSQL As String, Optional pTipo As String = "D")
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
  
  If pTipo = "D" Then
    'Departamentos
    For i = 1 To 3
      .Col = i
      Select Case i
       Case 1 'Codigo
          .Text = rs!Cod_Departamento
          .TextTip = TextTipFixed
          .TextTipDelay = 1000
          .CellNote = "Registrado: " & rs!registro_usuario & vbCrLf & "Fecha: " & rs!registro_fecha & vbCrLf & vbCrLf _
                    & "Modificado: " & rs!Modifica_Usuario & vbCrLf & "Fecha: " & rs!Modifica_Fecha
       Case 2 'Descripcion
          .Text = rs!Descripcion
      
       Case 3 'Unidad
          .Text = rs!cod_unidad
          .TextTip = TextTipFixed
          .TextTipDelay = 1000
          .CellNote = fxgCntUnidad(rs!cod_unidad)
      End Select
    Next i
  
  
  Else
   'Secciones
      For i = 1 To 3
      .Col = i
      Select Case i
       Case 1 'Codigo
          .Text = rs!Cod_Seccion
          .TextTip = TextTipFixed
          .TextTipDelay = 1000
          .CellNote = "Registrado: " & rs!registro_usuario & vbCrLf & "Fecha: " & rs!registro_fecha & vbCrLf & vbCrLf _
                    & "Modificado: " & rs!Modifica_Usuario & vbCrLf & "Fecha: " & rs!Modifica_Fecha
       Case 2 'Descripcion
          .Text = rs!Descripcion
       Case 3 'Centro de Costo
          .Text = rs!cod_centro_costo
          .TextTip = TextTipFixed
          .TextTipDelay = 1000
          .CellNote = fxgCntCentroCostos(rs!cod_centro_costo)
       
      End Select
    Next i
  End If
  
  
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



Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select coalesce(count(*),0) as Existe from Activos_departamentos " _
       & " where cod_departamento = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL, 0)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into Activos_departamentos(cod_departamento,descripcion,cod_unidad,registro_usuario,registro_fecha) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & Trim(vGrid.Text) & "','" & glogon.Usuario & "',getdate())"
  

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Departamento : " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update Activos_departamentos set descripcion = '" & vGrid.Text & "',cod_unidad = '"
 vGrid.Col = 3
 strSQL = strSQL & Trim(vGrid.Text) & "',Modifica_Usuario = '" & glogon.Usuario & "', Modifica_Fecha = Getdate()" _
        & " where cod_departamento = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Departamento : " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

If ssTab.Tab = 0 Then
 'Departamentos
    strSQL = "select * from Activos_departamentos" _
          & " order by cod_departamento"
    Call sbCargaGridLocal(vGrid, strSQL, "D")


Else
 'Secciones
    vPaso = True
    strSQL = "select rtrim(cod_departamento) + ' - ' + rtrim(descripcion) as Departamento from Activos_departamentos order by cod_departamento"
    Call OpenRecordSet(rs, strSQL, 0)
    
    cbo.Clear
    Do While Not rs.EOF
     cbo.AddItem rs!departamento
     rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
      rs.MoveFirst
      cbo.Text = rs!departamento
    End If
    rs.Close

    strSQL = "select * from Activos_Secciones where cod_departamento = '" & SIFGlobal.fxCodText(cbo.Text) & "'" _
          & " order by cod_seccion"
    Call sbCargaGridLocal(vGridSec, strSQL, "S")

End If


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

'Descripcion de la Unidad
If vGrid.ActiveCol = 3 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.TextTip = TextTipFixed
  vGrid.TextTipDelay = 1000
  vGrid.CellNote = fxgCntUnidad(vGrid.Text)
End If

'Consulta Unidades
If vGrid.ActiveCol = 3 And KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cod_unidad"
  gBusquedas.Consulta = "select cod_unidad as Unidad,Descripcion from CntX_Unidades"
  gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
  gBusquedas.Orden = "cod_unidad"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow

  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    vGrid.Text = gBusquedas.Resultado
    vGrid.TextTip = TextTipFixed
    vGrid.TextTipDelay = 1000
    vGrid.CellNote = gBusquedas.Resultado2
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
  strSQL = "delete Activos_Departamentos where cod_departamento = '" & vGrid.Text & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Departamento : " & vGrid.Text)
    
  vGrid.DeleteRows vGrid.ActiveRow, 1
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxGuardarSeccion() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardarSeccion = 0

With vGridSec

    .Row = .ActiveRow
    .Col = 1
    
    strSQL = "select coalesce(count(*),0) as Existe from Activos_secciones" _
           & " where cod_seccion = '" & .Text & "' and cod_departamento = '" _
           & SIFGlobal.fxCodText(cbo.Text) & "'"
    Call OpenRecordSet(rs, strSQL, 0)
    
    If rs!Existe = 0 Then 'Insertar
      If Trim(.Text) = "" Then Exit Function
      
      strSQL = "insert into Activos_secciones(cod_departamento,cod_seccion,descripcion,cod_centro_costo,registro_usuario,registro_fecha) values('" _
             & SIFGlobal.fxCodText(cbo.Text) & "','" & UCase(.Text) & "','"
      .Col = 2
      strSQL = strSQL & .Text & "','"
      .Col = 3
      strSQL = strSQL & .Text & "','" & glogon.Usuario & "',getdate())"
    
      Call ConectionExecute(strSQL)
    
      .Col = 1
       Call Bitacora("Registra", "Sección : " & .Text & " - Dept.:" & SIFGlobal.fxCodText(cbo.Text))
    
    Else 'Actualizar
    
     .Col = 2
     strSQL = "update Activos_secciones set descripcion = '" & .Text & "',cod_centro_costo = '"
     .Col = 3
     strSQL = strSQL & .Text & "',Modifica_Usuario = '" & glogon.Usuario & "',Modifica_Fecha = getdate()" _
            & " where cod_departamento = '" & SIFGlobal.fxCodText(cbo.Text) & "' and cod_seccion = '"
     .Col = 1
     strSQL = strSQL & .Text & "'"
     Call ConectionExecute(strSQL)
    
     .Col = 1
      Call Bitacora("Modifica", "Sección : " & .Text & " - Dept.:" & SIFGlobal.fxCodText(cbo.Text))
    
    End If
    rs.Close

End With

fxGuardarSeccion = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vGridSec_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, i As Integer

On Error GoTo vError

If vGridSec.ActiveCol = vGridSec.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardarSeccion
  If i = 0 Then Exit Sub
  vGridSec.Row = vGridSec.ActiveRow
  vGridSec.Col = 1
  If vGridSec.MaxRows <= vGridSec.ActiveRow Then
    vGridSec.MaxRows = vGridSec.MaxRows + 1
    vGridSec.Row = vGridSec.MaxRows
  End If
End If

'Descripcion del Centro de Costos
If vGridSec.ActiveCol = 3 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGridSec.Col = vGridSec.ActiveCol
  vGridSec.Row = vGridSec.ActiveRow
  vGridSec.TextTip = TextTipFixed
  vGridSec.TextTipDelay = 1000
  vGridSec.CellNote = fxgCntCentroCostos(vGridSec.Text)
End If

'Consulta Centro de Costos
If vGridSec.ActiveCol = 3 And KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cod_centro_Costo"
  gBusquedas.Consulta = "select cod_centro_Costo as Centro,Descripcion from cntx_centro_costos"
  gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace & " and Activo = 1"
  gBusquedas.Orden = "cod_centro_Costo"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  
  vGridSec.Col = vGridSec.ActiveCol
  vGridSec.Row = vGridSec.ActiveRow

  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    vGridSec.Text = gBusquedas.Resultado
    vGridSec.TextTip = TextTipFixed
    vGridSec.TextTipDelay = 1000
    vGridSec.CellNote = gBusquedas.Resultado2
  End If
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGridSec.MaxRows = vGridSec.MaxRows + 1
    vGridSec.InsertRows vGridSec.ActiveRow, 1
    vGridSec.Row = vGridSec.ActiveRow
End If

'Borrar Línea
If KeyCode = vbKeyDelete Then
  vGridSec.Row = vGridSec.ActiveRow
  vGridSec.Col = 1
  strSQL = "delete Activos_Secciones where cod_departamento = '" & SIFGlobal.fxCodText(cbo.Text) _
        & "' and cod_seccion = '" & vGridSec.Text & "'"
  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "Sección : " & vGridSec.Text & " - Dept.:" & SIFGlobal.fxCodText(cbo.Text))
    
  vGridSec.DeleteRows vGridSec.ActiveRow, 1
  vGridSec.MaxRows = vGridSec.MaxRows - 1
  If vGridSec.MaxRows = 0 Then vGridSec.MaxRows = 1
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub
