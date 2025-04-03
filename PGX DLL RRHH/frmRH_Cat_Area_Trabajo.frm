VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.ShortcutBar.v20.2.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmRH_Cat_Area_Trabajo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RRHH: Centro, Departamento y Secciones"
   ClientHeight    =   9420
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   14805
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   14805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   8055
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   14535
      _Version        =   1310722
      _ExtentX        =   25638
      _ExtentY        =   14208
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Centro de Trabajo"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Departamentos y Secciones"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "vgDept"
      Item(1).Control(1)=   "scDepartamentos"
      Item(1).Control(2)=   "scSecciones"
      Item(1).Control(3)=   "vgSecc"
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   7332
         Left            =   3360
         TabIndex        =   2
         Top             =   480
         Width           =   8532
         _Version        =   524288
         _ExtentX        =   15049
         _ExtentY        =   12933
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
         MaxCols         =   484
         ScrollBars      =   2
         SpreadDesigner  =   "frmRH_Cat_Area_Trabajo.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vgDept 
         Height          =   6852
         Left            =   -69880
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   7212
         _Version        =   524288
         _ExtentX        =   12721
         _ExtentY        =   12086
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "frmRH_Cat_Area_Trabajo.frx":05D1
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vgSecc 
         Height          =   6852
         Left            =   -62440
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   6972
         _Version        =   524288
         _ExtentX        =   12298
         _ExtentY        =   12086
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   4
         ScrollBars      =   2
         SpreadDesigner  =   "frmRH_Cat_Area_Trabajo.frx":0C24
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption scSecciones 
         Height          =   372
         Left            =   -62680
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   7212
         _Version        =   1310722
         _ExtentX        =   12721
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Secciones para Departamento: "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption scDepartamentos 
         Height          =   372
         Left            =   -69880
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   7212
         _Version        =   1310722
         _ExtentX        =   12721
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Departamentos para Centro: "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   492
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   6732
      _Version        =   1310722
      _ExtentX        =   11874
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Areas de Trabajo"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   -240
      Top             =   0
      Width           =   17772
   End
End
Attribute VB_Name = "frmRH_Cat_Area_Trabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub Form_Activate()
vModulo = 23
End Sub

Private Sub sbCargaGridLocal(ByRef pGrid As Object, strSQL As String, Optional pTipo As String = "D")
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String
Dim strUltimaSeleccion As String



Me.MousePointer = vbHourglass

On Error GoTo vError

pGrid.MaxRows = 0
pGrid.MaxRows = 1
pGrid.Row = pGrid.MaxRows

Call OpenRecordSet(rs, strSQL, 0)

With pGrid
Do While Not rs.EOF
  .Row = pGrid.MaxRows
  .Col = 1
  
  If pTipo = "D" Then
    'Departamentos
    For i = 1 To 5
      .Col = i
      Select Case i
       Case 2 'Codigo
          .Text = rs!Cod_Departamento
          .TextTip = TextTipFixed
          .TextTipDelay = 1000
          .CellNote = "Registrado: " & rs!Registro_Usuario & vbCrLf & "Fecha: " & rs!Registro_Fecha & vbCrLf & vbCrLf
       Case 3 'Descripcion
          .Text = rs!Descripcion
      
       Case 4 'Unidad
          .Text = rs!cod_unidad
          .TextTip = TextTipFixed
          .TextTipDelay = 1000
          .CellNote = fxgCntUnidad(rs!cod_unidad)
      
      
       Case 5 'Activo
          .Value = rs!ACTIVO
      End Select
    Next i
  
  
  Else
   'Secciones
      For i = 1 To 4
      .Col = i
      Select Case i
       Case 1 'Codigo
          .Text = rs!Cod_Seccion
          .TextTip = TextTipFixed
          .TextTipDelay = 1000
          .CellNote = "Registrado: " & rs!Registro_Usuario & vbCrLf & "Fecha: " & rs!Registro_Fecha & vbCrLf & vbCrLf
'                    & "Modificado: " & rs!Modifica_Usuario & vbCrLf & "Fecha: " & rs!Modifica_Fecha
       Case 2 'Descripcion
          .Text = rs!Descripcion
       Case 3 'Centro de Costo
          .Text = rs!cod_centro_costo
          .TextTip = TextTipFixed
          .TextTipDelay = 1000
          .CellNote = fxgCntCentroCostos(rs!cod_centro_costo)
       Case 4 'Activo
          .Value = rs!ACTIVO
       
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



Private Sub sbConsulta()
Dim strSQL As String

vPaso = True

tcMain.Item(0).Selected = True

scDepartamentos.Tag = ""
scDepartamentos.Caption = "Departamentos para el Centro: "
vgDept.MaxRows = 0

scSecciones.Tag = ""
scSecciones.Caption = "Secciones para el Departamento: "
vgSecc.MaxRows = 0

strSQL = "select 0, COD_CENTRO,descripcion,Activo" _
      & " from RH_CENTRO_TRABAJO" _
      & " order by COD_CENTRO"
Call sbCargaGrid(vGrid, 4, strSQL)

vPaso = False

End Sub


Private Sub sbDepartamentos()
Dim strSQL As String

vPaso = True

tcMain.Item(1).Selected = True

scSecciones.Tag = ""
scSecciones.Caption = "Secciones para el Departamento: "
vgSecc.MaxRows = 0

strSQL = "select *" _
      & " from RH_DEPARTAMENTOS Where COD_CENTRO = '" & scDepartamentos.Tag & "'" _
      & " order by COD_DEPARTAMENTO"
Call sbCargaGridLocal(vgDept, strSQL, "D")

vPaso = False

End Sub

Private Sub sbSecciones()
Dim strSQL As String

vPaso = True

tcMain.Item(1).Selected = True


strSQL = "select *" _
      & " from RH_SECCIONES" _
      & " Where COD_CENTRO = '" & scDepartamentos.Tag & "' and COD_DEPARTAMENTO = '" & scSecciones.Tag _
      & "' order by COD_SECCION"
Call sbCargaGridLocal(vgSecc, strSQL, "S")

vPaso = False

End Sub


Private Sub Form_Load()

vModulo = 23

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call sbConsulta

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
vGrid.Col = 2

strSQL = "select isnull(count(*),0) as Existe from RH_CENTRO_TRABAJO " _
       & " where COD_CENTRO = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into RH_CENTRO_TRABAJO(COD_CENTRO,DESCRIPCION,ACTIVO, REGISTRO_USUARIO, REGISTRO_FECHA) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  vGrid.Col = 2
  Call Bitacora("Registra", "Centro de Trabajo: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 3
 strSQL = "update RH_CENTRO_TRABAJO set descripcion = '" & vGrid.Text & "',Activo = "
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Value & " where COD_CENTRO = '"
 vGrid.Col = 2
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Centro de Trabajo: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vgDept_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Or Col = 5 Then Exit Sub

With vgDept
    .Row = Row
    .Col = 2
    scSecciones.Tag = .Text
    .Col = 3
    scSecciones.Caption = "Secciones para el Departamento: " & .Text
End With

Call sbSecciones

End Sub

Private Function fxGuardarDept() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

With vgDept

fxGuardarDept = 0
.Row = .ActiveRow
.Col = 2

strSQL = "select isnull(count(*),0) as Existe from RH_departamentos" _
       & " where cod_centro = '" & scDepartamentos.Tag & "' and cod_departamento = '" & .Text & "'"
Call OpenRecordSet(rs, strSQL, 0)

If rs!Existe = 0 Then 'Insertar
  If Trim(.Text) = "" Then Exit Function
  
  strSQL = "insert into RH_departamentos(cod_centro,cod_departamento,descripcion,cod_unidad,Activo,registro_usuario,registro_fecha) values('" _
         & scDepartamentos.Tag & "','" & .Text & "','"
  .Col = 3
  strSQL = strSQL & .Text & "','"
  .Col = 4
  strSQL = strSQL & Trim(.Text) & "',"
  .Col = 5
  strSQL = strSQL & .Value & ",'" & glogon.Usuario & "',dbo.myGetdate())"
  

  glogon.Conection.Execute strSQL

  .Col = 2
  Call Bitacora("Registra", "Departamento: " & .Text & " -Centro: " & scDepartamentos.Tag)

Else 'Actualizar

 .Col = 3
 strSQL = "update RH_departamentos set descripcion = '" & .Text & "',cod_unidad = '"
 .Col = 4
 strSQL = strSQL & Trim(.Text) & "', Activo = "
 .Col = 5
 strSQL = strSQL & .Value & " where cod_Centro = '" & scDepartamentos.Tag & "' and cod_departamento = '"
 .Col = 2
 strSQL = strSQL & .Text & "'"
 
 glogon.Conection.Execute strSQL

 .Col = 2
 Call Bitacora("Modifica", "Departamento: " & .Text & " -Centro: " & scDepartamentos.Tag)

End If
rs.Close

fxGuardarDept = 1

End With

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub vgDept_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, i As Integer

On Error GoTo vError

With vgDept

If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardarDept
  If i = 0 Then Exit Sub
  .Row = .ActiveRow
  .Col = 1
  If .MaxRows <= .ActiveRow Then
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows
  End If
End If

'Descripcion de la Unidad
If .ActiveCol = 4 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  .Col = .ActiveCol
  .Row = .ActiveRow
  .TextTip = TextTipFixed
  .TextTipDelay = 1000
  .CellNote = fxgCntUnidad(.Text)
End If

'Consulta Unidades
If .ActiveCol = 4 And KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cod_unidad"
  gBusquedas.Consulta = "select cod_unidad as Unidad,Descripcion from CntX_Unidades"
  gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
  gBusquedas.Orden = "cod_unidad"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  
  .Col = .ActiveCol
  .Row = .ActiveRow

  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    .Text = gBusquedas.Resultado
    .TextTip = TextTipFixed
    .TextTipDelay = 1000
    .CellNote = gBusquedas.Resultado2
  End If

End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    .MaxRows = .MaxRows + 1
    .InsertRows .ActiveRow, 1
    .Row = .ActiveRow
End If

'Borrar Línea
If KeyCode = vbKeyDelete Then
  .Row = .ActiveRow
  .Col = 2
  strSQL = "delete RH_Departamentos" _
         & " where cod_departamento = '" & .Text & "' and Cod_Centro = '" & scDepartamentos.Tag & "'"
  glogon.Conection.Execute strSQL
  
  Call Bitacora("Elimina", "Departamento: " & .Text & " -Centro: " & scDepartamentos.Tag)
    
  .DeleteRows .ActiveRow, 1
  .MaxRows = .MaxRows - 1
  If .MaxRows = 0 Then .MaxRows = 1
End If

End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Or Col = 4 Then Exit Sub

vGrid.Row = Row
vGrid.Col = 2

scDepartamentos.Tag = vGrid.Text
vGrid.Col = 3
scDepartamentos.Caption = "Departamentos para el Centro: " & vGrid.Text

Call sbDepartamentos

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String


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

'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 2
        strSQL = "delete RH_CENTRO_TRABAJO where COD_CENTRO = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 2
        Call Bitacora("Elimina", "Centro de Trabajo: " & vGrid.Text)
        
        Call sbConsulta
     
     End If
End If


End Sub




Private Function fxGuardarSeccion() As Long
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

fxGuardarSeccion = 0

With vgSecc

    .Row = .ActiveRow
    .Col = 1
    
    strSQL = "select isnull(count(*),0) as Existe" _
           & " from RH_Secciones" _
           & " where cod_seccion = '" & .Text _
           & "'  and cod_departamento = '" & scSecciones.Tag _
           & "'  and cod_Centro = '" & scDepartamentos.Tag & "'"
    Call OpenRecordSet(rs, strSQL, 0)
    
    If rs!Existe = 0 Then 'Insertar
      If Trim(.Text) = "" Then Exit Function
      
      strSQL = "insert into RH_Secciones(cod_Centro, cod_Departamento, cod_Seccion,descripcion,cod_centro_costo,Activo, registro_usuario,registro_fecha) values('" _
             & scDepartamentos.Tag & "','" & scSecciones.Tag & "','" & .Text & "','"
      .Col = 2
      strSQL = strSQL & .Text & "','"
      .Col = 3
      strSQL = strSQL & .Text & "',"
      .Col = 4
      strSQL = strSQL & .Value & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
    
      glogon.Conection.Execute strSQL
    
      .Col = 1
       Call Bitacora("Registra", "Sección: " & .Text & " -Dept.:" & scSecciones.Tag & " -Centro: " & scDepartamentos.Tag)
    
    Else 'Actualizar
    
     .Col = 2
     strSQL = "update RH_Secciones set descripcion = '" & .Text & "',cod_centro_costo = '"
     .Col = 3
     strSQL = strSQL & .Text & "', Activo = "
     .Col = 4
     strSQL = strSQL & .Value & " where cod_Centro = '" & scDepartamentos.Tag _
            & "' and cod_departamento = '" & scSecciones.Tag & "' and cod_seccion = '"
     .Col = 1
     strSQL = strSQL & .Text & "'"
     
     glogon.Conection.Execute strSQL
    
     .Col = 1
      Call Bitacora("Modifica", "Sección: " & .Text & " -Dept.:" & scSecciones.Tag & " -Centro: " & scDepartamentos.Tag)
    
    End If
    rs.Close

End With

fxGuardarSeccion = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub vgSecc_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, i As Integer

On Error GoTo vError

With vgSecc

If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardarSeccion
  If i = 0 Then Exit Sub
  .Row = .ActiveRow
  .Col = 1
  If .MaxRows <= .ActiveRow Then
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows
  End If
End If

'Descripcion del Centro de Costos
If .ActiveCol = 3 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  .Col = .ActiveCol
  .Row = .ActiveRow
  .TextTip = TextTipFixed
  .TextTipDelay = 1000
  .CellNote = fxgCntCentroCostos(.Text)
End If

'Consulta Centro de Costos
If .ActiveCol = 3 And KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cod_centro_Costo"
  gBusquedas.Consulta = "select cod_centro_Costo as Centro,Descripcion from cntx_centro_costos"
  gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace & " and Activo = 1"
  gBusquedas.Orden = "cod_centro_Costo"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  
  .Col = .ActiveCol
  .Row = .ActiveRow

  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    .Text = gBusquedas.Resultado
    .TextTip = TextTipFixed
    .TextTipDelay = 1000
    .CellNote = gBusquedas.Resultado2
  End If
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    .MaxRows = .MaxRows + 1
    .InsertRows .ActiveRow, 1
    .Row = .ActiveRow
End If

 

'Borrar Línea
If KeyCode = vbKeyDelete Then
  .Row = .ActiveRow
  .Col = 1
  strSQL = "delete RH_SECCIONES" _
        & " where cod_centro = '" & scDepartamentos.Tag & "' and  cod_departamento = '" & scSecciones.Tag _
        & "' and cod_seccion = '" & .Text & "'"
  glogon.Conection.Execute strSQL

  Call Bitacora("Elimina", "Sección: " & .Text & " -Dept.: " & scSecciones.Tag & "-Centro: " & scDepartamentos.Tag)
    
  .DeleteRows .ActiveRow, 1
  .MaxRows = .MaxRows - 1
  If .MaxRows = 0 Then .MaxRows = 1
End If


End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

