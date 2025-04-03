VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmActivos_Proveedores 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Listado de Proveedores"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   7920
      Top             =   240
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5652
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   8892
      _Version        =   524288
      _ExtentX        =   15684
      _ExtentY        =   9970
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
      SpreadDesigner  =   "frmActivos_Proveedores.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton btnImportar 
      Height          =   492
      Left            =   7440
      TabIndex        =   2
      ToolTipText     =   "Importa Catálogo de Cuentas por Pagar"
      Top             =   600
      Width           =   1572
      _Version        =   1441792
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Importar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmActivos_Proveedores.frx":0612
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedores"
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
      Height          =   372
      Left            =   1800
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
Attribute VB_Name = "frmActivos_Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vProveedor As String

Private Sub sbGrid_Load()
Dim strSQL As String

Me.MousePointer = vbHourglass

strSQL = "select * from Activos_proveedores" _
      & " order by cod_Proveedor"
Call sbCargaGridLocal(vGrid, strSQL)

Me.MousePointer = vbDefault

End Sub

Private Sub btnImportar_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "insert into ACTIVOS_PROVEEDORES(COD_PROVEEDOR,DESCRIPCION, ACTIVO, REGISTRO_FECHA, REGISTRO_USUARIO)" _
       & "  ( select convert(varchar(20), cod_proveedor), DESCRIPCION, 1, getdate(),'" & glogon.Usuario & "'" _
       & "      from CXP_PROVEEDORES" _
       & "     where convert(varchar(20), cod_proveedor) not in(select COD_PROVEEDOR  from ACTIVOS_PROVEEDORES)" _
       & "       and ESTADO = 'A'" _
       & "   )"

Call ConectionExecute(strSQL)

Call Form_Load

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
vModulo = 36

End Sub

Private Sub Form_Load()


vModulo = 36

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

strSQL = "select coalesce(count(*),0) as Existe from Activos_proveedores " _
       & " where cod_Proveedor = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL, 0)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into Activos_proveedores(cod_Proveedor,descripcion,activo,registro_usuario,registro_fecha) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',getdate())"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Proveedor : " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update Activos_proveedores set descripcion = '" & vGrid.Text & "',Activo = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & ",modifica_usuario = '" & glogon.Usuario & "',modifica_fecha = getdate()" _
        & " where cod_Proveedor = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Modifica", "Proveedor : " & vGrid.Text)

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

Call OpenRecordSet(rs, strSQL, 0)

With pGrid
Do While Not rs.EOF
  .Row = pGrid.MaxRows
  .Col = 1
  
    For i = 1 To 3
      .Col = i
      Select Case i
       Case 1 'Codigo
          .Text = rs!COD_PROVEEDOR
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



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

On Error GoTo vError

strSQL = "insert into ACTIVOS_PROVEEDORES(COD_PROVEEDOR,DESCRIPCION, ACTIVO, REGISTRO_FECHA, REGISTRO_USUARIO)" _
       & "  ( select convert(varchar(20), cod_proveedor), DESCRIPCION, 1, getdate(),'" & glogon.Usuario & "'" _
       & "      from CXP_PROVEEDORES" _
       & "     where convert(varchar(20), cod_proveedor) not in(select COD_PROVEEDOR  from ACTIVOS_PROVEEDORES)" _
       & "       and ESTADO = 'A'" _
       & "   )"

Call ConectionExecute(strSQL)

Call Form_Load

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbGrid_Load
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
  strSQL = "delete Activos_Proveedores where cod_Proveedor = '" & vGrid.Text & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Proveedor : " & vGrid.Text)
    
  vGrid.DeleteRows vGrid.ActiveRow, 1
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub






