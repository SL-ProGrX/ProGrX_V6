VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCntX_TiposCuentas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipos de Cuentas"
   ClientHeight    =   6120
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9708
   HelpContextID   =   3
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9708
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4812
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   9252
      _Version        =   524288
      _ExtentX        =   16320
      _ExtentY        =   8488
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
      MaxCols         =   494
      ScrollBars      =   2
      SpreadDesigner  =   "frmCntX_TiposCuentas.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Cuentas Contables"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   7452
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmCntX_TiposCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strUltimaSeleccion As String

Private Sub Form_Activate()
vModulo = 20
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 20

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

strUltimaSeleccion = "ACTIVOS"

strSQL = "select tipo_cuenta,descripcion,clasificacion,prioridad from CntX_Tipos_Cuentas" _
       & " where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta & " order by prioridad,descripcion"
Call sbCargaGridLocal(vGrid, 4, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Sub sbCargaComboTiposCuenta(vCol As Integer, vRow As Long, vGrid As Object)
Dim strResultado As String

vGrid.Col = vCol
vGrid.Row = vRow
vGrid.CellType = 8

strResultado = "ACTIVOS"
strResultado = strResultado & Chr$(9) & "PASIVOS"
strResultado = strResultado & Chr$(9) & "CAPITAL/PATRIMONIO"
strResultado = strResultado & Chr$(9) & "INGRESOS"
strResultado = strResultado & Chr$(9) & "GASTOS"
strResultado = strResultado & Chr$(9) & "ORDEN - DEUDORAS"
strResultado = strResultado & Chr$(9) & "ORDEN - ACREEDORAS"

vGrid.TypeComboBoxList = strResultado
vGrid.TypeComboBoxEditable = False
vGrid.Text = strUltimaSeleccion

End Sub

Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

Call sbCargaComboTiposCuenta(3, vGrid.MaxRows, vGrid)

vGrid.Row = vGrid.MaxRows

Call OpenRecordSet(rs, strSQL, 0)
Do While rs.EOF = False
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    If i <> 3 Then
        vGrid.Text = CStr(rs.Fields(i - 1).Value & "")
    Else
        vGrid.Text = fxCntX_TiposCuentas(CStr(rs.Fields(i - 1).Value))
    End If
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  Call sbCargaComboTiposCuenta(3, vGrid.MaxRows, vGrid)
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

End Sub

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

vGrid.Col = 3
strUltimaSeleccion = vGrid.Text

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

rs.Open "select isnull(count(*),0) as Total from CntX_Tipos_Cuentas where tipo_cuenta = '" _
        & vGrid.Text & "' and COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta, glogon.Conection, adOpenStatic

If rs!Total = 0 Then 'Insertar
  strSQL = "insert into CntX_Tipos_Cuentas(tipo_cuenta,COD_CONTABILIDAD,descripcion,clasificacion,prioridad) values('"
  vGrid.Col = 1
  strSQL = strSQL & UCase(vGrid.Text) & "'," & gCntX_Parametros.CodigoConta & ",'"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & fxCntX_TiposCuentas(vGrid.Text) & "','"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Text & "')"
  
  Call ConectionExecute(strSQL, 0)

  vGrid.Col = 2
  
  Call Bitacora("Registra", "Tipo de Cuenta : " & vGrid.Text & " Conta." & gCntX_Parametros.CodigoConta)
  
  fxGuardar = 1

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update CntX_Tipos_Cuentas set descripcion = '" & UCase(vGrid.Text) & "',Clasificacion = '"
 vGrid.Col = 3
 strSQL = strSQL & fxCntX_TiposCuentas(vGrid.Text) & "',Prioridad = '"
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Text & "'" _
        & " where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta & " and tipo_cuenta = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL, 0)
 
 fxGuardar = 1
End If

rs.Close

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
 ' vGrid.Text = i
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    Call sbCargaComboTiposCuenta(3, vGrid.MaxRows, vGrid)
  End If
End If

'Reporte
If KeyCode = vbKeyF5 Then
    Call sbCntX_Reportes_Catalogos("Tipos_Cuentas")
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
        vGrid.Col = 1
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete CntX_Tipos_Cuentas where tipo_cuenta = '" & vGrid.Text _
               & "' and COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta
        Call ConectionExecute(strSQL, 0)
        Call Bitacora("Elimina", "Tipo Cuenta : " & vGrid.Text & " Conta." & gCntX_Parametros.CodigoConta)
        
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
     
     End If
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


