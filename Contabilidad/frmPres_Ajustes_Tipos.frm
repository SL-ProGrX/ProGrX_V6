VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmPres_Ajustes_Tipos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipos de Ajustes"
   ClientHeight    =   7236
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   12684
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7236
   ScaleWidth      =   12684
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5892
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   12372
      _Version        =   524288
      _ExtentX        =   21823
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
      MaxCols         =   7
      ScrollBars      =   2
      SpreadDesigner  =   "frmPres_Ajustes_Tipos.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Ajustes Presupuestarios"
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
      Width           =   7452
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmPres_Ajustes_Tipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 12
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 12

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

strSQL = "select cod_ajuste,descripcion,ACTIVO,ajuste_libre_positivo, ajuste_libre_negativo, ajuste_entre_cuentas, ajuste_cta_dif_Naturaleza" _
       & " from pres_tipos_ajustes" _
       & " order by descripcion"
Call sbCargaGrid(vGrid, 7, strSQL, True)

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



strSQL = "select isnull(count(*),0) as Existe from pres_tipos_ajustes where cod_ajuste = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)


If rs!Existe = 0 Then 'Insertar
  strSQL = "insert into pres_tipos_ajustes(cod_ajuste,descripcion,ACTIVO,ajuste_libre_positivo, ajuste_libre_negativo, ajuste_entre_cuentas, ajuste_cta_dif_Naturaleza) values('"
  vGrid.Col = 1
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 6
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 7
  strSQL = strSQL & vGrid.Value & ")"
  
  Call ConectionExecute(strSQL, 0)

  vGrid.Col = 2
  
  Call Bitacora("Registra", "Tipo Ajuste Presupuetario: " & vGrid.Text)
  
  fxGuardar = 1

Else 'Actualizar
'cod_ajuste,descripcion,ACTIVO,ajuste_libre_positivo, ajuste_libre_negativo, ajuste_entre_cuentas, ajuste_cta_dif_Naturaleza
 vGrid.Col = 2
 strSQL = "update pres_tipos_ajustes set descripcion = '" & vGrid.Text & "',Activo = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & ",ajuste_libre_positivo = "
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Value & ",ajuste_libre_negativo = "
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Value & ",ajuste_entre_cuentas = "
 vGrid.Col = 6
 strSQL = strSQL & vGrid.Value & ",ajuste_cta_dif_Naturaleza = "
 vGrid.Col = 7
 strSQL = strSQL & vGrid.Value & " where cod_Ajuste = '"
 
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL, 0)
 
 
 Call Bitacora("Actualiza", "Tipo Ajuste Presupuetario: " & vGrid.Text)
 
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
'  vGrid.Text = i
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
        vGrid.Col = 1
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete pres_tipos_ajustes where cod_ajuste = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL, 0)
        
         Call Bitacora("Elimina", "Tipo Ajuste Presupuetario: " & vGrid.Text)
        
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
     
     End If
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




