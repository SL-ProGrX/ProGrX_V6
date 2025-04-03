VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCC_CA_Entidades 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargos Automáticos: Entidades"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   11835
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4932
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   11652
      _Version        =   524288
      _ExtentX        =   20553
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
      MaxCols         =   486
      ScrollBars      =   2
      SpreadDesigner  =   "frmCC_CA_Entidades.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Entidades Autorizadas para Cargo Automático"
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
      Height          =   480
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   7095
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmCC_CA_Entidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 10
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 10

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "select cod_entidad,descripcion,NUMERO_AFILIADO,formato,cod_cuenta,activo from prm_ca_Entidad" _
       & " order by cod_entidad"
Call sbCargaGridLocal(vGrid, 6, strSQL)

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
     Case 1
        vGrid.Text = CStr(rs!cod_entidad)
     Case 2
        vGrid.Text = CStr(rs!Descripcion)
     Case 3
        vGrid.Text = CStr(rs!NUMERO_AFILIADO & "")
     Case 4
        vGrid.Text = CStr(rs!Formato & "")
     Case 5
        vGrid.Text = fxgCntCuentaFormato(True, rs!cod_cuenta & "")
     Case 6
        vGrid.Value = rs!activo
    End Select
  
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

Me.MousePointer = vbDefault

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCuenta As String

On Error GoTo vError

vGrid.Col = 1
fxGuardar = 0

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 5
vCuenta = fxgCntCuentaFormato(False, vGrid.Text)

vGrid.Col = 1


strSQL = "select isnull(count(*),0) as Existe from prm_ca_Entidad" _
       & " where cod_entidad = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  
  strSQL = "insert into prm_ca_Entidad(cod_entidad,descripcion,NUMERO_AFILIADO,formato,cod_cuenta,activo,registro_Fecha,registro_usuario) values('" & vGrid.Text
  vGrid.Col = 2
  strSQL = strSQL & "','" & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Text & "','" & vCuenta & "',"
  vGrid.Col = 6
  strSQL = strSQL & vGrid.Value & ",dbo.mygetdate(),'" & glogon.Usuario & "')"
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  
  Call Bitacora("Registra", "Cargos Automáticos - Entidad: " & vGrid.Text)
  
Else 'Actualizar

    vGrid.Col = 2
    strSQL = "update prm_ca_Entidad set descripcion = '" & vGrid.Text & "',NUMERO_AFILIADO = '"
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Text & "', Formato = '"
    vGrid.Col = 4
    strSQL = strSQL & vGrid.Text & "',cod_cuenta = '" & vCuenta & "',Activo = "
    vGrid.Col = 6
    strSQL = strSQL & vGrid.Value
    vGrid.Col = 1
    strSQL = strSQL & " where cod_entidad = '" & vGrid.Text & "'"
    Call ConectionExecute(strSQL)
 
    vGrid.Col = 1
   
    Call Bitacora("Modifica", "Cargos Automáticos - Entidad: " & vGrid.Text)
 
End If

rs.Close
fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


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

If vGrid.ActiveCol = 5 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = fxgCntCuentaFormato(True, vGrid.Text)
End If


If vGrid.ActiveCol = 5 And KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gCuenta
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
        
        strSQL = "delete prm_ca_Entidad where cod_entidad = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Cargos Automáticos - Entidad: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

