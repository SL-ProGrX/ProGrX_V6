VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCxC_CargosTipos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Cargos/Ingresos"
   ClientHeight    =   6516
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   10284
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6516
   ScaleWidth      =   10284
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4932
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   10092
      _Version        =   524288
      _ExtentX        =   17801
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
      MaxCols         =   485
      ScrollBars      =   2
      SpreadDesigner  =   "frmCxC_CargosTipos.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Cargos e Ingresos"
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
      Height          =   480
      Index           =   0
      Left            =   2004
      TabIndex        =   1
      Top             =   360
      Width           =   6852
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmCxC_CargosTipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 31
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 31

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "select cod_cargo,descripcion,case when Tipo = 'I' then 'Ingreso' else 'Cargo' end as 'Tipo',cod_cuenta,activo from cxc_cargos" _
       & " order by cod_cargo"
Call sbCargaGridLocal(vGrid, 5, strSQL)

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
    vGrid.col = i
    Select Case i
     Case 1
        vGrid.Text = CStr(rs!COD_CARGO)
     Case 2
        vGrid.Text = CStr(rs!Descripcion)
     Case 3
        vGrid.Text = CStr(rs!Tipo)
     Case 4
        vGrid.Text = fxgCntCuentaFormato(True, rs!cod_cuenta)
     Case 5
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
Dim vCuenta As String, vTipo As String

On Error GoTo vError

vGrid.col = 1
fxGuardar = 0

vGrid.Row = vGrid.ActiveRow
vGrid.col = 3
vTipo = Mid(vGrid.Text, 1, 1)
vGrid.col = 4
vCuenta = fxgCntCuentaFormato(False, vGrid.Text)

vGrid.col = 1


strSQL = "select isnull(count(*),0) as Existe from cxc_cargos" _
       & " where cod_cargo = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  
  strSQL = "insert into cxc_cargos(cod_cargo,descripcion,Tipo,cod_cuenta,activo) values('" & vGrid.Text
  vGrid.col = 2
  strSQL = strSQL & "','" & vGrid.Text & "','" & vTipo & "','" & vCuenta & "',"
  vGrid.col = 5
  strSQL = strSQL & vGrid.Value & ")"
  Call ConectionExecute(strSQL)

  vGrid.col = 1
  
  Call Bitacora("Registra", "Tipo de Cargo de CxC: " & vGrid.Text)
  
Else 'Actualizar

    vGrid.col = 2
    strSQL = "update cxc_cargos set descripcion = '" & vGrid.Text & "', Tipo = '" & vTipo _
           & "',cod_cuenta = '" & vCuenta & "',Activo = "
    vGrid.col = 5
    strSQL = strSQL & vGrid.Value
    vGrid.col = 1
    strSQL = strSQL & " where cod_cargo = '" & vGrid.Text & "'"
    Call ConectionExecute(strSQL)
 
    vGrid.col = 1
   
    Call Bitacora("Modifica", "Tipo de Cargo de CxC: " & vGrid.Text)
 
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
  vGrid.col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

If vGrid.ActiveCol = 4 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = fxgCntCuentaFormato(True, vGrid.Text)
End If


If vGrid.ActiveCol = 4 And KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  vGrid.col = vGrid.ActiveCol
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
        vGrid.col = 1
        
        strSQL = "delete cxc_cargos where cod_cargo = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Tipo de Cargo de CxC: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
