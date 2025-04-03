VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmIVR_Cat_Comisiones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SCGI Tipos de Comisiones"
   ClientHeight    =   6396
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   13788
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6396
   ScaleWidth      =   13788
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5172
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   13572
      _Version        =   524288
      _ExtentX        =   23939
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
      MaxCols         =   488
      ScrollBars      =   2
      SpreadDesigner  =   "frmIVR_Cat_Comisiones.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   492
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   6732
      _Version        =   1310720
      _ExtentX        =   11874
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Tipos de Comisiones"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.8
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
      Left            =   0
      Top             =   0
      Width           =   13932
   End
End
Attribute VB_Name = "frmIVR_Cat_Comisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 22
End Sub

Private Sub sbConsulta()
Dim strSQL As String

 
strSQL = "select * from vIVR_COMISIONES_TIPOS" _
      & " order by COD_COMISION"
Call sbCargaGridLocal(vGrid, 8, strSQL)

End Sub




Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset
Dim i As Integer

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
 
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
     Case 1
       vGrid.Text = rs!COD_COMISION
     
     Case 2
       vGrid.Text = rs!Descripcion
        
     Case 3 'Cuenta
       vGrid.Text = rs!COD_CUENTA_MASK
       
        vGrid.TextTip = TextTipFixed
        vGrid.TextTipDelay = 1000
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = rs!COD_CUENTA_DESC
        
     
     Case 4 'Tipo Valor
       vGrid.Text = rs!Tipo_Valor
     
     Case 5 'Valor
       vGrid.Text = CStr(rs!Valor)
     
     
     Case 6 'Aplicar En
       vGrid.Text = CStr(rs!Aplica_Desc)
     
     Case 7 'Sumar en
       vGrid.Text = CStr(rs!Sumar_Desc)
     
     Case 8 'Activa
       vGrid.Value = rs!ACTIVO
     
     Case Else
    End Select
  Next i
  
    vGrid.Col = vGrid.MaxCols
    vGrid.TextTip = TextTipFixed
    vGrid.TextTipDelay = 1000
    vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
    vGrid.CellNote = "Usuario: " & IIf(IsNull(rs!Registro_Usuario), "...!", rs!Registro_Usuario) _
                     & vbCrLf & "Fecha: " & IIf(IsNull(rs!Registro_Fecha), "...!", rs!Registro_Fecha) _
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop
rs.Close

   
Me.MousePointer = vbDefault

End Sub





Private Sub Form_Load()

vModulo = 22

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
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from IVR_COMISIONES_TIPOS " _
       & " where COD_COMISION = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into IVR_COMISIONES_TIPOS(COD_COMISION,DESCRIPCION, COD_CUENTA, TIPO, VALOR" _
         & ", APLICA_EN, SUMAN_EN, ACTIVO, REGISTRO_USUARIO, REGISTRO_FECHA) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "','"
  vGrid.Col = 4
  strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "',"
  vGrid.Col = 5
  strSQL = strSQL & CDbl(vGrid.Text) & ",'"
  vGrid.Col = 6
  strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "','"
  vGrid.Col = 7
  strSQL = strSQL & UCase(Mid(vGrid.Text, 1, 2)) & "',"
  vGrid.Col = 8
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Tipo de Comisiones:  " & vGrid.Text)

Else 'Actualizar

  vGrid.Col = 2
  strSQL = "update IVR_COMISIONES_TIPOS set descripcion = '" & vGrid.Text & "', COD_CUENTA = '"
  vGrid.Col = 3
  strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "', TIPO = '"
  vGrid.Col = 4
  strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "', VALOR = "
  vGrid.Col = 5
  strSQL = strSQL & CDbl(vGrid.Text) & ", APLICA_EN = '"
  vGrid.Col = 6
  strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "', SUMAN_EN = '"
  vGrid.Col = 7
  strSQL = strSQL & UCase(Mid(vGrid.Text, 1, 2)) & "', ACTIVO = "
  vGrid.Col = 8
  strSQL = strSQL & vGrid.Value & " where COD_COMISION = '"
  vGrid.Col = 1
  strSQL = strSQL & vGrid.Text & "'"
  
  Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Tipo de Comisiones:  " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


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

'Cuenta
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 3 Then
    gCuenta = ""
    frmCntX_ConsultaCuentas.Show vbModal
    If gCuenta <> "" Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 3
        vGrid.Text = fxgCntCuentaFormato(True, gCuenta, 0)
    End If
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete IVR_COMISIONES_TIPOS where COD_COMISION = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Tipo de Comisiones:  " & vGrid.Text)
        
        Call sbConsulta
     
     End If
End If

End Sub
