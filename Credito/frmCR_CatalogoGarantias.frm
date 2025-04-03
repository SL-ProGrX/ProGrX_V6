VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmCR_CatalogoGarantias 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Garantías de Crédito"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   15720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   15720
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   15495
      _Version        =   524288
      _ExtentX        =   27331
      _ExtentY        =   8916
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
      MaxCols         =   496
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_CatalogoGarantias.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Garantías"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   4812
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   16455
   End
End
Attribute VB_Name = "frmCR_CatalogoGarantias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Activate()
vModulo = 3
End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
     Case 1
        vGrid.Text = rs!Garantia
     Case 2
        vGrid.Text = rs!Descripcion
     Case 3
        vGrid.Text = rs!formulario
     Case 4
        vGrid.Value = rs!maximos_utiliza
     Case 5
        vGrid.Text = CStr(rs!maximos_monto)
     Case 6
        vGrid.Text = rs!prioridad
     Case 7
        vGrid.Text = rs!Cta_Mask
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = rs!Cta_Desc & ""
        vGrid.TextTip = TextTipFixed
     Case 8
        vGrid.Text = CStr(rs!Porc_Mitigador)
    
     Case 9
        vGrid.Text = CStr(rs!REF_Plazo)
     Case 10
        vGrid.Text = CStr(rs!REF_Tasa)
     Case 11
        vGrid.Value = rs!V_Disponible
    
    
    End Select
  
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

Me.MousePointer = vbDefault

End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 3
vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "select G.*, isnull(Cta.Descripcion,'') as 'Cta_Desc', isnull(Cta.Cod_Cuenta_Mask,'') as 'Cta_Mask'" _
       & " from crd_garantia_tipos G left join vCNTX_CUENTAS_LOCAL Cta on Cta.cod_Cuenta = G.cod_cuenta_incobrable" _
       & " order by G.garantia"
Call sbCargaGridLocal(vGrid, 11, strSQL)


End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from crd_garantia_tipos " _
       & " where garantia = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into crd_garantia_tipos(garantia,descripcion,formulario,maximos_utiliza,maximos_monto,prioridad" _
         & ",cod_cuenta_incobrable, porc_Mitigador, ref_plazo, ref_tasa, v_disponible) values('" _
         & vGrid.Text & "','"
  vGrid.col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.col = 3
  strSQL = strSQL & UCase(vGrid.Text) & "',"
  vGrid.col = 4
  strSQL = strSQL & vGrid.Value & ","
  vGrid.col = 5
  strSQL = strSQL & CCur(IIf(vGrid.Text = "", 0, vGrid.Text)) & ",'"
  vGrid.col = 6
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.col = 7
  strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "',"
  vGrid.col = 8
  strSQL = strSQL & CCur(vGrid.Text) & ","
  vGrid.col = 9
  strSQL = strSQL & CLng(vGrid.Text) & ","
  vGrid.col = 10
  strSQL = strSQL & CCur(vGrid.Text) & ","
  vGrid.col = 11
  strSQL = strSQL & vGrid.Value & ")"
  

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Tipo de Garantía : " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update crd_garantia_tipos set descripcion = '" & vGrid.Text & "',formulario = '"
 vGrid.col = 3
 strSQL = strSQL & UCase(vGrid.Text) & "',maximos_utiliza = "
 vGrid.col = 4
 strSQL = strSQL & vGrid.Value & ",maximos_monto = "
 vGrid.col = 5
 strSQL = strSQL & CCur(IIf(vGrid.Text = "", 0, vGrid.Text)) & ", Prioridad = '"
 vGrid.col = 6
 strSQL = strSQL & vGrid.Text & "',cod_cuenta_incobrable = '"
 vGrid.col = 7
 strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "', porc_Mitigador = "
 vGrid.col = 8
 strSQL = strSQL & CCur(vGrid.Text) & ", ref_Plazo = "
 vGrid.col = 9
 strSQL = strSQL & CCur(vGrid.Text) & ", ref_Tasa = "
 vGrid.col = 10
 strSQL = strSQL & CCur(vGrid.Text) & ", v_disponible = "
 vGrid.col = 11
 strSQL = strSQL & vGrid.Value & " Where garantia = '"
 
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Tipo de Garantía : " & vGrid.Text)

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

'Formato de Cuenta Contable
If vGrid.ActiveCol = 7 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = fxgCntCuentaFormato(True, vGrid.Text)
End If

'Consulta Cuentas Contables
If vGrid.ActiveCol = 7 And KeyCode = vbKeyF4 Then
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
        strSQL = "delete crd_garantia_Tipos where Garantia = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Tipo de Garantía : " & vGrid.Text)
        
        Call Form_Load

     End If
End If


End Sub
