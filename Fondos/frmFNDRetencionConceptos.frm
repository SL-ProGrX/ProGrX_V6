VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmFNDRetencionConceptos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conceptos de Retenciones"
   ClientHeight    =   7092
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   9108
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7092
   ScaleWidth      =   9108
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5532
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   8532
      _Version        =   524288
      _ExtentX        =   15049
      _ExtentY        =   9758
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
      SpreadDesigner  =   "frmFNDRetencionConceptos.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Conceptos para Retención de ahorros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   3
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   6372
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10932
   End
End
Attribute VB_Name = "frmFNDRetencionConceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 18
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 18
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

strSQL = "select C.RETENCION_CODIGO,C.descripcion,C.Activo" _
       & ",C.cod_Cuenta,CntX.cod_Cuenta_Mask,CntX.descripcion as CtaDesc" _
       & " from FND_RETENCION_CONCEPTOS C left join CntX_cuentas CntX on CntX.cod_Cuenta = C.cod_cuenta and CntX.cod_contabilidad = " & GLOBALES.gEnlace _
       & " order by C.RETENCION_CODIGO"
Call sbCargaGridLocal(vGrid, 4, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)


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
    vGrid.Col = i
    Select Case i
     Case 1
        vGrid.Text = rs!RETENCION_CODIGO
     Case 2
        vGrid.Text = rs!Descripcion
     Case 3
        vGrid.Value = rs!activo
     Case 4
        vGrid.Text = rs!cod_Cuenta_Mask & ""
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = rs!CtaDesc & ""
        vGrid.TextTip = TextTipFixed
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
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If Trim(vGrid.Text) = "" Then
   MsgBox "No se especifico ningún código....verifique..!!!", vbExclamation
   Exit Function
End If

strSQL = "select isnull(count(*),0) as Existe from FND_RETENCION_CONCEPTOS where RETENCION_CODIGO = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  vGrid.Col = 1
  strSQL = "insert into FND_RETENCION_CONCEPTOS(RETENCION_CODIGO,descripcion,Activo,cod_cuenta) values('"
  strSQL = strSQL & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",'"
  vGrid.Col = 4
  strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "')"
  

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Concepto de Retención de Planes: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update FND_RETENCION_CONCEPTOS set descripcion = '" & vGrid.Text & "',Activo = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & ",cod_cuenta = '"
 vGrid.Col = 4
 strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "' where RETENCION_CODIGO = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 
 Call Bitacora("Modifica", "Concepto de Retención de Planes: " & vGrid.Text)

End If
rs.Close
fxGuardar = 1


Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

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
If vGrid.ActiveCol = 4 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = fxgCntCuentaFormato(True, vGrid.Text)
End If

'Consulta Cuentas Contables
If vGrid.ActiveCol = 4 And KeyCode = vbKeyF4 Then
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

'Borrar una linea
If KeyCode = vbKeyDelete Then

        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1

       If vGrid.Text = "" Then Exit Sub

     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Col = 1
        strSQL = "delete FND_RETENCION_CONCEPTOS where RETENCION_CODIGO = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        
        Call Bitacora("Elimina", "Concepto de Retención de Planes: " & vGrid.Text)
        
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
        
     End If
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
