VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Begin VB.Form frmPreaSubExtras 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Expediente : xx"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5532
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   6012
      _Version        =   524288
      _ExtentX        =   10604
      _ExtentY        =   9758
      _StockProps     =   64
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
      SpreadDesigner  =   "frmPreaSubExtras.frx":0000
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   4320
      TabIndex        =   2
      Top             =   7080
      Width           =   1575
      _Version        =   1310722
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0.00"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Totales ..:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rebajo de Extras"
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
      Height          =   492
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   3852
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmPreaSubExtras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsMensajes As New ProGrX_EstudioCrd.clsEstudioMensajes
Private clsEntidad As New ProGrX_EstudioCrd.clsEntidad
Private clsNull As New ProGrX_EstudioCrd.clsNull
Private RebajoExtras As Currency
Private vChanged As Boolean
Dim mTipoExtra As String, mTipoExtraLista As String, vPaso As Boolean


Private Sub sbCalculaTotales()
Dim i As Integer, curMonto As Currency

curMonto = 0

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.Col = 3
  curMonto = curMonto + IIf((vGrid.Text = ""), 0, vGrid.Text)
Next i
 
txtMonto.Text = Format(curMonto, "Standard")

End Sub


Private Sub CargarGrid()
Dim sql As String

On Error GoTo error

sql = "spCRDPreaDETALLE_EXTRAS_TxExpediente " & fxFormatearValor(gPreAnalisis.Expediente, Caracter)
Call sbCargaGridLocal(vGrid, 3, sql)
        
salir:
    Exit Sub

error:
    Call cMensaje.deError("Ocurrió un erro en visual basic al traer la información solicitada. Error " & Err.Description)
End Sub



Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String

Me.MousePointer = vbHourglass

vPaso = True

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1


vGrid.Row = vGrid.MaxRows
rs.CursorLocation = adUseServer

Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  vGrid.Col = 2
  vGrid.CellType = CellTypeComboBox
  vGrid.TypeComboBoxList = mTipoExtraLista
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = mTipoExtra
    
  
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
   Select Case i
     Case 2
        vGrid.Text = CStr(rs!TipoExtra)
     Case Else
        vGrid.Text = CStr(rs.Fields(i - 1).Value) & ""
   End Select
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext
Loop
rs.Close

vGrid.Row = vGrid.MaxRows
  
vPaso = False
  
vGrid.Col = 2
vGrid.CellType = 8
vGrid.TypeComboBoxList = mTipoExtraLista
vGrid.TypeComboBoxEditable = False
vGrid.Text = mTipoExtra
   
Call sbCalculaTotales

Me.MousePointer = vbDefault

End Sub


Public Function fxAgregaColleccion(ByVal Id As Integer, ByVal Expediente As String, ByVal CodExtra As String, ByVal Monto As Double) As String
Dim Vcoleccion As New Collection

On Error GoTo error

With Vcoleccion
    .Add fxFormatearValor(Id, Numerico)
    .Add fxFormatearValor(Expediente, Caracter)
    .Add fxFormatearValor(CodExtra, Caracter)
    .Add fxFormatearValor(Monto, Numerico)
End With

fxAgregaColleccion = fxFormatearValuesCollection(Vcoleccion)

Exit Function

error:
    MsgBox fxSys_Error_Handler(Err.Description)

End Function

Public Function fxAgregaColleccionBorrar(ByVal Id As String, ByVal Expediente As String) As String
Dim Vcoleccion As New Collection

On Error GoTo error

With Vcoleccion
    .Add fxFormatearValor(Id, Numerico)
    .Add fxFormatearValor(Expediente, Caracter)

End With

fxAgregaColleccionBorrar = fxFormatearValuesCollection(Vcoleccion)

Exit Function
error:
    MsgBox fxSys_Error_Handler(Err.Description)
End Function

Private Sub sbNuevoRegistro()
    vGrid.Col = 1
    vGrid.Row = vGrid.MaxRows
    

    vGrid.Col = 2
    vGrid.CellType = CellTypeComboBox
    
    vGrid.TypeComboBoxList = mTipoExtraLista
    vGrid.TypeComboBoxEditable = False
    vGrid.Text = mTipoExtra
    
    vGrid.Action = 0
    vGrid.SetFocus
End Sub

Private Sub SbGuardaRegistro()
    If (vChanged = True) Then
        If (vGrid.Row = vGrid.MaxRows) Then
            If (fxInsertar(vGrid.Row) = True) Then
                vChanged = False
            End If
        Else
            If (fxModificar(vGrid.ActiveRow) = True) Then
                vChanged = False
            End If
        End If
    End If
End Sub

Private Function fxInsertar(ByRef fila As Integer) As Boolean
Dim vID As String
Dim vCodExtra As String
Dim vDesExtra As String
Dim vMonto As Double

On Error GoTo vError

     fxInsertar = False
     If Not ValidaEstadoPreanalisis(gPreAnalisis.ESTADO) Then
      Exit Function
     End If
     
    vGrid.Row = fila
    vID = 0
    vMonto = 0
    
    vGrid.Row = fila
    vGrid.Col = 3
        If Val(vGrid.Text) > 0 Then
            vMonto = CDbl(vGrid.Text)
        Else
            MsgBox "Monto es requerido.", vbExclamation, gMsgTitulo
            Me.MousePointer = vbDefault
            Exit Function
        End If
    
    vGrid.Col = 2
    If vGrid.Text = "" Then
       vCodExtra = SIFGlobal.fxCodText(mTipoExtra)
    Else
       vCodExtra = SIFGlobal.fxCodText(vGrid.Text)
       mTipoExtra = vGrid.Text
    End If
    
    If Len(vCodExtra) = 0 Then
        MsgBox "Es requerido seleccionar una tipo de extra.", vbExclamation, gMsgTitulo
        Me.MousePointer = vbDefault
        Exit Function
    End If

    
 Me.MousePointer = vbHourglass
    
    clsEntidad.tablaName = "spCRDPreaDETALLE_EXTRAS"
    If (clsEntidad.fxAgregar(fxAgregaColleccion(vID, gPreAnalisis.Expediente, vCodExtra, vMonto))) Then
        fxInsertar = True
        vGrid.Col = 1
        vGrid.Lock = True
        vGrid.MaxRows = vGrid.MaxRows + 1
       
        glogon.strSQL = "select max(IDX) as IDX from CRD_PREA_DETALLE_EXTRAS  Where cod_preanalisis = " & fxFormatearValor(gPreAnalisis.Expediente, Caracter)
        If (execSql(glogon.strSQL, True)) Then
            vGrid.Col = 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = glogon.Recordset!IdX
        Else
            MsgBox "Error obteniendo el consecutivo del registro ingresado", vbExclamation, gMsgTitulo
        End If
    End If
    
Call sbCalculaTotales
    
Me.MousePointer = vbDefault

Exit Function

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo
End Function

Private Function fxModificar(fila As Long) As Boolean
Dim vID As Integer
Dim vCodExtra As String
Dim vNuevaDesExtra As String
Dim vMonto As Double

On Error GoTo vError

    
fxModificar = False
 If Not ValidaEstadoPreanalisis(gPreAnalisis.ESTADO) Then
  Exit Function
 End If

vGrid.Row = fila
vGrid.Col = 1
    vID = vGrid.Text
vGrid.Col = 3
    vMonto = vGrid.Text

vGrid.Col = 2
If vGrid.Text = "" Then
    vCodExtra = SIFGlobal.fxCodText(mTipoExtra)
Else
    vCodExtra = SIFGlobal.fxCodText(vGrid.Text)
End If



Me.MousePointer = vbHourglass

clsEntidad.tablaName = "spCRDPreaDETALLE_EXTRAS"

If (clsEntidad.fxModificar(fxAgregaColleccion(vID, gPreAnalisis.Expediente, vCodExtra, vMonto))) Then
    vGrid.Col = 1
    fxModificar = True
Else
    MsgBox "No se pudo actualizar la información seleccionada.", vbExclamation, gMsgTitulo
End If

Me.MousePointer = vbDefault
Exit Function

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo

End Function

Private Function fxBorrar(fila As Long) As Boolean
Dim vID As Integer
On Error GoTo vError
 Me.MousePointer = vbHourglass
 
    fxBorrar = False
     If Not ValidaEstadoPreanalisis(gPreAnalisis.ESTADO) Then
      GoTo salir
     End If

    vGrid.Col = 1
    vGrid.Row = fila
    vID = Val(vGrid.Text)
    
    clsEntidad.tablaName = "spCRDPreaDETALLE_EXTRAS"
    If (MsgBox("¿ Desea borrar la información seleccionada?", vbQuestion + vbYesNo, gMsgTitulo) = vbYes) Then
        If clsEntidad.fxRemover(fxAgregaColleccionBorrar(vID, gPreAnalisis.Expediente)) Then
            vGrid.Col = 1
            fxBorrar = True
            
                vGrid.Action = 5
                vGrid.MaxRows = vGrid.MaxRows - 1
                vGrid.Row = fila
                vGrid.Col = 1
                vGrid.Action = 0
                
                Call sbCalculaTotales
        End If
        vGrid.SetFocus
    End If

salir:
    Me.MousePointer = vbDefault
    Exit Function
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo
    
    Resume salir
End Function


Private Sub vGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim sql As String
On Error GoTo error
    If ((Shift And vbCtrlMask) <> 0) And (KeyCode = vbKeyS) Then
        UnLoad Me
        DoEvents
        
    ElseIf ((Shift And vbCtrlMask) <> 0) And (KeyCode = vbKeyN) Then
        Call sbNuevoRegistro
            
    ElseIf ((Shift And vbCtrlMask) <> 0) And (KeyCode = vbKeyG) Then
        Call SbGuardaRegistro
            
    ElseIf (KeyCode = vbKeyDelete) Then
        Call fxBorrar(vGrid.ActiveRow)
    End If
    
salir:
    Exit Sub
error:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

'Carga Listado de Tipos de Extras
mTipoExtra = ""
mTipoExtraLista = ""


Me.Caption = "Expediente: " & gPreAnalisis.Expediente

Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

strSQL = "select rtrim(cod_Extras) + ' - ' + rtrim(descripcion) as 'TipoExtra'" _
       & " from CRD_PREA_TIPOS_EXTRAS order by cod_Extras"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And mTipoExtra = "" Then mTipoExtra = rs!TipoExtra

Do While Not rs.EOF
  If Len(mTipoExtraLista) = 0 Then
    mTipoExtraLista = Chr$(9) & rs!TipoExtra
  Else
    mTipoExtraLista = mTipoExtraLista & Chr$(9) & rs!TipoExtra
  End If
  rs.MoveNext
Loop
rs.Close


vChanged = False
Call CargarGrid

End Sub

Private Sub Form_Unload(Cancel As Integer)

GLOBALES.gTag = txtMonto.Text

End Sub


Private Sub vGrid_Advance(ByVal AdvanceNext As Boolean)
Dim fila As Integer
Dim wColValue As Variant

wColValue = vGrid.Text

fila = vGrid.ActiveRow

If (vGrid.MaxRows > 1) And (vGrid.ActiveRow = 1) Then
    Exit Sub
ElseIf (vGrid.ActiveRow = vGrid.MaxRows) Then
    vGrid.Col = 1
    vGrid.Row = vGrid.ActiveRow
    If (fxInsertar(fila) = True) Then
        vGrid.Col = 1
        vGrid.Row = vGrid.MaxRows
        vGrid.Action = 0
        vChanged = False
        
        vGrid.Col = 2
        vGrid.CellType = CellTypeComboBox
        
        vGrid.TypeComboBoxList = mTipoExtraLista
        vGrid.TypeComboBoxEditable = False
        vGrid.Text = mTipoExtra
    End If
    
End If

End Sub

Private Sub vGrid_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    vChanged = True
End Sub

Private Sub vGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

On Error GoTo vError

If (NewCol = -1) Then Exit Sub
    If (vChanged = True) Then
        If (Row = NewRow) Then
            Exit Sub
        Else
            If (Row = vGrid.MaxRows) Then
                If (fxInsertar(Val(Row)) = True) Then
                    vChanged = False
                End If
            Else
                If fxModificar(Row) Then
                    vChanged = False
                End If
            End If
        End If
    End If

vError:

End Sub
