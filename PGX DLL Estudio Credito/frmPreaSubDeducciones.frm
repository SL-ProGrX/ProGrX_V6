VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Begin VB.Form frmPreaSubDeducciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Expediente : xx"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnImport 
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   6360
      Width           =   4815
      _Version        =   1310720
      _ExtentX        =   8493
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Importar Cuotas de Créditos/Recaudos Vigentes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4935
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   10815
      _Version        =   524288
      _ExtentX        =   19076
      _ExtentY        =   8705
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
      MaxCols         =   4
      ScrollBars      =   2
      SpreadDesigner  =   "frmPreaSubDeducciones.frx":0000
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtCuota 
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   6360
      Width           =   1695
      _Version        =   1310720
      _ExtentX        =   2990
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   6360
      Width           =   1815
      _Version        =   1310720
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -360
      TabIndex        =   2
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Deducciones de la Colilla"
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
      Width           =   5892
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmPreaSubDeducciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsMensajes As New ProGrX_EstudioCrd.clsEstudioMensajes
Private clsEntidad As New ProGrX_EstudioCrd.clsEntidad
Private clsNull As New ProGrX_EstudioCrd.clsNull
Public Monto As Double
Public m_NumeroDepagos As Integer
Public dtpCorte As Date

Private vChanged As Boolean

Private Sub sbTrerNumeroDepagos()

On Error GoTo error

m_NumeroDepagos = 0

If gPreAnalisis.Institucion = "-1" Or gPreAnalisis.Institucion = "" Then
    gPreAnalisis.Institucion = "0"
End If
  
   
glogon.strSQL = "select  dbo.fxCRDPreaNumPagos(" & fxFormatearValor(GLOBALES.gTag2, fecha) & "," & fxFormatearValor(gPreAnalisis.Institucion, Caracter) & " )"
If execSql(glogon.strSQL) Then
    If Trim(glogon.Recordset(0) & "") <> "" Then
    m_NumeroDepagos = glogon.Recordset(0)
    End If
End If

Exit Sub

error:
    Call cMensaje.deError("Ocurrió un erro en visual basic al traer la información solicitada. Error " & Err.Description)
End Sub

Private Sub CargarGrid()
Dim sql As String

On Error GoTo error

vGrid.MaxCols = 4
vGrid.Col = 1
vGrid.Lock = True

sql = "spCRDPreaDETALLE_DEDUC_TxExpediente " & fxFormatearValor(gPreAnalisis.Expediente, Caracter)
Call sbCargaGrid(vGrid, 4, sql)

Call sbCalculaTotales

salir:
    Exit Sub
error:
    Call cMensaje.deError("Ocurrió un erro en visual basic al traer la información solicitada. Error " & Err.Description)
End Sub

Public Function fxAgregaColleccion(ByVal pId As Integer, ByVal pExpediente As String, ByVal pCuota_Colilla As Double _
                , ByVal pCuota_Mensual As Double, Optional pDetalle As String = "") As String
Dim Vcoleccion As New Collection

On Error GoTo error

With Vcoleccion
    .Add fxFormatearValor(pId, Numerico)
    .Add fxFormatearValor(pExpediente, Caracter)
    .Add fxFormatearValor(pCuota_Colilla, Numerico)
    .Add fxFormatearValor(pCuota_Mensual, Numerico)
    .Add fxFormatearValor(pDetalle, Caracter)
    
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
Dim vCuota_Colilla As Double
Dim vCuota_Mensual As Double
Dim vMonto As Double

On Error GoTo vError
 

    
    fxInsertar = False
    
    If Not ValidaEstadoPreanalisis(gPreAnalisis.ESTADO) Then
      Exit Function
    End If
    
    vGrid.Row = fila
    vID = 0
    vCuota_Colilla = 0
    vCuota_Mensual = 0
    vGrid.Col = 2
        If Val(vGrid.Text) > 0 Then
            vCuota_Colilla = CDbl(vGrid.Text)
        Else
            MsgBox "El monto de la cuota por planilla es requerido.", vbExclamation, gMsgTitulo
            Exit Function
        End If
        
    vGrid.Col = 3
        If Val(vGrid.Text) > 0 Then
            vCuota_Mensual = CDbl(vGrid.Text)
        Else
            MsgBox "El monto de la cuota mensual es requerido.", vbExclamation, gMsgTitulo
            Exit Function
        End If

    
 Me.MousePointer = vbHourglass
    vGrid.Col = 4
 
    clsEntidad.tablaName = "spCRDPreaDETALLE_DEDUC"
    If (clsEntidad.fxAgregar(fxAgregaColleccion(vID, gPreAnalisis.Expediente, vCuota_Colilla, vCuota_Mensual, vGrid.Text))) Then
        fxInsertar = True
        vGrid.MaxRows = vGrid.MaxRows + 1
       glogon.strSQL = "select max(IDX) as IDX from CRD_PREA_DETALLE_DEDUC  Where cod_preanalisis = " & fxFormatearValor(gPreAnalisis.Expediente, Caracter)

        If (execSql(glogon.strSQL, True)) Then
            vGrid.Col = 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = glogon.Recordset!IdX
        Else
            MsgBox "Error obteniendo el consecutivo del registro ingresado", vbExclamation, gMsgTitulo
        End If
    End If
    
    Me.MousePointer = vbDefault
    Exit Function

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo
End Function

Private Function fxModificar(fila As Long) As Boolean
Dim vID As String
Dim vNewCuota_Colilla As Double
Dim vNewCuota_Mensual As Double
Dim vMonto As Double

On Error GoTo vError

    
    fxModificar = False
    If Not ValidaEstadoPreanalisis(gPreAnalisis.ESTADO) Then
      Exit Function
    End If
    
    
Me.MousePointer = vbHourglass
    
    vNewCuota_Colilla = 0
    vNewCuota_Mensual = 0
    vGrid.Row = fila
    
    vGrid.Col = 1
        vID = vGrid.Text
     vGrid.Col = 2
        If Val(vGrid.Text) > 0 Then
            vNewCuota_Colilla = CDbl(vGrid.Text)
        Else
            MsgBox "El monto de la cuota por planilla es requerido.", vbExclamation, gMsgTitulo
            Exit Function
        End If
    vGrid.Col = 3
        If Val(vGrid.Text) > 0 Then
            vNewCuota_Mensual = CDbl(vGrid.Text)
        Else
            MsgBox "El monto de la cuota mensual es requerido.", vbExclamation, gMsgTitulo
            Exit Function
        End If
        
        
    vGrid.Col = 4
    
    
        clsEntidad.tablaName = "spCRDPreaDETALLE_DEDUC"
    If (clsEntidad.fxModificar(fxAgregaColleccion(vID, gPreAnalisis.Expediente, vNewCuota_Colilla, vNewCuota_Mensual, vGrid.Text))) Then
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
 
    
    fxBorrar = False
    If Not ValidaEstadoPreanalisis(gPreAnalisis.ESTADO) Then
        Exit Function
    End If
    
    
 Me.MousePointer = vbHourglass
    
    vGrid.Col = 1
    vGrid.Row = fila
    vID = Val(vGrid.Text)
    
    clsEntidad.tablaName = "spCRDPreaDETALLE_DEDUC"
    If (MsgBox("¿ Desea borrar la información seleccionada?", vbQuestion + vbYesNo, gMsgTitulo) = vbYes) Then
        If clsEntidad.fxRemover(fxAgregaColleccionBorrar(vID, gPreAnalisis.Expediente)) Then
            vGrid.Col = 1
            fxBorrar = True
            vGrid.Action = 5
            vGrid.MaxRows = vGrid.MaxRows - 1
            vGrid.Row = fila
            vGrid.Col = 1
            vGrid.Action = 0
        End If
        vGrid.SetFocus
    End If

    Me.MousePointer = vbDefault
    Exit Function

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo

End Function


Private Sub btnImport_Click()
Dim strSQL As String

On Error GoTo vError
    
Me.MousePointer = vbHourglass
    
strSQL = "exec spCRDPreaImportCreditosVigentes '" & GLOBALES.gTag & "', " & m_NumeroDepagos
Call ConectionExecute(strSQL)

Call CargarGrid

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo

End Sub

Private Sub vGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim sql As String

On Error GoTo vError
    
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
    
    
    Call sbCalculaTotales
    
    Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo
End Sub

Private Sub Form_Load()
          
frmPreaSubDeducciones.Caption = "Expediente: " & GLOBALES.gTag
frmPreaSubDeducciones.dtpCorte = GLOBALES.gTag2


Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

 vChanged = False
 Call sbTrerNumeroDepagos
 Call CargarGrid
End Sub


Private Sub sbCalculaTotales()
Dim i As Integer, curCuota As Currency, curMonto As Currency

curCuota = 0
curMonto = 0

For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.Col = 2 'Monto
    curMonto = curMonto + IIf((vGrid.Text = ""), 0, vGrid.Text)
    vGrid.Col = 3 'Cuota
    curCuota = curCuota + IIf((vGrid.Text = ""), 0, vGrid.Text)
Next i


txtCuota.Text = Format(curCuota, "Standard")
txtMonto.Text = Format(curMonto, "Standard")

End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    GLOBALES.gTag = txtMonto.Text
    GLOBALES.gTag2 = txtCuota.Text
    
End Sub

Private Sub vGrid_Advance(ByVal AdvanceNext As Boolean)
    Dim fila As Integer
    Dim wColValue As Variant
    
    wColValue = vGrid.Text
    vGrid.Text = CStr(wColValue)
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
        End If
    End If

End Sub

Private Sub vGrid_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    vChanged = True
End Sub

Private Sub vGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim d As Double
    If (NewCol = -1) Then Exit Sub
        If (vChanged = True) Then
            If (Row = NewRow) Then
                If Col = 2 Then
                    vGrid.Col = Col
                    vGrid.Row = Row
                    If vGrid.Text = "" Then
                         vGrid.Text = 0
                    End If
                    d = vGrid.Text
                    vGrid.Col = NewCol
                    vGrid.Row = NewRow
                    vGrid.Text = CStr(d * m_NumeroDepagos)
                End If
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
End Sub

Sub RecalculaGrid() '' Procedimiento para recalcular el grid al final
    
    Dim d As Double
    Dim i As Integer
    
    vGrid.Col = 2
    For i = 1 To vGrid.MaxRows
        vGrid.Row = i
        If Val(vGrid.Text) > 0 Then
            d = vGrid.Text
            vGrid.Col = 3
            vGrid.Text = CStr(d * m_NumeroDepagos)
            vGrid.Col = 2
        End If
    Next i
    


    
End Sub


