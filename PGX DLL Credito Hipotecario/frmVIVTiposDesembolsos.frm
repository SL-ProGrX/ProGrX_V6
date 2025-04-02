VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmVivTiposDesembolsos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipos Desembolsos"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   12615
      _Version        =   524288
      _ExtentX        =   22251
      _ExtentY        =   9551
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      DisplayRowHeaders=   0   'False
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
      MaxRows         =   498
      ScrollBars      =   2
      SpreadDesigner  =   "frmVIVTiposDesembolsos.frx":0000
      VScrollSpecialType=   2
      Appearance      =   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Desembolsos (Hipotecario)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   7332
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmVivTiposDesembolsos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cambioDatos As Boolean

Private Sub Form_Load()

vModulo = 3

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


Call sbCargarGrid

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Sub sbCargarGrid()
On Error GoTo error

If ObjConsultar.fxTraerTiposDesembolsos Then
    Call sbLlenaGrid(vGrid, 10)
Else
    vGrid.ClearSelection
    vGrid.MaxCols = 10
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
End If

salir:
    Exit Sub
error:
    Call cMensaje.deError("Ocurrió un erro en visual basic al traer la información solicitada. Error " & Err.Description)
    
End Sub

Private Sub sbLlenaGrid(vGrid As Object, vGridMaxCol As Integer)
Dim i As Integer
Dim vregistros As Integer

On Error GoTo error

vregistros = glogon.Recordset.RecordCount
vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows

Do While Not glogon.Recordset.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    If glogon.Recordset.Fields(i - 1).Type = adDBTimeStamp Then
        vGrid.Text = Format((glogon.Recordset.Fields(i - 1).Value), "yyyy/mm/dd")
    Else
        Select Case i
        
            Case 1, 2, 9
                vGrid.Text = CStr(IIf(IsNull(glogon.Recordset.Fields(i - 1).Value), "", glogon.Recordset.Fields(i - 1).Value))
            Case 8
                vGrid.Value = CCur(IIf(IsNull(glogon.Recordset.Fields(i - 1).Value), 0, glogon.Recordset.Fields(i - 1).Value))
            Case 10
                If CStr(IIf(IsNull(glogon.Recordset.Fields(i - 1).Value), "A", glogon.Recordset.Fields(i - 1).Value)) = "A" Then
                    vGrid.Value = 1
                Else
                    vGrid.Value = 0
                End If
                
            Case Else
            
                If CStr(IIf(IsNull(glogon.Recordset.Fields(i - 1).Value), "", glogon.Recordset.Fields(i - 1).Value)) = True Then
                    vGrid.Value = 1
                Else
                    vGrid.Value = 0
                End If
        End Select
    End If
    vGrid.Col = 1
    vGrid.Action = 0
    vGrid.SetActiveCell 1, vGrid.ActiveRow
    vGrid.Lock = True
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  glogon.Recordset.MoveNext
Loop
glogon.Recordset.Close
Exit Sub

error:
 cMensaje.deError ("Ocurrió un error al construir el grig para mostrar la información solicitada. Error:" & Err.Description)

End Sub

Private Sub sbNuevoRegistro()
    vGrid.Col = 1
    vGrid.SetActiveCell 1, vGrid.MaxRows
    vGrid.SetFocus
End Sub

Private Function fxValidaDatos(ByVal fila As Long) As Boolean

On Error GoTo error

fxValidaDatos = False

ReDim gParametros(1 To 11)

vGrid.Row = fila
vGrid.Col = 1
    gParametros(1) = vGrid.Text 'Codigo
vGrid.Col = 2
    gParametros(2) = vGrid.Text  'Descripción
vGrid.Col = 3
    gParametros(3) = vGrid.Value 'NivelDesembolso
vGrid.Col = 4
    gParametros(4) = vGrid.Value 'Nivel Formaliza
vGrid.Col = 5
    gParametros(5) = vGrid.Value 'Aplica Ingenieros
vGrid.Col = 6
    gParametros(6) = vGrid.Value 'Aplica Abogados
vGrid.Col = 7
    gParametros(7) = vGrid.Value 'Aplica Interes
vGrid.Col = 8
    gParametros(8) = CCur(vGrid.Value) 'Porcentaje a cobrar

vGrid.Col = 9
    If Len(Trim(vGrid.Text)) = 0 Then
        gParametros(9) = ObjNull.NullString  'Número de cuenta
    Else
        gParametros(9) = vGrid.Text 'Número de cuenta
    End If
    
vGrid.Col = 10
    If vGrid.Value = 0 Then
        gParametros(10) = "I"  'Estado Inactivo
    Else
        gParametros(10) = "A" 'Estado Activo
    End If
    


If (Len(Trim(gParametros(1))) = 0) Then
    Me.MousePointer = vbDefault
    MsgBox ("Debe de ingresar un código válido para el tipo desembolso.")
    vGrid.SetActiveCell 1, vGrid.Row
    m_cambioDatos = False
    Exit Function
ElseIf (Len(Trim(gParametros(2))) = 0) Then
    Me.MousePointer = vbDefault
    MsgBox ("Debe de ingresar una descripción para el tipo desembolso.")
    vGrid.SetActiveCell 1, vGrid.Row
    m_cambioDatos = False
    Exit Function
    
ElseIf (Len(Trim(gParametros(8))) = 0) Then
    Me.MousePointer = vbDefault
    MsgBox ("Debe de ingresar un porcenta a cobrar para el tipo desembolso.")
    vGrid.SetActiveCell 1, vGrid.Row
    m_cambioDatos = False
    Exit Function
ElseIf (Trim(gParametros(9))) = ObjNull.NullString Then
    Me.MousePointer = vbDefault
    MsgBox ("Debe de ingresar un número de cuenta para el tipo desembolso, Presione F4 para consultar cuentas disponibles.")
    vGrid.SetActiveCell 1, vGrid.Row
    m_cambioDatos = False
    Exit Function

End If

gParametros(11) = glogon.Usuario

fxValidaDatos = True
salir:
    Exit Function
error:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function

Private Sub sbGuardaRegistro()
On Error GoTo error

If (m_cambioDatos = True) Then
    If (vGrid.Row = vGrid.MaxRows) Then
        Call sbAgregar(vGrid.Row)
    Else
        Call sbModificar(vGrid.ActiveRow)
    End If
End If

salir:
Exit Sub
error:
    Call cMensaje.deError("Ocurrió un erro en visual basic al registar la información ingresada. Error " & Err.Description)
End Sub

Private Sub sbModificar(fila As Long)
On Error GoTo error

Me.MousePointer = vbHourglass

If fxValidaDatos(fila) = False Then Exit Sub

gParametros(9) = fxgCntCuentaFormato(False, gParametros(9), 0)

If ObjAgregar.fxTiposDesembolsos(1, gParametros(1), gParametros(2), gParametros(3), gParametros(4), _
                                            gParametros(5), gParametros(6), gParametros(7), gParametros(8), _
                                            gParametros(9), gParametros(10), glogon.Usuario, "1900/01/01") Then
    
    Call Bitacora("MODIFICA", "Vivienda tipo desembolso: " & gParametros(1))

    m_cambioDatos = False
    vGrid.Col = 1
End If

Me.MousePointer = vbDefault

salir:
    Exit Sub
error:
Me.MousePointer = vbDefault
    Call ObjMensajes.deError("Ocurrió un error en visual basic al actualización la información ingresada. Error " & Err.Description)
End Sub

Private Sub sbAgregar(ByRef fila As Integer)

Dim vCodigo As String

On Error GoTo vError

Me.MousePointer = vbHourglass

If m_cambioDatos = False Then Exit Sub
If fxValidaDatos(fila) = False Then Exit Sub
    
    
    gParametros(9) = fxgCntCuentaFormato(False, gParametros(9), 0)
    
    
    If ObjAgregar.fxTiposDesembolsos(-1, gParametros(1), gParametros(2), gParametros(3), gParametros(4), _
                                            gParametros(5), gParametros(6), gParametros(7), gParametros(8), _
                                            gParametros(9), gParametros(10), gParametros(11), "1900/01/01") Then
        m_cambioDatos = False
        
        Call Bitacora("REGISTRA", "Vivienda tipo desembolso: " & gParametros(1))
        
        vGrid.MaxRows = vGrid.MaxRows + 1
        vGrid.SetActiveCell 1, vGrid.MaxRows
        
    End If


Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar(ByRef fila As Integer)

On Error GoTo error

ReDim gParametros(1 To 1)

vGrid.Row = fila
vGrid.Col = 1
gParametros(1) = vGrid.Text

If Len(vGrid.Text) = 0 Then
    Me.MousePointer = vbDefault
    MsgBox ("Debe de ingresar un código válido.")
    vGrid.SetActiveCell 1, vGrid.Row
    Exit Sub
End If

    
If ObjMensajes.deDatos("08") = vbYes Then
    If ObjBorrar.fxTipoDesembolso(gParametros(1)) Then
        vGrid.DeleteRows vGrid.Row, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Col = 1
        vGrid.SetActiveCell 1, vGrid.ActiveRow
    End If
End If
Me.MousePointer = vbDefault
vGrid.SetFocus
salir:
    Exit Sub
error:
    Call ObjMensajes.deError("Ocurrió un error en visual basic al borrar la información seleccionada. Error " & Err.Description)
End Sub

Private Sub vGrid_Advance(ByVal AdvanceNext As Boolean)
Dim fila As Integer
fila = vGrid.ActiveRow
    If (vGrid.MaxRows > 1) And (vGrid.ActiveRow = 1) Then
        Exit Sub
    ElseIf (vGrid.ActiveRow = vGrid.MaxRows) Then
        vGrid.Col = 1
        vGrid.Row = vGrid.ActiveRow
        Call sbAgregar(fila)
    End If
End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 And vGrid.ActiveCol = 9 Then
        gBusquedas.Resultado = ""
        
        Call sbgCntCuentaConsulta("D")
        
        If gBusquedas.Resultado = "" Then Exit Sub
        m_cambioDatos = True
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = vGrid.ActiveCol
        vGrid.Text = gBusquedas.Resultado
        
        If (vGrid.Row = vGrid.MaxRows) Then
            Call sbAgregar(vGrid.ActiveRow)
        Else
            Call sbModificar(vGrid.ActiveRow)
        End If
        
        
End If
End Sub

Private Sub vGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

If (NewCol = -1) Then Exit Sub
    If (m_cambioDatos = True) Then
        If (Row = NewRow) Then
        Exit Sub
    Else
        
        If (Row = vGrid.MaxRows) Then
            Call sbAgregar(Val(Row))
        Else
            Call sbModificar(Val(Row))
        End If
    End If
End If
End Sub

Private Sub vGrid_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    m_cambioDatos = True
End Sub

Private Sub vGrid_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

Dim sql As String
    If ((Shift And vbCtrlMask) <> 0) And (KeyCode = vbKeyS) Then
        UnLoad Me

    ElseIf ((Shift And vbCtrlMask) <> 0) And (KeyCode = vbKeyN) Then
        Call sbNuevoRegistro
            
    ElseIf ((Shift And vbCtrlMask) <> 0) And (KeyCode = vbKeyG) Then
        Call sbGuardaRegistro
            
    ElseIf (KeyCode = vbKeyDelete) Then
        If (vGrid.ActiveRow <> vGrid.MaxRows) Then
            vGrid.Col = 1
            Call sbBorrar(vGrid.ActiveRow)
        End If
         
    End If
    
salir:
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    Call ObjMensajes.deError(Err.Description)

End Sub

