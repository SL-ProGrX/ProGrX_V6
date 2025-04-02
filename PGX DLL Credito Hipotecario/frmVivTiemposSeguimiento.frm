VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.ShortcutBar.v19.1.0.ocx"
Begin VB.Form frmVivTiemposSeguimiento 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tiempos Seguimiento de Garantias"
   ClientHeight    =   3636
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3636
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2532
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   7572
      _Version        =   524288
      _ExtentX        =   13356
      _ExtentY        =   4466
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
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   0
      SpreadDesigner  =   "frmVivTiemposSeguimiento.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption lblTitulo 
      Height          =   732
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   8772
      _Version        =   1245185
      _ExtentX        =   15473
      _ExtentY        =   1291
      _StockProps     =   14
      Caption         =   "Utilice este formulario para establecer los tiempos para realizar el trámite de una garantia segú el profesional"
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
   End
End
Attribute VB_Name = "frmVivTiemposSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_cambioDatos As Boolean
Private Sub sbCargarGrid()
On Error GoTo error
'crea tiempos por defecto si no exiten
glogon.Conection.Execute "exec spCRDVivTiemposSeguimiento"

If ObjConsultar.fxTraerTiemposSeguimiento Then
    Call sbLlenaGrid(vGrid, 4)

End If

salir:
    Exit Sub
error:
    Call cMensaje.deError("Ocurrió un erro en visual basic al traer la información solicitada. Error " & Err.Description)
    
End Sub

Private Sub sbLlenaGrid(vGrid As Object, vGridMaxCol As Integer)
Dim i As Integer
On Error GoTo error
vGrid.MaxCols = vGridMaxCol
vGrid.Row = 1
Do While Not glogon.Recordset.EOF
  For i = 1 To vGrid.MaxCols - 1
    vGrid.Col = i + 1
    vGrid.Text = Trim(CStr(glogon.Recordset.Fields(i - 1).Value))
  Next i
    
  If vGrid.Row = 3 Then
    vGrid.Row = vGrid.Row + 2
  Else
    vGrid.Row = vGrid.Row + 1
  End If
  glogon.Recordset.MoveNext
Loop
glogon.Recordset.Close
vGrid.ColWidth(5) = 0
vGrid.ColWidth(6) = 0
Exit Sub
error:

 cMensaje.deError ("Ocurrió un error al construir el grig para mostrar la información solicitada. Error:" & Err.Description)

End Sub
Private Function fxGuardar() As Long
Dim vProfesional As String
Dim vProceso As String
Dim vTMax As Integer
Dim vTAlerta As Integer

Dim strSQL As String, rs As New ADODB.Recordset
On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow

vGrid.Col = 3 'Tiempo Máximo
vTMax = vGrid.Text
vGrid.Col = 4 'Tiempo Alerta
vTAlerta = vGrid.Text
vGrid.Col = 5 'Proceso
vProceso = vGrid.Text
vGrid.Col = 6 'Profesional
vProfesional = vGrid.Text

glogon.strSQL = " Update dbo.ViviendaTiemposSeguimiento " & _
         "    SET TiempoMaximo = " & vTMax & " , " & _
         "        TiempoAlerta = " & vTAlerta & " " & _
         " WHERE Profesional = '" & vProfesional & "' AND Proceso = '" & vProceso & "'"
   

If execSql(glogon.strSQL) Then
    fxGuardar = 1
    Call sbCargarGrid
    m_cambioDatos = False
    
End If

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub Form_Activate()
vModulo = 3 'Modulo de Credito
End Sub

Private Sub vGrid_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
m_cambioDatos = True
End Sub

Private Sub vGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim i As Integer

If (m_cambioDatos = True) Then
    If (Row = NewRow) Then
        Exit Sub
    Else
        If (Row = 4) Then
            Exit Sub
        Else
            i = fxGuardar
            If i = 0 Then Exit Sub
            
            vGrid.Row = vGrid.ActiveRow
       End If
    End If
End If

End Sub


Private Sub Form_Load()
vGrid.AppearanceStyle = fxGridStyle
vModulo = 3 'Modulo de Credito
'Inicializa Seguridad
Call Formularios(Me)
Call RefrescaTags(Me)

Call sbCargarGrid
End Sub

