VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmPreaParametros 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   14835
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2415
      Left            =   240
      TabIndex        =   3
      Top             =   5880
      Width           =   14535
      _Version        =   1572864
      _ExtentX        =   25638
      _ExtentY        =   4260
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   21
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4215
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   14535
      _Version        =   524288
      _ExtentX        =   25638
      _ExtentY        =   7435
      _StockProps     =   64
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
      MaxCols         =   7
      ScrollBars      =   2
      SpreadDesigner  =   "frmPreaParametros.frx":0000
      AppearanceStyle =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5520
      Width           =   14535
      _Version        =   1572864
      _ExtentX        =   25638
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Histórico de Cambios"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parámetros de Estudio de Créditos"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   14895
   End
End
Attribute VB_Name = "frmPreaParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub sbGrid_Load()

On Error GoTo vError

vPaso = True

strSQL = "select 0, cod_parametro,descripcion,valor, FechaActualiza, UsuarioActualiza, Valor_Anterior" _
      & " from Crd_Prea_parametros" _
      & " order by cod_parametro"
Call sbCargaGrid(vGrid, 7, strSQL)

vGrid.MaxRows = vGrid.MaxRows - 1

vPaso = False

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sblsw_Load(pRow As Long)

If vPaso Then Exit Sub
If pRow <= 0 Then Exit Sub

On Error GoTo vError

vGrid.Row = pRow
vGrid.Col = 2
strSQL = "select Top 100 IdHistorico, CodParametro, Valor, FechaActualiza, UsuarioActualiza " _
       & "  From CRD_PREA_PARAMETROS_HISTORICO" _
       & " Where CodParametro = '" & vGrid.Text & "'" _
       & " order by IdHistorico desc"
Call OpenRecordSet(rs, strSQL)

With lsw.ListItems
 .Clear
 
 Do While Not rs.EOF
    Set itmX = .Add(, , rs!IdHistorico)
        itmX.SubItems(1) = rs!CodParametro
        itmX.SubItems(2) = rs!FechaActualiza & ""
        itmX.SubItems(3) = rs!UsuarioActualiza & ""
        itmX.SubItems(4) = rs!Valor & ""
  
  rs.MoveNext
 Loop
End With


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Activate()
vModulo = 3 'Modulo de Credito
End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 3 'Modulo de Credito

'Inicializa Seguridad
Call Formularios(Me)
Call RefrescaTags(Me)

vGrid.AppearanceStyle = AppearanceStyleVisualStyles
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

strSQL = "exec spCRDPreaParametros"
Call ConectionExecute(strSQL)


With lsw.ColumnHeaders
    .Add , , "Id", 1200
    .Add , , "Parámetro", 1200, vbCenter
    .Add , , "Fecha", 3500
    .Add , , "Usuario", 3500
    .Add , , "Valor", 3500, vbCenter
End With

Call sbGrid_Load

End Sub


Private Function fxGuardar() As Long
Dim pValor As String

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow


vGrid.Col = 4
pValor = vGrid.Text

vPaso = True

strSQL = "update Crd_Prea_parametros set valor_anterior = valor,  FechaActualiza = getdate(), UsuarioActualiza = '" & glogon.Usuario & "',  valor = '" & pValor & "'"
vGrid.Col = 2
strSQL = strSQL & " where cod_parametro = '" & vGrid.Text & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Modifica", "Parametro de PreAnalisis Cod : " & vGrid.Text & ", Valor: " & pValor)

MsgBox "Parámetro " & vGrid.Text & ", actualizado satisfactoriamente!", vbInformation


strSQL = "insert CRD_PREA_PARAMETROS_HISTORICO(CodParametro, Valor, FechaActualiza, UsuarioActualiza) " _
       & "values('" & vGrid.Text & "', '" & pValor & "', getdate(), '" & glogon.Usuario & "')"
Call ConectionExecute(strSQL)


vPaso = False
fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub


If Col = 1 Then
    Call sblsw_Load(Row)
End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = 4 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
End If

End Sub

