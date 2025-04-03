VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAH_Excedentes_Aplicadores 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aplicadores de Excedentes"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   8775
      _Version        =   524288
      _ExtentX        =   15478
      _ExtentY        =   9975
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
      MaxCols         =   478
      ScrollBars      =   2
      SpreadDesigner  =   "frmAH_Excedentes_Aplicadores.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   9855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Encargos de Aplicar Excedentes y su distribución"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   5172
   End
End
Attribute VB_Name = "frmAH_Excedentes_Aplicadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub Form_Activate()
vModulo = 2
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 2

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vGrid.AppearanceStyle = fxGridStyle

      
strSQL = "select A.USUARIO,  A.ACTIVO, A.CARGA, A.REAL, A.PROYECTADO, A.PRORRATEADO" _
        & "  from EXC_APLICADORES A left join USUARIOS U on A.USUARIO = U.NOMBRE" _
        & " Where U.ESTADO = 'A'" _
        & " order by A.ACTIVO desc, A.USUARIO"

Call sbCargaGrid(vGrid, 6, strSQL)


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Function fxGuardar() As Long

Dim vCuenta As String, vTipo As String

On Error GoTo vError

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

fxGuardar = 0


If Trim(vGrid.Text) <> "" Then 'Insertar
  
'spExc_Aplicadores_Add(@A_Usuario varchar(30), @Mov char(1) = 'A', @Usuario varchar(30)
'        ,  @Activo smallint,  @Carga smallint,  @Real smallint,  @Proyectado smallint,  @Prorrateado smallint)

  strSQL = "exec spExc_Aplicadores_Add '" & vGrid.Text & "', 'A', '" & glogon.Usuario & "', "
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Value & ", "
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ", "
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ", "
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ", "
  vGrid.Col = 6
  strSQL = strSQL & vGrid.Value
  
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Excedentes> Usuario Aplicador: " & vGrid.Text)

End If

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
        
        strSQL = "delete EXC_APLICADORES where Usuario = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Excedentes: Usuario Aplicador: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
