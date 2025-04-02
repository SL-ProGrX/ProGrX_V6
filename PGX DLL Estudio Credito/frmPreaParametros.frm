VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmPreaParametros 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4812
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   8052
      _Version        =   524288
      _ExtentX        =   14203
      _ExtentY        =   8488
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
      MaxCols         =   3
      ScrollBars      =   2
      SpreadDesigner  =   "frmPreaParametros.frx":0000
      AppearanceStyle =   1
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
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmPreaParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

strSQL = "select cod_parametro,descripcion,valor from Crd_Prea_parametros" _
      & " order by cod_parametro"
Call sbCargaGrid(vGrid, 3, strSQL)

vGrid.MaxRows = vGrid.MaxRows - 1

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
  vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1


vGrid.Col = 3
strSQL = "update Crd_Prea_parametros set valor = '" & vGrid.Text & "'"
vGrid.Col = 1
strSQL = strSQL & " where cod_parametro = '" & vGrid.Text & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Modifica", "Parametro de PreAnalisis Cod : " & vGrid.Text)

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
End If

End Sub









