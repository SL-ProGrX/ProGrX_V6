VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmSIF_Parametros 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SIF: Parámetros del Sistema"
   ClientHeight    =   6888
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   9528
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6888
   ScaleWidth      =   9528
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5532
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   9252
      _Version        =   524288
      _ExtentX        =   16320
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
      MaxCols         =   496
      ScrollBars      =   2
      SpreadDesigner  =   "frmSIF_Parametros.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parámetros del Sistema"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmSIF_Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 10
End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 10
vGrid.AppearanceStyle = fxGridStyle
imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "exec spSIFParametros"
Call ConectionExecute(strSQL)

strSQL = "select cod_parametro,descripcion,valor from SIF_PARAMETROS" _
      & " order by cod_parametro"
Call sbCargaGrid(vGrid, 3, strSQL)

vGrid.MaxRows = vGrid.MaxRows - 1

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim vValor As String
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
  vGrid.Row = vGrid.ActiveRow
vGrid.Col = 3
vValor = vGrid.Text
vGrid.Col = 1

strSQL = "update SIF_PARAMETROS set valor = '" & vValor _
       & "' where cod_parametro = '" & vGrid.Text & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Modifica", "Parametro del SIF : " & vGrid.Text)

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
End If

End Sub









