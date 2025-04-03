VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCR_Prioridad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prioridades de Deducción de Planillas por Línea de Crédito/Recaudo"
   ClientHeight    =   9156
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10128
   Icon            =   "frmCR_Prioridad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9156
   ScaleWidth      =   10128
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   7932
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   9852
      _Version        =   524288
      _ExtentX        =   17378
      _ExtentY        =   13991
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
      MaxCols         =   6
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_Prioridad.frx":030A
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "*Aplica si el modelo de priorización no es por Garantía"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   7320
      TabIndex        =   2
      Top             =   240
      Width           =   2652
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Prioridad de deducción de Línea vía Planilla "
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
      Height          =   480
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   6132
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmCR_Prioridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 3
End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 3
vGrid.AppearanceStyle = fxGridStyle
 
imgBanner.Picture = frmContenedor.imgBanner_01.Picture


strSQL = "select CODIGO,DESCRIPCION, LINEA_INTERNA, CASE WHEN RETENCION = 'N' AND POLIZA = 'N' THEN 1 ELSE 0 END" _
       & ",CASE WHEN CONVENIO = 'S' THEN 1 ELSE 0 end,PRIORIDAD" _
       & " from Catalogo order by Prioridad asc"
Call sbCargaGrid(vGrid, 6, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

 vGrid.col = 6
 strSQL = "update Catalogo set Prioridad = " & vGrid.Text & " where codigo = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Prioridad deducción x Línea: " & vGrid.Text)

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
End Sub


