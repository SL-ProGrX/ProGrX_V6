VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmPosReparacionParametros 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Repación : Parámetros "
   ClientHeight    =   5820
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   7560
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4452
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   7332
      _Version        =   524288
      _ExtentX        =   12933
      _ExtentY        =   7853
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
      MaxCols         =   487
      ScrollBars      =   2
      SpreadDesigner  =   "frmPosReparacionParametros.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parámetros de Servicios de Reparación"
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
      Height          =   612
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10932
   End
End
Attribute VB_Name = "frmPosReparacionParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 33
End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 33
vGrid.AppearanceStyle = fxGridStyle


Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "exec spPosReparacionParametros"
Call ConectionExecute(strSQL)

strSQL = "select cod_parametro,descripcion,valor from POS_Reparacion_Parametros" _
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
vGrid.col = 1


vGrid.col = 3
strSQL = "update POS_Reparacion_Parametros set valor = '" & vGrid.Text & "'"
vGrid.col = 1
strSQL = strSQL & " where cod_parametro = '" & vGrid.Text & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Modifica", "Parametro de POS Serv.Reparacion: " & vGrid.Text)

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

If vGrid.ActiveCol = 3 And KeyCode = vbKeyF4 Then
  Select Case vGrid.ActiveRow
    Case 2, 3 'Bodegas
        gBusquedas.Convertir = "N"
        gBusquedas.Columna = "cod_bodega"
        gBusquedas.Orden = "cod_bodega"
        gBusquedas.Consulta = "select cod_bodega,descripcion from pv_bodegas"
        gBusquedas.Filtro = ""
        frmBusquedas.Show vbModal
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = vGrid.ActiveCol
        vGrid.Text = gBusquedas.Resultado
     
     Case 4 'Causas de Entrada
        gBusquedas.Convertir = "N"
        gBusquedas.Columna = "cod_entsal"
        gBusquedas.Orden = "cod_entsal"
        gBusquedas.Consulta = "select cod_entsal,descripcion,tipo,cod_cuenta from pv_entrada_salida"
        gBusquedas.Filtro = " and Tipo = 'E'"
        frmBusquedas.Show vbModal
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = vGrid.ActiveCol
        vGrid.Text = gBusquedas.Resultado
     
     Case 5 'Causas de Salida
        gBusquedas.Convertir = "N"
        gBusquedas.Columna = "cod_entsal"
        gBusquedas.Orden = "cod_entsal"
        gBusquedas.Consulta = "select cod_entsal,descripcion,tipo,cod_cuenta from pv_entrada_salida"
        gBusquedas.Filtro = " and Tipo = 'S'"
        frmBusquedas.Show vbModal
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = vGrid.ActiveCol
        vGrid.Text = gBusquedas.Resultado
     


   End Select
End If


End Sub









