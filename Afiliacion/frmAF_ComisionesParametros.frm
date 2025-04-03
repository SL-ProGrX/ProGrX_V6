VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_ComisionesParametros 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros de Comisiones para Afiliación/Renuncias"
   ClientHeight    =   7356
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   8652
   Icon            =   "frmAF_ComisionesParametros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7356
   ScaleWidth      =   8652
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5892
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   8412
      _Version        =   524288
      _ExtentX        =   14838
      _ExtentY        =   10393
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
      SpreadDesigner  =   "frmAF_ComisionesParametros.frx":000C
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Parámetros de Comisiones de Afiliación/Renuncias"
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
      Height          =   732
      Left            =   1884
      TabIndex        =   0
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   8652
   End
End
Attribute VB_Name = "frmAF_ComisionesParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 1
End Sub


Private Sub Form_Load()
Dim strSQL As String


vModulo = 1
vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "exec spAFIComisionesParametros"
Call ConectionExecute(strSQL)

strSQL = "select cod_parametro,descripcion,valor from AFI_COMISIONES_PARAMETROS" _
      & " order by cod_parametro"
Call sbCargaGrid(vGrid, 3, strSQL)

vGrid.MaxRows = vGrid.MaxRows - 1

End Sub


Private Function fxValida() As Boolean
Dim vParametro As String, vMensaje As String
Dim vUnidad As String


vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

vMensaje = ""
vParametro = Trim(vGrid.Text)

vGrid.col = 3
Select Case vParametro
  Case "01" 'Cuenta Contable
     If Not fxgCntCuentaValida(vGrid.Text) Then
        vMensaje = vMensaje & vbCrLf & " - Cuenta Contable no es válida...!"
     End If
  Case "19" 'Tesoreria Unidad
     If Not fxgTESValidaDatos("UNIDAD", vGrid.Text) Then
        vMensaje = vMensaje & vbCrLf & " - Código de Unidad no existe o se encuentra desactivado...!"
     End If
  Case "20" 'Tesoreria Centro de Costo
     vUnidad = fxgAFIParametroComision("19")
     If Not fxgTESValidaDatos("CC", vGrid.Text, vUnidad) Then
        vMensaje = vMensaje & vbCrLf & " - Código de Centro de Costo no existe o se encuentra desactivado, o no ha sido asignado a esta unidad: " & vUnidad & "...!"
     End If
  Case "21" 'Tesoreria Conceptos
     If Not fxgTESValidaDatos("CONCEPTO", vGrid.Text) Then
        vMensaje = vMensaje & vbCrLf & " - Código de Concepto no existe o se encuentra desactivado...!"
     End If
End Select


If Len(vMensaje) > 0 Then
   MsgBox vMensaje, vbExclamation
   fxValida = False
Else
   fxValida = True
End If


End Function

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0

If Not fxValida Then
   Exit Function
End If

vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

vGrid.col = 3
strSQL = "update AFI_COMISIONES_PARAMETROS set valor = '" & vGrid.Text & "'"
vGrid.col = 1
strSQL = strSQL & " where cod_parametro = '" & vGrid.Text & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Modifica", "Parametro de Comisiones de Afiliación : " & vGrid.Text)

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, vParametro As String

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
End If

'Activa Busquedas
If KeyCode = vbKeyF4 Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 1
  vParametro = Trim(vGrid.Text)

    Select Case vParametro
      Case "01" 'Cuenta Contable
         Call sbgCntCuentaConsulta
         vGrid.col = 3
         vGrid.Text = gBusquedas.Resultado
      Case "19" 'Tesoreria Unidad
         Call sbgTESBusqueda("UNIDAD")
         vGrid.col = 3
         vGrid.Text = gBusquedas.Resultado
      Case "20" 'Tesoreria Centro de Costo
         vParametro = fxgAFIParametroComision("19")
         Call sbgTESBusqueda("CC", vParametro)
         vGrid.col = 3
         vGrid.Text = gBusquedas.Resultado
      Case "21" 'Tesoreria Conceptos
         Call sbgTESBusqueda("CONCEPTO")
         vGrid.col = 3
         vGrid.Text = gBusquedas.Resultado
    
    End Select
End If

End Sub









