VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmRH_Cat_Incapacidades 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RRHH: Tipos de Incapacidades"
   ClientHeight    =   6930
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   10695
      _Version        =   524288
      _ExtentX        =   18865
      _ExtentY        =   9763
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
      MaxCols         =   486
      ScrollBars      =   2
      SpreadDesigner  =   "frmRH_Cat_Incapacidades.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   492
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   6732
      _Version        =   1310723
      _ExtentX        =   11874
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Tipos de Incapacidades"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmRH_Cat_Incapacidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 23
End Sub

Private Sub sbConsulta()
Dim strSQL As String

strSQL = "select COD_INCAPACIDAD,descripcion,REQUIERE_AUTORIZACION,PORC_PATRONO,COD_CONCEPTO,ACTIVA from RH_INCAPACIDADES_TIPOS" _
      & " order by COD_INCAPACIDAD"
Call sbCargaGrid(vGrid, 6, strSQL)

End Sub

Private Sub Form_Load()

vModulo = 23

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call sbConsulta

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from RH_INCAPACIDADES_TIPOS " _
       & " where COD_INCAPACIDAD = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into RH_INCAPACIDADES_TIPOS(COD_INCAPACIDAD,DESCRIPCION,REQUIERE_AUTORIZACION,PORC_PATRONO, COD_CONCEPTO, ACTIVA, REGISTRO_USUARIO, REGISTRO_FECHA) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.col = 3
  strSQL = strSQL & vGrid.Value & ","
  vGrid.col = 4
  strSQL = strSQL & CCur(vGrid.Text) & ",'"
  vGrid.col = 5
  strSQL = strSQL & vGrid.Text & "', "
  vGrid.col = 6
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Tipo de Incapacidad: " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update RH_INCAPACIDADES_TIPOS set descripcion = '" & vGrid.Text & "',REQUIERE_AUTORIZACION = "
 vGrid.col = 3
 strSQL = strSQL & vGrid.Value & ", PORC_PATRONO = "
 vGrid.col = 4
 strSQL = strSQL & CCur(vGrid.Text) & ", COD_CONCEPTO = '"
 vGrid.col = 5
 strSQL = strSQL & vGrid.Text & "', ACTIVA = "
 vGrid.col = 6
 strSQL = strSQL & vGrid.Value & " where COD_INCAPACIDAD = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Tipo de Incapacidad: " & vGrid.Text)

End If
rs.Close

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

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        strSQL = "delete RH_INCAPACIDADES_TIPOS where COD_INCAPACIDAD = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Tipo de Incapacidad: " & vGrid.Text)
        
        Call sbConsulta
     
     End If
End If

'Borrar una linea
If KeyCode = vbKeyF4 Then
        
    vGrid.Row = vGrid.ActiveRow
    vGrid.col = 5
    gBusquedas.Col1Name = "Concepto"
    gBusquedas.Col2Name = "Descripción"
    
    gBusquedas.Columna = "COD_CONCEPTO"
    gBusquedas.Orden = "COD_CONCEPTO"
    gBusquedas.Consulta = "Select COD_CONCEPTO, DESCRIPCION FROM RH_CONCEPTOS"
    gBusquedas.Filtro = " AND SUMAR_EN = 'HRS_INCAPACIDAD'"
    frmBusquedas.Show vbModal
    
    vGrid.Text = gBusquedas.Resultado
    
End If



End Sub

