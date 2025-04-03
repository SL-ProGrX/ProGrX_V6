VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmIVR_Cat_Instrumentos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SCGI Instrumentos de Inversión"
   ClientHeight    =   6990
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   13020
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   12615
      _Version        =   524288
      _ExtentX        =   22251
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
      MaxCols         =   488
      ScrollBars      =   2
      SpreadDesigner  =   "frmIVR_Cat_Instrumentos.frx":0000
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
      Caption         =   "Tipos de Instrumentos "
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
      Width           =   13092
   End
End
Attribute VB_Name = "frmIVR_Cat_Instrumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 22
End Sub

Private Sub sbConsulta()
Dim strSQL As String

strSQL = "select COD_INSTRUMENTO, DESCRIPCION, TIPO_DESC, Form_ItmX" _
      & ", BASE_INTERESES, VALORIZA, DIAS_VALORIZA, ACTIVO" _
      & " from vIVR_INSTRUMENTOS" _
      & " order by COD_INSTRUMENTO"
Call sbCargaGrid(vGrid, 8, strSQL)

End Sub

Private Sub Form_Load()

vModulo = 22

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
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from IVR_INSTRUMENTOS " _
       & " where COD_INSTRUMENTO = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into IVR_INSTRUMENTOS(COD_INSTRUMENTO,DESCRIPCION, TIPO, TIPO_FORM, BASE_INTERESES" _
         & ", VALORIZA, DIAS_VALORIZA, ACTIVO, NOTAS, REGISTRO_USUARIO, REGISTRO_FECHA) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "','"
  vGrid.Col = 4
  strSQL = strSQL & SIFGlobal.fxCodText(vGrid.Text) & "',"
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Text & ","
  vGrid.Col = 6
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 7
  strSQL = strSQL & vGrid.Text & ","
 
  vGrid.Col = 8
  strSQL = strSQL & vGrid.Value & ",'','" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Tipo de Instrumento:  " & vGrid.Text)

Else 'Actualizar

  vGrid.Col = 2
  strSQL = "update IVR_INSTRUMENTOS set descripcion = '" & vGrid.Text & "',TIPO = '"
  vGrid.Col = 3
  strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "', TIPO_FORM = '"
  vGrid.Col = 4
  strSQL = strSQL & SIFGlobal.fxCodText(vGrid.Text) & "', BASE_INTERESES = "
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Text & ", VALORIZA = "
  vGrid.Col = 6
  strSQL = strSQL & vGrid.Value & ", DIAS_VALORIZA = "
  vGrid.Col = 7
  strSQL = strSQL & vGrid.Text & ", ACTIVO = "
 
  vGrid.Col = 8
  strSQL = strSQL & vGrid.Value & " where COD_INSTRUMENTO = '"
  vGrid.Col = 1
  strSQL = strSQL & vGrid.Text & "'"
 
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Modifica", "Tipo de Instrumento:  " & vGrid.Text)

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
        vGrid.Col = 1
        strSQL = "delete IVR_INSTRUMENTOS where COD_INSTRUMENTO = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Tipo de Instrumento:  " & vGrid.Text)
        
        Call sbConsulta
     
     End If
End If


End Sub



