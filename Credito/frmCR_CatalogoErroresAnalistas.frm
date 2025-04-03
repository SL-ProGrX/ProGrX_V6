VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCR_CatalogoErroresAnalistas 
   Caption         =   "Definición Errores Revisión Analistas Crédito"
   ClientHeight    =   6756
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   11364
   LinkTopic       =   "Form1"
   ScaleHeight     =   6756
   ScaleWidth      =   11364
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5412
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   11052
      _Version        =   524288
      _ExtentX        =   19494
      _ExtentY        =   9546
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
      MaxCols         =   485
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_CatalogoErroresAnalistas.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de errores en la revisiones de los analistas de crédito"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   9252
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmCR_CatalogoErroresAnalistas"
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
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture
 
Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "SELECT ID_ERROR,DESCRIPCION,MENSAJE,ACTIVO FROM  CRD_ANALISIS_ERRORES" _
       & " order by ID_ERROR"
Call sbCargaGrid(vGrid, 4, strSQL)


End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from CRD_ANALISIS_ERRORES " _
       & " where ID_ERROR = " & vGrid.Text
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into CRD_ANALISIS_ERRORES(ID_ERROR,DESCRIPCION,MENSAJE,ACTIVO) values(" _
         & UCase(vGrid.Text) & ",'"
  vGrid.col = 2
  strSQL = strSQL & Trim(vGrid.Text) & "','"
  vGrid.col = 3
  strSQL = strSQL & Trim(vGrid.Text) & "','"
  vGrid.col = 4
  strSQL = strSQL & vGrid.Value & "')"

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Errores Analistas : " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update CRD_ANALISIS_ERRORES set DESCRIPCION = '" & Trim(UCase(vGrid.Text)) & "',MENSAJE = '"
 vGrid.col = 3
 strSQL = strSQL & Trim(vGrid.Text) & "',ACTIVO = '"
 vGrid.col = 4
 strSQL = strSQL & vGrid.Value & "' where ID_ERROR = "

 vGrid.col = 1
 strSQL = strSQL & vGrid.Text
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Errores Analistas : " & vGrid.Text)

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

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este error", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        strSQL = "delete CRD_ANALISIS_ERRORES where ID_ERROR = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Elimina", "Tipo de Garantía : " & vGrid.Text)
        
        vGrid.col = 1
        strSQL = "SELECT ID_ERROR,DESCRIPCION,MENSAJE,ACTIVO FROM  CRD_ANALISIS_ERRORES" _
               & " order by ID_ERROR"
        Call sbCargaGrid(vGrid, 4, strSQL)
     End If
End If


End Sub

