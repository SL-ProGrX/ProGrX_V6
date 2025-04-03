VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCajas_EfectivoDenominaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cajas: Denominaciones del Efectivo"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   8370
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5172
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8052
      _Version        =   524288
      _ExtentX        =   14203
      _ExtentY        =   9123
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
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
      MaxCols         =   495
      ScrollBars      =   2
      SpreadDesigner  =   "frmCajas_EfectivoDenominaciones.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   5652
      _Version        =   1310723
      _ExtentX        =   9975
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16579836
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16579836
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Divisa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmCajas_EfectivoDenominaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mDivisa As String, vPaso As Boolean


Private Sub cbo_Click()
Dim strSQL As String

If vPaso Or cbo.ListCount = 0 Then Exit Sub

mDivisa = cbo.ItemData(cbo.ListIndex)

strSQL = "select Denominacion,case when Tipo = 'M' Then 'Moneda' else 'Billete' end as 'Tipo', descripcion,Activa from CAJAS_EFECTIVO_DENOMINACIONES" _
       & " where cod_divisa = '" & mDivisa & "'" _
       & " order by Denominacion desc"
Call sbCargaGrid(vGrid, 4, strSQL)

End Sub

Private Sub Form_Activate()
vModulo = 5
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset, vDivisa As String

vModulo = 5
vGrid.AppearanceStyle = fxGridStyle


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True
 
strSQL = "select rtrim(cod_divisa) as 'IdX' , rtrim(descripcion) as 'ItmX', Divisa_local" _
       & " From CntX_Divisas where COD_CONTABILIDAD = " & GLOBALES.gEnlace _
       & " order by divisa_local,cod_divisa"
Call OpenRecordSet(rs, strSQL)
cbo.Clear

Do While Not rs.EOF
  cbo.AddItem rs!itmX & ""
  cbo.ItemData(cbo.ListCount - 1) = CStr(rs!IdX)
 
  If rs!divisa_local = 1 Then vDivisa = rs!itmX
  rs.MoveNext
Loop
rs.Close
vPaso = False

If vDivisa <> "" Then
    cbo.Text = vDivisa
End If

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

strSQL = "select isnull(count(*),0) as Existe from CAJAS_EFECTIVO_DENOMINACIONES " _
       & " where Denominacion = " & CCur(vGrid.Text) & " and cod_divisa = '" & mDivisa & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert CAJAS_EFECTIVO_DENOMINACIONES(cod_divisa,Denominacion,Tipo,descripcion,Activa,Registro_Usuario,Registro_Fecha)" _
         & " values('" & mDivisa & "'," & CCur(vGrid.Text) & ",'"
  vGrid.col = 2
  strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "','"
  vGrid.col = 3
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.col = 4
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
  
  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Denominación del Efectivo: " & vGrid.Text & " Divisa: " & mDivisa)

Else 'Actualizar

 vGrid.col = 3
 strSQL = "update CAJAS_EFECTIVO_DENOMINACIONES set descripcion = '" & vGrid.Text & "',Activa = "
 vGrid.col = 4
 strSQL = strSQL & vGrid.Value & " where Denominacion = "
 vGrid.col = 1
 strSQL = strSQL & CCur(vGrid.Text) & " and cod_divisa = '" & mDivisa & "'"
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Denominación del Efectivo: " & vGrid.Text & " Divisa: " & mDivisa)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

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
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        strSQL = "delete CAJAS_EFECTIVO_DENOMINACIONES where Denominacion = '" & vGrid.Text _
               & "' and cod_divisa = '" & mDivisa & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Denominación del Efectivo: " & vGrid.Text & " Divisa: " & mDivisa)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

