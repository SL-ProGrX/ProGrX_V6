VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCajas_Tipos_Cambios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cajas: Tipos de cambios"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   8850
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4932
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   6972
      _Version        =   524288
      _ExtentX        =   12298
      _ExtentY        =   8700
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
      MaxCols         =   494
      ScrollBars      =   2
      SpreadDesigner  =   "frmCajas_Tipos_Cambios.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   5772
      _Version        =   1310723
      _ExtentX        =   10186
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   1440
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
Attribute VB_Name = "frmCajas_Tipos_Cambios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub cbo_Click()
Dim strSQL As String

If Not vPaso Then Exit Sub
 
strSQL = "SELECT ID_Cambio,TC_Compra,TC_Venta,Inicio,Corte,Variacion" _
       & " FROM CAJAS_DIVISAS_TIPO_CAMBIO where cod_divisa = '" & cbo.ItemData(cbo.ListIndex) _
       & "' and COD_CONTABILIDAD = " & GLOBALES.gEnlace _
       & " order by id_Cambio desc"

Call sbCargaGrid(vGrid, vGrid.MaxCols, strSQL)


End Sub

Private Sub Form_Activate()
vModulo = 5

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 5

vGrid.AppearanceStyle = fxGridStyle
imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = False
 

strSQL = "select rtrim(cod_divisa) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " From CntX_Divisas where COD_CONTABILIDAD = " & GLOBALES.gEnlace _
       & " and divisa_local = 0 order by cod_divisa"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

vGrid.MaxCols = 6
vGrid.MaxRows = 1

vPaso = True

Call cbo_Click

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

If vGrid.Text = "" Then
    vGrid.col = 1
    strSQL = "select (isnull(max(id_Cambio),0) + 1) as Ultimo from CAJAS_DIVISAS_TIPO_CAMBIO" _
           & " where cod_divisa ='" & cbo.ItemData(cbo.ListIndex) & "' and COD_CONTABILIDAD = " & GLOBALES.gEnlace
    Call OpenRecordSet(rs, strSQL)
    vGrid.Value = rs!ultimo
    rs.Close

    strSQL = "insert into CAJAS_DIVISAS_TIPO_CAMBIO(ID_Cambio,COD_CONTABILIDAD,cod_divisa,usuario,fecha,tc_Compra" _
           & ",tc_venta,Inicio,Corte,variacion) values(" & vGrid.Value & "," & GLOBALES.gEnlace & ",'" _
           & cbo.ItemData(cbo.ListIndex) & "','" & glogon.Usuario & "',dbo.MyGetdate(),"
    vGrid.col = 2
    strSQL = strSQL & vGrid.Value & ","
    vGrid.col = 3
    strSQL = strSQL & vGrid.Value & ",'"
    vGrid.col = 4
    strSQL = strSQL & Format(vGrid.Text, "yyyy/mm/dd") & " 00:00:00','"
    vGrid.col = 5
    strSQL = strSQL & Format(vGrid.Text, "yyyy/mm/dd") & " 23:59:59',"
    vGrid.col = 6
    strSQL = strSQL & CCur(vGrid.Text) & ")"
    
    Call ConectionExecute(strSQL)
    
    'Bitacora
    vGrid.col = 1
    strSQL = "ID-" & vGrid.Text & " Divisa : " & cbo.ItemData(cbo.ListIndex)
    strSQL = strSQL & " Conta." & GLOBALES.gEnlace
    
    vGrid.col = 2
    Call Bitacora("Registra", "Tipo Cambio : " & strSQL)
  
   Else 'Actualizar
       
       vGrid.col = 2
       strSQL = "update CAJAS_DIVISAS_TIPO_CAMBIO set tc_Compra = " & vGrid.Value
       vGrid.col = 3
       strSQL = strSQL & ",tc_Venta = " & vGrid.Value
       vGrid.col = 4
       strSQL = strSQL & ",inicio = '" & Format(vGrid.Text, "yyyy/mm/dd") & " 00:00:00',corte = '"
       vGrid.col = 5
       strSQL = strSQL & Format(vGrid.Text, "yyyy/mm/dd") & " 23:59:59',variacion = "
       vGrid.col = 6
       strSQL = strSQL & CCur(vGrid.Text) _
              & " where COD_CONTABILIDAD = " & GLOBALES.gEnlace _
              & " and cod_divisa = '" & cbo.ItemData(cbo.ListIndex) & "' and Id_Cambio = "
       vGrid.col = 1
       strSQL = strSQL & vGrid.Value
            
       Call ConectionExecute(strSQL)
            
      'Bitacora
      vGrid.col = 1
      strSQL = "ID-" & vGrid.Text & " Divisa : " & cbo.ItemData(cbo.ListIndex) _
             & " Conta." & GLOBALES.gEnlace
      
      Call Bitacora("Modifica", "Tipo Cambio : " & strSQL)
       
 End If

 vGrid.col = 1
 fxGuardar = 1
 
Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, vTemp(6) As Variant, x As Integer
Dim lng As Long, strSQL As String

'Guarda Linea
If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


If KeyCode = vbKeyDelete Then
  vGrid.col = 1
  If vGrid.Text <> "" Then
    strSQL = "delete CAJAS_DIVISAS_TIPO_CAMBIO where COD_CONTABILIDAD = " & GLOBALES.gEnlace _
           & " and cod_divisa = '" & cbo.ItemData(cbo.ListIndex) & "' and ID_Cambio = " & vGrid.Text
    Call ConectionExecute(strSQL)
    
     strSQL = "ID-" & vGrid.Text & " Divisa : " & cbo.ItemData(cbo.ListIndex) _
            & " Conta." & GLOBALES.gEnlace
     Call Bitacora("Elimina", "Tipo Cambio : " & strSQL)
    
  End If
  
  
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 6
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To vGrid.MaxCols
        vGrid.col = x
        vTemp(x) = vGrid.Text
     Next x
     
     vGrid.Row = lng
     For x = 1 To vGrid.MaxCols
       vGrid.col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
  
End If


End Sub



