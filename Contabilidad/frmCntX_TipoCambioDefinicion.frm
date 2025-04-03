VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCntX_TipoCambioDefinicion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Definición de Tipos de Cambios / Para Ingreso de Asientos MultiDivisas"
   ClientHeight    =   7128
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7128
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5292
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   6972
      _Version        =   524288
      _ExtentX        =   12298
      _ExtentY        =   9335
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
      SpreadDesigner  =   "frmCntX_TipoCambioDefinicion.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   5532
      _Version        =   1245187
      _ExtentX        =   9758
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtLineas 
      Height          =   310
      Left            =   6960
      TabIndex        =   4
      Top             =   6720
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   547
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "50"
      Alignment       =   2
      Appearance      =   2
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Left            =   5640
      TabIndex        =   3
      Top             =   6720
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Lineas:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Divisa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
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
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmCntX_TipoCambioDefinicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub cbo_Click()
Dim strSQL As String

If Not vPaso Then Exit Sub
 
If Not IsNumeric(txtLineas.Text) Then
    txtLineas.Text = 50
End If

 
strSQL = "SELECT TOP " & CLng(txtLineas.Text) & " ID_Cambio,TC_Compra,TC_Venta,Inicio,Corte,Variacion" _
       & " FROM CntX_Divisas_Tipo_Cambio where cod_divisa = '" & cbo.ItemData(cbo.ListIndex) _
       & "' and COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & " order by id_Cambio desc"

Call sbCargaGrid(vGrid, vGrid.MaxCols, strSQL)



End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

vGrid.AppearanceStyle = fxGridStyle

vGrid.MaxCols = 6
vGrid.MaxRows = 1

vPaso = False

strSQL = "select rtrim(cod_divisa) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " From CntX_Divisas where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and divisa_local = 0"
Call sbCbo_Llena_New(cbo, strSQL, False, True)


vPaso = True

Call cbo_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim vDivisa As String, vTC_Compra As Currency, vTC_Venta As Currency
Dim vTC_Corte As Date


'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow


vDivisa = cbo.ItemData(cbo.ListIndex)

vGrid.Col = 2
vTC_Compra = vGrid.Value
vGrid.Col = 3
vTC_Venta = vGrid.Value
vGrid.Col = 5
vTC_Corte = Format(vGrid.Text, "yyyy/mm/dd")
 
vGrid.Col = 1
If vGrid.Text = "" Then
    vGrid.Col = 1
    strSQL = "select (isnull(max(id_Cambio),0) + 1) as Ultimo from CntX_Divisas_Tipo_Cambio" _
           & " where cod_divisa ='" & vDivisa & "' and COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta
    Call OpenRecordSet(rs, strSQL, 0)
    vGrid.Value = rs!ultimo
    rs.Close

    strSQL = "insert into CntX_Divisas_Tipo_Cambio(ID_Cambio,COD_CONTABILIDAD,cod_divisa,usuario,fecha,tc_Compra" _
           & ",tc_venta,Inicio,Corte,variacion) values(" & vGrid.Value & "," & gCntX_Parametros.CodigoConta & ",'" _
           & vDivisa & "','" & glogon.Usuario & "',getdate(),"
    vGrid.Col = 2
    strSQL = strSQL & vGrid.Value & ","
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Value & ",'"
    vGrid.Col = 4
    strSQL = strSQL & Format(vGrid.Text, "yyyy/mm/dd") & "','"
    vGrid.Col = 5
    strSQL = strSQL & Format(vGrid.Text, "yyyy/mm/dd") & " 23:59:59',"
    vGrid.Col = 6
    strSQL = strSQL & CCur(vGrid.Text) & ")"
    
    Call ConectionExecute(strSQL, 0)
    
    'Bitacora
    vGrid.Col = 1
    strSQL = "ID-" & vGrid.Text & " Divisa : " & vDivisa
    strSQL = strSQL & " Conta." & gCntX_Parametros.CodigoConta
    
    vGrid.Col = 2
    Call Bitacora("Registra", "Tipo Cambio : " & strSQL)
  
   Else 'Actualizar
       
       vGrid.Col = 2
       strSQL = "update CntX_Divisas_Tipo_Cambio set tc_Compra = " & vGrid.Value
       vGrid.Col = 3
       strSQL = strSQL & ",tc_Venta = " & vGrid.Value
       vGrid.Col = 4
       strSQL = strSQL & ",inicio = '" & Format(vGrid.Text, "yyyy/mm/dd") & "',corte = '"
       vGrid.Col = 5
       strSQL = strSQL & Format(vGrid.Text, "yyyy/mm/dd") & "',variacion = "
       vGrid.Col = 6
       strSQL = strSQL & CCur(vGrid.Text) _
              & " where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
              & " and cod_divisa = '" & vDivisa & "' and Id_Cambio = "
       vGrid.Col = 1
       strSQL = strSQL & vGrid.Value
            
       Call ConectionExecute(strSQL, 0)
            
      'Bitacora
      vGrid.Col = 1
      strSQL = "ID-" & vGrid.Text & " Divisa : " & vDivisa _
             & " Conta." & gCntX_Parametros.CodigoConta
      
      Call Bitacora("Modifica", "Tipo Cambio : " & strSQL)
       
 End If

 vGrid.Col = 1
 fxGuardar = 1
 
'Actualiza Ultimo TC en la Moneda Foranea
strSQL = "exec spCntX_DivisasTC_Update " & gCntX_Parametros.CodigoConta & ",'" & vDivisa & "','" & Format(vTC_Corte, "yyyy/mm/dd") _
       & "'," & vTC_Compra & "," & vTC_Venta
Call ConectionExecute(strSQL, 0)
 
 
Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function




Private Sub txtLineas_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Call cbo_Click
End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, vTemp(6) As Variant, x As Integer
Dim lng As Long, strSQL As String

'Guarda Linea
If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
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
  vGrid.Col = 1
  If vGrid.Text <> "" Then
    strSQL = "delete CntX_Divisas_Tipo_Cambio where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
           & " and cod_divisa = '" & cbo.ItemData(cbo.ListIndex) & "' and ID_Cambio = " & vGrid.Text
    Call ConectionExecute(strSQL, 0)
    
     strSQL = "ID-" & vGrid.Text & " Divisa : " & cbo.ItemData(cbo.ListIndex) _
            & " Conta." & gCntX_Parametros.CodigoConta
     Call Bitacora("Elimina", "Tipo Cambio : " & strSQL)
    
  End If
  
  
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 6
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To vGrid.MaxCols
        vGrid.Col = x
        vTemp(x) = vGrid.Text
     Next x
     
     vGrid.Row = lng
     For x = 1 To vGrid.MaxCols
       vGrid.Col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
  
End If


End Sub

