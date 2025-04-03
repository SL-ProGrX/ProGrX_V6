VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCxPTiposProv 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Proveedores"
   ClientHeight    =   6336
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6336
   ScaleWidth      =   8952
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5052
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   8532
      _Version        =   524288
      _ExtentX        =   15050
      _ExtentY        =   8911
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
      MaxCols         =   484
      ScrollBars      =   2
      SpreadDesigner  =   "frmTiposProveedores.frx":0000
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
      _Version        =   1245187
      _ExtentX        =   11874
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Tipos de Proveedores"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.8
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
Attribute VB_Name = "frmCxPTiposProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 30
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 30

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "select cod_clasificacion,descripcion,NIT_Codigo,Activo from cxp_prov_clas" _
      & " order by cod_clasificacion"
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

strSQL = "select isnull(count(*),0) as Existe from cxp_prov_clas " _
       & " where cod_clasificacion = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into cxp_prov_clas(cod_clasificacion,descripcion,Nit_Codigo,Activo) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.col = 2
  strSQL = strSQL & UCase(vGrid.Text) & "','"
  vGrid.col = 3
  strSQL = strSQL & UCase(vGrid.Text) & "',"
  vGrid.col = 4
  strSQL = strSQL & vGrid.Value & ")"

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Tipo de Proveedor : " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update cxp_prov_clas set descripcion = '" & vGrid.Text & "',NIT_CODIGO = '"
 vGrid.col = 3
 strSQL = strSQL & vGrid.Text & "', Activo = "
 vGrid.col = 4
 strSQL = strSQL & vGrid.Value & " where cod_clasificacion = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Tipo de Proveedor : " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function
'
'Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
'Dim i As Integer, strSQL As String
'
'On Error Resume Next
'
'Select Case UCase(Button.Key)
'  Case "NUEVO"
'    vGrid.MaxRows = vGrid.MaxRows + 1
'
'  Case "BORRAR"
'     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
'     If i = 6 Then
'        vGrid.Row = vGrid.ActiveRow
'        vGrid.col = 1
'        strSQL = "delete cxp_prov_clas where cod_clasificacion = '" & vGrid.Text & "'"
'        Call ConectionExecute(strSQL)
'        strSQL = vGrid.Text
'        vGrid.col = 1
'        Call Bitacora("Elimina", "Tipo de Proveedor : " & vGrid.Text)
'
'        vGrid.col = 1
'        strSQL = "select cod_clasificacion,descripcion,nit_Codigo,Activo from cxp_prov_clas" _
'              & " order by cod_clasificacion"
'        Call sbCargaGrid(vGrid, 4, strSQL)
'     End If
'
'  Case "REPORTES"
''     Call sbReportes("Caracteristicas", Me)
'
'  Case "AYUDA"
'        frmContenedor.CD.HelpContext = Me.HelpContextID
'        frmContenedor.CD.ShowHelp
'
'End Select
'
'End Sub


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
        strSQL = "delete cxp_prov_clas where cod_clasificacion = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Tipo de Proveedor : " & vGrid.Text)
        
        vGrid.col = 1
        strSQL = "select cod_clasificacion,descripcion,nit_Codigo,Activo from cxp_prov_clas" _
              & " order by cod_clasificacion"
        Call sbCargaGrid(vGrid, 4, strSQL)
     
     End If
End If


End Sub

