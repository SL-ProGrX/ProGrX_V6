VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmUS_Formularios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formularios"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   10575
      _Version        =   524288
      _ExtentX        =   18648
      _ExtentY        =   13145
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
      ScrollBars      =   2
      SpreadDesigner  =   "frmUS_Formularios.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   345
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   7335
      _Version        =   1441793
      _ExtentX        =   12938
      _ExtentY        =   609
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "M�dulo"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frmUS_Formularios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub cbo_Click()
Dim strSQL As String

If vPaso Then Exit Sub

    strSQL = "select Formulario,Descripcion from US_formularios" _
           & " where modulo = " & cbo.ItemData(cbo.ListIndex) _
           & " order by formulario"
    Call sbCargaGrid(vGrid, 2, strSQL)

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 13

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True

strSQL = "select Nombre as 'ItmX', Modulo as 'IdX' from us_modulos order by modulo"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

vPaso = False

Call cbo_Click

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim vExiste As Long
'Guarda la informaci�n de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

    'Saca Ultimo ID
    strSQL = "select isnull(count(*),0) as 'Existe' from US_formularios " _
           & " where modulo = " & cbo.ItemData(cbo.ListIndex) _
           & " and formulario = '" & vGrid.Text & "'"
    Call OpenRecordSet(rs, strSQL)
    vExiste = rs!Existe
    rs.Close



If vExiste = 0 Then
    vGrid.Col = 1
    strSQL = "insert us_formularios(modulo,formulario,descripcion) values(" & cbo.ItemData(cbo.ListIndex) _
           & ",'" & vGrid.Text & "','"
    vGrid.Col = 2
    strSQL = strSQL & vGrid.Text & "')"
    
    Call ConectionExecute(strSQL, 1)
  
    vGrid.Col = 1
    Call Bitacora("Registra", "Formulario: " & vGrid.Text)
        
    'Inserta Opcion Default de Acceso al Menu
    strSQL = "select isnull(max(cod_Opcion),0) + 1 as 'Consec' from US_Opciones"
    Call OpenRecordSet(rs, strSQL)
    vExiste = rs!Consec
    rs.Close
    
    
    strSQL = "insert us_Opciones(modulo,formulario,cod_Opcion,Opcion,Opcion_descripcion,registro_Fecha,registro_usuario)" _
           & "  values(" & cbo.ItemData(cbo.ListIndex) & ",'" & vGrid.Text & "'," & vExiste & ",'MenuAccess','Acceso al Formulario'" _
           & ",getdate(),'" & glogon.Usuario & "')"
    Call ConectionExecute(strSQL, 1)
  
   Else 'Actualizar

    vGrid.Col = 1
    strSQL = "update US_formularios set formulario = '" & vGrid.Text & "',descripcion = '"
    vGrid.Col = 2
    strSQL = strSQL & vGrid.Text & "' where Formulario = '"
    vGrid.Col = 1
    strSQL = strSQL & vGrid.Text & "' and Modulo = " & cbo.ItemData(cbo.ListIndex)
    
    Call ConectionExecute(strSQL, 1)
    
    vGrid.Col = 1
    Call Bitacora("Modifica", "Formulario: " & vGrid.Text)
    
   End If

   fxGuardar = 1
   
Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
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
        vGrid.Col = 1
        
        strSQL = "delete US_ROL_PERMISOS where cod_opcion in(select cod_opcion from US_Opciones where formulario = '" & vGrid.Text & "' and Modulo = " & cbo.ItemData(cbo.ListIndex) & ")"
        Call ConectionExecute(strSQL, 1)
        
        
        strSQL = "delete US_Opciones where formulario = '" & vGrid.Text & "' and Modulo = " & cbo.ItemData(cbo.ListIndex)
        Call ConectionExecute(strSQL, 1)
        
        strSQL = "delete US_Formularios where formulario = '" & vGrid.Text & "' and Modulo = " & cbo.ItemData(cbo.ListIndex)
        Call ConectionExecute(strSQL, 1)

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Formulario: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If




End Sub





