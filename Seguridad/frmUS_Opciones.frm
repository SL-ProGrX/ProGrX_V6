VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmUS_Opciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento : Opciones"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12390
   HelpContextID   =   1003
   Icon            =   "frmUS_Opciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   12390
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   7455
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   3975
      _Version        =   1441793
      _ExtentX        =   7011
      _ExtentY        =   13150
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   7455
      Left            =   4200
      TabIndex        =   0
      Top             =   1320
      Width           =   8055
      _Version        =   524288
      _ExtentX        =   14208
      _ExtentY        =   13150
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
      SpreadDesigner  =   "frmUS_Opciones.frx":030A
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   345
      Left            =   2760
      TabIndex        =   1
      Top             =   120
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
   Begin XtremeSuiteControls.FlatEdit txtFormulario 
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   600
      Width           =   7335
      _Version        =   1441793
      _ExtentX        =   12938
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Formulario"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Módulo"
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
Attribute VB_Name = "frmUS_Opciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub cbo_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If vPaso Then Exit Sub

txtFormulario.Text = ""
vGrid.MaxRows = 0
lsw.ListItems.Clear

strSQL = "select Formulario,Descripcion from US_formularios" _
    & " where modulo = " & cbo.ItemData(cbo.ListIndex) _
    & " order by formulario"
Call OpenRecordSet(rs, strSQL, 1)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Formulario)
  rs.MoveNext
Loop
rs.Close

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 13

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True

strSQL = "select Nombre as 'ItmX', Modulo as 'IdX' from us_modulos order by modulo"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

vPaso = False

With lsw.ColumnHeaders
    .Clear
    .Add , , "", lsw.Width - 100
End With

Call cbo_Click

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim vExiste As Long
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If vGrid.Text = "" Then
    strSQL = "select isnull(max(cod_Opcion),0) + 1 as 'Consec' from US_Opciones"
    Call OpenRecordSet(rs, strSQL)
    vGrid.Text = rs!Consec
    rs.Close
    
    strSQL = "insert us_Opciones(modulo,formulario,cod_Opcion,Opcion,Opcion_descripcion,registro_Fecha,registro_usuario)" _
           & "  values(" & cbo.ItemData(cbo.ListIndex) & ",'" & txtFormulario.Text & "'," & vGrid.Text & ",'"
    vGrid.Col = 2
    strSQL = strSQL & vGrid.Text & "','"
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Text & "',getdate(),'" & glogon.Usuario & "')"
    
    Call ConectionExecute(strSQL, 1)
  
    vGrid.Col = 1
    Call Bitacora("Registra", "Opción de Sistema: " & vGrid.Text)
    
  
  
   Else 'Actualizar

    vGrid.Col = 2
    strSQL = "update us_Opciones set Opcion = '" & vGrid.Text & "',Opcion_descripcion = '"
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Text & "' where Formulario = '" & txtFormulario.Text & "' and Cod_Opcion = "
    vGrid.Col = 1
    strSQL = strSQL & vGrid.Text & " and Modulo = " & cbo.ItemData(cbo.ListIndex)
    
    Call ConectionExecute(strSQL, 1)
    
    vGrid.Col = 1
    Call Bitacora("Modifica", "Opción de Sistema: " & vGrid.Text)
    
   End If

   fxGuardar = 1
   
Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub
If lsw.ListItems.Count <= 0 Then Exit Sub

Dim strSQL As String

txtFormulario.Text = Item.Text

strSQL = "select cod_Opcion, Opcion,Opcion_Descripcion from US_Opciones" _
       & " where formulario = '" & txtFormulario.Text & "' and Modulo = " & cbo.ItemData(cbo.ListIndex)
Call sbCargaGrid(vGrid, 3, strSQL)

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String


If txtFormulario.Text = "" Then Exit Sub


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
        
        strSQL = "delete US_Opciones where formulario = '" & txtFormulario.Text & "' and Modulo = " & cbo.ItemData(cbo.ListIndex) _
               & " and Cod_Opcion = " & vGrid.Text
        Call ConectionExecute(strSQL, 1)


        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Opción de Sistema: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

End Sub






