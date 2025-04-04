VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCR_Prendas_Tipos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipos de Prendas"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   6120
      Width           =   11055
      _Version        =   1572864
      _ExtentX        =   19500
      _ExtentY        =   4471
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnSel 
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   5
      Top             =   5640
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Marcas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   11055
      _Version        =   524288
      _ExtentX        =   19500
      _ExtentY        =   7011
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
      SpreadDesigner  =   "frmCR_Prendas_Tipos.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltro 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   4095
      _Version        =   1572864
      _ExtentX        =   7223
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnSel 
      Height          =   375
      Index           =   1
      Left            =   5280
      TabIndex        =   6
      Top             =   5640
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Fuente Poder"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnSel 
      Height          =   375
      Index           =   2
      Left            =   6600
      TabIndex        =   7
      Top             =   5640
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Extas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnSel 
      Height          =   375
      Index           =   3
      Left            =   7560
      TabIndex        =   8
      Top             =   5640
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Presentación"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnSel 
      Height          =   375
      Index           =   4
      Left            =   8760
      TabIndex        =   9
      Top             =   5640
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Pólizas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeShortcutBar.ShortcutCaption scTipo 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   11055
      _Version        =   1572864
      _ExtentX        =   19500
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "(Seleccione un Tipo de Prenda)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   492
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   6732
      _Version        =   1572864
      _ExtentX        =   11874
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Tipos de Prendas"
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
Attribute VB_Name = "frmCR_Prendas_Tipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub sbConsulta_Asignacion()
Dim pTipo As String

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

If scTipo.Tag = "" Then
  Exit Sub
End If

vPaso = True

txtFiltro.Text = fxSysCleanTxtInject(txtFiltro.Text)

Select Case True
    Case btnSel(0).Checked  'Marcas
        pTipo = "MAR"
    Case btnSel(1).Checked 'Cobustible
        pTipo = "COB"
    Case btnSel(2).Checked 'Extras
        pTipo = "EXT"
    Case btnSel(3).Checked 'Presentación
        pTipo = "PRE"
    Case btnSel(4).Checked 'Pólizas
        pTipo = "POL"
    Case Else
        pTipo = ""
End Select


strSQL = "exec spCrd_Prendas_Cat_List_Asignacion '" & scTipo.Tag & "', '" & pTipo & "', '" & txtFiltro.Text & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!IdX)
      itmX.SubItems(1) = rs!itmX
      itmX.SubItems(2) = rs!registro_Usuario & ""
      itmX.SubItems(3) = rs!Registro_Fecha & ""
      itmX.Checked = IIf((rs!Asignado = 1), True, False)
  rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub

Private Sub btnSel_Click(Index As Integer)
Dim i As Integer

For i = 0 To btnSel.Count - 1
    btnSel.Item(i).Checked = False
Next i

btnSel.Item(Index).Checked = True


Call sbConsulta_Asignacion
End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub sbConsulta()

vPaso = True

strSQL = "select TIPO_PRENDA, DESCRIPCION, FORMULARIO, PORC_COBERTURA, ACTIVA, '...'" _
      & " from CRD_PRENDAS_TIPOS" _
      & " order by TIPO_PRENDA"
Call sbCargaGrid(vGrid, 6, strSQL)

vPaso = False

End Sub

Private Sub Form_Load()

vModulo = 3

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Descripción", 3000
    .Add , , "Usuario", 2100, vbCenter
    .Add , , "Fecha", 2100, vbCenter
End With

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

strSQL = "select isnull(count(*),0) as Existe from CRD_PRENDAS_TIPOS " _
       & " where TIPO_PRENDA = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into CRD_PRENDAS_TIPOS(TIPO_PRENDA,DESCRIPCION, FORMULARIO, PORC_COBERTURA" _
         & ", ACTIVA, REGISTRO_USUARIO, REGISTRO_FECHA) values('" _
         & UCase(vGrid.Text) & "', '"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "', '"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 4
  strSQL = strSQL & CCur(vGrid.Text) & ","
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Tipo de Prenda: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update CRD_PRENDAS_TIPOS set descripcion = '" & vGrid.Text & "', FORMULARIO = '"
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Text & "', PORC_COBERTURA = "
 
 vGrid.Col = 4
 strSQL = strSQL & CCur(vGrid.Text) & ", ACTIVA = "
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Value & " where TIPO_PRENDA = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Tipo de Prenda: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

Dim pTipo As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case True
    Case btnSel(0).Checked 'Marcas
        pTipo = "MAR"
    Case btnSel(1).Checked 'Cobustible
        pTipo = "COB"
    Case btnSel(2).Checked 'Extras
        pTipo = "EXT"
    Case btnSel(3).Checked 'Presentación
        pTipo = "PRE"
    Case btnSel(4).Checked 'Pólizas
        pTipo = "POL"
    Case Else
        pTipo = ""
End Select

strSQL = "exec spCrd_Prendas_Cat_List_Asignacion_Add '" & scTipo.Tag & "', '" & pTipo _
       & "', '" & Item.Text & "', '" & glogon.Usuario & "', '" & IIf((Item.Checked), "A", "E") & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub



Private Sub txtFiltro_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   Call sbConsulta_Asignacion
End If
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

vGrid.Row = Row
vGrid.Col = 1

scTipo.Tag = vGrid.Text
vGrid.Col = 2
scTipo.Caption = vGrid.Text

Call btnSel_Click(0)

End Sub

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
        strSQL = "delete CRD_PRENDAS_TIPOS where TIPO_PRENDA = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Tipo de Prenda: " & vGrid.Text)
        
        Call sbConsulta
     
     End If
End If


End Sub

