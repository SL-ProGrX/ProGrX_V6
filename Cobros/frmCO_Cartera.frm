VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmCO_Cartera 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cartera de Cobro"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13545
   Icon            =   "frmCO_Cartera.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   13545
   Begin XtremeSuiteControls.PushButton cmdModifica 
      Height          =   252
      Left            =   11760
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Modifica"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6252
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   13332
      _Version        =   1441793
      _ExtentX        =   23516
      _ExtentY        =   11028
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   3
      Item(0).Caption =   "Definición"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Asignación Metodo No.1"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "lsw"
      Item(1).Control(1)=   "lswCat"
      Item(1).Control(2)=   "lbl"
      Item(2).Caption =   "Asignación Metodo No.2"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "lswMet2"
      Item(2).Control(1)=   "chkTodos"
      Item(2).Control(2)=   "cboX"
      Item(2).Control(3)=   "ShortcutCaption1"
      Begin XtremeSuiteControls.ListView lswMet2 
         Height          =   5772
         Left            =   -63400
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1441793
         _ExtentX        =   11663
         _ExtentY        =   10181
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswCat 
         Height          =   5412
         Left            =   -63400
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1441793
         _ExtentX        =   11663
         _ExtentY        =   9546
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5412
         Left            =   -69880
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   6492
         _Version        =   1441793
         _ExtentX        =   11451
         _ExtentY        =   9546
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   612
         Left            =   -66160
         TabIndex        =   7
         Top             =   1560
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4254
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Asignar todos las Líneas de Crédito esta Cartera ?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         Appearance      =   16
         Alignment       =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5532
         Left            =   2040
         TabIndex        =   2
         Top             =   480
         Width           =   8892
         _Version        =   524288
         _ExtentX        =   15684
         _ExtentY        =   9758
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
         MaxCols         =   496
         ScrollBars      =   2
         SpreadDesigner  =   "frmCO_Cartera.frx":030A
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboX 
         Height          =   312
         Left            =   -70000
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   6372
         _Version        =   1441793
         _ExtentX        =   11245
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
         Text            =   "ComboBox1"
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   372
         Left            =   -70000
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1441793
         _ExtentX        =   11663
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Carteras de Créditos disponibles:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption lbl 
         Height          =   372
         Left            =   -69880
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   13092
         _Version        =   1441793
         _ExtentX        =   23093
         _ExtentY        =   656
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.PushButton cmdActualiza 
      Height          =   252
      Left            =   11760
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Actualiza"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdReporte 
      Height          =   252
      Left            =   11760
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Reporte"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Definición de Carteras"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   6852
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13812
   End
End
Attribute VB_Name = "frmCO_Cartera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String, vPaso As Boolean


Private Sub cboX_Click()
If vPaso Then Exit Sub
Call sbLswMet2_Load
End Sub

Private Sub chkTodos_Click()
Dim i As Byte, y As Integer
Dim pCodigo As String, strSQL As String

i = MsgBox("Esta seguro que desea Marcar o DesMarcar todos los códigos para esta categoria?", vbYesNo)
If i = vbNo Then Exit Sub

Me.MousePointer = vbHourglass

pCodigo = cboX.ItemData(cboX.ListIndex)

vPaso = True

With lswMet2.ListItems

    For y = 1 To .Count
       If chkTodos.Value = vbChecked Then
          If Not .Item(y).Checked Then
             'Insertar
              strSQL = "insert CBR_CLASIFICACION_DETALLE(COD_CLASIFICACION,codigo) values('" _
                     & pCodigo & "','" & .Item(y).Text & "')"
              Call ConectionExecute(strSQL)
          End If
       
       Else
          If .Item(y).Checked Then
            'Eliminar
            strSQL = "delete CBR_CLASIFICACION_DETALLE where COD_CLASIFICACION = '" _
                   & pCodigo & "' and codigo = '" & .Item(y).Text & "'"
            Call ConectionExecute(strSQL)
          End If
       End If
       .Item(y).Checked = chkTodos.Value
    Next y

End With

vPaso = False

Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
Dim strSQL As String


vModulo = 4
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

tcMain.Item(0).Selected = True


With lsw.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Descripción", lsw.Width - (1200 + 250)
End With

With lswCat.ColumnHeaders
    .Clear
    .Add , , "Línea", 1200
    .Add , , "Descripción", lswCat.Width - (1200 + 250)
End With

With lswMet2.ColumnHeaders
    .Clear
    .Add , , "Línea", 1200
    .Add , , "Descripción", lswMet2.Width - (1200 + 250)
End With


strSQL = "select COD_CLASIFICACION,descripcion,estado from CBR_CLASIFICACION_CARTERA" _
       & " order by COD_CLASIFICACION"
Call sbCargaGrid(vGrid, 3, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

vGrid.Enabled = cmdModifica.Enabled
lswCat.Enabled = cmdActualiza.Enabled
lswMet2.Enabled = cmdActualiza.Enabled
chkTodos.Enabled = cmdActualiza.Enabled
End Sub

Private Function fxExiste(vCod As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from CBR_CLASIFICACION_CARTERA" _
       & " where COD_CLASIFICACION = '" & vCod & "'"
Call OpenRecordSet(rs, strSQL)
fxExiste = IIf((rs!Existe = 1), True, False)
rs.Close
End Function


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

If vGrid.Text = "" Then Exit Function


If Not fxExiste(vGrid.Text) Then
   vGrid.col = 1
   strSQL = "insert CBR_CLASIFICACION_CARTERA(COD_CLASIFICACION,descripcion,estado)" _
          & " values('" & vGrid.Text & "','"
   vGrid.col = 2
   strSQL = strSQL & vGrid.Text & "',"
   vGrid.col = 3
   strSQL = strSQL & vGrid.Value & ")"
   
   Call ConectionExecute(strSQL)
   vGrid.col = 1
   Call Bitacora("Registra", "Cartera Crédito Código: " & vGrid.Text)
   
 Else 'Actualizar
    vGrid.col = 2
    strSQL = "update CBR_CLASIFICACION_CARTERA set descripcion = '" & vGrid.Text
    vGrid.col = 3
    strSQL = strSQL & "',estado = " & vGrid.Value
    vGrid.col = 1
    strSQL = strSQL & " where COD_CLASIFICACION = '" & vGrid.Text & "'"
    
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Cartera Crédito Código: " & vGrid.Text)
    
End If

Exit Function
   
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

   
End Function

Private Sub sbLswCat_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

strSQL = "select R.*,A.codigo as Existe" _
       & " from CBR_CLASIFICACION_CARTERA R left Join CBR_CLASIFICACION_DETALLE A " _
       & " on R.COD_CLASIFICACION = A.COD_CLASIFICACION and A.codigo = '" _
       & vCodigo & "' order by existe desc,R.COD_CLASIFICACION"
Call OpenRecordSet(rs, strSQL, 0)

lswCat.ListItems.Clear

vPaso = True

Do While Not rs.EOF
  Set itmX = lswCat.ListItems.Add(, , rs!COD_CLASIFICACION)
      itmX.SubItems(1) = rs!Descripcion & ""
      itmX.Checked = IIf(IsNull(rs!Existe), False, True)
      If itmX.Checked Then itmX.ForeColor = vbBlue
      
  rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault

End Sub


Private Sub sbLswMet2_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

strSQL = "select R.codigo,R.descripcion,A.codigo as Existe" _
       & " from Catalogo R left Join CBR_CLASIFICACION_DETALLE A " _
       & " on R.codigo = A.codigo and A.COD_CLASIFICACION  = '" _
       & cboX.ItemData(cboX.ListIndex) & "' order by existe desc,R.codigo"

Call OpenRecordSet(rs, strSQL)

lswMet2.ListItems.Clear

vPaso = True

Do While Not rs.EOF
  Set itmX = lswMet2.ListItems.Add(, , rs!Codigo)
      itmX.SubItems(1) = rs!Descripcion & ""
      itmX.Checked = IIf(IsNull(rs!Existe), False, True)
      If itmX.Checked Then itmX.ForeColor = vbBlue
  rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

vCodigo = Item.Text
lbl.Caption = Item.Text & " - " & Item.SubItems(1)

Call sbLswCat_Load

End Sub


Private Sub lswCat_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

Dim strSQL As String
    
On Error GoTo vError

If Item.Checked Then
    strSQL = "insert CBR_CLASIFICACION_DETALLE(COD_CLASIFICACION,codigo) values('" _
           & Item.Text & "','" & vCodigo & "')"
Else
    strSQL = "delete CBR_CLASIFICACION_DETALLE where COD_CLASIFICACION = '" _
           & Item.Text & "' and codigo = '" & vCodigo & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswMet2_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub


Dim strSQL As String, xCodigo As String

On Error GoTo vError

xCodigo = cboX.ItemData(cboX.ListIndex)

If Item.Checked Then
    strSQL = "insert CBR_CLASIFICACION_DETALLE(COD_CLASIFICACION,codigo) values('" _
           & xCodigo & "','" & Item.Text & "')"
Else
    strSQL = "delete CBR_CLASIFICACION_DETALLE where COD_CLASIFICACION = '" _
           & xCodigo & "' and codigo = '" & Item.Text & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If Item.Index = 0 Then Exit Sub

Me.MousePointer = vbHourglass

vPaso = True

Select Case Item.Index
 Case 1
    vCodigo = ""
    lbl.Caption = ""
    lswCat.ListItems.Clear
    
    strSQL = "select codigo,descripcion from catalogo order by codigo"
    Call OpenRecordSet(rs, strSQL, 0)
    lsw.ListItems.Clear
    Do While Not rs.EOF
     Set itmX = lsw.ListItems.Add(, , rs!Codigo)
         itmX.SubItems(1) = rs!Descripcion & ""
     rs.MoveNext
    Loop
    rs.Close
  
  Case 2
    strSQL = "select rtrim(cod_clasificacion) as 'IdX', rtrim(descripcion) as 'ItmX'" _
           & " from CBR_CLASIFICACION_CARTERA order by cod_clasificacion"
    
    Call sbCbo_Llena_New(cboX, strSQL, False, True)
    
    lswMet2.ListItems.Clear
   
End Select

vPaso = False

Me.MousePointer = vbDefault

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

If vGrid.ActiveCol = 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = vGrid.Text
End If

End Sub



