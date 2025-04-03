VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCntX_Unidades 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Unidades de Negocios"
   ClientHeight    =   7908
   ClientLeft      =   48
   ClientTop       =   408
   ClientWidth     =   13896
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7908
   ScaleWidth      =   13896
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6372
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   13692
      _Version        =   1245187
      _ExtentX        =   24151
      _ExtentY        =   11239
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
      ItemCount       =   2
      Item(0).Caption =   "Unidades"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Centro de Costos"
      Item(1).ControlCount=   5
      Item(1).Control(0)=   "Label1(0)"
      Item(1).Control(1)=   "Label2"
      Item(1).Control(2)=   "lblX"
      Item(1).Control(3)=   "lsw"
      Item(1).Control(4)=   "lswAsg"
      Begin XtremeSuiteControls.ListView lswAsg 
         Height          =   4932
         Left            =   -63160
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   6372
         _Version        =   1245187
         _ExtentX        =   11239
         _ExtentY        =   8700
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
         Height          =   4932
         Left            =   -69520
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   6252
         _Version        =   1245187
         _ExtentX        =   11028
         _ExtentY        =   8700
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5652
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   13332
         _Version        =   524288
         _ExtentX        =   23516
         _ExtentY        =   9970
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
         MaxCols         =   8
         ScrollBars      =   2
         SpreadDesigner  =   "frmCntX_Unidades.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   -63160
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   6372
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Unidades"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   -69520
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   6252
      End
      Begin VB.Label Label1 
         Caption         =   $"frmCntX_Unidades.frx":07D9
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   0
         Left            =   -69760
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   9012
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unidades de Negocios"
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
      Height          =   372
      Index           =   1
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   7692
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   14292
   End
End
Attribute VB_Name = "frmCntX_Unidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub Form_Activate()
vModulo = 20
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 20

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

tcMain.Item(0).Selected = True

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1100
    .Add , , "Descripción", 3650
End With

With lswAsg.ColumnHeaders
    .Clear
    .Add , , "Id", 1100
    .Add , , "Descripción", 3650
End With
lswAsg.Checkboxes = True

strSQL = "select cod_unidad,descripcion,Nivel,unidad_omision,reporta_renta,activa, Cta_Renta, Cta_Renta_Gasto" _
       & " from CntX_Unidades where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & " order by cod_unidad"
Call sbCargaGridLocal(vGrid, 8, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub




Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

Me.MousePointer = vbHourglass

On Error GoTo vError

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows

Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    Select Case i
      Case 7, 8
        vGrid.Col = i
        vGrid.Text = fxCntX_CuentaFormato(True, CStr(rs.Fields(i - 1).Value & ""))
      Case Else
        vGrid.Col = i
        vGrid.Text = rs.Fields(i - 1).Value & ""
     End Select
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

vError:

Me.MousePointer = vbDefault

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

rs.Open "select isnull(count(*),0) as Total from CntX_Unidades where cod_unidad = '" _
        & vGrid.Text & "' and COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta, glogon.Conection, adOpenStatic

If rs!Total = 0 Then 'Insertar
  strSQL = "insert into CntX_Unidades(cod_unidad,COD_CONTABILIDAD,descripcion,nivel,unidad_omision,reporta_renta,activa" _
         & ",Cta_Renta, Cta_Renta_Gasto) values('"
  vGrid.Col = 1
  strSQL = strSQL & UCase(vGrid.Text) & "'," & gCntX_Parametros.CodigoConta & ",'"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & ","
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 6
  strSQL = strSQL & vGrid.Value & ",'"
  vGrid.Col = 7
  strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "','"
  vGrid.Col = 8
  strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "')"
 
   
  Call ConectionExecute(strSQL, 0)

  vGrid.Col = 2
  
  Call Bitacora("Registra", "Unidad de Negocio: " & vGrid.Text & " Conta." & gCntX_Parametros.CodigoConta)
  
  fxGuardar = 1

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update CntX_Unidades set descripcion = '" & vGrid.Text & "',nivel = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Text & ",unidad_omision = "
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Value & ", activa = "
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Value & ", reporta_renta = "
 vGrid.Col = 6
 strSQL = strSQL & vGrid.Value & ", Cta_Renta = '"
 vGrid.Col = 7
 strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "', Cta_Renta_Gasto = '"
 vGrid.Col = 8
 strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) _
        & "' where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta & " and cod_unidad = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL, 0)
 
 Call Bitacora("Modifica", "Unidad de Negocio: " & vGrid.Text & " Conta." & gCntX_Parametros.CodigoConta)
 
 fxGuardar = 1
End If

rs.Close

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbCargaAsignacion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
  
Me.MousePointer = vbHourglass
  
lswAsg.ListItems.Clear
vPaso = True

strSQL = "select C.*,A.cod_centro_costo as ExisteX" _
       & " from CntX_Centro_costos C left join CntX_Unidades_CC A on C.cod_centro_costo = A.cod_centro_costo" _
       & " and C.cod_contabilidad = A.cod_contabilidad and A.cod_unidad = '" & lblX.Tag _
       & "' and A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " Where C.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " order by ExisteX desc,C.cod_centro_costo"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lswAsg.ListItems.Add(, , rs!cod_centro_costo)
     itmX.SubItems(1) = rs!Descripcion
     itmX.Checked = IIf(IsNull(rs!ExisteX), vbUnchecked, vbChecked)
 If itmX.Checked Then itmX.ForeColor = vbBlue
 rs.MoveNext
Loop
rs.Close

vPaso = False
Me.MousePointer = vbDefault
  
End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub
If lsw.ListItems.Count = 0 Then Exit Sub

lblX.Tag = Item.Text
lblX.Caption = Item.SubItems(1)

Call sbCargaAsignacion
End Sub

Private Sub lswAsg_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lswAsg_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
  strSQL = "insert CntX_unidades_cc(cod_unidad,cod_centro_costo,cod_contabilidad) values('" & lblX.Tag _
         & "','" & Item.Text & "'," & gCntX_Parametros.CodigoConta & ")"
Else
  strSQL = "delete CntX_unidades_cc where cod_unidad = '" & lblX.Tag _
         & "' and cod_centro_costo = '" & Item.Text & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta

End If

Call ConectionExecute(strSQL, 0)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Select Case Item.Index
  Case 0
    strSQL = "select cod_unidad,descripcion,Nivel,unidad_omision,reporta_renta,activa" _
           & " from CntX_Unidades where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
           & " order by cod_unidad"
    Call sbCargaGrid(vGrid, 6, strSQL)

  Case 1
       lsw.ListItems.Clear
       lswAsg.ListItems.Clear
       lblX.Caption = ">> Seleccione una Unidad <<"
       lblX.Tag = "(x)"
               
               
       vPaso = True
       
       strSQL = "select cod_unidad,descripcion from CntX_Unidades where Activa = 1 and cod_contabilidad = " & gCntX_Parametros.CodigoConta
       Call OpenRecordSet(rs, strSQL, 0)
       Do While Not rs.EOF
         Set itmX = lsw.ListItems.Add(, , rs!cod_unidad)
             itmX.SubItems(1) = rs!Descripcion
         rs.MoveNext
       Loop
       rs.Close
        
       vPaso = False
        
End Select

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If


'Reporte
If KeyCode = vbKeyF5 Then
    Call sbCntX_Reportes_Catalogos("Unidades")
End If


'Consulta Cuentas
If KeyCode = vbKeyF4 And (vGrid.ActiveCol = 7 Or vGrid.ActiveCol = 8) Then
  frmCntX_ConsultaCuentas.Show vbModal
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gCuenta
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
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete CntX_Unidades where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
               & " and cod_unidad = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL, 0)
        
        Call Bitacora("Elimina", "Unidad Negocio: " & vGrid.Text & " Conta." & gCntX_Parametros.CodigoConta)
        
        
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
     
     End If
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub
