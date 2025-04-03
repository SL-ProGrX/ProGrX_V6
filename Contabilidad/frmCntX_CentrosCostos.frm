VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCntX_CentrosCostos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Centros de Costos"
   ClientHeight    =   7788
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   10236
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7788
   ScaleWidth      =   10236
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6372
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   9972
      _Version        =   1245187
      _ExtentX        =   17590
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
      Item(0).Caption =   "Centro de Costos"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Unidades"
      Item(1).ControlCount=   5
      Item(1).Control(0)=   "Label1(0)"
      Item(1).Control(1)=   "Label2"
      Item(1).Control(2)=   "lblX"
      Item(1).Control(3)=   "lsw"
      Item(1).Control(4)=   "lswAsg"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4932
         Left            =   -69880
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   4812
         _Version        =   1245187
         _ExtentX        =   8488
         _ExtentY        =   8700
         _StockProps     =   77
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.ListView lswAsg 
         Height          =   4932
         Left            =   -64960
         TabIndex        =   2
         Top             =   1440
         Visible         =   0   'False
         Width           =   4932
         _Version        =   1245187
         _ExtentX        =   8700
         _ExtentY        =   8700
         _StockProps     =   77
         BackColor       =   -2147483643
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5652
         Left            =   720
         TabIndex        =   4
         Top             =   600
         Width           =   8412
         _Version        =   524288
         _ExtentX        =   14838
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
         MaxCols         =   490
         ScrollBars      =   2
         SpreadDesigner  =   "frmCntX_CentrosCostos.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label Label1 
         Caption         =   $"frmCntX_CentrosCostos.frx":05BE
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
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   9012
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Centros de Costos"
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
         Left            =   -69880
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   4812
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
         Left            =   -64960
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   4932
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Centros de Costos"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   7692
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   1104
      Left            =   0
      Top             =   0
      Width           =   10344
   End
End
Attribute VB_Name = "frmCntX_CentrosCostos"
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


strSQL = "select cod_centro_costo,descripcion,activo" _
       & " from CntX_Centro_Costos where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & " order by cod_centro_costo"
Call sbCargaGrid(vGrid, 3, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

rs.Open "select isnull(count(*),0) as Total from CntX_Centro_Costos where cod_centro_costo = '" _
        & vGrid.Text & "' and COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta, glogon.Conection, adOpenStatic

If rs!Total = 0 Then 'Insertar
  strSQL = "insert into CntX_Centro_Costos(cod_centro_costo,COD_CONTABILIDAD,descripcion,activo) values('"
  vGrid.Col = 1
  strSQL = strSQL & UCase(vGrid.Text) & "'," & gCntX_Parametros.CodigoConta & ",'"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ")"
   
  Call ConectionExecute(strSQL, 0)

  vGrid.Col = 2
  
  Call Bitacora("Registra", "Centro de Costo: " & vGrid.Text & " Conta." & gCntX_Parametros.CodigoConta)
  
  fxGuardar = 1

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update CntX_Centro_Costos set descripcion = '" & vGrid.Text & "',Activo = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & " where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta & " and cod_centro_costo = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL, 0)
 
 Call Bitacora("Modifica", "Centro de Costo: " & vGrid.Text & " Conta." & gCntX_Parametros.CodigoConta)
 
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
  
strSQL = "select C.*,A.cod_unidad as ExisteX" _
       & " from CntX_Unidades C left join CntX_Unidades_CC A on C.cod_unidad = A.cod_unidad" _
       & " and C.cod_contabilidad = A.cod_contabilidad and A.cod_centro_costo = '" & lblX.Tag _
       & "' and A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " Where C.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " order by ExisteX desc,C.cod_unidad"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lswAsg.ListItems.Add(, , rs!cod_unidad)
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
  strSQL = "insert CntX_unidades_cc(cod_unidad,cod_centro_costo,cod_contabilidad) values('" & Item.Text _
         & "','" & lblX.Tag & "'," & gCntX_Parametros.CodigoConta & ")"
Else
  strSQL = "delete CntX_unidades_cc where cod_unidad = '" & Item.Text _
         & "' and cod_centro_costo = '" & lblX.Tag & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta

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
        strSQL = "select cod_centro_costo,descripcion,activo" _
               & " from CntX_Centro_Costos where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
               & " order by cod_centro_costo"
        Call sbCargaGrid(vGrid, 3, strSQL)

  Case 1
       lsw.ListItems.Clear
       lswAsg.ListItems.Clear
       lblX.Caption = ">> Seleccione un Centro de Costo <<"
       lblX.Tag = "(x)"
       
       vPaso = True
       
       strSQL = "select cod_centro_costo,descripcion from CntX_Centro_Costos where Activo = 1 and cod_contabilidad = " & gCntX_Parametros.CodigoConta
       Call OpenRecordSet(rs, strSQL, 0)
       Do While Not rs.EOF
         Set itmX = lsw.ListItems.Add(, , rs!cod_centro_costo)
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
        strSQL = "delete CntX_Centro_Costos where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
               & " and cod_centro_costo = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL, 0)
        
        Call Bitacora("Elimina", "Centro de Costo: " & vGrid.Text & " Conta." & gCntX_Parametros.CodigoConta)
        
        
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
     
     End If
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


